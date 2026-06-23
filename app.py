from __future__ import annotations

import io
import json
import os
import re
import zipfile
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP
from html import escape
from urllib.error import HTTPError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

from flask import Flask, Response, jsonify, request
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024
app.config["MAX_FORM_MEMORY_SIZE"] = 25 * 1024 * 1024


@dataclass
class Stay:
    bill_no: str | None
    registration_no: str | None
    folio_no: str | None
    guest_name: str | None
    room_no: str | None
    departure_date: date | None
    check_in_date: date | None
    check_out_date: date | None
    number_of_pax: int | None
    number_of_nights: int | None
    rate_type: str | None
    tariff_before_tax: Decimal
    tax_amount: Decimal
    food_amount: Decimal
    bar_amount: Decimal
    laundry_amount: Decimal
    call_amount: Decimal
    misc_amount: Decimal
    total_amount: Decimal
    amount_paid: Decimal
    payment_method: str | None
    flags: list[str]


def clean(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def dec(value) -> Decimal:
    try:
        return Decimal(clean(value).replace(",", "") or "0")
    except Exception:
        return Decimal("0")


def num(value, default=0) -> int:
    try:
        return int(dec(value))
    except Exception:
        return default


def date_value(value, datemode=None) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)) and datemode is not None:
        import xlrd

        try:
            return xlrd.xldate_as_datetime(value, datemode).date()
        except Exception:
            return None
    text = clean(value)
    if not text:
        return None
    for fmt in ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d", "%m/%d/%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    return None


def money(value) -> str:
    amount = Decimal(str(value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return f"{amount:,.2f}"


def supabase_configured() -> bool:
    return bool(os.environ.get("SUPABASE_URL") and os.environ.get("SUPABASE_SERVICE_ROLE_KEY"))


def supabase_request(method: str, table: str, payload=None, query: dict | None = None, prefer: str | None = None):
    if not supabase_configured():
        raise RuntimeError("Supabase is not configured.")
    base = os.environ["SUPABASE_URL"].rstrip("/")
    key = os.environ["SUPABASE_SERVICE_ROLE_KEY"]
    url = f"{base}/rest/v1/{table}"
    if query:
        url += "?" + urlencode(query)
    body = json.dumps(payload).encode("utf-8") if payload is not None else None
    headers = {
        "apikey": key,
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/json",
    }
    if prefer:
        headers["Prefer"] = prefer
    req = Request(url, data=body, headers=headers, method=method)
    try:
        with urlopen(req, timeout=20) as response:
            text = response.read().decode("utf-8")
            return json.loads(text) if text else None
    except HTTPError as exc:
        detail = exc.read().decode("utf-8", "ignore")
        raise RuntimeError(f"Supabase request failed: {exc.code} {detail}") from exc


def excel_rows(filename: str, data: bytes):
    lower = filename.lower()
    if lower.endswith(".xls"):
        import xlrd

        book = xlrd.open_workbook(file_contents=data)
        for sheet in book.sheets():
            for row_idx in range(sheet.nrows):
                yield [sheet.cell_value(row_idx, col) for col in range(sheet.ncols)], book.datemode
        return
    if lower.endswith(".xlsx"):
        from openpyxl import load_workbook

        workbook = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
        for sheet in workbook.worksheets:
            for row in sheet.iter_rows(values_only=True):
                yield list(row), None
        return
    raise ValueError("Please upload an .xls or .xlsx file.")


def parse_payment_method(text: str) -> str | None:
    match = re.search(r"\[\s*([^:\]]+)\s*:", clean(text))
    return match.group(1).strip() if match else None


def parse_payment_amount(text: str) -> Decimal | None:
    match = re.search(r":\s*([0-9,]+(?:\.[0-9]+)?)\s*\]", clean(text))
    return dec(match.group(1)) if match else None


def parse_stays(filename: str, data: bytes) -> list[Stay]:
    stays = []
    rows = list(excel_rows(filename, data))
    index = 0
    while index < len(rows):
        row, datemode = rows[index]
        if len(row) < 17 or not clean(row[0]).upper().startswith("BN"):
            index += 1
            continue
        details = rows[index + 1][0] if index + 1 < len(rows) else []
        settlement = rows[index + 2][0] if index + 2 < len(rows) else []
        departure = date_value(row[3], datemode)
        nights = max(num(details[2] if len(details) > 2 else None, 0), 0)
        if not departure:
            index += 1
            continue
        check_in = departure - timedelta(days=nights) if nights else departure
        tariff = dec(row[5])
        tax = dec(row[10])
        food = dec(row[11])
        bar = dec(row[12])
        laundry = dec(row[13])
        call = dec(row[14])
        misc = dec(row[15])
        total = dec(row[16])
        paid = dec(settlement[0] if len(settlement) > 0 else None)
        payment_text = settlement[1] if len(settlement) > 1 else ""
        bracket_paid = parse_payment_amount(payment_text)
        if paid == Decimal("0") and bracket_paid is not None:
            paid = bracket_paid
        expected_total = tariff + tax + food + bar + laundry + call + misc
        flags = []
        if abs(paid - total) > Decimal("0.01"):
            flags.append("PAYMENT_MISMATCH")
        if abs(expected_total - total) > Decimal("0.01"):
            flags.append("TOTAL_MISMATCH")
        stays.append(
            Stay(
                bill_no=clean(row[0]) or None,
                registration_no=clean(row[1]) or None,
                folio_no=clean(row[2]) or None,
                guest_name=clean(row[4]).strip() or None,
                room_no=clean(details[0] if len(details) > 0 else None) or None,
                departure_date=departure,
                check_in_date=check_in,
                check_out_date=departure,
                number_of_pax=num(details[1] if len(details) > 1 else None, 0) or None,
                number_of_nights=nights or None,
                rate_type=clean(details[3] if len(details) > 3 else None) or None,
                tariff_before_tax=tariff,
                tax_amount=tax,
                food_amount=food,
                bar_amount=bar,
                laundry_amount=laundry,
                call_amount=call,
                misc_amount=misc,
                total_amount=total,
                amount_paid=paid,
                payment_method=parse_payment_method(payment_text),
                flags=flags,
            )
        )
        index += 4
    if not stays:
        raise ValueError("No guest rows were found. Check that the file is the Sales Bill Register export.")
    return sorted(stays, key=lambda s: (s.check_in_date, s.room_no or "", s.guest_name or ""))


def stay_record(stay: Stay) -> dict:
    return {
        "bill_no": stay.bill_no,
        "registration_no": stay.registration_no,
        "folio_no": stay.folio_no,
        "guest_name": stay.guest_name,
        "room_no": stay.room_no,
        "departure_date": stay.departure_date.isoformat(),
        "check_in_date": stay.check_in_date.isoformat(),
        "check_out_date": stay.check_out_date.isoformat(),
        "number_of_pax": stay.number_of_pax,
        "number_of_nights": stay.number_of_nights,
        "rate_type": stay.rate_type,
        "tariff_before_tax": float(stay.tariff_before_tax),
        "tax_amount": float(stay.tax_amount),
        "food_amount": float(stay.food_amount),
        "bar_amount": float(stay.bar_amount),
        "laundry_amount": float(stay.laundry_amount),
        "call_amount": float(stay.call_amount),
        "misc_amount": float(stay.misc_amount),
        "total_other_charges": float(stay.food_amount + stay.bar_amount + stay.laundry_amount + stay.call_amount + stay.misc_amount),
        "room_revenue": float(stay.tariff_before_tax),
        "total_amount": float(stay.total_amount),
        "amount_paid": float(stay.amount_paid),
        "payment_method": stay.payment_method,
        "flags": stay.flags,
    }


def save_import_to_supabase(filename: str, stays: list[Stay], sales: list[dict], summary_data: dict) -> bool:
    if not supabase_configured():
        return False
    batch = supabase_request(
        "POST",
        "hotel_import_batches",
        {
            "source_filename": filename,
            "stay_count": len(stays),
            "sale_row_count": len(sales),
            "total_sales": str(dec(summary_data.get("total_sales"))),
            "total_kelestarian": str(dec(summary_data.get("total_kelestarian"))),
        },
        query={"select": "id"},
        prefer="return=representation",
    )
    batch_id = batch[0]["id"] if batch else None
    rows = []
    for stay in stays:
        row = stay_record(stay)
        row["import_batch_id"] = batch_id
        row["updated_at"] = datetime.utcnow().isoformat() + "Z"
        rows.append(row)
    for index in range(0, len(rows), 500):
        supabase_request(
            "POST",
            "guest_stays",
            rows[index : index + 500],
            query={"on_conflict": "bill_no"},
            prefer="resolution=merge-duplicates",
        )
    return True


def load_stays_from_supabase() -> list[dict]:
    if not supabase_configured():
        return []
    return supabase_request(
        "GET",
        "guest_stays",
        query={
            "select": "*",
            "order": "check_in_date.asc,room_no.asc,guest_name.asc",
            "limit": "10000",
        },
    ) or []


def records_to_stays(records: list[dict]) -> list[Stay]:
    stays = []
    for item in records:
        check_in = date_value(item.get("check_in_date"))
        check_out = date_value(item.get("check_out_date"))
        if not check_in or not check_out:
            continue
        nights = max(num(item.get("number_of_nights"), 0), 0)
        total = dec(item.get("total_amount"))
        stays.append(
            Stay(
                bill_no=clean(item.get("bill_no")) or None,
                registration_no=clean(item.get("registration_no")) or None,
                folio_no=clean(item.get("folio_no")) or None,
                guest_name=clean(item.get("guest_name")) or None,
                room_no=clean(item.get("room_no")),
                departure_date=check_out,
                check_in_date=check_in,
                check_out_date=check_out,
                number_of_pax=num(item.get("number_of_pax"), 0) or None,
                number_of_nights=nights or None,
                rate_type=clean(item.get("rate_type")) or None,
                tariff_before_tax=dec(item.get("tariff_before_tax")),
                tax_amount=dec(item.get("tax_amount")),
                food_amount=dec(item.get("food_amount")),
                bar_amount=dec(item.get("bar_amount")),
                laundry_amount=dec(item.get("laundry_amount")),
                call_amount=dec(item.get("call_amount")),
                misc_amount=dec(item.get("misc_amount")),
                total_amount=total,
                amount_paid=dec(item.get("amount_paid")),
                payment_method=clean(item.get("payment_method")) or None,
                flags=item.get("flags") or [],
            )
        )
    if not stays:
        raise ValueError("No imported stays found. Please import the Excel file again.")
    return sorted(stays, key=lambda s: (s.check_in_date, s.room_no or "", s.guest_name or ""))


def expanded_sales(stays: list[Stay], fee_rate: Decimal) -> list[dict]:
    rows = []
    for stay in stays:
        if not stay.check_in_date or not stay.number_of_nights:
            continue
        total_kelestarian = fee_rate * Decimal(stay.number_of_nights)
        paid_date = stay.check_in_date.strftime("%d/%m/%Y")
        for offset in range(stay.number_of_nights):
            day = stay.check_in_date + timedelta(days=offset)
            is_check_in_day = offset == 0
            rows.append(
                {
                    "date": day.isoformat(),
                    "display_date": day.strftime("%d/%m/%Y"),
                    "room_no": stay.room_no,
                    "guest_name": stay.guest_name,
                    "stay_progress": f"{offset + 1}/{stay.number_of_nights}",
                    "nights": stay.number_of_nights,
                    "multi_night": stay.number_of_nights > 1,
                    "price": money(stay.amount_paid if is_check_in_day else Decimal("0")),
                    "total_paid": money(stay.amount_paid),
                    "amount_before_tax": money(stay.tariff_before_tax if is_check_in_day else Decimal("0")),
                    "room_revenue": money(stay.tariff_before_tax if is_check_in_day else Decimal("0")),
                    "kelestarian": money(total_kelestarian if is_check_in_day else Decimal("0")),
                    "payment_status": "Collected" if is_check_in_day else f"Paid on {paid_date}",
                    "paid_date": stay.check_in_date.isoformat(),
                    "is_paid_continuation": not is_check_in_day,
                    "bill_no": stay.bill_no,
                    "payment_method": stay.payment_method,
                    "rate_type": stay.rate_type,
                    "flags": stay.flags,
                }
            )
    return sorted(rows, key=lambda r: (r["date"], r["room_no"], r["guest_name"]))


def summary(stays: list[Stay], fee_rate: Decimal) -> dict:
    sales = expanded_sales(stays, fee_rate)
    total_fee = sum((dec(row["kelestarian"]) for row in sales), Decimal("0"))
    total_sales = sum((s.total_amount for s in stays), Decimal("0"))
    days = sorted({row["date"] for row in sales})
    daily_revenue = defaultdict(lambda: {"rooms": 0, "total_amount": Decimal("0"), "kelestarian": Decimal("0")})
    payment_breakdown = defaultdict(lambda: {"count": 0, "total_amount": Decimal("0")})
    room_revenue_report = defaultdict(lambda: {"nights": 0, "room_revenue": Decimal("0"), "total_amount": Decimal("0")})
    occupancy_by_date = defaultdict(set)
    revenue_by_rate_type = defaultdict(lambda: {"count": 0, "total_amount": Decimal("0")})
    issues = []
    for row in sales:
        day = row["display_date"]
        daily_revenue[day]["rooms"] += 1
        daily_revenue[day]["total_amount"] += dec(row["price"])
        daily_revenue[day]["kelestarian"] += dec(row["kelestarian"])
        occupancy_by_date[day].add(row["room_no"])
    for stay in stays:
        payment = stay.payment_method or "Unknown"
        payment_breakdown[payment]["count"] += 1
        payment_breakdown[payment]["total_amount"] += stay.amount_paid
        room_key = stay.room_no or "Unknown"
        room_revenue_report[room_key]["nights"] += stay.number_of_nights or 0
        room_revenue_report[room_key]["room_revenue"] += stay.tariff_before_tax
        room_revenue_report[room_key]["total_amount"] += stay.total_amount
        rate_type = stay.rate_type or "Unknown"
        revenue_by_rate_type[rate_type]["count"] += 1
        revenue_by_rate_type[rate_type]["total_amount"] += stay.total_amount
        if stay.flags:
            issues.append(
                {
                    "bill_no": stay.bill_no,
                    "guest_name": stay.guest_name,
                    "room_no": stay.room_no,
                    "flags": stay.flags,
                    "total_amount": money(stay.total_amount),
                    "amount_paid": money(stay.amount_paid),
                }
            )
    total_nights = sum((s.number_of_nights or 0) for s in stays)
    average_length = Decimal(total_nights) / Decimal(len(stays)) if stays else Decimal("0")
    top_guests = sorted(stays, key=lambda s: s.total_amount, reverse=True)[:10]
    return {
        "stays": len(stays),
        "sale_rows": len(sales),
        "days": len(days),
        "first_date": days[0] if days else "",
        "last_date": days[-1] if days else "",
        "total_sales": money(total_sales),
        "total_kelestarian": money(total_fee),
        "daily_revenue_summary": [
            {"date": key, "rooms": value["rooms"], "total_amount": money(value["total_amount"]), "kelestarian": money(value["kelestarian"])}
            for key, value in daily_revenue.items()
        ],
        "payment_method_breakdown": [
            {"payment_method": key, "count": value["count"], "total_amount": money(value["total_amount"])}
            for key, value in sorted(payment_breakdown.items())
        ],
        "room_revenue_report": [
            {"room_no": key, "nights": value["nights"], "room_revenue": money(value["room_revenue"]), "total_amount": money(value["total_amount"])}
            for key, value in sorted(room_revenue_report.items())
        ],
        "occupancy_by_date": [
            {"date": key, "occupied_rooms": len(value)}
            for key, value in occupancy_by_date.items()
        ],
        "average_length_of_stay": money(average_length),
        "revenue_before_tax": money(sum((s.tariff_before_tax for s in stays), Decimal("0"))),
        "total_tax_collected": money(sum((s.tax_amount for s in stays), Decimal("0"))),
        "top_guests_by_spend": [
            {"guest_name": s.guest_name, "room_no": s.room_no, "total_amount": money(s.total_amount)}
            for s in top_guests
        ],
        "revenue_by_rate_type": [
            {"rate_type": key, "count": value["count"], "total_amount": money(value["total_amount"])}
            for key, value in sorted(revenue_by_rate_type.items())
        ],
        "data_quality_issues": issues,
    }


def styles():
    s = getSampleStyleSheet()
    s.add(ParagraphStyle(name="TitleBlue", parent=s["Title"], alignment=TA_CENTER, fontName="Helvetica-Bold", fontSize=14, leading=17, textColor=colors.HexColor("#1F4E79"), spaceAfter=4))
    s.add(ParagraphStyle(name="SubTitle", parent=s["Normal"], alignment=TA_CENTER, fontName="Helvetica-Bold", fontSize=10, leading=13, spaceAfter=12))
    s.add(ParagraphStyle(name="Small", parent=s["Normal"], alignment=TA_LEFT, fontName="Helvetica", fontSize=8, leading=10))
    s.add(ParagraphStyle(name="SmallBold", parent=s["Small"], fontName="Helvetica-Bold"))
    return s


def p(text, style):
    return Paragraph(escape(str(text)), style)


def info_table(fields, s):
    rows = []
    for i in range(0, len(fields), 2):
        row = []
        for label, value in fields[i : i + 2]:
            row += [p(label, s["SmallBold"]), p(value, s["Small"])]
        while len(row) < 4:
            row.append("")
        rows.append(row)
    table = Table(rows, colWidths=[3.25 * cm, 5.0 * cm, 3.25 * cm, 5.0 * cm])
    table.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#B7C3D0")), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#F2F6FA")), ("BACKGROUND", (2, 0), (2, -1), colors.HexColor("#F2F6FA")), ("PADDING", (0, 0), (-1, -1), 5)]))
    return table


def build_pdf(story, page_size=A4):
    out = io.BytesIO()
    doc = SimpleDocTemplate(out, pagesize=page_size, leftMargin=1.2 * cm, rightMargin=1.2 * cm, topMargin=1.2 * cm, bottomMargin=1.2 * cm)
    doc.build(story)
    return out.getvalue()


def form_fields(settings, month_label="", date_label=""):
    return [
        ("Nama Premis Penginapan", settings.get("premise_name", "")),
        ("No. Lesen Perniagaan (PBT)", settings.get("license_no", "")),
        ("No. Siri Sijil", settings.get("certificate_no", "")),
        ("Kod Kategori Premis", settings.get("category_code", "")),
        ("Alamat Premis Penginapan", settings.get("address", "")),
        ("Bulan & Tahun Pelaporan" if month_label else "Tarikh Hari Daftar Masuk", month_label or date_label),
        ("Wakil Untuk Dihubungi", settings.get("contact_name", "")),
        ("No. Telefon / Emel", settings.get("contact", "")),
    ]


def confirmation(story, s, label, amount):
    story += [Spacer(1, 10), p("PENGESAHAN OLEH PENGUSAHA PREMIS PENGINAPAN", s["SmallBold"]), p("Saya mengesahkan bahawa maklumat ini adalah BENAR dan TEPAT berdasarkan rekod kutipan SEBENAR bagi tempoh yang dilaporkan.", s["Small"]), Spacer(1, 8), p(f"{label} (RM): {money(amount)}", s["Small"]), p(f"Tarikh: {datetime.now().strftime('%d/%m/%Y')}", s["Small"]), Spacer(1, 28), p("................................................", s["Small"]), p("Cop Rasmi & Tandatangan:", s["Small"])]


def form_b_pdf(stays: list[Stay], settings: dict, fee_rate: Decimal) -> bytes:
    s = styles()
    first_day = min(stay.check_in_date for stay in stays if stay.check_in_date)
    grouped = defaultdict(list)
    for row in expanded_sales(stays, fee_rate):
        grouped[row["display_date"]].append(row)
    story = [p("LAPORAN PENYATA KUTIPAN FI KELESTARIAN NEGERI SELANGOR", s["TitleBlue"]), p("(BULANAN)", s["SubTitle"]), info_table(form_fields(settings, month_label=first_day.strftime("%B %Y")), s), Spacer(1, 12)]
    rows = [["Tarikh", "Jumlah Bilik (Unit)", "Bilangan Malam", "Jumlah Kutipan (RM)"]]
    total_rooms = total_nights = 0
    total_fee = Decimal("0")
    for day in sorted(grouped.keys(), key=lambda d: datetime.strptime(d, "%d/%m/%Y")):
        rooms = len(grouped[day])
        fee = sum((dec(row["kelestarian"]) for row in grouped[day]), Decimal("0"))
        total_rooms += rooms
        total_nights += rooms
        total_fee += fee
        rows.append([day, str(rooms), str(rooms), money(fee)])
    rows.append(["Jumlah", str(total_rooms), str(total_nights), money(total_fee)])
    table = Table(rows, repeatRows=1, colWidths=[4.3 * cm, 4.3 * cm, 4.3 * cm, 4.3 * cm])
    table.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#8EA4BA")), ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#D9EAF7")), ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#E2F0D9")), ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"), ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"), ("FONTSIZE", (0, 0), (-1, -1), 8), ("ALIGN", (0, 0), (-1, -1), "CENTER"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("PADDING", (0, 0), (-1, -1), 5)]))
    story.append(table)
    confirmation(story, s, "Jumlah Kutipan Bulanan", total_fee)
    return build_pdf(story)


def form_c_pdf(stays: list[Stay], settings: dict, fee_rate: Decimal) -> bytes:
    s = styles()
    grouped = defaultdict(list)
    for row in expanded_sales(stays, fee_rate):
        grouped[row["display_date"]].append(row)
    story = []
    for idx, day in enumerate(sorted(grouped.keys(), key=lambda d: datetime.strptime(d, "%d/%m/%Y"))):
        if idx:
            story.append(PageBreak())
        rows_for_day = grouped[day]
        day_fee = sum((dec(row["kelestarian"]) for row in rows_for_day), Decimal("0"))
        story += [p("LAPORAN TRANSAKSI PENGGUNAAN BILIK", s["TitleBlue"]), p("(HARIAN)", s["SubTitle"]), info_table(form_fields(settings, date_label=day), s), Spacer(1, 10)]
        rows = [["Bil.", "Nama", "No. Bilik", "Tinggal", "Harga Dibayar (RM)", "Fi Kelestarian (RM)"]]
        for row_no, item in enumerate(rows_for_day, 1):
            rows.append([str(row_no), p(item["guest_name"], s["Small"]), item["room_no"], item["stay_progress"], item["payment_status"] if item["is_paid_continuation"] else item["price"], item["kelestarian"]])
        rows.append(["", "Jumlah Keseluruhan", "", "", "", money(day_fee)])
        table = Table(rows, repeatRows=1, colWidths=[1.0 * cm, 7.2 * cm, 2.0 * cm, 2.0 * cm, 3.4 * cm, 3.4 * cm])
        table.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#8EA4BA")), ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#D9EAF7")), ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#E2F0D9")), ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"), ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"), ("FONTSIZE", (0, 0), (-1, -1), 7), ("ALIGN", (0, 0), (-1, -1), "CENTER"), ("ALIGN", (1, 1), (1, -2), "LEFT"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("PADDING", (0, 0), (-1, -1), 4)]))
        story.append(table)
        confirmation(story, s, "Jumlah Kutipan Harian", day_fee)
    return build_pdf(story, page_size=landscape(A4))


PAGE = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>KL Guest Hotel Sales</title>
<style>
:root{font-family:Inter,ui-sans-serif,system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif;background:#f2f6fb;color:#071a36}*{box-sizing:border-box}body{margin:0}.app{min-height:100vh;display:grid;grid-template-columns:278px 1fr}.sidebar{background:#12233b;color:#fff;padding:28px 18px;display:flex;flex-direction:column}.brand{display:flex;gap:14px;align-items:center;margin-bottom:30px}.logo{width:48px;height:48px;border-radius:9px;background:#fff;color:#0d3265;display:grid;place-items:center;font-weight:900}.brand span{display:block;color:#b5cae8;font-size:13px}.nav-label{font-size:12px;letter-spacing:.12em;text-transform:uppercase;color:#8aa6c8;font-weight:900}.nav{display:grid;gap:8px;margin-top:12px}.nav button{background:transparent;color:#b8d0ee;text-align:left;border:0;padding:14px;border-radius:8px;font-size:17px;cursor:pointer}.nav button.active{background:#1d3a60;color:#fff}.account{border-top:1px solid #314966;margin-top:auto;padding-top:28px;display:flex;gap:12px;align-items:center}.avatar{width:42px;height:42px;border-radius:999px;background:#3478d4;display:grid;place-items:center;font-weight:900}.main{padding:28px 34px 42px}.top{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:22px}.eyebrow{margin:0 0 8px;color:#7587a0;text-transform:uppercase;letter-spacing:.14em;font-size:13px;font-weight:900}h1{font-size:32px;margin:0}h2{font-size:20px;margin:0}.badge{background:#e8f1fd;color:#1c5a9d;border-radius:999px;padding:9px 16px;font-weight:900;font-size:13px}.notice{background:#e8f2ff;border:1px solid #bdd6fb;color:#0b4f9a;border-radius:9px;padding:16px 18px;margin-bottom:18px}.cards{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:18px}.metric,.panel{background:#fff;border:1px solid #d7e0eb;border-radius:10px;box-shadow:0 8px 24px rgba(10,31,68,.04)}.metric{padding:16px}.metric span{display:block;color:#71839c;font-size:13px}.metric strong{display:block;margin-top:4px;font-size:22px}.panel{padding:22px;margin-bottom:18px}.drop{height:118px;border:1.5px dashed #9cb9dc;border-radius:9px;background:#f8fbff;display:grid;place-items:center;text-align:center;color:#1b5fab;font-weight:900;cursor:pointer}.drop.drag{background:#e8f2ff;border-color:#1b5fab}.drop small{display:block;color:#71839c;font-weight:600;margin-top:5px}.drop input{display:none}.toolbar{display:flex;justify-content:space-between;align-items:center;gap:12px;margin:14px 0;flex-wrap:wrap}.toolbar input{max-width:320px}.filters{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin:0 0 14px}.filters button{border:1px solid #c7d7ea;background:#fff;color:#1a4f8d;border-radius:7px;padding:9px 12px;font-weight:900;cursor:pointer}.filters button.active{background:#1d5fa7;color:#fff;border-color:#1d5fa7}.filters input{width:150px}.settings-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:14px}label{display:grid;gap:7px;font-size:13px;font-weight:800;color:#2f4360}input{width:100%;border:1px solid #c9d5e3;border-radius:7px;padding:11px 12px;font:inherit}button.primary{border:0;border-radius:8px;background:#10233e;color:#fff;font-weight:900;padding:12px 18px;cursor:pointer}button.primary:disabled{opacity:.45;cursor:not-allowed}.view{display:none}.view.active{display:block}.table-wrap{overflow:auto;border:1px solid #dde6f1;border-radius:8px;background:#fff}table{width:100%;border-collapse:collapse;font-size:13px}th,td{padding:10px 12px;border-bottom:1px solid #e7edf5;text-align:left;white-space:nowrap}th{background:#f5f8fc;color:#516985;font-size:12px;text-transform:uppercase;letter-spacing:.05em}.day-row td{background:#edf4fc;color:#153a63;font-weight:900;font-size:14px}.multi td{background:#fff7dd}.paid-continuation td{background:#eaf8ef;color:#16633a}.paid-pill{display:inline-block;background:#dff5e8;color:#16633a;border:1px solid #a9dfbf;border-radius:999px;padding:4px 9px;font-weight:900}.empty{color:#71839c;padding:22px;border:1px dashed #c9d8ea;border-radius:8px;background:#f9fbfe}.report-grid{display:grid;grid-template-columns:1fr;gap:18px}.wide{grid-column:span 2}@media(max-width:1050px){.app{grid-template-columns:1fr}.sidebar{display:none}.cards,.settings-grid{grid-template-columns:1fr}.main{padding:20px}.top{flex-direction:column;gap:12px}}
</style>
</head>
<body>
<main class="app">
  <aside class="sidebar">
    <div class="brand"><div class="logo">KL</div><div><strong>KL Guest Hotel</strong><span>Sales & Lampiran</span></div></div>
    <div class="nav-label">Workspace</div>
    <nav class="nav">
      <button data-view="dashboard" class="active">Dashboard</button>
      <button data-view="b">Laporan B</button>
      <button data-view="c">Laporan C</button>
      <button data-view="settings">Settings</button>
    </nav>
    <div class="account"><div class="avatar">A</div><div><strong>Admin</strong><br><span style="color:#a6bddb">@admin</span></div></div>
  </aside>
  <section class="main">
    <header class="top"><div><p class="eyebrow">Sales Tracking Workspace</p><h1 id="pageTitle">Dashboard</h1></div><div class="badge">ADMIN</div></header>
    <section class="notice" id="notice">Drag in the Sales Bill Register Excel file. The dashboard will remember every guest stay and expand multi-night bills by day.</section>
    <section class="cards">
      <article class="metric"><span>Imported stays</span><strong id="mStays">0</strong></article>
      <article class="metric"><span>Sales rows</span><strong id="mRows">0</strong></article>
      <article class="metric"><span>Total sales</span><strong id="mSales">RM 0.00</strong></article>
      <article class="metric"><span>Kelestarian</span><strong id="mFee">RM 0.00</strong></article>
    </section>
    <section class="panel">
        <label class="drop" id="drop"><span id="dropText">Choose or drop Sales Bill Register Excel<small>.xls or .xlsx</small></span><input id="file" type="file" accept=".xls,.xlsx"></label>
    </section>
    <section id="dashboard" class="view active">
      <div class="panel"><div class="toolbar"><h2>Daily sales ledger</h2><input id="search" placeholder="Search guest or room"></div><div class="filters"><button data-range="today">Today</button><button data-range="last7">Last 7 days</button><button data-range="month" class="active">This month</button><button data-range="custom">Custom date</button><input id="startDate" type="date"><input id="endDate" type="date"></div><div id="ledger" class="empty">No imported check-ins yet.</div></div>
      <div class="panel"><h2>Sales reports</h2><br><div id="analytics" class="empty">Import Excel first.</div></div>
    </section>
    <section id="b" class="view">
      <div class="panel"><div class="toolbar"><h2>Laporan B preview</h2><button class="primary" id="downloadB" disabled>Download PDF</button></div><div id="previewB" class="empty">Import Excel first.</div></div>
    </section>
    <section id="c" class="view">
      <div class="panel"><div class="toolbar"><h2>Laporan C preview</h2><button class="primary" id="downloadC" disabled>Download PDF</button></div><div id="previewC" class="empty">Import Excel first.</div></div>
    </section>
    <section id="settings" class="view">
      <div class="panel"><h2>Reporting settings</h2><br><div class="settings-grid">
        <label>Nama Premis Penginapan<input name="premise_name" class="setting" placeholder="KL Guest Hotel"></label>
        <label>No. Lesen Perniagaan (PBT)<input name="license_no" class="setting" placeholder="MBPJ-0000"></label>
        <label>No. Siri Sijil<input name="certificate_no" class="setting"></label>
        <label>Kod Kategori Premis<input name="category_code" class="setting" placeholder="Hotel 1-3 bintang"></label>
        <label>Wakil Untuk Dihubungi<input name="contact_name" class="setting"></label>
        <label>No. Telefon / Emel<input name="contact" class="setting"></label>
        <label>Fi per bilik/malam (RM)<input name="fee_rate" class="setting" type="number" step="0.01" value="5.00"></label>
        <label class="wide">Alamat Premis Penginapan<input name="address" class="setting"></label>
      </div></div>
    </section>
  </section>
</main>
<script>
let stays=JSON.parse(localStorage.getItem("stays")||"[]");
let sales=JSON.parse(localStorage.getItem("sales")||"[]");
let summary=JSON.parse(localStorage.getItem("summary")||"{}");
let rangeMode=localStorage.getItem("rangeMode")||"month";
const q=s=>document.querySelector(s), qa=s=>[...document.querySelectorAll(s)];
function settings(){let o={};qa(".setting").forEach(i=>o[i.name]=i.value);return o}
function asMoney(v){return Number(String(v||"0").replace(/,/g,""))||0}
function iso(d){let z=n=>String(n).padStart(2,"0");return d.getFullYear()+"-"+z(d.getMonth()+1)+"-"+z(d.getDate())}
function rangeBounds(){let now=new Date(),start=new Date(now),end=new Date(now);start.setHours(0,0,0,0);end.setHours(0,0,0,0);if(rangeMode==="last7"){start.setDate(start.getDate()-6)}else if(rangeMode==="month"){start=new Date(now.getFullYear(),now.getMonth(),1);end=new Date(now.getFullYear(),now.getMonth()+1,0)}else if(rangeMode==="custom"){let s=q("#startDate").value,e=q("#endDate").value;return {start:s||"0000-01-01",end:e||"9999-12-31"}}return {start:iso(start),end:iso(end)}}
function visibleRows(){let bounds=rangeBounds(),term=q("#search").value.toLowerCase();return sales.filter(r=>r.date>=bounds.start&&r.date<=bounds.end).filter(r=>(String(r.guest_name||"")+" "+String(r.room_no||"")).toLowerCase().includes(term))}
function groupRows(rows){let html='<div class="table-wrap"><table><thead><tr><th>Room</th><th>Guest</th><th>Stay</th><th>Amount Paid</th><th>Kelestarian</th><th>Payment</th><th>Rate Type</th></tr></thead><tbody>';let cur='';rows.forEach(r=>{let paid=r.is_paid_continuation?`<span class="paid-pill">${r.payment_status||("Paid on "+(r.display_date||""))}</span>`:`RM ${r.price||"0.00"}`;let fee=r.is_paid_continuation?"":`RM ${r.kelestarian||"0.00"}`;if(r.display_date!==cur){cur=r.display_date;html+=`<tr class="day-row"><td colspan="7">${cur}</td></tr>`}html+=`<tr class="${r.is_paid_continuation?'paid-continuation':(r.multi_night?'multi':'')}"><td>${r.room_no||''}</td><td>${r.guest_name||''}</td><td>${r.stay_progress||''}</td><td>${paid}</td><td>${fee}</td><td>${r.payment_method||''}</td><td>${r.rate_type||''}</td></tr>`});return html+'</tbody></table></div>'}
function table(rows,heads,mapper){return '<div class="table-wrap"><table><thead><tr>'+heads.map(h=>`<th>${h}</th>`).join('')+'</tr></thead><tbody>'+rows.map(mapper).join('')+'</tbody></table></div>'}
function collectionBreakdown(rows){let map={};rows.filter(r=>!r.is_paid_continuation).forEach(r=>{let key=r.payment_method||"Unknown";map[key]??={payment_method:key,bills:0,amount:0,kelestarian:0};map[key].bills++;map[key].amount+=asMoney(r.price);map[key].kelestarian+=asMoney(r.kelestarian)});return Object.values(map).sort((a,b)=>b.amount-a.amount)}
function render(){qa(".filters button").forEach(b=>b.classList.toggle("active",b.dataset.range===rangeMode));let filtered=visibleRows();let visibleSales=filtered.reduce((t,r)=>t+asMoney(r.price),0),visibleFee=filtered.reduce((t,r)=>t+asMoney(r.kelestarian),0);q("#mStays").textContent=summary.stays||0;q("#mRows").textContent=filtered.length;q("#mSales").textContent='RM '+visibleSales.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2});q("#mFee").textContent='RM '+visibleFee.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2});q("#downloadB").disabled=!stays.length;q("#downloadC").disabled=!stays.length;q("#ledger").className=filtered.length?'':'empty';q("#ledger").innerHTML=filtered.length?groupRows(filtered):'No rows in this date range.';let byDay={};filtered.forEach(r=>{byDay[r.display_date]??={rooms:0,fee:0};byDay[r.display_date].rooms++;byDay[r.display_date].fee+=asMoney(r.kelestarian)});let bRows=Object.entries(byDay).map(([d,v])=>({d,...v}));q("#previewB").className=bRows.length?'':'empty';q("#previewB").innerHTML=bRows.length?table(bRows,['Tarikh','Bilik Dipaparkan','Malam Dipaparkan','Kutipan Pada Hari Ini'],r=>`<tr><td>${r.d}</td><td>${r.rooms}</td><td>${r.rooms}</td><td>RM ${r.fee.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})}</td></tr>`):'No rows in this date range.';q("#previewC").className=filtered.length?'':'empty';q("#previewC").innerHTML=filtered.length?groupRows(filtered):'No rows in this date range.';let analytics=q("#analytics"),collections=collectionBreakdown(filtered);if(!summary.payment_method_breakdown){analytics.className='empty';analytics.innerHTML='Import Excel first.'}else{analytics.className='';analytics.innerHTML='<div class="cards"><article class="metric"><span>Collected in range</span><strong>RM '+visibleSales.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})+'</strong></article><article class="metric"><span>Kelestarian in range</span><strong>RM '+visibleFee.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})+'</strong></article><article class="metric"><span>Collection methods</span><strong>'+collections.length+'</strong></article><article class="metric"><span>Data issues</span><strong>'+(summary.data_quality_issues||[]).length+'</strong></article></div><h2>Payment collection methods</h2><br>'+ (collections.length?table(collections,['Payment Method','Bills','Collected','Kelestarian'],r=>`<tr><td>${r.payment_method}</td><td>${r.bills}</td><td>RM ${r.amount.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})}</td><td>RM ${r.kelestarian.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})}</td></tr>`):'<div class="empty">No collections in this date range.</div>') + '<br><h2>Revenue by rate type</h2><br>'+table(summary.revenue_by_rate_type,['Rate Type','Bills','Amount'],r=>`<tr><td>${r.rate_type}</td><td>${r.count}</td><td>RM ${r.total_amount}</td></tr>`) }}
async function importFile(file){q("#dropText").innerHTML=file.name+'<small>Importing...</small>';let fd=new FormData();fd.append('file',file);fd.append('fee_rate',settings().fee_rate||'5.00');let res=await fetch('/api/import',{method:'POST',body:fd});let data=await res.json();if(!res.ok){q("#notice").textContent=data.error||'Import failed';return}stays=data.stays;sales=data.sales;summary=data.summary;localStorage.setItem('stays',JSON.stringify(stays));localStorage.setItem('sales',JSON.stringify(sales));localStorage.setItem('summary',JSON.stringify(summary));q("#notice").textContent=`Imported ${summary.stays} stays and expanded them into ${summary.sale_rows} daily sales rows${data.saved?' and saved them to Supabase.':'.'}`;q("#dropText").innerHTML=file.name+'<small>Imported</small>';render()}
async function loadHistory(){let res=await fetch('/api/history?fee_rate='+(settings().fee_rate||'5.00'));let data=await res.json();if(!res.ok||!data.enabled||!data.stays.length)return;stays=data.stays;sales=data.sales;summary=data.summary;localStorage.setItem('stays',JSON.stringify(stays));localStorage.setItem('sales',JSON.stringify(sales));localStorage.setItem('summary',JSON.stringify(summary));q("#notice").textContent=`Loaded ${summary.stays} saved stays from Supabase.`;render()}
async function download(kind){let fd=new FormData();fd.append('kind',kind);fd.append('stays',JSON.stringify(stays));Object.entries(settings()).forEach(([k,v])=>fd.append(k,v));let res=await fetch('/api/report',{method:'POST',body:fd});if(!res.ok){q("#notice").textContent=await res.text();return}let blob=await res.blob(),url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=kind==='b'?'laporan-b.pdf':'laporan-c.pdf';a.click();URL.revokeObjectURL(url)}
qa(".nav button").forEach(b=>b.onclick=()=>{qa(".nav button").forEach(x=>x.classList.remove('active'));b.classList.add('active');qa(".view").forEach(v=>v.classList.remove('active'));q("#"+b.dataset.view).classList.add('active');q("#pageTitle").textContent=b.textContent});
q("#file").onchange=e=>e.target.files[0]&&importFile(e.target.files[0]);q("#search").oninput=render;q("#downloadB").onclick=()=>download('b');q("#downloadC").onclick=()=>download('c');
qa(".filters button").forEach(b=>b.onclick=()=>{rangeMode=b.dataset.range;localStorage.setItem("rangeMode",rangeMode);render()});
q("#startDate").value=localStorage.getItem("startDate")||"";q("#endDate").value=localStorage.getItem("endDate")||"";
["#startDate","#endDate"].forEach(id=>q(id).onchange=()=>{rangeMode="custom";localStorage.setItem("rangeMode",rangeMode);localStorage.setItem(id.slice(1),q(id).value);render()});
let drop=q("#drop");['dragenter','dragover'].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.add('drag')}));['dragleave','drop'].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.remove('drag')}));drop.addEventListener('drop',e=>{if(e.dataTransfer.files[0])importFile(e.dataTransfer.files[0])});
qa(".setting").forEach(i=>{i.value=localStorage.getItem('set_'+i.name)||i.value;i.oninput=()=>localStorage.setItem('set_'+i.name,i.value)});
render();
loadHistory();
</script>
</body>
</html>"""


@app.get("/")
def index():
    return PAGE


@app.post("/api/import")
def api_import():
    upload = request.files.get("file")
    if not upload:
        return jsonify({"error": "Please upload an Excel file."}), 400
    try:
        fee_rate = dec(request.form.get("fee_rate", "5.00"))
        stays = parse_stays(upload.filename, upload.read())
        sales = expanded_sales(stays, fee_rate)
        summary_data = summary(stays, fee_rate)
        saved = False
        save_error = ""
        try:
            saved = save_import_to_supabase(upload.filename, stays, sales, summary_data)
        except Exception as exc:
            save_error = str(exc)
        return jsonify({"stays": [stay_record(s) for s in stays], "sales": sales, "summary": summary_data, "saved": saved, "save_error": save_error})
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400


@app.get("/api/history")
def api_history():
    try:
        if not supabase_configured():
            return jsonify({"enabled": False, "stays": [], "sales": [], "summary": {}})
        fee_rate = dec(request.args.get("fee_rate", "5.00"))
        records = load_stays_from_supabase()
        if not records:
            return jsonify({"enabled": True, "stays": [], "sales": [], "summary": {}})
        stays = records_to_stays(records)
        return jsonify({"enabled": True, "stays": [stay_record(s) for s in stays], "sales": expanded_sales(stays, fee_rate), "summary": summary(stays, fee_rate)})
    except Exception as exc:
        return jsonify({"enabled": supabase_configured(), "error": str(exc)}), 400


@app.post("/api/report")
def api_report():
    try:
        kind = request.form.get("kind", "b")
        stays = records_to_stays(json.loads(request.form.get("stays", "[]")))
        fee_rate = dec(request.form.get("fee_rate", "5.00"))
        settings = {k: escape(request.form.get(k, "").strip()) for k in request.form.keys()}
        pdf = form_b_pdf(stays, settings, fee_rate) if kind == "b" else form_c_pdf(stays, settings, fee_rate)
        name = "laporan-b.pdf" if kind == "b" else "laporan-c.pdf"
        return Response(pdf, mimetype="application/pdf", headers={"Content-Disposition": f"attachment; filename={name}"})
    except Exception as exc:
        return Response(str(exc), status=400)


if __name__ == "__main__":
    app.run(debug=True)
