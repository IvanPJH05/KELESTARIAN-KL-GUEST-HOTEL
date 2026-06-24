from __future__ import annotations

import io
import json
import os
import re
import zipfile
from calendar import month_name, monthrange
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP
from urllib.error import HTTPError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

from flask import Flask, Response, jsonify, request
from reportlab.lib import colors
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas

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


def verification_value(value) -> str:
    return clean(value).upper()


def classify_verified_stays(stays: list[Stay], existing_records: list[dict]) -> tuple[list[dict], list[dict]]:
    existing_by_folio = {
        verification_value(row.get("folio_no")): row
        for row in existing_records
        if verification_value(row.get("folio_no"))
    }
    existing_by_bill = {
        verification_value(row.get("bill_no")): row
        for row in existing_records
        if verification_value(row.get("bill_no"))
    }
    accepted_by_folio = {}
    accepted_bill_to_folio = {}
    reviews = []
    for stay in stays:
        incoming = stay_record(stay)
        folio = verification_value(stay.folio_no)
        bill = verification_value(stay.bill_no)
        incoming["folio_no"] = folio or None
        incoming["bill_no"] = bill or None
        folio_match = accepted_by_folio.get(folio) or existing_by_folio.get(folio)
        bill_match = existing_by_bill.get(bill)
        accepted_bill_folio = accepted_bill_to_folio.get(bill)
        existing_bill_folio = verification_value(bill_match.get("folio_no")) if bill_match else ""
        reason = ""
        existing = None
        if not folio or not bill:
            reason = "MISSING_FOLIO_OR_BILL"
        elif folio_match and verification_value(folio_match.get("bill_no")) != bill:
            reason = "FOLIO_MATCHES_DIFFERENT_BILL"
            existing = folio_match
        elif accepted_bill_folio and accepted_bill_folio != folio:
            reason = "BILL_MATCHES_DIFFERENT_FOLIO"
            existing = accepted_by_folio.get(accepted_bill_folio)
        elif bill_match and existing_bill_folio != folio:
            reason = "BILL_MATCHES_DIFFERENT_FOLIO"
            existing = bill_match
        if reason:
            reviews.append(
                {
                    "review_key": f"{folio or '(missing)'}|{bill or '(missing)'}",
                    "folio_no": folio or None,
                    "incoming_bill_no": bill or None,
                    "existing_folio_no": verification_value((existing or {}).get("folio_no")) or None,
                    "existing_bill_no": verification_value((existing or {}).get("bill_no")) or None,
                    "reason": reason,
                    "incoming_record": incoming,
                    "existing_record": existing,
                    "status": "pending",
                }
            )
            continue
        accepted_by_folio[folio] = incoming
        accepted_bill_to_folio[bill] = folio
    return list(accepted_by_folio.values()), reviews


def save_import_to_supabase(filename: str, stays: list[Stay], sales: list[dict], summary_data: dict) -> dict:
    if not supabase_configured():
        return {"saved": False, "accepted_count": 0, "review_count": 0}
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
    existing_records = load_stays_from_supabase()
    rows, reviews = classify_verified_stays(stays, existing_records)
    for row in rows:
        row["import_batch_id"] = batch_id
        row["updated_at"] = datetime.utcnow().isoformat() + "Z"
    for review in reviews:
        review["import_batch_id"] = batch_id
        review["source_filename"] = filename
        review["updated_at"] = datetime.utcnow().isoformat() + "Z"
    for index in range(0, len(rows), 500):
        supabase_request(
            "POST",
            "guest_stays",
            rows[index : index + 500],
            query={"on_conflict": "folio_no"},
            prefer="resolution=merge-duplicates",
        )
    for index in range(0, len(reviews), 500):
        supabase_request(
            "POST",
            "guest_stay_review_queue",
            reviews[index : index + 500],
            query={"on_conflict": "review_key"},
            prefer="resolution=merge-duplicates",
        )
    return {"saved": True, "accepted_count": len(rows), "review_count": len(reviews)}


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


def load_review_queue_from_supabase() -> list[dict]:
    if not supabase_configured():
        return []
    return supabase_request(
        "GET",
        "guest_stay_review_queue",
        query={
            "select": "id,folio_no,incoming_bill_no,existing_folio_no,existing_bill_no,reason,incoming_record,existing_record,source_filename,status,created_at,updated_at",
            "status": "eq.pending",
            "order": "updated_at.desc",
            "limit": "1000",
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


CREST_PATH = os.path.join(os.path.dirname(__file__), "assets", "selangor-crest.jpeg")
GREY = colors.HexColor("#D9D9D9")
BODY_FONT = "Helvetica"
BOLD_FONT = "Helvetica-Bold"


def _text(c, x, y, value, size=11, bold=False, align="left"):
    value = clean(value)
    c.setFillColor(colors.black)
    c.setFont(BOLD_FONT if bold else BODY_FONT, size)
    if align == "center":
        c.drawCentredString(x, y, value)
    elif align == "right":
        c.drawRightString(x, y, value)
    else:
        c.drawString(x, y, value)


def _fit_size(value, width, size, bold=False, minimum=7.5):
    value = clean(value)
    font = BOLD_FONT if bold else BODY_FONT
    while size > minimum and stringWidth(value, font, size) > width:
        size -= 0.25
    return size


def _wrap_lines(value, width, size=10.5, bold=False, max_lines=2):
    words = clean(value).split()
    font = BOLD_FONT if bold else BODY_FONT
    lines = []
    current = ""
    for word in words:
        candidate = f"{current} {word}".strip()
        if current and stringWidth(candidate, font, size) > width:
            lines.append(current)
            current = word
        else:
            current = candidate
    if current:
        lines.append(current)
    return lines[:max_lines]


def _outer(c, label, x, right, bottom, top=743):
    c.setStrokeColor(colors.black)
    c.setLineWidth(0.55)
    c.rect(x, bottom, right - x, top - bottom, stroke=1, fill=0)
    _text(c, right - 22, top + 3, label, 12, True, "right")


def _crest(c, x=281.5, y=645, width=49, height=72):
    if os.path.exists(CREST_PATH):
        c.drawImage(ImageReader(CREST_PATH), x, y, width, height, preserveAspectRatio=True, mask="auto")


def _grid(c, x, top, widths, row_heights, fills=None):
    y = top
    total_width = sum(widths)
    for row_no, height in enumerate(row_heights):
        y -= height
        if fills and fills[row_no]:
            c.setFillColor(fills[row_no])
            c.rect(x, y, total_width, height, stroke=0, fill=1)
        c.setStrokeColor(colors.black)
        c.setLineWidth(0.45)
        c.rect(x, y, total_width, height, stroke=1, fill=0)
        cursor = x
        for width in widths[:-1]:
            cursor += width
            c.line(cursor, y, cursor, y + height)
    return y


def _cell_text(c, x, y, width, height, value, size=10, bold=False, align="center"):
    size = _fit_size(value, width - 7, size, bold)
    baseline = y + (height - size) / 2 + 1
    if align == "left":
        _text(c, x + 4, baseline, value, size, bold)
    else:
        _text(c, x + width / 2, baseline, value, size, bold, "center")


def _b_fields(c, labels, values):
    label_x = 62.5
    colon_x = 280.5
    box_x = 288.5
    box_width = 266.5
    top = 570
    heights = [20.5, 20.5, 20.5, 41, 20.5, 20.5, 20.5, 20.5]
    y = top
    for index, (label, value, height) in enumerate(zip(labels, values, heights)):
        y -= height
        label_size = _fit_size(label, colon_x - label_x - 8, 11.5)
        _text(c, label_x, y + (height - label_size) / 2 + 1, label, label_size)
        _text(c, colon_x, y + (height - 11) / 2 + 1, ":", 11)
        c.rect(box_x, y, box_width, height, stroke=1, fill=0)
        if index == 3:
            lines = _wrap_lines(value, box_width - 12, 10.5, max_lines=2)
            for line_no, line in enumerate(lines):
                _text(c, box_x + 6, y + height - 14 - line_no * 14, line, 10.5)
        else:
            size = _fit_size(value, box_width - 12, 10.5)
            _text(c, box_x + 6, y + (height - size) / 2 + 1, value, size)
    return y


B_WIDTHS = [87, 142.5, 126, 143]


def _b_table(c, top, rows, include_header=False, include_total=None):
    rendered = []
    fills = []
    if include_header:
        rendered.append(["Tarikh", "Jumlah Bilik (Unit)", "Bilangan Malam", "Jumlah Kutipan (RM)"])
        fills.append(GREY)
    rendered.extend(rows)
    fills.extend([None] * len(rows))
    if include_total is not None:
        rendered.append(include_total)
        fills.append(GREY)
    heights = [17.5] * len(rendered)
    bottom = _grid(c, 56.5, top, B_WIDTHS, heights, fills)
    y = top
    for row_no, row in enumerate(rendered):
        y -= heights[row_no]
        cursor = 56.5
        for col_no, value in enumerate(row):
            _cell_text(c, cursor, y, B_WIDTHS[col_no], heights[row_no], value, 10.5, bool(fills[row_no]))
            cursor += B_WIDTHS[col_no]
    return bottom


def _b_confirmation(c, total_fee, table_bottom):
    heading_y = table_bottom - 27
    _text(c, 56.5, heading_y, "C. PENGESAHAN OLEH PENGUSAHA PREMIS PENGINAPAN", 12, True)
    _text(c, 56.5, heading_y - 26, "Saya mengesahkan bahawa maklumat ini adalah BENAR dan TEPAT berdasarkan rekod", 11)
    _text(c, 56.5, heading_y - 41, "kutipan SEBENAR bagi bulan yang dilaporkan.", 11)
    _text(c, 56.5, heading_y - 71, f"Jumlah Kutipan Bulanan (RM) : {money(total_fee)}", 11.5, True)
    _text(c, 56.5, heading_y - 88, f"Tarikh : {datetime.now().strftime('%d/%m/%Y')}", 11.5, True)
    _text(c, 56.5, heading_y - 157, "................................................", 11, True)
    _text(c, 56.5, heading_y - 172, "Cop Rasmi & Tandatangan:", 11.5, True)
    divider = heading_y - 188
    c.line(50.5, divider, 561, divider)
    _text(c, 56.5, divider - 22, "D. UNTUK KEGUNAAN PEJABAT PBT SAHAJA", 12, True)
    _text(c, 56.5, divider - 50, "(SEMAKAN & PENGESAHAN PBT)", 12, True)
    pbt = ["Tarikh Diterima", "Ulasan / Catatan", "Jumlah Kutipan Bulanan (RM)", "Caj lewat (RM) (Sekiranya ada)"]
    for index, label in enumerate(pbt):
        y = divider - 83 - index * 17
        _text(c, 62, y, label, 11)
        _text(c, 240, y, ": ................................................", 11)
    _text(c, 56.5, divider - 202, "................................................", 11, True)
    _text(c, 56.5, divider - 217, "Cop Rasmi & Tandatangan:", 11.5, True)


def report_filename(kind: str, stays: list[Stay]) -> str:
    check_in_dates = sorted({stay.check_in_date for stay in stays if stay.check_in_date})
    if kind == "b":
        months = sorted({(day.year, day.month) for day in check_in_dates})
        labels = [f"{month_name[month]} {year}" for year, month in months]
        period = labels[0] if len(labels) == 1 else f"{labels[0]} to {labels[-1]}"
        return f"Lampiran B {period}.pdf"
    labels = [f"{day.day:02d} {month_name[day.month]} {day.year}" for day in check_in_dates]
    period = labels[0] if len(labels) == 1 else f"{labels[0]} to {labels[-1]}"
    return f"Lampiran C {period}.pdf"


def filter_stays_for_report(stays: list[Stay], kind: str, report_month="", report_start="", report_end="") -> list[Stay]:
    if kind == "b" and report_month:
        if not re.fullmatch(r"\d{4}-\d{2}", report_month):
            raise ValueError("Please select a valid month for Lampiran B.")
        year, month = (int(part) for part in report_month.split("-"))
        selected = [stay for stay in stays if stay.check_in_date and stay.check_in_date.year == year and stay.check_in_date.month == month]
    elif kind == "c" and (report_start or report_end):
        try:
            start = date.fromisoformat(report_start or report_end)
            end = date.fromisoformat(report_end or report_start)
        except ValueError as exc:
            raise ValueError("Please select valid dates for Lampiran C.") from exc
        if start > end:
            raise ValueError("Lampiran C start date must be before or equal to the end date.")
        selected = [stay for stay in stays if stay.check_in_date and start <= stay.check_in_date <= end]
    else:
        selected = stays
    if not selected:
        label = "month" if kind == "b" else "date range"
        raise ValueError(f"No guest check-ins were found for the selected {label}.")
    return selected


def form_b_pdf(stays: list[Stay], settings: dict, fee_rate: Decimal) -> bytes:
    check_in_rows = [row for row in expanded_sales(stays, fee_rate) if not row["is_paid_continuation"]]
    months = defaultdict(list)
    for row in check_in_rows:
        day = date.fromisoformat(row["date"])
        months[(day.year, day.month)].append(row)
    out = io.BytesIO()
    c = canvas.Canvas(out, pagesize=LETTER)
    for month_index, ((year, month), month_rows) in enumerate(sorted(months.items())):
        if month_index:
            c.showPage()
        daily = defaultdict(list)
        for row in month_rows:
            daily[date.fromisoformat(row["date"]).day].append(row)
        values = [
            settings.get("premise_name", ""), settings.get("license_no", ""),
            settings.get("certificate_no", ""), settings.get("address", ""),
            settings.get("contact_name", ""), settings.get("contact", ""),
            settings.get("category_code", ""), f"{month_name[month]} {year}",
        ]
        labels = ["Nama Premis Penginapan", "No. Lesen Perniagaan (PBT)", "No. Siri Sijil", "Alamat Premis Penginapan", "Nama wakil premis untuk dihubungi", "No. Telefon / Emel", "Kod Kategori Premis Penginapan", "Bulan & Tahun Pelaporan"]
        _outer(c, "LAMPIRAN B", 50.5, 561, 14.5)
        _crest(c)
        _text(c, LETTER[0] / 2, 620, "LAPORAN PENYATA KUTIPAN FI KELESTARIAN NEGERI SELANGOR (BULANAN)", 12, True, "center")
        _text(c, 56.5, 589, "A. MAKLUMAT PREMIS PENGINAPAN", 12, True)
        _b_fields(c, labels, values)
        _text(c, 56.5, 365, "B. MAKLUMAT KUTIPAN FI KELESTARIAN", 12, True)
        day_rows = []
        total_rooms = total_nights = 0
        total_fee = Decimal("0")
        for day_no in range(1, monthrange(year, month)[1] + 1):
            collected = daily.get(day_no, [])
            rooms = len(collected)
            nights = sum(item["nights"] for item in collected)
            fee = sum((dec(item["kelestarian"]) for item in collected), Decimal("0"))
            total_rooms += rooms
            total_nights += nights
            total_fee += fee
            day_rows.append([f"{day_no}/{month:02d}/{year}", str(rooms) if rooms else "", str(nights) if nights else "", money(fee) if fee else ""])
        _b_table(c, 347.5, day_rows[:18], include_header=True)
        c.showPage()
        _outer(c, "LAMPIRAN B", 50.5, 561, 29)
        bottom = _b_table(c, 743, day_rows[18:], include_total=["Jumlah", str(total_rooms), str(total_nights), money(total_fee)])
        _b_confirmation(c, total_fee, bottom)
    c.save()
    return out.getvalue()


C_WIDTHS = [29.5, 276.5, 85, 71, 107.5]
C_HEADERS = ["Bil.", "Nama", "Jumlah Bilik\n(Unit)", "Bilangan\nMalam", "Jumlah Fi\nKelestarian (RM)"]
C_MANUAL_ROWS = 5


def _c_header_cells(c, x, y, height=40):
    cursor = x
    for width, label in zip(C_WIDTHS, C_HEADERS):
        lines = label.split("\n")
        baseline = y + height / 2 + (len(lines) - 1) * 6 - 4
        for line_no, line in enumerate(lines):
            size = _fit_size(line, width - 6, 11.5, True, 8.5)
            _text(c, cursor + width / 2, baseline - line_no * 13, line, size, True, "center")
        cursor += width


def _c_rows(c, top, items, capacity, start_number, include_header=False):
    header_height = 40 if include_header else 0
    row_height = 17.5
    padded = items + [None] * max(0, capacity - len(items))
    fills = ([GREY] if include_header else []) + [None] * len(padded)
    heights = ([header_height] if include_header else []) + [row_height] * len(padded)
    bottom = _grid(c, 21, top, C_WIDTHS, heights, fills)
    y = top
    if include_header:
        y -= header_height
        _c_header_cells(c, 21, y, header_height)
    for local_no, item in enumerate(padded):
        y -= row_height
        if item is None:
            continue
        values = [str(start_number + local_no), item["guest_name"], "1", str(item["nights"]), item["kelestarian"]]
        cursor = 21
        for col_no, value in enumerate(values):
            _cell_text(c, cursor, y, C_WIDTHS[col_no], row_height, value, 10.5, False, "left" if col_no == 1 else "center")
            cursor += C_WIDTHS[col_no]
    return bottom


def _c_first_page(c, settings, day_label, items):
    _outer(c, "LAMPIRAN C", 15.5, 596, 35.5)
    _crest(c, 281.5, 664, 49, 72)
    _text(c, LETTER[0] / 2, 641, "LAPORAN TRANSAKSI PENGGUNAAN BILIK (HARIAN)", 12, True, "center")
    labels = ["Nama Premis Penginapan", "No. Rujukan Lesen (PBT)", "No. Siri Sijil", "Tarikh Hari Daftar Masuk (Check-In)"]
    values = [settings.get("premise_name", ""), settings.get("license_no", ""), settings.get("certificate_no", ""), day_label]
    box_tops = [635, 606, 576.5, 547.5]
    for label, value, box_top in zip(labels, values, box_tops):
        size = _fit_size(label, 215, 12)
        _text(c, 40.5, box_top - 12, label, size)
        _text(c, 263, box_top - 12, ":", 12)
        c.rect(270.5, box_top - 15, 306.5, 15, stroke=1, fill=0)
        value_size = _fit_size(value, 294, 10.5)
        _text(c, 276.5, box_top - 12, value, value_size)
    _c_rows(c, 513.5, items, 25, 1, include_header=True)


def _c_confirmation(c, table_bottom, rows_for_day):
    total_rooms = len(rows_for_day)
    total_nights = sum(row["nights"] for row in rows_for_day)
    total_fee = sum((dec(row["kelestarian"]) for row in rows_for_day), Decimal("0"))
    height = 21.5
    c.setFillColor(GREY)
    c.rect(21, table_bottom - height, sum(C_WIDTHS), height, stroke=0, fill=1)
    c.setStrokeColor(colors.black)
    c.rect(21, table_bottom - height, sum(C_WIDTHS), height, stroke=1, fill=0)
    cursor = 21 + sum(C_WIDTHS[:2])
    for width in C_WIDTHS[2:]:
        c.line(cursor, table_bottom - height, cursor, table_bottom)
        cursor += width
    _text(c, 21 + sum(C_WIDTHS[:2]) / 2, table_bottom - 15, "Jumlah Keseluruhan (Unit @ RM)", 11.5, True, "center")
    cursor = 21 + sum(C_WIDTHS[:2])
    for width, value in zip(C_WIDTHS[2:], [str(total_rooms), str(total_nights), money(total_fee)]):
        _cell_text(c, cursor, table_bottom - height, width, height, value, 10.5, True)
        cursor += width
    box_top = table_bottom - height - 11
    c.line(15.5, box_top, 596, box_top)
    _text(c, 21, box_top - 34, "PENGESAHAN OLEH PENGUSAHA PREMIS PENGINAPAN", 12, True)
    _text(c, 21, box_top - 57, "Saya mengesahkan bahawa maklumat ini adalah BENAR dan TEPAT berdasarkan rekod kutipan", 11)
    _text(c, 21, box_top - 72, "SEBENAR bagi hari yang dilaporkan.", 11, True)
    _text(c, 21, box_top - 101, f"Jumlah Kutipan Harian (RM) : {money(total_fee)}", 11.5, True)
    _text(c, 21, box_top - 118, f"Tarikh : {datetime.now().strftime('%d/%m/%Y')}", 11.5, True)
    _text(c, 21, box_top - 185, "................................................", 11, True)
    _text(c, 21, box_top - 200, "Cop Rasmi & Tandatangan:", 11.5, True)


def form_c_pdf(stays: list[Stay], settings: dict, fee_rate: Decimal) -> bytes:
    grouped = defaultdict(list)
    for row in expanded_sales(stays, fee_rate):
        if not row["is_paid_continuation"]:
            grouped[row["date"]].append(row)
    out = io.BytesIO()
    c = canvas.Canvas(out, pagesize=LETTER)
    first_report = True
    for day_key in sorted(grouped):
        if not first_report:
            c.showPage()
        first_report = False
        rows_for_day = grouped[day_key]
        day_label = date.fromisoformat(day_key).strftime("%d/%m/%Y")
        _c_first_page(c, settings, day_label, rows_for_day[:25])
        remaining = rows_for_day[25:]
        row_number = 26
        while len(remaining) > 16:
            c.showPage()
            _outer(c, "LAMPIRAN C", 15.5, 596, 35.5)
            chunk_size = min(35, len(remaining) - 16)
            chunk = remaining[:chunk_size]
            _c_rows(c, 743, chunk, len(chunk), row_number, include_header=True)
            row_number += len(chunk)
            remaining = remaining[len(chunk):]
        c.showPage()
        _outer(c, "LAMPIRAN C", 15.5, 596, 134)
        bottom = _c_rows(c, 743, remaining, len(remaining) + C_MANUAL_ROWS, row_number, include_header=False)
        _c_confirmation(c, bottom, rows_for_day)
    c.save()
    return out.getvalue()


PAGE = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>KL Guest Hotel Sales</title>
<style>
:root{font-family:Inter,ui-sans-serif,system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif;background:#f2f6fb;color:#071a36}*{box-sizing:border-box}body{margin:0}.app{min-height:100vh;display:grid;grid-template-columns:278px 1fr}.sidebar{background:#12233b;color:#fff;padding:28px 18px;display:flex;flex-direction:column}.brand{display:flex;gap:14px;align-items:center;margin-bottom:30px}.logo{width:48px;height:48px;border-radius:9px;background:#fff;color:#0d3265;display:grid;place-items:center;font-weight:900}.brand span{display:block;color:#b5cae8;font-size:13px}.nav-label{font-size:12px;letter-spacing:.12em;text-transform:uppercase;color:#8aa6c8;font-weight:900}.nav{display:grid;gap:8px;margin-top:12px}.nav button{background:transparent;color:#b8d0ee;text-align:left;border:0;padding:14px;border-radius:8px;font-size:17px;cursor:pointer}.nav button.active{background:#1d3a60;color:#fff}.account{border-top:1px solid #314966;margin-top:auto;padding-top:28px;display:flex;gap:12px;align-items:center}.avatar{width:42px;height:42px;border-radius:999px;background:#3478d4;display:grid;place-items:center;font-weight:900}.main{padding:28px 34px 42px}.top{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:22px}.eyebrow{margin:0 0 8px;color:#7587a0;text-transform:uppercase;letter-spacing:.14em;font-size:13px;font-weight:900}h1{font-size:32px;margin:0}h2{font-size:20px;margin:0}.badge{background:#e8f1fd;color:#1c5a9d;border-radius:999px;padding:9px 16px;font-weight:900;font-size:13px}.notice{background:#e8f2ff;border:1px solid #bdd6fb;color:#0b4f9a;border-radius:9px;padding:16px 18px;margin-bottom:18px}.cards{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:18px}.metric,.panel{background:#fff;border:1px solid #d7e0eb;border-radius:10px;box-shadow:0 8px 24px rgba(10,31,68,.04)}.metric{padding:16px}.metric span{display:block;color:#71839c;font-size:13px}.metric strong{display:block;margin-top:4px;font-size:22px}.panel{padding:22px;margin-bottom:18px}.drop{height:118px;border:1.5px dashed #9cb9dc;border-radius:9px;background:#f8fbff;display:grid;place-items:center;text-align:center;color:#1b5fab;font-weight:900;cursor:pointer}.drop.drag{background:#e8f2ff;border-color:#1b5fab}.drop small{display:block;color:#71839c;font-weight:600;margin-top:5px}.drop input{display:none}.toolbar{display:flex;justify-content:space-between;align-items:center;gap:12px;margin:14px 0;flex-wrap:wrap}.toolbar input{max-width:320px}.filters,.report-controls{display:flex;gap:8px;align-items:end;flex-wrap:wrap;margin:0 0 14px}.filters button{border:1px solid #c7d7ea;background:#fff;color:#1a4f8d;border-radius:7px;padding:9px 12px;font-weight:900;cursor:pointer}.filters button.active{background:#1d5fa7;color:#fff;border-color:#1d5fa7}.filters input{width:150px}.report-controls label{min-width:190px}.settings-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:14px}label{display:grid;gap:7px;font-size:13px;font-weight:800;color:#2f4360}input{width:100%;border:1px solid #c9d5e3;border-radius:7px;padding:11px 12px;font:inherit}button.primary{border:0;border-radius:8px;background:#10233e;color:#fff;font-weight:900;padding:12px 18px;cursor:pointer}button.primary:disabled{opacity:.45;cursor:not-allowed}.view{display:none}.view.active{display:block}.table-wrap{overflow:auto;border:1px solid #dde6f1;border-radius:8px;background:#fff}table{width:100%;border-collapse:collapse;font-size:13px}th,td{padding:10px 12px;border-bottom:1px solid #e7edf5;text-align:left;white-space:nowrap}th{background:#f5f8fc;color:#516985;font-size:12px;text-transform:uppercase;letter-spacing:.05em}.day-row td{background:#edf4fc;color:#153a63;font-weight:900;font-size:14px}.multi td{background:#fff7dd}.paid-continuation td{background:#eaf8ef;color:#16633a}.paid-pill{display:inline-block;background:#dff5e8;color:#16633a;border:1px solid #a9dfbf;border-radius:999px;padding:4px 9px;font-weight:900}.review-reason{color:#9b2c2c;font-weight:900}.empty{color:#71839c;padding:22px;border:1px dashed #c9d8ea;border-radius:8px;background:#f9fbfe}.report-grid{display:grid;grid-template-columns:1fr;gap:18px}.wide{grid-column:span 2}@media(max-width:1050px){.app{grid-template-columns:1fr}.sidebar{display:none}.cards,.settings-grid{grid-template-columns:1fr}.main{padding:20px}.top{flex-direction:column;gap:12px}}
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
      <button data-view="review">Manual Review</button>
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
      <div class="panel"><div class="toolbar"><h2>Laporan B preview</h2><button class="primary" id="downloadB" disabled>Download PDF</button></div><div class="report-controls"><label>Month to print<input id="reportMonthB" type="month"></label></div><div id="previewB" class="empty">Import Excel first.</div></div>
    </section>
    <section id="c" class="view">
      <div class="panel"><div class="toolbar"><h2>Laporan C preview</h2><button class="primary" id="downloadC" disabled>Download PDF</button></div><div class="report-controls"><label>First date<input id="reportStartC" type="date"></label><label>Last date<input id="reportEndC" type="date"></label></div><div id="previewC" class="empty">Import Excel first.</div></div>
    </section>
    <section id="review" class="view">
      <div class="panel"><div class="toolbar"><div><h2>Manual Review</h2><small>Folio and Bill Number mismatches are kept out of the guest database.</small></div><button class="primary" id="refreshReview">Refresh</button></div><div id="reviewQueue" class="empty">No records need manual review.</div></div>
    </section>
    <section id="settings" class="view">
      <div class="panel"><h2>Reporting settings</h2><br><div class="settings-grid">
        <label>Nama Premis Penginapan<input name="premise_name" class="setting" placeholder="KL Guest Hotel"></label>
        <label>No. Lesen Perniagaan (PBT)<input name="license_no" class="setting" placeholder="MBPJ-0000"></label>
        <label>No. Siri Sijil<input name="certificate_no" class="setting"></label>
        <label>Kod Kategori Premis<input name="category_code" class="setting" placeholder="Hotel 1-3 bintang"></label>
        <label>Wakil Untuk Dihubungi<input name="contact_name" class="setting"></label>
        <label>No. Telefon / Emel<input name="contact" class="setting" value="012-205-0039 / hueyjiunphang@gmail.com"></label>
        <label>Fi per bilik/malam (RM)<input name="fee_rate" class="setting" type="number" step="0.01" value="5.00"></label>
        <label class="wide">Alamat Premis Penginapan<input name="address" class="setting" value="8, Jalan AU 1a/4c, Taman Keramat Permai, 54200 Kuala Lumpur, Federal Territory of Kuala Lumpur"></label>
      </div></div>
    </section>
  </section>
</main>
<script>
let stays=JSON.parse(localStorage.getItem("stays")||"[]");
let sales=JSON.parse(localStorage.getItem("sales")||"[]");
let summary=JSON.parse(localStorage.getItem("summary")||"{}");
let rangeMode=localStorage.getItem("rangeMode")||"month";
let reviewQueue=[];
const q=s=>document.querySelector(s), qa=s=>[...document.querySelectorAll(s)];
function settings(){let o={};qa(".setting").forEach(i=>o[i.name]=i.value);return o}
function asMoney(v){return Number(String(v||"0").replace(/,/g,""))||0}
function html(v){return String(v??"").replace(/[&<>"']/g,ch=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[ch]))}
function iso(d){let z=n=>String(n).padStart(2,"0");return d.getFullYear()+"-"+z(d.getMonth()+1)+"-"+z(d.getDate())}
function rangeBounds(){let now=new Date(),start=new Date(now),end=new Date(now);start.setHours(0,0,0,0);end.setHours(0,0,0,0);if(rangeMode==="last7"){start.setDate(start.getDate()-6)}else if(rangeMode==="month"){start=new Date(now.getFullYear(),now.getMonth(),1);end=new Date(now.getFullYear(),now.getMonth()+1,0)}else if(rangeMode==="custom"){let s=q("#startDate").value,e=q("#endDate").value;return {start:s||"0000-01-01",end:e||"9999-12-31"}}return {start:iso(start),end:iso(end)}}
function dataBounds(){let dates=sales.map(r=>r.date).filter(Boolean).sort();return dates.length?{start:dates[0],end:dates.at(-1)}:null}
function showImportedRange(){let bounds=dataBounds();if(!bounds)return "";rangeMode="custom";q("#startDate").value=bounds.start;q("#endDate").value=bounds.end;localStorage.setItem("rangeMode",rangeMode);localStorage.setItem("startDate",bounds.start);localStorage.setItem("endDate",bounds.end);return ` Showing ${bounds.start.split("-").reverse().join("/")} to ${bounds.end.split("-").reverse().join("/")}.`}
function visibleRows(){let bounds=rangeBounds(),term=q("#search").value.toLowerCase();return sales.filter(r=>r.date>=bounds.start&&r.date<=bounds.end).filter(r=>(String(r.guest_name||"")+" "+String(r.room_no||"")).toLowerCase().includes(term))}
function groupRows(rows){let html='<div class="table-wrap"><table><thead><tr><th>Room</th><th>Guest</th><th>Stay</th><th>Amount Paid</th><th>Kelestarian</th><th>Payment</th><th>Rate Type</th></tr></thead><tbody>';let cur='';rows.forEach(r=>{let paid=r.is_paid_continuation?`<span class="paid-pill">${r.payment_status||("Paid on "+(r.display_date||""))}</span>`:`RM ${r.price||"0.00"}`;let fee=r.is_paid_continuation?"":`RM ${r.kelestarian||"0.00"}`;if(r.display_date!==cur){cur=r.display_date;html+=`<tr class="day-row"><td colspan="7">${cur}</td></tr>`}html+=`<tr class="${r.is_paid_continuation?'paid-continuation':(r.multi_night?'multi':'')}"><td>${r.room_no||''}</td><td>${r.guest_name||''}</td><td>${r.stay_progress||''}</td><td>${paid}</td><td>${fee}</td><td>${r.payment_method||''}</td><td>${r.rate_type||''}</td></tr>`});return html+'</tbody></table></div>'}
function table(rows,heads,mapper){return '<div class="table-wrap"><table><thead><tr>'+heads.map(h=>`<th>${h}</th>`).join('')+'</tr></thead><tbody>'+rows.map(mapper).join('')+'</tbody></table></div>'}
function collectionBreakdown(rows){let map={};rows.filter(r=>!r.is_paid_continuation).forEach(r=>{let key=r.payment_method||"Unknown";map[key]??={payment_method:key,bills:0,amount:0,kelestarian:0};map[key].bills++;map[key].amount+=asMoney(r.price);map[key].kelestarian+=asMoney(r.kelestarian)});return Object.values(map).sort((a,b)=>b.amount-a.amount)}
function setReportDefaults(){let dates=stays.map(s=>s.check_in_date).filter(Boolean).sort(),latest=dates.at(-1);if(!latest)return;if(!q("#reportMonthB").value)q("#reportMonthB").value=localStorage.getItem("reportMonthB")||latest.slice(0,7);if(!q("#reportStartC").value)q("#reportStartC").value=localStorage.getItem("reportStartC")||latest;if(!q("#reportEndC").value)q("#reportEndC").value=localStorage.getItem("reportEndC")||latest}
function selectedCheckIns(kind){let rows=sales.filter(r=>!r.is_paid_continuation);if(kind==='b'){let month=q("#reportMonthB").value;return rows.filter(r=>r.date.slice(0,7)===month)}let start=q("#reportStartC").value,end=q("#reportEndC").value;return rows.filter(r=>(!start||r.date>=start)&&(!end||r.date<=end))}
function renderReviewQueue(){let el=q("#reviewQueue");el.className=reviewQueue.length?'':'empty';el.innerHTML=reviewQueue.length?table(reviewQueue,['Folio No.','Incoming Bill','Existing Bill','Guest','Reason','Source'],r=>`<tr><td>${html(r.folio_no)}</td><td>${html(r.incoming_bill_no)}</td><td>${html(r.existing_bill_no)}</td><td>${html((r.incoming_record||{}).guest_name)}</td><td class="review-reason">${html(String(r.reason||'').replaceAll('_',' '))}</td><td>${html(r.source_filename)}</td></tr>`):'No records need manual review.'}
function render(){setReportDefaults();qa(".filters button").forEach(b=>b.classList.toggle("active",b.dataset.range===rangeMode));let filtered=visibleRows();let visibleSales=filtered.reduce((t,r)=>t+asMoney(r.price),0),visibleFee=filtered.reduce((t,r)=>t+asMoney(r.kelestarian),0);q("#mStays").textContent=summary.stays||0;q("#mRows").textContent=filtered.length;q("#mSales").textContent='RM '+visibleSales.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2});q("#mFee").textContent='RM '+visibleFee.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2});q("#downloadB").disabled=!stays.length;q("#downloadC").disabled=!stays.length;q("#ledger").className=filtered.length?'':'empty';q("#ledger").innerHTML=filtered.length?groupRows(filtered):'No rows in this date range.';let bSelected=selectedCheckIns('b'),byDay={};bSelected.forEach(r=>{byDay[r.display_date]??={rooms:0,nights:0,fee:0};byDay[r.display_date].rooms++;byDay[r.display_date].nights+=Number(r.nights||0);byDay[r.display_date].fee+=asMoney(r.kelestarian)});let bRows=Object.entries(byDay).map(([d,v])=>({d,...v}));q("#previewB").className=bRows.length?'':'empty';q("#previewB").innerHTML=bRows.length?table(bRows,['Tarikh','Jumlah Bilik','Bilangan Malam','Jumlah Kutipan'],r=>`<tr><td>${r.d}</td><td>${r.rooms}</td><td>${r.nights}</td><td>RM ${r.fee.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})}</td></tr>`):'No check-ins in the selected month.';let cSelected=selectedCheckIns('c');q("#previewC").className=cSelected.length?'':'empty';q("#previewC").innerHTML=cSelected.length?groupRows(cSelected):'No check-ins in the selected date range.';let analytics=q("#analytics"),collections=collectionBreakdown(filtered);if(!summary.payment_method_breakdown){analytics.className='empty';analytics.innerHTML='Import Excel first.'}else{analytics.className='';analytics.innerHTML='<div class="cards"><article class="metric"><span>Collected in range</span><strong>RM '+visibleSales.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})+'</strong></article><article class="metric"><span>Kelestarian in range</span><strong>RM '+visibleFee.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})+'</strong></article><article class="metric"><span>Collection methods</span><strong>'+collections.length+'</strong></article><article class="metric"><span>Data issues</span><strong>'+(summary.data_quality_issues||[]).length+'</strong></article></div><h2>Payment collection methods</h2><br>'+ (collections.length?table(collections,['Payment Method','Bills','Collected','Kelestarian'],r=>`<tr><td>${r.payment_method}</td><td>${r.bills}</td><td>RM ${r.amount.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})}</td><td>RM ${r.kelestarian.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})}</td></tr>`):'<div class="empty">No collections in this date range.</div>') + '<br><h2>Revenue by rate type</h2><br>'+table(summary.revenue_by_rate_type,['Rate Type','Bills','Amount'],r=>`<tr><td>${r.rate_type}</td><td>${r.count}</td><td>RM ${r.total_amount}</td></tr>`) }}
async function importFile(file){q("#dropText").innerHTML=html(file.name)+'<small>Importing...</small>';let fd=new FormData();fd.append('file',file);fd.append('fee_rate',settings().fee_rate||'5.00');let res=await fetch('/api/import',{method:'POST',body:fd});let data=await res.json();if(!res.ok){q("#notice").textContent=data.error||'Import failed';return}stays=data.stays;sales=data.sales;summary=data.summary;localStorage.setItem('stays',JSON.stringify(stays));localStorage.setItem('sales',JSON.stringify(sales));localStorage.setItem('summary',JSON.stringify(summary));let rangeNote=showImportedRange();let message=`Imported ${summary.stays} stays. ${data.accepted_count||0} verified record(s) were saved.`+rangeNote;if(data.review_count)message+=` ${data.review_count} mismatch(es) were sent to Manual Review.`;if(data.save_error)message+=` Database warning: ${data.save_error}`;q("#notice").textContent=message;q("#dropText").innerHTML=html(file.name)+'<small>Imported</small>';render();loadReviewQueue()}
async function loadHistory(){let res=await fetch('/api/history?fee_rate='+(settings().fee_rate||'5.00'));let data=await res.json();if(!res.ok||!data.enabled||!data.stays.length)return;stays=data.stays;sales=data.sales;summary=data.summary;localStorage.setItem('stays',JSON.stringify(stays));localStorage.setItem('sales',JSON.stringify(sales));localStorage.setItem('summary',JSON.stringify(summary));let rangeNote=visibleRows().length?"":showImportedRange();q("#notice").textContent=`Loaded ${summary.stays} saved stays from Supabase.`+rangeNote;render()}
async function loadReviewQueue(){let res=await fetch('/api/review-queue'),data=await res.json();reviewQueue=res.ok&&data.enabled?(data.records||[]):[];renderReviewQueue()}
async function download(kind){let fd=new FormData();fd.append('kind',kind);fd.append('stays',JSON.stringify(stays));if(kind==='b'){let month=q("#reportMonthB").value;if(!month){q("#notice").textContent='Select a month for Lampiran B.';return}fd.append('report_month',month)}else{let start=q("#reportStartC").value,end=q("#reportEndC").value;if(!start||!end){q("#notice").textContent='Select the first and last date for Lampiran C.';return}if(start>end){q("#notice").textContent='Lampiran C first date must not be after the last date.';return}fd.append('report_start',start);fd.append('report_end',end)}Object.entries(settings()).forEach(([k,v])=>fd.append(k,v));let res=await fetch('/api/report',{method:'POST',body:fd});if(!res.ok){q("#notice").textContent=await res.text();return}let blob=await res.blob(),url=URL.createObjectURL(blob),a=document.createElement('a'),disposition=res.headers.get('Content-Disposition')||'',match=disposition.match(/filename="?([^";]+)"?/i);a.href=url;a.download=match?match[1]:(kind==='b'?'Lampiran B.pdf':'Lampiran C.pdf');a.click();URL.revokeObjectURL(url)}
qa(".nav button").forEach(b=>b.onclick=()=>{qa(".nav button").forEach(x=>x.classList.remove('active'));b.classList.add('active');qa(".view").forEach(v=>v.classList.remove('active'));q("#"+b.dataset.view).classList.add('active');q("#pageTitle").textContent=b.textContent});
q("#file").onchange=e=>e.target.files[0]&&importFile(e.target.files[0]);q("#search").oninput=render;q("#downloadB").onclick=()=>download('b');q("#downloadC").onclick=()=>download('c');
q("#refreshReview").onclick=loadReviewQueue;
qa(".filters button").forEach(b=>b.onclick=()=>{rangeMode=b.dataset.range;localStorage.setItem("rangeMode",rangeMode);render()});
q("#startDate").value=localStorage.getItem("startDate")||"";q("#endDate").value=localStorage.getItem("endDate")||"";
["#startDate","#endDate"].forEach(id=>q(id).onchange=()=>{rangeMode="custom";localStorage.setItem("rangeMode",rangeMode);localStorage.setItem(id.slice(1),q(id).value);render()});
["#reportMonthB","#reportStartC","#reportEndC"].forEach(id=>q(id).onchange=()=>{localStorage.setItem(id.slice(1),q(id).value);render()});
let drop=q("#drop");['dragenter','dragover'].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.add('drag')}));['dragleave','drop'].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.remove('drag')}));drop.addEventListener('drop',e=>{if(e.dataTransfer.files[0])importFile(e.dataTransfer.files[0])});
qa(".setting").forEach(i=>{i.value=localStorage.getItem('set_'+i.name)||i.value;i.oninput=()=>localStorage.setItem('set_'+i.name,i.value)});
render();
loadHistory();
loadReviewQueue();
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
        save_result = {"saved": False, "accepted_count": 0, "review_count": 0}
        save_error = ""
        try:
            save_result = save_import_to_supabase(upload.filename, stays, sales, summary_data)
        except Exception as exc:
            save_error = str(exc)
        response_stays = stays
        if supabase_configured():
            verified_records = load_stays_from_supabase()
            if verified_records:
                response_stays = records_to_stays(verified_records)
                sales = expanded_sales(response_stays, fee_rate)
                summary_data = summary(response_stays, fee_rate)
        return jsonify({"stays": [stay_record(s) for s in response_stays], "sales": sales, "summary": summary_data, **save_result, "save_error": save_error})
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


@app.get("/api/review-queue")
def api_review_queue():
    try:
        if not supabase_configured():
            return jsonify({"enabled": False, "records": []})
        return jsonify({"enabled": True, "records": load_review_queue_from_supabase()})
    except Exception as exc:
        return jsonify({"enabled": supabase_configured(), "error": str(exc), "records": []}), 400


@app.post("/api/report")
def api_report():
    try:
        kind = request.form.get("kind", "b")
        stays = records_to_stays(json.loads(request.form.get("stays", "[]")))
        stays = filter_stays_for_report(
            stays,
            kind,
            request.form.get("report_month", ""),
            request.form.get("report_start", ""),
            request.form.get("report_end", ""),
        )
        fee_rate = dec(request.form.get("fee_rate", "5.00"))
        settings = {k: clean(request.form.get(k, "")) for k in request.form.keys()}
        pdf = form_b_pdf(stays, settings, fee_rate) if kind == "b" else form_c_pdf(stays, settings, fee_rate)
        name = report_filename(kind, stays)
        return Response(pdf, mimetype="application/pdf", headers={"Content-Disposition": f'attachment; filename="{name}"'})
    except Exception as exc:
        return Response(str(exc), status=400)


if __name__ == "__main__":
    app.run(debug=True)
