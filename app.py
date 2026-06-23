from __future__ import annotations

import io
import json
import os
import re
from calendar import month_name
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP

from flask import Flask, Response, jsonify, request
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024


@dataclass
class Stay:
    voucher: str
    room_no: str
    room_type: str
    rate: Decimal
    guest_name: str
    checkin_date: date
    checkin_time: str
    checkout_date: date | None
    nights: int
    booking_code: str
    customer_type: str


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


def clean_guest_name(value) -> str:
    """Remove OTA reference numbers printed before guest names."""
    name = clean(value)
    return re.sub(r"^[\s'\"]*(?:\d+\s*[-:/]?\s*)+", "", name).strip()


def customer_type(booking_code, guest_name="") -> str:
    """Use the hotel's source code first; name shape is only a legacy fallback."""
    code = clean(booking_code).upper()
    if code == "WLKIN":
        return "Walk-in"
    if code == "XPDIA":
        return "Online Booking"
    return "Online Booking" if re.match(r"^\s*\d+", clean(guest_name)) else "Walk-in"


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


def parse_stays(filename: str, data: bytes) -> list[Stay]:
    stays = []
    for row, datemode in excel_rows(filename, data):
        if len(row) < 14:
            continue
        checkin = date_value(row[10], datemode)
        checkout = date_value(row[12], datemode)
        voucher = clean(row[0])
        if not voucher or not checkin:
            continue
        exported_nights = max(num(row[13], 1), 1)
        calendar_nights = (checkout - checkin).days if checkout and checkout > checkin else 0
        nights = calendar_nights or exported_nights
        booking_code = clean(row[4])
        original_name = clean(row[5])
        stays.append(
            Stay(
                voucher=voucher,
                room_no=clean(row[1]),
                room_type=clean(row[2]),
                rate=dec(row[3]),
                guest_name=clean_guest_name(original_name),
                checkin_date=checkin,
                checkin_time=clean(row[11]),
                checkout_date=checkout,
                nights=nights,
                booking_code=booking_code,
                customer_type=customer_type(booking_code, original_name),
            )
        )
    if not stays:
        raise ValueError("No guest rows were found. Check that the file is the hotel check-in export.")
    return sorted(stays, key=lambda s: (s.checkin_date, s.room_no, s.guest_name))


def stay_record(stay: Stay) -> dict:
    return {
        "voucher": stay.voucher,
        "room_no": stay.room_no,
        "room_type": stay.room_type,
        "rate": str(stay.rate),
        "guest_name": stay.guest_name,
        "checkin_date": stay.checkin_date.isoformat(),
        "checkin_time": stay.checkin_time,
        "checkout_date": stay.checkout_date.isoformat() if stay.checkout_date else "",
        "nights": stay.nights,
        "booking_code": stay.booking_code,
        "customer_type": stay.customer_type,
    }


def records_to_stays(records: list[dict]) -> list[Stay]:
    stays = []
    for item in records:
        checkin = date_value(item.get("checkin_date"))
        if not checkin:
            continue
        original_name = clean(item.get("guest_name"))
        booking_code = clean(item.get("booking_code"))
        imported_type = clean(item.get("customer_type"))
        stays.append(
            Stay(
                voucher=clean(item.get("voucher")),
                room_no=clean(item.get("room_no")),
                room_type=clean(item.get("room_type")),
                rate=dec(item.get("rate")),
                guest_name=clean_guest_name(original_name),
                checkin_date=checkin,
                checkin_time=clean(item.get("checkin_time")),
                checkout_date=date_value(item.get("checkout_date")),
                nights=max(num(item.get("nights"), 1), 1),
                booking_code=booking_code,
                customer_type=imported_type or customer_type(booking_code, original_name),
            )
        )
    if not stays:
        raise ValueError("No imported stays found. Please import the Excel file again.")
    return sorted(stays, key=lambda s: (s.checkin_date, s.room_no, s.guest_name))


def expanded_sales(stays: list[Stay], fee_rate: Decimal) -> list[dict]:
    rows = []
    for stay in stays:
        for offset in range(stay.nights):
            day = stay.checkin_date + timedelta(days=offset)
            payment_collected = offset == 0
            full_fee = fee_rate * stay.nights
            row_state = "single" if stay.nights == 1 else "payment" if payment_collected else "paid"
            rows.append(
                {
                    "date": day.isoformat(),
                    "display_date": day.strftime("%d/%m/%Y"),
                    "room_no": stay.room_no,
                    "guest_name": stay.guest_name,
                    "customer_type": stay.customer_type,
                    "booking_code": stay.booking_code,
                    "stay_progress": f"{offset + 1}/{stay.nights}",
                    "nights": stay.nights,
                    "multi_night": stay.nights > 1,
                    "payment_collected": payment_collected,
                    "row_state": row_state,
                    "row_class": "" if row_state == "single" else f"{row_state}-row",
                    "price": money(stay.rate) if payment_collected else "PAID",
                    "kelestarian": money(full_fee) if payment_collected else "PAID",
                    "actual_price": money(stay.rate) if payment_collected else "0.00",
                    "actual_kelestarian": money(full_fee) if payment_collected else "0.00",
                    "voucher": stay.voucher,
                }
            )
    return sorted(rows, key=lambda r: (r["date"], r["room_no"], r["guest_name"]))


def summary(stays: list[Stay], fee_rate: Decimal) -> dict:
    sales = expanded_sales(stays, fee_rate)
    total_fee = sum((Decimal(row["actual_kelestarian"].replace(",", "")) for row in sales), Decimal("0"))
    total_sales = sum((s.rate for s in stays), Decimal("0"))
    days = sorted({row["date"] for row in sales})
    return {
        "stays": len(stays),
        "sale_rows": len(sales),
        "days": len(days),
        "first_date": days[0] if days else "",
        "last_date": days[-1] if days else "",
        "total_sales": money(total_sales),
        "total_kelestarian": money(total_fee),
    }


def report_filename(kind: str, stays: list[Stay]) -> str:
    checkin_dates = sorted({stay.checkin_date for stay in stays})
    if kind == "b":
        months = sorted({(day.year, day.month) for day in checkin_dates})
        labels = [f"{month_name[month]}-{year}" for year, month in months]
        period = labels[0] if len(labels) == 1 else f"{labels[0]}-to-{labels[-1]}"
        return f"Lampiran-B-{period}.pdf"
    labels = [day.strftime("%d-%m-%Y") for day in checkin_dates]
    period = labels[0] if len(labels) == 1 else f"{labels[0]}-to-{labels[-1]}"
    return f"Lampiran-C-{period}.pdf"


CREST_PATH = os.path.join(os.path.dirname(__file__), "assets", "selangor-crest.jpeg")
GREY = colors.HexColor("#D9D9D9")


def _text(c, x, y, value, size=9, bold=False, align="left"):
    value = clean(value)
    font = "Helvetica-Bold" if bold else "Helvetica"
    c.setFillColor(colors.black)
    c.setFont(font, size)
    if align == "center":
        c.drawCentredString(x, y, value)
    elif align == "right":
        c.drawRightString(x, y, value)
    else:
        c.drawString(x, y, value)


def _fit(value, width, size=7, bold=False):
    value = clean(value)
    font = "Helvetica-Bold" if bold else "Helvetica"
    if stringWidth(value, font, size) <= width:
        return value
    while value and stringWidth(value + "...", font, size) > width:
        value = value[:-1]
    return value.rstrip() + "..."


def _outer(c, label, x=48, right=547, bottom=38, top=803):
    c.setStrokeColor(colors.black)
    c.setLineWidth(0.55)
    c.rect(x, bottom, right - x, top - bottom, stroke=1, fill=0)
    _text(c, right - 20, top + 4, label, 12, True, "right")


def _crest(c, y=705):
    if os.path.exists(CREST_PATH):
        c.drawImage(ImageReader(CREST_PATH), 273, y, 49, 68, preserveAspectRatio=True, mask="auto")


def _field_rows(c, labels, values, x, top, label_width, value_width, row_height, font_size=9):
    for index, (label, value) in enumerate(zip(labels, values)):
        y = top - (index + 1) * row_height
        _text(c, x, y + 6, label, font_size)
        _text(c, x + label_width - 8, y + 6, ":", font_size)
        c.rect(x + label_width, y, value_width, row_height, stroke=1, fill=0)
        _text(c, x + label_width + 5, y + 6, _fit(value, value_width - 10, font_size), font_size)


def _grid(c, x, top, widths, row_heights, fills=None):
    y = top
    total_width = sum(widths)
    for row_no, height in enumerate(row_heights):
        y -= height
        if fills and fills[row_no]:
            c.setFillColor(fills[row_no])
            c.rect(x, y, total_width, height, stroke=0, fill=1)
        c.setStrokeColor(colors.black)
        c.rect(x, y, total_width, height, stroke=1, fill=0)
        cursor = x
        for width in widths[:-1]:
            cursor += width
            c.line(cursor, y, cursor, y + height)
    return y


def _cell_text(c, x, y, width, height, value, size=7, bold=False, align="center"):
    value = _fit(value, width - 6, size, bold)
    baseline = y + (height - size) / 2 + 1
    if align == "left":
        _text(c, x + 3, baseline, value, size, bold)
    else:
        _text(c, x + width / 2, baseline, value, size, bold, "center")


def _b_table(c, x, top, rows, include_header=False, include_total=False):
    widths = [87, 142, 126, 144]
    rendered = []
    fills = []
    if include_header:
        rendered.append(["Tarikh", "Jumlah Bilik (Unit)", "Bilangan Malam", "Jumlah Kutipan (RM)"])
        fills.append(GREY)
    rendered.extend(rows)
    fills.extend([None] * len(rows))
    if include_total:
        rendered.append(include_total)
        fills.append(GREY)
    heights = [19] * len(rendered)
    bottom = _grid(c, x, top, widths, heights, fills)
    y = top
    for row_no, row in enumerate(rendered):
        y -= heights[row_no]
        cursor = x
        for col_no, value in enumerate(row):
            _cell_text(c, cursor, y, widths[col_no], heights[row_no], value, 8, row_no == 0 and include_header or fills[row_no] == GREY)
            cursor += widths[col_no]
    return bottom


def _b_confirmation(c, total_fee):
    _text(c, 54, 494, "C. PENGESAHAN OLEH PENGUSAHA PREMIS PENGINAPAN", 11, True)
    _text(c, 54, 472, "Saya mengesahkan bahawa maklumat ini adalah BENAR dan TEPAT berdasarkan rekod", 8.8)
    _text(c, 54, 458, "kutipan SEBENAR bagi bulan yang dilaporkan.", 8.8)
    _text(c, 54, 430, f"Jumlah Kutipan Bulanan (RM) : {money(total_fee)}", 9.5, True)
    _text(c, 54, 414, f"Tarikh : {datetime.now().strftime('%d/%m/%Y')}", 9.5, True)
    _text(c, 54, 350, "................................................", 9, True)
    _text(c, 54, 336, "Cop Rasmi & Tandatangan:", 9, True)
    c.line(48, 310, 547, 310)
    _text(c, 54, 288, "D. UNTUK KEGUNAAN PEJABAT PBT SAHAJA", 11, True)
    _text(c, 54, 260, "(SEMAKAN & PENGESAHAN PBT)", 10.5, True)
    pbt = ["Tarikh Diterima", "Ulasan / Catatan", "Jumlah Kutipan Bulanan (RM)", "Caj lewat (RM) (Sekiranya ada)"]
    for i, label in enumerate(pbt):
        _text(c, 60, 230 - i * 18, label, 9)
        _text(c, 238, 230 - i * 18, ": ................................................", 9)
    _text(c, 54, 104, "................................................", 9, True)
    _text(c, 54, 90, "Cop Rasmi & Tandatangan:", 9, True)


def form_b_pdf(stays: list[Stay], settings: dict, fee_rate: Decimal) -> bytes:
    sales = [row for row in expanded_sales(stays, fee_rate) if row["payment_collected"]]
    months = defaultdict(list)
    for row in sales:
        day = date.fromisoformat(row["date"])
        months[(day.year, day.month)].append(row)
    out = io.BytesIO()
    c = canvas.Canvas(out, pagesize=A4)
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
        _outer(c, "LAMPIRAN B")
        _crest(c)
        _text(c, A4[0] / 2, 674, "LAPORAN PENYATA KUTIPAN FI KELESTARIAN NEGERI SELANGOR (BULANAN)", 10.5, True, "center")
        _text(c, 54, 642, "A. MAKLUMAT PREMIS PENGINAPAN", 11, True)
        _field_rows(c, labels, values, 60, 616, 225, 250, 20, 8.5)
        _text(c, 54, 430, "B. MAKLUMAT KUTIPAN FI KELESTARIAN", 11, True)
        table_rows = []
        for day_no in range(1, 32):
            collected = daily.get(day_no, [])
            rooms = len(collected)
            nights = sum(item["nights"] for item in collected)
            fee = sum((Decimal(item["actual_kelestarian"].replace(",", "")) for item in collected), Decimal("0"))
            table_rows.append([f"{day_no}/{month:02d}/{year}", str(rooms) if rooms else "", str(nights) if nights else "", money(fee) if fee else ""])
        _b_table(c, 48, 408, table_rows[:18], include_header=True)
        c.showPage()
        _outer(c, "LAMPIRAN B")
        total_rooms = len(month_rows)
        total_nights = sum(row["nights"] for row in month_rows)
        total_fee = sum((Decimal(row["actual_kelestarian"].replace(",", "")) for row in month_rows), Decimal("0"))
        _b_table(c, 48, 803, table_rows[18:], include_total=["Jumlah", str(total_rooms), str(total_nights), money(total_fee)])
        _b_confirmation(c, total_fee)
    c.save()
    return out.getvalue()


C_WIDTHS = [30, 275, 85, 75, 100]
C_HEADERS = ["Bil.", "Nama", "Jumlah Bilik (Unit)", "Bilangan Malam", "Jumlah Fi Kelestarian (RM)"]


def _c_header_cells(c, x, y, height=44):
    cursor = x
    for width, label in zip(C_WIDTHS, C_HEADERS):
        words = label.split()
        lines = []
        current = ""
        for word in words:
            candidate = (current + " " + word).strip()
            if stringWidth(candidate, "Helvetica-Bold", 7) > width - 5 and current:
                lines.append(current)
                current = word
            else:
                current = candidate
        if current:
            lines.append(current)
        start = y + height / 2 + (len(lines) - 1) * 4 - 2
        for line_no, line in enumerate(lines):
            _text(c, cursor + width / 2, start - line_no * 8, line, 7, True, "center")
        cursor += width


def _c_rows(c, x, top, items, capacity, start_number, include_header=False):
    header_height = 44 if include_header else 0
    row_height = 18
    fills = ([GREY] if include_header else []) + [None for item in items]
    padded = items + [None] * (capacity - len(items))
    fills += [None] * (capacity - len(items))
    heights = ([header_height] if include_header else []) + [row_height] * capacity
    _grid(c, x, top, C_WIDTHS, heights, fills)
    y = top
    if include_header:
        y -= header_height
        _c_header_cells(c, x, y, header_height)
    for local_no, item in enumerate(padded):
        y -= row_height
        if not item:
            continue
        values = [str(start_number + local_no), item["guest_name"], "1", str(item["nights"]), item["kelestarian"]]
        cursor = x
        for col_no, value in enumerate(values):
            _cell_text(c, cursor, y, C_WIDTHS[col_no], row_height, value, 6.4 if col_no == 1 else 7, False, "left" if col_no == 1 else "center")
            cursor += C_WIDTHS[col_no]
    return y


def _c_first_page(c, settings, day_label, items):
    _outer(c, "LAMPIRAN C", x=15, right=580)
    _crest(c, 708)
    _text(c, A4[0] / 2, 686, "LAPORAN TRANSAKSI PENGGUNAAN BILIK (HARIAN)", 11.5, True, "center")
    labels = ["Nama Premis Penginapan", "No. Rujukan Lesen (PBT)", "No. Siri Sijil", "Tarikh Hari Daftar Masuk (Check-In)"]
    values = [settings.get("premise_name", ""), settings.get("license_no", ""), settings.get("certificate_no", ""), day_label]
    _field_rows(c, labels, values, 40, 670, 245, 280, 29, 9)
    _c_rows(c, 15, 545, items, 24, 1, include_header=True)


def _c_confirmation(c, x, top, rows_for_day):
    total_rooms = len(rows_for_day)
    total_nights = sum(row["nights"] for row in rows_for_day)
    total_fee = sum((Decimal(row["actual_kelestarian"].replace(",", "")) for row in rows_for_day), Decimal("0"))
    height = 26
    c.setFillColor(GREY)
    c.rect(x, top - height, sum(C_WIDTHS), height, stroke=0, fill=1)
    c.setStrokeColor(colors.black)
    c.rect(x, top - height, sum(C_WIDTHS), height, stroke=1, fill=0)
    cursor = x + sum(C_WIDTHS[:2])
    for width in C_WIDTHS[2:-1]:
        c.line(cursor, top - height, cursor, top)
        cursor += width
    c.line(cursor, top - height, cursor, top)
    _text(c, x + sum(C_WIDTHS[:2]) / 2, top - 17, "Jumlah Keseluruhan (Unit @ RM)", 8.5, True, "center")
    totals = [str(total_rooms), str(total_nights), money(total_fee)]
    cursor = x + sum(C_WIDTHS[:2])
    for width, value in zip(C_WIDTHS[2:], totals):
        _cell_text(c, cursor, top - height, width, height, value, 7, True)
        cursor += width
    box_top = top - height - 10
    c.line(15, box_top, 580, box_top)
    _text(c, 22, box_top - 32, "PENGESAHAN OLEH PENGUSAHA PREMIS PENGINAPAN", 11, True)
    _text(c, 22, box_top - 54, "Saya mengesahkan bahawa maklumat ini adalah BENAR dan TEPAT berdasarkan rekod kutipan", 8.5)
    _text(c, 22, box_top - 69, "SEBENAR bagi hari yang dilaporkan.", 8.5, True)
    _text(c, 22, box_top - 98, f"Jumlah Kutipan Harian (RM) : {money(total_fee)}", 9.5, True)
    _text(c, 22, box_top - 114, f"Tarikh : {datetime.now().strftime('%d/%m/%Y')}", 9.5, True)
    _text(c, 22, box_top - 180, "................................................", 9, True)
    _text(c, 22, box_top - 195, "Cop Rasmi & Tandatangan:", 9, True)


def form_c_pdf(stays: list[Stay], settings: dict, fee_rate: Decimal) -> bytes:
    grouped = defaultdict(list)
    for row in expanded_sales(stays, fee_rate):
        if row["payment_collected"]:
            grouped[row["date"]].append(row)
    out = io.BytesIO()
    c = canvas.Canvas(out, pagesize=A4)
    first_report = True
    for day_key in sorted(grouped):
        if not first_report:
            c.showPage()
        first_report = False
        rows_for_day = grouped[day_key]
        day_label = date.fromisoformat(day_key).strftime("%d/%m/%Y")
        _c_first_page(c, settings, day_label, rows_for_day[:24])
        remaining = rows_for_day[24:]
        row_number = 25
        while len(remaining) > 19:
            c.showPage()
            _outer(c, "LAMPIRAN C", x=15, right=580)
            chunk = remaining[: min(38, len(remaining) - 19)]
            _c_rows(c, 15, 790, chunk, 38, row_number, include_header=True)
            row_number += len(chunk)
            remaining = remaining[len(chunk):]
        c.showPage()
        _outer(c, "LAMPIRAN C", x=15, right=580)
        bottom = _c_rows(c, 15, 803, remaining, 19, row_number, include_header=False)
        _c_confirmation(c, 15, bottom, rows_for_day)
    c.save()
    return out.getvalue()


PAGE = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>KL Guest Hotel Sales</title>
<style>
:root{font-family:Inter,ui-sans-serif,system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif;background:#f2f6fb;color:#071a36}*{box-sizing:border-box}body{margin:0}.app{min-height:100vh;display:grid;grid-template-columns:278px 1fr}.sidebar{background:#12233b;color:#fff;padding:28px 18px;display:flex;flex-direction:column}.brand{display:flex;gap:14px;align-items:center;margin-bottom:30px}.logo{width:48px;height:48px;border-radius:9px;background:#fff;color:#0d3265;display:grid;place-items:center;font-weight:900}.brand span{display:block;color:#b5cae8;font-size:13px}.nav-label{font-size:12px;letter-spacing:.12em;text-transform:uppercase;color:#8aa6c8;font-weight:900}.nav{display:grid;gap:8px;margin-top:12px}.nav button{background:transparent;color:#b8d0ee;text-align:left;border:0;padding:14px;border-radius:8px;font-size:17px;cursor:pointer}.nav button.active{background:#1d3a60;color:#fff}.account{border-top:1px solid #314966;margin-top:auto;padding-top:28px;display:flex;gap:12px;align-items:center}.avatar{width:42px;height:42px;border-radius:999px;background:#3478d4;display:grid;place-items:center;font-weight:900}.main{padding:28px 34px 42px}.top{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:22px}.eyebrow{margin:0 0 8px;color:#7587a0;text-transform:uppercase;letter-spacing:.14em;font-size:13px;font-weight:900}h1{font-size:32px;margin:0}h2{font-size:20px;margin:0}.badge{background:#e8f1fd;color:#1c5a9d;border-radius:999px;padding:9px 16px;font-weight:900;font-size:13px}.notice{background:#e8f2ff;border:1px solid #bdd6fb;color:#0b4f9a;border-radius:9px;padding:16px 18px;margin-bottom:18px}.cards{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:18px}.metric,.panel{background:#fff;border:1px solid #d7e0eb;border-radius:10px;box-shadow:0 8px 24px rgba(10,31,68,.04)}.metric{padding:16px}.metric span{display:block;color:#71839c;font-size:13px}.metric strong{display:block;margin-top:4px;font-size:22px}.panel{padding:22px;margin-bottom:18px}.drop{height:118px;border:1.5px dashed #9cb9dc;border-radius:9px;background:#f8fbff;display:grid;place-items:center;text-align:center;color:#1b5fab;font-weight:900;cursor:pointer}.drop.drag{background:#e8f2ff;border-color:#1b5fab}.drop small{display:block;color:#71839c;font-weight:600;margin-top:5px}.drop input{display:none}.toolbar{display:flex;justify-content:space-between;align-items:center;gap:12px;margin:14px 0}.toolbar input{max-width:320px}.settings-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:14px}label{display:grid;gap:7px;font-size:13px;font-weight:800;color:#2f4360}input{width:100%;border:1px solid #c9d5e3;border-radius:7px;padding:11px 12px;font:inherit}button.primary{border:0;border-radius:8px;background:#10233e;color:#fff;font-weight:900;padding:12px 18px;cursor:pointer}button.primary:disabled{opacity:.45;cursor:not-allowed}.view{display:none}.view.active{display:block}.table-wrap{overflow:auto;border:1px solid #dde6f1;border-radius:8px;background:#fff}table{width:100%;border-collapse:collapse;font-size:13px}th,td{padding:10px 12px;border-bottom:1px solid #e7edf5;text-align:left;white-space:nowrap}th{background:#f5f8fc;color:#516985;font-size:12px;text-transform:uppercase;letter-spacing:.05em}.day-row td{background:#edf4fc;color:#153a63;font-weight:900;font-size:14px}.payment-row td{background:#fff2cc}.paid-row td{background:#ddebf7}.legend{display:flex;gap:16px;align-items:center;color:#516985;font-size:12px}.legend i{display:inline-block;width:14px;height:14px;border:1px solid #9aa9b8;vertical-align:-2px;margin-right:5px}.legend .yellow{background:#fff2cc}.legend .blue{background:#ddebf7}.empty{color:#71839c;padding:22px;border:1px dashed #c9d8ea;border-radius:8px;background:#f9fbfe}.report-grid{display:grid;grid-template-columns:1fr;gap:18px}.wide{grid-column:span 2}@media(max-width:1050px){.app{grid-template-columns:1fr}.sidebar{display:none}.cards,.settings-grid{grid-template-columns:1fr}.main{padding:20px}.top{flex-direction:column;gap:12px}}
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
    <section class="notice" id="notice">Drag in the hotel Excel file. The dashboard will remember every check-in and expand multi-night stays by day.</section>
    <section class="cards">
      <article class="metric"><span>Imported stays</span><strong id="mStays">0</strong></article>
      <article class="metric"><span>Sales rows</span><strong id="mRows">0</strong></article>
      <article class="metric"><span>Total sales</span><strong id="mSales">RM 0.00</strong></article>
      <article class="metric"><span>Kelestarian</span><strong id="mFee">RM 0.00</strong></article>
    </section>
    <section class="panel">
      <label class="drop" id="drop"><span id="dropText">Choose or drop guest check-in Excel<small>.xls or .xlsx</small></span><input id="file" type="file" accept=".xls,.xlsx"></label>
    </section>
    <section id="dashboard" class="view active">
      <div class="panel"><div class="toolbar"><div><h2>Daily sales ledger</h2><div class="legend"><span><i class="yellow"></i>Payment collected</span><span><i class="blue"></i>Already paid</span></div></div><input id="search" placeholder="Search guest, type or room"></div><div id="ledger" class="empty">No imported check-ins yet.</div></div>
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
const SCHEMA_VERSION='2';
if(localStorage.getItem('reportSchema')!==SCHEMA_VERSION){localStorage.removeItem('stays');localStorage.removeItem('sales');localStorage.removeItem('summary');localStorage.setItem('reportSchema',SCHEMA_VERSION)}
let stays=JSON.parse(localStorage.getItem("stays")||"[]");
let sales=JSON.parse(localStorage.getItem("sales")||"[]");
let summary=JSON.parse(localStorage.getItem("summary")||"{}");
const q=s=>document.querySelector(s), qa=s=>[...document.querySelectorAll(s)];
function settings(){let o={};qa(".setting").forEach(i=>o[i.name]=i.value);return o}
function amount(v){return v==='PAID'?'PAID':'RM '+v}
function groupRows(rows){let html='<div class="table-wrap"><table><thead><tr><th>Room</th><th>Guest</th><th>Customer Type</th><th>Stay</th><th>Price Paid</th><th>Kelestarian</th></tr></thead><tbody>';let cur='';rows.forEach(r=>{if(r.display_date!==cur){cur=r.display_date;html+=`<tr class="day-row"><td colspan="6">${cur}</td></tr>`}let rowClass=r.row_class??(r.multi_night?(r.payment_collected?'payment-row':'paid-row'):'');html+=`<tr class="${rowClass}"><td>${r.room_no}</td><td>${r.guest_name}</td><td>${r.customer_type}</td><td>${r.stay_progress}</td><td>${amount(r.price)}</td><td>${amount(r.kelestarian)}</td></tr>`});return html+'</tbody></table></div>'}
function table(rows,heads,mapper){return '<div class="table-wrap"><table><thead><tr>'+heads.map(h=>`<th>${h}</th>`).join('')+'</tr></thead><tbody>'+rows.map(mapper).join('')+'</tbody></table></div>'}
function render(){q("#mStays").textContent=summary.stays||0;q("#mRows").textContent=summary.sale_rows||0;q("#mSales").textContent='RM '+(summary.total_sales||'0.00');q("#mFee").textContent='RM '+(summary.total_kelestarian||'0.00');q("#downloadB").disabled=!stays.length;q("#downloadC").disabled=!stays.length;let term=q("#search").value.toLowerCase();let filtered=sales.filter(r=>(r.guest_name+' '+r.customer_type+' '+r.room_no).toLowerCase().includes(term));q("#ledger").className=filtered.length?'':'empty';q("#ledger").innerHTML=filtered.length?groupRows(filtered):'No imported check-ins yet.';let byDay={};sales.filter(r=>r.payment_collected).forEach(r=>{byDay[r.display_date]??={rooms:0,nights:0,fee:0};byDay[r.display_date].rooms++;byDay[r.display_date].nights+=Number(r.nights);byDay[r.display_date].fee+=Number(r.actual_kelestarian.replace(/,/g,''))});let bRows=Object.entries(byDay).map(([d,v])=>({d,...v}));q("#previewB").className=bRows.length?'':'empty';q("#previewB").innerHTML=bRows.length?table(bRows,['Tarikh','Jumlah Bilik','Bilangan Malam','Jumlah Kutipan'],r=>`<tr><td>${r.d}</td><td>${r.rooms}</td><td>${r.nights}</td><td>RM ${r.fee.toFixed(2)}</td></tr>`):'Import Excel first.';q("#previewC").className=sales.length?'':'empty';q("#previewC").innerHTML=sales.length?groupRows(sales):'Import Excel first.'}
async function importFile(file){q("#dropText").innerHTML=file.name+'<small>Importing...</small>';let fd=new FormData();fd.append('file',file);let res=await fetch('/api/import',{method:'POST',body:fd});let data=await res.json();if(!res.ok){q("#notice").textContent=data.error||'Import failed';return}stays=data.stays;sales=data.sales;summary=data.summary;localStorage.setItem('stays',JSON.stringify(stays));localStorage.setItem('sales',JSON.stringify(sales));localStorage.setItem('summary',JSON.stringify(summary));q("#notice").textContent=`Imported ${summary.stays} stays and expanded them into ${summary.sale_rows} daily sales rows.`;q("#dropText").innerHTML=file.name+'<small>Imported</small>';render()}
async function download(kind){let fd=new FormData();fd.append('kind',kind);fd.append('stays',JSON.stringify(stays));Object.entries(settings()).forEach(([k,v])=>fd.append(k,v));let res=await fetch('/api/report',{method:'POST',body:fd});if(!res.ok){q("#notice").textContent=await res.text();return}let blob=await res.blob(),url=URL.createObjectURL(blob),a=document.createElement('a'),disposition=res.headers.get('Content-Disposition')||'',match=disposition.match(/filename="?([^";]+)"?/i);a.href=url;a.download=match?match[1]:(kind==='b'?'Lampiran-B.pdf':'Lampiran-C.pdf');a.click();URL.revokeObjectURL(url)}
qa(".nav button").forEach(b=>b.onclick=()=>{qa(".nav button").forEach(x=>x.classList.remove('active'));b.classList.add('active');qa(".view").forEach(v=>v.classList.remove('active'));q("#"+b.dataset.view).classList.add('active');q("#pageTitle").textContent=b.textContent});
q("#file").onchange=e=>e.target.files[0]&&importFile(e.target.files[0]);q("#search").oninput=render;q("#downloadB").onclick=()=>download('b');q("#downloadC").onclick=()=>download('c');
let drop=q("#drop");['dragenter','dragover'].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.add('drag')}));['dragleave','drop'].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.remove('drag')}));drop.addEventListener('drop',e=>{if(e.dataTransfer.files[0])importFile(e.dataTransfer.files[0])});
qa(".setting").forEach(i=>{i.value=localStorage.getItem('set_'+i.name)||i.value;i.oninput=()=>localStorage.setItem('set_'+i.name,i.value)});
render();
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
        fee_rate = Decimal("5.00")
        stays = parse_stays(upload.filename, upload.read())
        return jsonify({"stays": [stay_record(s) for s in stays], "sales": expanded_sales(stays, fee_rate), "summary": summary(stays, fee_rate)})
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400


@app.post("/api/report")
def api_report():
    try:
        kind = request.form.get("kind", "b")
        stays = records_to_stays(json.loads(request.form.get("stays", "[]")))
        fee_rate = dec(request.form.get("fee_rate", "5.00"))
        settings = {k: clean(request.form.get(k, "")) for k in request.form.keys()}
        pdf = form_b_pdf(stays, settings, fee_rate) if kind == "b" else form_c_pdf(stays, settings, fee_rate)
        name = report_filename(kind, stays)
        return Response(pdf, mimetype="application/pdf", headers={"Content-Disposition": f"attachment; filename={name}"})
    except Exception as exc:
        return Response(str(exc), status=400)


if __name__ == "__main__":
    app.run(debug=True)
