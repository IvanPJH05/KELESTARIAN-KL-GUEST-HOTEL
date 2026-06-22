from __future__ import annotations

import io
import re
import zipfile
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP
from html import escape
from typing import Iterable

from flask import Flask, Response, request
from docx import Document
from docx.enum.section import WD_SECTION_START
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

app = Flask(__name__, static_folder="public", static_url_path="/static")
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024


@dataclass
class GuestStay:
    voucher: str
    room_no: str
    room_type: str
    rate: Decimal
    guest_name: str
    checkin_date: date
    checkin_time: str
    checkout_date: date | None
    nights: int


def money(value: Decimal | int | float) -> str:
    amount = Decimal(str(value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return f"{amount:,.2f}"


def clean_text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def parse_date(value, datemode: int | None = None) -> date | None:
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
    text = clean_text(value)
    if not text:
        return None
    for fmt in ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d", "%m/%d/%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    return None


def parse_decimal(value) -> Decimal:
    text = clean_text(value).replace(",", "")
    if not text:
        return Decimal("0")
    try:
        return Decimal(text)
    except Exception:
        return Decimal("0")


def parse_int(value, default: int = 0) -> int:
    try:
        return int(Decimal(clean_text(value)))
    except Exception:
        return default


def rows_from_excel(filename: str, data: bytes) -> Iterable[tuple[list, int | None]]:
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

    raise ValueError("Please upload an .xls or .xlsx guest check-in file.")


def parse_guest_stays(filename: str, data: bytes) -> list[GuestStay]:
    stays: list[GuestStay] = []
    for row, datemode in rows_from_excel(filename, data):
        if len(row) < 14:
            continue
        voucher = clean_text(row[0])
        checkin_date = parse_date(row[10], datemode)
        if not voucher or not checkin_date:
            continue
        nights = max(parse_int(row[13], 1), 1)
        stays.append(
            GuestStay(
                voucher=voucher,
                room_no=clean_text(row[1]),
                room_type=clean_text(row[2]),
                rate=parse_decimal(row[3]),
                guest_name=clean_text(row[5]),
                checkin_date=checkin_date,
                checkin_time=clean_text(row[11]),
                checkout_date=parse_date(row[12], datemode),
                nights=nights,
            )
        )
    if not stays:
        raise ValueError("No guest rows were found. Please check that the Excel file matches the guest check-in export format.")
    return sorted(stays, key=lambda stay: (stay.checkin_date, stay.room_no, stay.guest_name))


def set_cell_shading(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def set_cell_text(cell, text: str, bold: bool = False, size: int = 9) -> None:
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER


def add_title(doc: Document, title: str, subtitle: str = "") -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(title)
    run.bold = True
    run.font.size = Pt(14)
    run.font.color.rgb = RGBColor(31, 78, 121)
    if subtitle:
        p2 = doc.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r2 = p2.add_run(subtitle)
        r2.bold = True
        r2.font.size = Pt(11)


def add_info_table(doc: Document, fields: list[tuple[str, str]], cols: int = 2) -> None:
    table = doc.add_table(rows=0, cols=cols * 2)
    table.style = "Table Grid"
    for i in range(0, len(fields), cols):
        row = table.add_row().cells
        for block in range(cols):
            idx = i + block
            if idx >= len(fields):
                continue
            label, value = fields[idx]
            set_cell_text(row[block * 2], label, bold=True, size=8)
            set_cell_text(row[block * 2 + 1], value, size=8)
            row[block * 2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
            row[block * 2 + 1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
    doc.add_paragraph()


def add_confirmation(doc: Document, total_label: str, total_amount: Decimal) -> None:
    doc.add_paragraph()
    p = doc.add_paragraph()
    r = p.add_run("PENGESAHAN OLEH PENGUSAHA PREMIS PENGINAPAN")
    r.bold = True
    r.font.size = Pt(10)
    p = doc.add_paragraph(
        "Saya mengesahkan bahawa maklumat ini adalah BENAR dan TEPAT berdasarkan rekod kutipan SEBENAR bagi tempoh yang dilaporkan."
    )
    p.paragraph_format.space_after = Pt(6)
    doc.add_paragraph(f"{total_label} (RM): {money(total_amount)}")
    doc.add_paragraph(f"Tarikh: {datetime.now().strftime('%d/%m/%Y')}")
    doc.add_paragraph("\n\n................................................")
    doc.add_paragraph("Cop Rasmi & Tandatangan:")


def setup_document() -> Document:
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(1.4)
    section.bottom_margin = Cm(1.4)
    section.left_margin = Cm(1.2)
    section.right_margin = Cm(1.2)
    styles = doc.styles
    styles["Normal"].font.name = "Arial"
    styles["Normal"].font.size = Pt(9)
    return doc


def build_form_b(stays: list[GuestStay], form: dict[str, str], fee_rate: Decimal) -> bytes:
    doc = setup_document()
    first_day = min(stay.checkin_date for stay in stays)
    month_label = first_day.strftime("%B %Y")
    daily: dict[date, list[GuestStay]] = defaultdict(list)
    for stay in stays:
        daily[stay.checkin_date].append(stay)

    add_title(doc, "LAPORAN PENYATA KUTIPAN FI KELESTARIAN NEGERI SELANGOR", "(BULANAN)")
    add_info_table(
        doc,
        [
            ("Nama Premis Penginapan", form.get("premise_name", "")),
            ("No. Lesen Perniagaan (PBT)", form.get("license_no", "")),
            ("No. Siri Sijil", form.get("certificate_no", "")),
            ("Kod Kategori Premis", form.get("category_code", "")),
            ("Alamat Premis Penginapan", form.get("address", "")),
            ("Bulan & Tahun Pelaporan", month_label),
            ("Wakil Untuk Dihubungi", form.get("contact_name", "")),
            ("No. Telefon / Emel", form.get("contact", "")),
        ],
    )

    table = doc.add_table(rows=1, cols=4)
    table.style = "Table Grid"
    headers = ["Tarikh", "Jumlah Bilik (Unit)", "Bilangan Malam", "Jumlah Kutipan (RM)"]
    for cell, header in zip(table.rows[0].cells, headers):
        set_cell_text(cell, header, bold=True)
        set_cell_shading(cell, "D9EAF7")

    total_rooms = 0
    total_nights = 0
    total_fee = Decimal("0")
    days_in_month = sorted(daily.keys())
    for day in days_in_month:
        day_stays = daily[day]
        rooms = len(day_stays)
        nights = sum(stay.nights for stay in day_stays)
        fee = Decimal(nights) * fee_rate
        total_rooms += rooms
        total_nights += nights
        total_fee += fee
        row = table.add_row().cells
        values = [day.strftime("%d/%m/%Y"), str(rooms), str(nights), money(fee)]
        for cell, value in zip(row, values):
            set_cell_text(cell, value)

    row = table.add_row().cells
    values = ["Jumlah", str(total_rooms), str(total_nights), money(total_fee)]
    for cell, value in zip(row, values):
        set_cell_text(cell, value, bold=True)
        set_cell_shading(cell, "E2F0D9")

    add_confirmation(doc, "Jumlah Kutipan Bulanan", total_fee)
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def build_form_c(stays: list[GuestStay], form: dict[str, str], fee_rate: Decimal) -> bytes:
    doc = setup_document()
    grouped: dict[date, list[GuestStay]] = defaultdict(list)
    for stay in stays:
        grouped[stay.checkin_date].append(stay)

    for page_idx, day in enumerate(sorted(grouped.keys())):
        if page_idx:
            doc.add_section(WD_SECTION_START.NEW_PAGE)
        day_stays = grouped[day]
        day_fee = sum(Decimal(stay.nights) * fee_rate for stay in day_stays)

        add_title(doc, "LAPORAN TRANSAKSI PENGGUNAAN BILIK", "(HARIAN)")
        add_info_table(
            doc,
            [
                ("Nama Premis Penginapan", form.get("premise_name", "")),
                ("No. Rujukan Lesen (PBT)", form.get("license_no", "")),
                ("No. Siri Sijil", form.get("certificate_no", "")),
                ("Tarikh Hari Daftar Masuk", day.strftime("%d/%m/%Y")),
            ],
        )

        table = doc.add_table(rows=1, cols=6)
        table.style = "Table Grid"
        headers = ["Bil.", "Nama", "No. Bilik", "Jumlah Bilik (Unit)", "Bilangan Malam", "Jumlah Fi Kelestarian (RM)"]
        for cell, header in zip(table.rows[0].cells, headers):
            set_cell_text(cell, header, bold=True, size=8)
            set_cell_shading(cell, "D9EAF7")

        for idx, stay in enumerate(day_stays, start=1):
            row = table.add_row().cells
            values = [str(idx), stay.guest_name, stay.room_no, "1", str(stay.nights), money(Decimal(stay.nights) * fee_rate)]
            for cell, value in zip(row, values):
                set_cell_text(cell, value, size=8)
            row[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT

        row = table.add_row().cells
        values = ["", "Jumlah Keseluruhan", "", str(len(day_stays)), str(sum(stay.nights for stay in day_stays)), money(day_fee)]
        for cell, value in zip(row, values):
            set_cell_text(cell, value, bold=True, size=8)
            set_cell_shading(cell, "E2F0D9")

        add_confirmation(doc, "Jumlah Kutipan Harian", day_fee)

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def safe_filename(text: str) -> str:
    return re.sub(r"[^A-Za-z0-9_.-]+", "-", text).strip("-") or "report"


@app.get("/")
def index() -> str:
    return """<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>KL Guest Hotel Fulfilment</title>
  <link rel="stylesheet" href="/static/styles.css">
</head>
<body>
  <main class="shell">
    <aside class="sidebar">
      <div class="brand">
        <span class="brand-mark">KL</span>
        <div>
          <strong>Guest Hotel</strong>
          <span>Fi Kelestarian Fulfilment</span>
        </div>
      </div>
      <nav>
        <a class="active" href="#">Report Fulfilment</a>
        <a href="#">Guest Uploads</a>
        <a href="#">Document Archive</a>
        <a href="#">Settings</a>
      </nav>
    </aside>
    <section class="workspace">
      <header class="topbar">
        <div>
          <p class="eyebrow">Monthly compliance workflow</p>
          <h1>Generate Lampiran B & C</h1>
        </div>
        <div class="status-pill">Ready for upload</div>
      </header>

      <section class="metrics" aria-label="Workflow summary">
        <article>
          <span>Input</span>
          <strong>Guest check-in list</strong>
        </article>
        <article>
          <span>Output</span>
          <strong>2 DOCX reports</strong>
        </article>
        <article>
          <span>Default fee</span>
          <strong>RM5 / night</strong>
        </article>
      </section>

      <form class="fulfilment-board" action="/generate" method="post" enctype="multipart/form-data">
        <section class="card upload-card">
          <div>
            <p class="section-label">Step 1</p>
            <h2>Upload guest list</h2>
          </div>
          <label class="dropzone">
            <span>Choose guest check-in Excel</span>
            <small>.xls or .xlsx export from the hotel system</small>
            <input name="guest_file" type="file" accept=".xls,.xlsx" required>
          </label>
          <div class="hint-row">
            <span>Reads room, guest, check-in date, checkout date, and nights.</span>
          </div>
        </section>

        <section class="card details-card">
          <div>
            <p class="section-label">Step 2</p>
            <h2>Premise details</h2>
          </div>
          <div class="grid">
            <label>Nama Premis Penginapan<input name="premise_name" placeholder="Hotel / premis"></label>
            <label>No. Lesen Perniagaan (PBT)<input name="license_no" placeholder="MBPJ-0000"></label>
            <label>No. Siri Sijil<input name="certificate_no" placeholder="Sijil"></label>
            <label>Kod Kategori Premis<input name="category_code" placeholder="Hotel 1-3 bintang"></label>
            <label>Wakil Untuk Dihubungi<input name="contact_name" placeholder="Nama wakil"></label>
            <label>No. Telefon / Emel<input name="contact" placeholder="Telefon atau emel"></label>
            <label>Fi per bilik/malam (RM)<input name="fee_rate" type="number" step="0.01" min="0" value="5.00"></label>
          </div>
          <label>Alamat Premis Penginapan<textarea name="address" rows="3" placeholder="Alamat penuh premis"></textarea></label>
        </section>

        <section class="action-strip">
          <div>
            <strong>Fulfilment package</strong>
            <span>Creates a ZIP with Lampiran B monthly summary and Lampiran C daily transactions.</span>
          </div>
          <button type="submit">Generate reports</button>
        </section>
      </form>
    </section>
  </main>
</body>
</html>"""


@app.post("/generate")
def generate() -> Response:
    upload = request.files.get("guest_file")
    if not upload or not upload.filename:
        return Response("Please upload the guest check-in Excel file.", status=400)

    try:
        fee_rate = parse_decimal(request.form.get("fee_rate", "5.00"))
        stays = parse_guest_stays(upload.filename, upload.read())
        form = {key: escape(request.form.get(key, "").strip()) for key in request.form.keys()}
        form_b = build_form_b(stays, form, fee_rate)
        form_c = build_form_c(stays, form, fee_rate)

        first_day = min(stay.checkin_date for stay in stays)
        suffix = safe_filename(first_day.strftime("%Y-%m"))
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as bundle:
            bundle.writestr(f"Lampiran-B-Fi-Kelestarian-{suffix}.docx", form_b)
            bundle.writestr(f"Lampiran-C-Transaksi-Bilik-Harian-{suffix}.docx", form_c)
        zip_buffer.seek(0)
        return Response(
            zip_buffer.getvalue(),
            mimetype="application/zip",
            headers={"Content-Disposition": f"attachment; filename=fi-kelestarian-reports-{suffix}.zip"},
        )
    except Exception as exc:
        return Response(str(exc), status=400)


if __name__ == "__main__":
    app.run(debug=True)
