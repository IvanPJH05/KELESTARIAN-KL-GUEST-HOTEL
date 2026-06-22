from __future__ import annotations

import io
import json
import re
import zipfile
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP
from html import escape

from flask import Flask, Response, jsonify, request
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

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
        voucher = clean(row[0])
        if not voucher or not checkin:
            continue
        stays.append(
            Stay(
                voucher=voucher,
                room_no=clean(row[1]),
                room_type=clean(row[2]),
                rate=dec(row[3]),
                guest_name=clean(row[5]),
                checkin_date=checkin,
                checkin_time=clean(row[11]),
                checkout_date=date_value(row[12], datemode),
                nights=max(num(row[13], 1), 1),
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
    }


def records_to_stays(records: list[dict]) -> list[Stay]:
    stays = []
    for item in records:
        checkin = date_value(item.get("checkin_date"))
        if not checkin:
            continue
        stays.append(
            Stay(
                voucher=clean(item.get("voucher")),
                room_no=clean(item.get("room_no")),
                room_type=clean(item.get("room_type")),
                rate=dec(item.get("rate")),
                guest_name=clean(item.get("guest_name")),
                checkin_date=checkin,
                checkin_time=clean(item.get("checkin_time")),
                checkout_date=date_value(item.get("checkout_date")),
                nights=max(num(item.get("nights"), 1), 1),
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
            rows.append(
                {
                    "date": day.isoformat(),
                    "display_date": day.strftime("%d/%m/%Y"),
                    "room_no": stay.room_no,
                    "guest_name": stay.guest_name,
                    "stay_progress": f"{offset + 1}/{stay.nights}",
                    "nights": stay.nights,
                    "multi_night": stay.nights > 1,
                    "price": money(stay.rate),
                    "kelestarian": money(fee_rate),
                    "voucher": stay.voucher,
                }
            )
    return sorted(rows, key=lambda r: (r["date"], r["room_no"], r["guest_name"]))


def summary(stays: list[Stay], fee_rate: Decimal) -> dict:
    sales = expanded_sales(stays, fee_rate)
    total_fee = Decimal(len(sales)) * fee_rate
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
    first_day = min(stay.checkin_date for stay in stays)
    grouped = defaultdict(list)
    for row in expanded_sales(stays, fee_rate):
        grouped[row["display_date"]].append(row)
    story = [p("LAPORAN PENYATA KUTIPAN FI KELESTARIAN NEGERI SELANGOR", s["TitleBlue"]), p("(BULANAN)", s["SubTitle"]), info_table(form_fields(settings, month_label=first_day.strftime("%B %Y")), s), Spacer(1, 12)]
    rows = [["Tarikh", "Jumlah Bilik (Unit)", "Bilangan Malam", "Jumlah Kutipan (RM)"]]
    total_rooms = total_nights = 0
    total_fee = Decimal("0")
    for day in sorted(grouped.keys(), key=lambda d: datetime.strptime(d, "%d/%m/%Y")):
        rooms = len(grouped[day])
        fee = Decimal(rooms) * fee_rate
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
        day_fee = Decimal(len(rows_for_day)) * fee_rate
        story += [p("LAPORAN TRANSAKSI PENGGUNAAN BILIK", s["TitleBlue"]), p("(HARIAN)", s["SubTitle"]), info_table(form_fields(settings, date_label=day), s), Spacer(1, 10)]
        rows = [["Bil.", "Nama", "No. Bilik", "Tinggal", "Harga Dibayar (RM)", "Fi Kelestarian (RM)"]]
        for row_no, item in enumerate(rows_for_day, 1):
            rows.append([str(row_no), p(item["guest_name"], s["Small"]), item["room_no"], item["stay_progress"], item["price"], item["kelestarian"]])
        rows.append(["", "Jumlah Keseluruhan", "", "", "", money(day_fee)])
        table = Table(rows, repeatRows=1, colWidths=[1.0 * cm, 7.0 * cm, 2.0 * cm, 2.2 * cm, 3.0 * cm, 3.0 * cm])
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
:root{font-family:Inter,ui-sans-serif,system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif;background:#f2f6fb;color:#071a36}*{box-sizing:border-box}body{margin:0}.app{min-height:100vh;display:grid;grid-template-columns:278px 1fr}.sidebar{background:#12233b;color:#fff;padding:28px 18px;display:flex;flex-direction:column}.brand{display:flex;gap:14px;align-items:center;margin-bottom:30px}.logo{width:48px;height:48px;border-radius:9px;background:#fff;color:#0d3265;display:grid;place-items:center;font-weight:900}.brand span{display:block;color:#b5cae8;font-size:13px}.nav-label{font-size:12px;letter-spacing:.12em;text-transform:uppercase;color:#8aa6c8;font-weight:900}.nav{display:grid;gap:8px;margin-top:12px}.nav button{background:transparent;color:#b8d0ee;text-align:left;border:0;padding:14px;border-radius:8px;font-size:17px;cursor:pointer}.nav button.active{background:#1d3a60;color:#fff}.account{border-top:1px solid #314966;margin-top:auto;padding-top:28px;display:flex;gap:12px;align-items:center}.avatar{width:42px;height:42px;border-radius:999px;background:#3478d4;display:grid;place-items:center;font-weight:900}.main{padding:28px 34px 42px}.top{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:22px}.eyebrow{margin:0 0 8px;color:#7587a0;text-transform:uppercase;letter-spacing:.14em;font-size:13px;font-weight:900}h1{font-size:32px;margin:0}h2{font-size:20px;margin:0}.badge{background:#e8f1fd;color:#1c5a9d;border-radius:999px;padding:9px 16px;font-weight:900;font-size:13px}.notice{background:#e8f2ff;border:1px solid #bdd6fb;color:#0b4f9a;border-radius:9px;padding:16px 18px;margin-bottom:18px}.cards{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:18px}.metric,.panel{background:#fff;border:1px solid #d7e0eb;border-radius:10px;box-shadow:0 8px 24px rgba(10,31,68,.04)}.metric{padding:16px}.metric span{display:block;color:#71839c;font-size:13px}.metric strong{display:block;margin-top:4px;font-size:22px}.panel{padding:22px;margin-bottom:18px}.drop{height:118px;border:1.5px dashed #9cb9dc;border-radius:9px;background:#f8fbff;display:grid;place-items:center;text-align:center;color:#1b5fab;font-weight:900;cursor:pointer}.drop.drag{background:#e8f2ff;border-color:#1b5fab}.drop small{display:block;color:#71839c;font-weight:600;margin-top:5px}.drop input{display:none}.toolbar{display:flex;justify-content:space-between;align-items:center;gap:12px;margin:14px 0}.toolbar input{max-width:320px}.settings-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:14px}label{display:grid;gap:7px;font-size:13px;font-weight:800;color:#2f4360}input{width:100%;border:1px solid #c9d5e3;border-radius:7px;padding:11px 12px;font:inherit}button.primary{border:0;border-radius:8px;background:#10233e;color:#fff;font-weight:900;padding:12px 18px;cursor:pointer}button.primary:disabled{opacity:.45;cursor:not-allowed}.view{display:none}.view.active{display:block}.table-wrap{overflow:auto;border:1px solid #dde6f1;border-radius:8px;background:#fff}table{width:100%;border-collapse:collapse;font-size:13px}th,td{padding:10px 12px;border-bottom:1px solid #e7edf5;text-align:left;white-space:nowrap}th{background:#f5f8fc;color:#516985;font-size:12px;text-transform:uppercase;letter-spacing:.05em}.day-row td{background:#edf4fc;color:#153a63;font-weight:900;font-size:14px}.multi td{background:#fff7dd}.empty{color:#71839c;padding:22px;border:1px dashed #c9d8ea;border-radius:8px;background:#f9fbfe}.report-grid{display:grid;grid-template-columns:1fr;gap:18px}.wide{grid-column:span 2}@media(max-width:1050px){.app{grid-template-columns:1fr}.sidebar{display:none}.cards,.settings-grid{grid-template-columns:1fr}.main{padding:20px}.top{flex-direction:column;gap:12px}}
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
      <div class="panel"><div class="toolbar"><h2>Daily sales ledger</h2><input id="search" placeholder="Search guest or room"></div><div id="ledger" class="empty">No imported check-ins yet.</div></div>
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
const q=s=>document.querySelector(s), qa=s=>[...document.querySelectorAll(s)];
function settings(){let o={};qa(".setting").forEach(i=>o[i.name]=i.value);return o}
function groupRows(rows){let html='<div class="table-wrap"><table><thead><tr><th>Room</th><th>Guest</th><th>Stay</th><th>Price Paid</th><th>Kelestarian</th></tr></thead><tbody>';let cur='';rows.forEach(r=>{if(r.display_date!==cur){cur=r.display_date;html+=`<tr class="day-row"><td colspan="5">${cur}</td></tr>`}html+=`<tr class="${r.multi_night?'multi':''}"><td>${r.room_no}</td><td>${r.guest_name}</td><td>${r.stay_progress}</td><td>RM ${r.price}</td><td>RM ${r.kelestarian}</td></tr>`});return html+'</tbody></table></div>'}
function table(rows,heads,mapper){return '<div class="table-wrap"><table><thead><tr>'+heads.map(h=>`<th>${h}</th>`).join('')+'</tr></thead><tbody>'+rows.map(mapper).join('')+'</tbody></table></div>'}
function render(){q("#mStays").textContent=summary.stays||0;q("#mRows").textContent=summary.sale_rows||0;q("#mSales").textContent='RM '+(summary.total_sales||'0.00');q("#mFee").textContent='RM '+(summary.total_kelestarian||'0.00');q("#downloadB").disabled=!stays.length;q("#downloadC").disabled=!stays.length;let term=q("#search").value.toLowerCase();let filtered=sales.filter(r=>(r.guest_name+' '+r.room_no).toLowerCase().includes(term));q("#ledger").className=filtered.length?'':'empty';q("#ledger").innerHTML=filtered.length?groupRows(filtered):'No imported check-ins yet.';let byDay={};sales.forEach(r=>{byDay[r.display_date]??={rooms:0,fee:0};byDay[r.display_date].rooms++;byDay[r.display_date].fee+=Number(r.kelestarian.replace(/,/g,''))});let bRows=Object.entries(byDay).map(([d,v])=>({d,...v}));q("#previewB").className=bRows.length?'':'empty';q("#previewB").innerHTML=bRows.length?table(bRows,['Tarikh','Jumlah Bilik','Bilangan Malam','Jumlah Kutipan'],r=>`<tr><td>${r.d}</td><td>${r.rooms}</td><td>${r.rooms}</td><td>RM ${r.fee.toFixed(2)}</td></tr>`):'Import Excel first.';q("#previewC").className=sales.length?'':'empty';q("#previewC").innerHTML=sales.length?groupRows(sales):'Import Excel first.'}
async function importFile(file){q("#dropText").innerHTML=file.name+'<small>Importing...</small>';let fd=new FormData();fd.append('file',file);let res=await fetch('/api/import',{method:'POST',body:fd});let data=await res.json();if(!res.ok){q("#notice").textContent=data.error||'Import failed';return}stays=data.stays;sales=data.sales;summary=data.summary;localStorage.setItem('stays',JSON.stringify(stays));localStorage.setItem('sales',JSON.stringify(sales));localStorage.setItem('summary',JSON.stringify(summary));q("#notice").textContent=`Imported ${summary.stays} stays and expanded them into ${summary.sale_rows} daily sales rows.`;q("#dropText").innerHTML=file.name+'<small>Imported</small>';render()}
async function download(kind){let fd=new FormData();fd.append('kind',kind);fd.append('stays',JSON.stringify(stays));Object.entries(settings()).forEach(([k,v])=>fd.append(k,v));let res=await fetch('/api/report',{method:'POST',body:fd});if(!res.ok){q("#notice").textContent=await res.text();return}let blob=await res.blob(),url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=kind==='b'?'laporan-b.pdf':'laporan-c.pdf';a.click();URL.revokeObjectURL(url)}
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
        settings = {k: escape(request.form.get(k, "").strip()) for k in request.form.keys()}
        pdf = form_b_pdf(stays, settings, fee_rate) if kind == "b" else form_c_pdf(stays, settings, fee_rate)
        name = "laporan-b.pdf" if kind == "b" else "laporan-c.pdf"
        return Response(pdf, mimetype="application/pdf", headers={"Content-Disposition": f"attachment; filename={name}"})
    except Exception as exc:
        return Response(str(exc), status=400)


if __name__ == "__main__":
    app.run(debug=True)
