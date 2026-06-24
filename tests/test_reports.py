import io
from datetime import date, timedelta
from decimal import Decimal

from pypdf import PdfReader

from app import C_MANUAL_ROWS, PAGE, Stay, expanded_sales, form_b_pdf, form_c_pdf, report_filename, summary


FEE = Decimal("5.00")
SETTINGS = {
    "premise_name": "KL GUEST HOTEL SDN BHD",
    "license_no": "L00619905",
    "certificate_no": "3567",
    "address": "8, Jalan AU 1a/4c, Taman Keramat Permai, 54200 Kuala Lumpur, Federal Territory of Kuala Lumpur",
    "contact_name": "Phang Huey Jiun",
    "contact": "012-205-0039 / hueyjiunphang@gmail.com",
    "category_code": "Hotel 1-3 bintang",
}


def stay(name, room, amount, nights, check_in=date(2026, 6, 1)):
    check_out = check_in + timedelta(days=nights)
    total = Decimal(amount)
    return Stay(
        bill_no=f"BN-{room}",
        registration_no=f"REG-{room}",
        folio_no=f"F-{room}",
        guest_name=name,
        room_no=room,
        departure_date=check_out,
        check_in_date=check_in,
        check_out_date=check_out,
        number_of_pax=1,
        number_of_nights=nights,
        rate_type="ROOM",
        tariff_before_tax=total,
        tax_amount=Decimal("0"),
        food_amount=Decimal("0"),
        bar_amount=Decimal("0"),
        laundry_amount=Decimal("0"),
        call_amount=Decimal("0"),
        misc_amount=Decimal("0"),
        total_amount=total,
        amount_paid=total,
        payment_method="CASH",
        flags=[],
    )


def scenarios():
    return [
        stay("WALK SINGLE", "101", "100.00", 1),
        stay("ONLINE SINGLE", "102", "120.00", 1),
        stay("WALK MULTI", "103", "300.00", 3),
        stay("ONLINE MULTI", "104", "240.00", 2),
    ]


def pdf_text(pdf):
    reader = PdfReader(io.BytesIO(pdf))
    return reader, "\n".join(page.extract_text() or "" for page in reader.pages)


def test_reporting_settings_include_default_contact_details():
    assert 'value="012-205-0039 / hueyjiunphang@gmail.com"' in PAGE
    assert 'value="8, Jalan AU 1a/4c, Taman Keramat Permai, 54200 Kuala Lumpur, Federal Territory of Kuala Lumpur"' in PAGE


def test_multi_night_rows_do_not_duplicate_collections():
    sales = expanded_sales(scenarios(), FEE)
    assert len(sales) == 7
    first_days = [row for row in sales if not row["is_paid_continuation"]]
    later_days = [row for row in sales if row["is_paid_continuation"]]
    assert len(first_days) == 4
    assert len(later_days) == 3
    assert sum(Decimal(row["kelestarian"].replace(",", "")) for row in sales) == Decimal("35.00")
    assert all(row["price"] == row["kelestarian"] == "0.00" for row in later_days)
    totals = summary(scenarios(), FEE)
    assert totals["total_sales"] == "760.00"
    assert totals["total_kelestarian"] == "35.00"


def test_official_reports_use_letter_pages_totals_and_readable_filenames():
    b_reader, b_text = pdf_text(form_b_pdf(scenarios(), SETTINGS, FEE))
    c_reader, c_text = pdf_text(form_c_pdf(scenarios(), SETTINGS, FEE))
    assert len(b_reader.pages) == 2
    assert len(c_reader.pages) == 2
    assert all((float(page.mediabox.width), float(page.mediabox.height)) == (612.0, 792.0) for page in [*b_reader.pages, *c_reader.pages])
    assert "LAPORAN PENYATA KUTIPAN FI KELESTARIAN NEGERI SELANGOR (BULANAN)" in b_text
    assert "Federal Territory of Kuala Lumpur" in b_text
    assert "Jumlah" in b_text and "35.00" in b_text
    assert "LAPORAN TRANSAKSI PENGGUNAAN BILIK (HARIAN)" in c_text
    assert "Jumlah Bilik" in c_text and "Bilangan" in c_text and "Malam" in c_text
    assert "No. Bilik" not in c_text
    assert "Paid on" not in c_text
    assert report_filename("b", scenarios()) == "Lampiran B June 2026.pdf"
    assert report_filename("c", scenarios()) == "Lampiran C 01 June 2026.pdf"


def test_lampiran_b_drops_days_that_do_not_exist_in_the_month():
    april = [stay("APRIL GUEST", "201", "100.00", 1, date(2026, 4, 30))]
    _, text = pdf_text(form_b_pdf(april, SETTINGS, FEE))
    assert "30/04/2026" in text
    assert "31/04/2026" not in text


def test_lampiran_c_keeps_five_manual_rows_after_overflow_data():
    guests = [stay(f"GUEST {index:02d}", str(200 + index), "100.00", 1) for index in range(1, 31)]
    reader, text = pdf_text(form_c_pdf(guests, SETTINGS, FEE))
    assert C_MANUAL_ROWS == 5
    assert len(reader.pages) == 2
    assert "GUEST 30" in text
    assert "150.00" in text
