from datetime import date
from decimal import Decimal
from pathlib import Path

from pypdf import PdfReader

from app import PAGE, Stay, expanded_sales, form_b_pdf, form_c_pdf, parse_stays, summary


FEE = Decimal("5.00")
SETTINGS = {
    "premise_name": "Audit Hotel",
    "license_no": "PBT-123",
    "certificate_no": "CERT-1",
    "address": "Selangor",
    "contact_name": "Manager",
    "contact": "0123456789",
    "category_code": "HOTEL",
}


def stay(name, code, room, rate, nights):
    kind = "Walk-in" if code == "WLKIN" else "Online Booking"
    return Stay(
        voucher=f"V-{room}",
        room_no=room,
        room_type="Room",
        rate=Decimal(rate),
        guest_name=name,
        checkin_date=date(2026, 6, 1),
        checkin_time="14:00",
        checkout_date=date(2026, 6, 1 + nights),
        nights=nights,
        booking_code=code,
        customer_type=kind,
    )


def scenarios():
    return [
        stay("WALK SINGLE", "WLKIN", "101", "100.00", 1),
        stay("ONLINE SINGLE", "XPDIA", "102", "120.00", 1),
        stay("WALK MULTI", "WLKIN", "103", "300.00", 3),
        stay("ONLINE MULTI", "XPDIA", "104", "240.00", 2),
    ]


def test_four_customer_and_stay_scenarios_do_not_duplicate_payments():
    sales = expanded_sales(scenarios(), FEE)
    assert len(sales) == 7
    assert [row["customer_type"] for row in sales if row["stay_progress"] == "1/1"] == [
        "Walk-in",
        "Online Booking",
    ]

    first_days = [row for row in sales if row["payment_collected"]]
    later_days = [row for row in sales if not row["payment_collected"]]
    assert len(first_days) == 4
    assert len(later_days) == 3
    single_nights = [row for row in sales if not row["multi_night"]]
    assert all(row["row_state"] == "single" and row["row_class"] == "" for row in single_nights)
    assert all(row["row_state"] == "payment" and row["row_class"] == "payment-row" for row in first_days if row["multi_night"])
    assert all(row["row_state"] == "paid" for row in later_days)
    assert all(row["row_class"] == "paid-row" for row in later_days)
    assert all(row["price"] == row["kelestarian"] == "PAID" for row in later_days)
    assert sum(Decimal(row["actual_price"]) for row in sales) == Decimal("760.00")
    assert sum(Decimal(row["actual_kelestarian"]) for row in sales) == Decimal("35.00")

    totals = summary(scenarios(), FEE)
    assert totals["total_sales"] == "760.00"
    assert totals["total_kelestarian"] == "35.00"
    assert "r.row_class??(r.multi_night?" in PAGE


def test_official_layout_pdfs_have_expected_pages_and_totals():
    b_pdf = form_b_pdf(scenarios(), SETTINGS, FEE)
    c_pdf = form_c_pdf(scenarios(), SETTINGS, FEE)
    b_reader = PdfReader(__import__("io").BytesIO(b_pdf))
    c_reader = PdfReader(__import__("io").BytesIO(c_pdf))
    assert len(b_reader.pages) == 2
    assert len(c_reader.pages) == 2  # official daily transaction form is grouped by check-in date
    b_text = "\n".join(page.extract_text() or "" for page in b_reader.pages)
    c_text = "\n".join(page.extract_text() or "" for page in c_reader.pages)
    assert "LAMPIRAN B" in b_text
    assert "35.00" in b_text
    assert "LAMPIRAN C" in c_text
    assert "Jumlah Bilik (Unit)" in c_text
    assert "Bilangan Malam" in c_text
    assert "Online Booking" not in c_text
    assert "PAID" not in c_text


def test_real_hotel_export_uses_column_e_and_strips_leading_numbers():
    source = Path(r"C:\Users\nickb\Downloads\Guest Checked In_20260622.xls")
    stays = parse_stays(source.name, source.read_bytes())
    assert stays
    assert {item.booking_code for item in stays}.issuperset({"WLKIN", "XPDIA"})
    assert all(not item.guest_name[:1].isdigit() for item in stays)
    assert all(item.customer_type == "Walk-in" for item in stays if item.booking_code == "WLKIN")
    assert all(item.customer_type == "Online Booking" for item in stays if item.booking_code == "XPDIA")
    assert any(item.booking_code == "WLKIN" and item.guest_name == "Jamaldi Saeden" for item in stays)
    assert any(item.booking_code == "XPDIA" and item.guest_name == "HELNA SARI" for item in stays)
    sales = expanded_sales(stays, FEE)
    assert len(sales) == 527
    assert sum(Decimal(row["actual_kelestarian"].replace(",", "")) for row in sales) == Decimal("2635.00")
    assert all(row["price"] == row["kelestarian"] == "PAID" for row in sales if not row["payment_collected"])
