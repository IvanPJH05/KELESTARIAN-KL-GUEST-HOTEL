import io
from datetime import date, timedelta
from decimal import Decimal

from pypdf import PdfReader

import app as hotel_app
from app import (
    C_MANUAL_ROWS,
    PAGE,
    Stay,
    classify_verified_stays,
    expanded_sales,
    filter_stays_for_report,
    form_b_pdf,
    form_c_pdf,
    report_filename,
    summary,
)


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


def stay(name, room, amount, nights, check_in=date(2026, 6, 1), folio_no=None, bill_no=None):
    check_out = check_in + timedelta(days=nights)
    total = Decimal(amount)
    return Stay(
        bill_no=bill_no or f"BN-{room}",
        registration_no=f"REG-{room}",
        folio_no=folio_no or f"FN-{room}",
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
    assert 'id="reportMonthB" type="month"' in PAGE
    assert 'id="reportStartC" type="date"' in PAGE
    assert 'id="reportEndC" type="date"' in PAGE
    assert 'data-view="review">Manual Review' in PAGE


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


def test_report_date_selection_filters_month_and_keeps_dates_on_separate_forms():
    june = stay("JUNE", "301", "100.00", 1, date(2026, 6, 30))
    july = stay("JULY", "302", "100.00", 1, date(2026, 7, 1))
    selected_b = filter_stays_for_report([june, july], "b", report_month="2026-07")
    assert [item.guest_name for item in selected_b] == ["JULY"]
    selected_c = filter_stays_for_report([june, july], "c", report_start="2026-06-30", report_end="2026-07-01")
    reader, text = pdf_text(form_c_pdf(selected_c, SETTINGS, FEE))
    assert len(reader.pages) == 4
    assert "30/06/2026" in text and "01/07/2026" in text


def test_folio_and_bill_pair_updates_matching_guest_and_queues_mismatches():
    existing = [{"id": "row-1", "folio_no": "FN30254", "bill_no": "BN100", "guest_name": "OLD NAME"}]
    matching = stay("UPDATED NAME", "401", "150.00", 1, folio_no="fn30254", bill_no="bn100")
    wrong_bill = stay("WRONG BILL", "402", "150.00", 1, folio_no="FN30254", bill_no="BN999")
    wrong_folio = stay("WRONG FOLIO", "403", "150.00", 1, folio_no="FN99999", bill_no="BN100")
    accepted, reviews = classify_verified_stays([matching, wrong_bill, wrong_folio], existing)
    assert len(accepted) == 1
    assert accepted[0]["folio_no"] == "FN30254"
    assert accepted[0]["bill_no"] == "BN100"
    assert accepted[0]["guest_name"] == "UPDATED NAME"
    assert {item["reason"] for item in reviews} == {
        "FOLIO_MATCHES_DIFFERENT_BILL",
        "BILL_MATCHES_DIFFERENT_FOLIO",
    }


def test_sales_bill_register_detail_folio_column_value_is_number_of_nights(monkeypatch):
    bill_row = [
        "BN99999",
        "REG99999",
        "FN99999",
        "23/06/2026",
        "  TEST GUEST",
        "200.00",
        "0.00",
        "0.00",
        "12.00",
        "0.00",
        "12.00",
        "0.00",
        "0.00",
        "0.00",
        "0.00",
        "0.00",
        "212.00",
    ]
    rows = [
        (bill_row, None),
        (["103", "2", "2", "WLKIN"], None),
        (["212.00", "[Cash : 212.00]"], None),
    ]
    monkeypatch.setattr(hotel_app, "excel_rows", lambda _filename, _data: iter(rows))

    parsed = hotel_app.parse_stays("Sales Bill Register.xls", b"fake excel bytes")

    assert len(parsed) == 1
    assert parsed[0].folio_no == "FN99999"
    assert parsed[0].number_of_nights == 2
    assert parsed[0].check_in_date == date(2026, 6, 21)
    assert parsed[0].check_out_date == date(2026, 6, 23)


def test_import_response_shows_only_verified_database_rows_after_save(monkeypatch):
    verified = stay("VERIFIED GUEST", "401", "150.00", 1, folio_no="FN30254", bill_no="BN100")
    rejected_mismatch = stay("REJECTED MISMATCH", "402", "150.00", 1, folio_no="FN30254", bill_no="BN999")

    monkeypatch.setattr(hotel_app, "parse_stays", lambda _filename, _data: [rejected_mismatch])
    monkeypatch.setattr(
        hotel_app,
        "save_import_to_supabase",
        lambda _filename, _stays, _sales, _summary: {"saved": True, "accepted_count": 0, "review_count": 1},
    )
    monkeypatch.setattr(hotel_app, "load_stays_from_supabase", lambda: [hotel_app.stay_record(verified)])
    monkeypatch.setattr(hotel_app, "supabase_configured", lambda: True)

    client = hotel_app.app.test_client()
    response = client.post(
        "/api/import",
        data={"fee_rate": "5.00", "file": (io.BytesIO(b"fake"), "Sales Bill Register.xls")},
        content_type="multipart/form-data",
    )

    assert response.status_code == 200
    payload = response.get_json()
    assert [item["guest_name"] for item in payload["stays"]] == ["VERIFIED GUEST"]
    assert "REJECTED MISMATCH" not in {item["guest_name"] for item in payload["stays"]}


def test_verified_rows_patch_existing_folios_and_insert_new_without_upsert(monkeypatch):
    calls = []

    def fake_supabase_request(method, table, payload=None, query=None, prefer=None):
        calls.append({"method": method, "table": table, "payload": payload, "query": query, "prefer": prefer})

    monkeypatch.setattr(hotel_app, "supabase_request", fake_supabase_request)
    existing = [{"id": "row-1", "folio_no": "FN100", "bill_no": "BN100"}]
    update_row = {"folio_no": "FN100", "bill_no": "BN100", "guest_name": "UPDATED"}
    insert_row = {"folio_no": "FN101", "bill_no": "BN101", "guest_name": "NEW"}

    hotel_app.save_verified_stay_rows([update_row, insert_row], existing)

    assert calls[0]["method"] == "PATCH"
    assert calls[0]["table"] == "guest_stays"
    assert calls[0]["query"] == {"id": "eq.row-1"}
    assert calls[1]["method"] == "POST"
    assert calls[1]["table"] == "guest_stays"
    assert calls[1]["payload"] == [insert_row]
    assert calls[1]["query"] is None
