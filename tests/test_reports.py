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
    dedupe_import_records,
    expanded_sales,
    find_existing_matches,
    filter_stays_for_report,
    form_b_pdf,
    form_c_pdf,
    historical_summary,
    ledger_sales,
    normalize_payment_method,
    report_filename,
    stay_record,
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
    assert 'id="historicalYear"></select>' in PAGE
    assert "All years" not in PAGE
    assert 'year=q("#historicalYear").value||localStorage.getItem("historicalYear")||"2025"' in PAGE
    assert 'APP_VERSION="dashboard-custom-ledger-20260625"' in PAGE
    assert 'id="reportMonthB" type="month"' in PAGE
    assert 'id="reportMonthC" type="month"' in PAGE
    assert 'id="reportStartC" type="date"' in PAGE
    assert 'id="reportEndC" type="date"' in PAGE
    assert 'data-view="review">Manual Review' in PAGE


def test_payment_methods_are_categorised():
    assert normalize_payment_method("Visa Debit Card") == "Card"
    assert normalize_payment_method("Master Card") == "Card"
    assert normalize_payment_method("ATM Card") == "Bank / ATM"
    assert normalize_payment_method("Online Booking") == "Online Booking"
    assert normalize_payment_method("Security Deposit") == "Security Deposit"


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


def test_kelestarian_rows_are_limited_to_calendar_year_2026():
    before = stay("BEFORE", "001", "100.00", 2, date(2025, 12, 29))
    crossing_start = stay("CROSS START", "002", "400.00", 4, date(2025, 12, 30))
    crossing_end = stay("CROSS END", "003", "300.00", 3, date(2026, 12, 30))
    after = stay("AFTER", "004", "100.00", 1, date(2027, 1, 1))

    sales = expanded_sales([before, crossing_start, crossing_end, after], FEE)

    assert [row["date"] for row in sales] == ["2026-12-30", "2026-12-31"]
    assert sum(Decimal(row["kelestarian"].replace(",", "")) for row in sales) == Decimal("10.00")
    assert sales[0]["kelestarian"] == "10.00"


def test_ledger_sales_can_display_historical_custom_dates_without_kelestarian():
    old = stay("OLD GUEST", "201", "120.00", 2, date(2024, 1, 10))
    new = stay("NEW GUEST", "202", "100.00", 1, date(2026, 1, 10))

    rows = ledger_sales([old, new], FEE)

    assert [row["date"] for row in rows] == ["2024-01-10", "2024-01-11", "2026-01-10"]
    assert rows[0]["price"] == "120.00"
    assert rows[0]["kelestarian"] == "0.00"
    assert rows[1]["payment_status"] == "Paid on 10/01/2024"
    assert rows[2]["kelestarian"] == "5.00"


def test_report_selection_excludes_pre_2026_checkins():
    crossing = stay("CROSS START", "002", "400.00", 4, date(2025, 12, 30))
    valid = stay("VALID", "003", "100.00", 2, date(2026, 1, 1))
    after = stay("AFTER", "004", "100.00", 1, date(2027, 1, 1))

    selected_b = filter_stays_for_report([crossing, valid, after], "b", report_month="2026-01")
    selected_c = filter_stays_for_report([crossing, valid, after], "c", report_start="2026-01-01", report_end="2026-01-02")

    assert [item.guest_name for item in selected_b] == ["VALID"]
    assert [item.guest_name for item in selected_c] == ["VALID"]


def test_historical_summary_excludes_2026_kelestarian_data():
    old = stay("OLD", "101", "100.00", 1, date(2025, 6, 1))
    kelestarian = stay("NEW", "102", "100.00", 1, date(2026, 6, 1))

    data = historical_summary([old, kelestarian], FEE)

    assert data["historical_stays"] == 1
    assert data["total_revenue"] == "100.00"
    assert data["yearly_revenue"] == [{"year": "2025", "stays": 1, "nights": 1, "revenue": "100.00"}]


def test_historical_summary_can_filter_one_year():
    old_2024 = stay("OLD 2024", "101", "100.00", 1, date(2024, 6, 1))
    old_2025 = stay("OLD 2025", "102", "200.00", 1, date(2025, 6, 1))

    data = historical_summary([old_2024, old_2025], FEE, 2024)

    assert data["selected_year"] == "2024"
    assert data["historical_stays"] == 1
    assert data["total_revenue"] == "100.00"
    assert data["yearly_revenue"] == [{"year": "2024", "stays": 1, "nights": 1, "revenue": "100.00"}]


def test_historical_api_defaults_to_latest_single_year(monkeypatch):
    calls = []

    def fake_load(start, end):
        calls.append((start, end))
        return [stay_record(stay("OLD 2025", "102", "200.00", 1, date(2025, 6, 1)))]

    monkeypatch.setattr(hotel_app, "supabase_configured", lambda: True)
    monkeypatch.setattr(hotel_app, "load_stays_by_checkout_range_from_supabase", fake_load)
    monkeypatch.setattr(hotel_app, "load_import_batches_from_supabase", lambda: [])

    response = hotel_app.app.test_client().get("/api/historical?fee_rate=5.00")
    payload = response.get_json()

    assert response.status_code == 200
    assert calls == [(date(2025, 1, 1), date(2025, 12, 31))]
    assert payload["summary"]["selected_year"] == "2025"
    assert payload["summary"]["yearly_revenue"] == [{"year": "2025", "stays": 1, "nights": 1, "revenue": "200.00"}]


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
    selected_c_month = filter_stays_for_report([june, july], "c", report_month="2026-07")
    assert [item.guest_name for item in selected_c_month] == ["JULY"]
    selected_c = filter_stays_for_report([june, july], "c", report_start="2026-06-30", report_end="2026-07-01")
    reader, text = pdf_text(form_c_pdf(selected_c, SETTINGS, FEE))
    assert len(reader.pages) == 4
    assert "30/06/2026" in text and "01/07/2026" in text


def test_report_pdfs_are_clipped_to_selected_month_and_date_range():
    crossing = stay("CROSS MONTH", "401", "200.00", 2, date(2026, 4, 30))

    b_reader, b_text = pdf_text(form_b_pdf([crossing], SETTINGS, FEE, date(2026, 4, 1), date(2026, 4, 30)))
    c_reader, c_text = pdf_text(form_c_pdf([crossing], SETTINGS, FEE, date(2026, 4, 30), date(2026, 4, 30)))

    assert len(b_reader.pages) == 2
    assert "April 2026" in b_text
    assert "30/04/2026" in b_text
    assert "01/05/2026" not in b_text
    assert len(c_reader.pages) == 2
    assert "30/04/2026" in c_text
    assert "01/05/2026" not in c_text
    assert report_filename("b", [crossing], date(2026, 4, 1), date(2026, 4, 30)) == "Lampiran B April 2026.pdf"


def test_api_report_uses_month_for_lampiran_c_pdf():
    crossing = stay("CROSS MONTH", "401", "200.00", 2, date(2026, 4, 30))
    client = hotel_app.app.test_client()

    response = client.post(
        "/api/report",
        data={
            "kind": "c",
            "stays": hotel_app.json.dumps([stay_record(crossing)]),
            "report_month": "2026-04",
            **SETTINGS,
            "fee_rate": "5.00",
        },
    )

    assert response.status_code == 200
    _, text = pdf_text(response.data)
    assert "30/04/2026" in text
    assert "01/05/2026" not in text


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


def test_import_response_returns_uploaded_rows_without_full_history_reload(monkeypatch):
    imported = stay("IMPORTED GUEST", "402", "150.00", 1, folio_no="FN30254", bill_no="BN999")

    monkeypatch.setattr(hotel_app, "parse_stays", lambda _filename, _data: [imported])
    monkeypatch.setattr(
        hotel_app,
        "save_import_to_supabase",
        lambda _filename, _stays, _sales, _summary: {"saved": True, "accepted_count": 1, "review_count": 0},
    )
    monkeypatch.setattr(
        hotel_app,
        "load_stays_from_supabase",
        lambda: (_ for _ in ()).throw(AssertionError("import should not load full history")),
    )
    monkeypatch.setattr(hotel_app, "supabase_configured", lambda: True)

    client = hotel_app.app.test_client()
    response = client.post(
        "/api/import",
        data={"fee_rate": "5.00", "file": (io.BytesIO(b"fake"), "Sales Bill Register.xls")},
        content_type="multipart/form-data",
    )

    assert response.status_code == 200
    payload = response.get_json()
    assert [item["guest_name"] for item in payload["stays"]] == ["IMPORTED GUEST"]
    assert payload["summary"]["stays"] == 1


def test_dashboard_history_loads_only_2026_checkins(monkeypatch):
    calls = []
    valid = stay_record(stay("VALID 2026", "601", "100.00", 1, date(2026, 6, 1)))

    def fake_load(start, end):
        calls.append((start, end))
        return [valid]

    monkeypatch.setattr(hotel_app, "supabase_configured", lambda: True)
    monkeypatch.setattr(hotel_app, "load_stays_by_checkin_range_from_supabase", fake_load)
    monkeypatch.setattr(
        hotel_app,
        "load_stays_from_supabase",
        lambda: (_ for _ in ()).throw(AssertionError("dashboard should not load the full archive")),
    )

    response = hotel_app.app.test_client().get("/api/history?fee_rate=5.00")
    payload = response.get_json()

    assert response.status_code == 200
    assert calls == [(date(2026, 1, 1), date(2026, 12, 31))]
    assert [item["guest_name"] for item in payload["stays"]] == ["VALID 2026"]
    assert payload["summary"]["stays"] == 1


def test_dashboard_history_custom_range_loads_selected_checkins(monkeypatch):
    calls = []
    old = stay_record(stay("OLD CUSTOM", "701", "120.00", 2, date(2024, 1, 10)))

    def fake_load(start, end):
        calls.append((start, end))
        return [old]

    monkeypatch.setattr(hotel_app, "supabase_configured", lambda: True)
    monkeypatch.setattr(hotel_app, "load_stays_by_checkin_range_from_supabase", fake_load)

    response = hotel_app.app.test_client().get("/api/history?fee_rate=5.00&start=2024-01-01&end=2024-01-31")
    payload = response.get_json()

    assert response.status_code == 200
    assert calls == [(date(2024, 1, 1), date(2024, 1, 31))]
    assert [item["date"] for item in payload["sales"]] == ["2024-01-10", "2024-01-11"]
    assert payload["sales"][0]["price"] == "120.00"
    assert payload["sales"][0]["kelestarian"] == "0.00"


def test_verified_rows_batch_update_existing_folios_and_insert_new(monkeypatch):
    calls = []

    def fake_supabase_request(method, table, payload=None, query=None, prefer=None):
        calls.append({"method": method, "table": table, "payload": payload, "query": query, "prefer": prefer})

    monkeypatch.setattr(hotel_app, "supabase_request", fake_supabase_request)
    existing = [{"id": "row-1", "folio_no": "FN100", "bill_no": "BN100"}]
    update_row = {"folio_no": "FN100", "bill_no": "BN100", "guest_name": "UPDATED"}
    insert_row = {"folio_no": "FN101", "bill_no": "BN101", "guest_name": "NEW"}

    hotel_app.save_verified_stay_rows([update_row, insert_row], existing)

    assert calls[0]["method"] == "POST"
    assert calls[0]["table"] == "guest_stays"
    assert calls[0]["query"] == {"on_conflict": "id"}
    assert calls[0]["prefer"] == "resolution=merge-duplicates,return=minimal"
    assert calls[0]["payload"] == [{**update_row, "id": "row-1"}]
    assert calls[1]["method"] == "POST"
    assert calls[1]["table"] == "guest_stays"
    assert calls[1]["payload"] == [insert_row]
    assert calls[1]["query"] is None


def test_verified_rows_batch_update_existing_bill_when_database_has_bill_constraint(monkeypatch):
    calls = []

    def fake_supabase_request(method, table, payload=None, query=None, prefer=None):
        calls.append({"method": method, "table": table, "payload": payload, "query": query, "prefer": prefer})

    monkeypatch.setattr(hotel_app, "supabase_request", fake_supabase_request)
    existing = [{"id": "row-bill", "folio_no": "FN35634", "bill_no": "BN35634"}]
    update_row = {"folio_no": "FN35634", "bill_no": "BN35634", "guest_name": "UPDATED"}

    hotel_app.save_verified_stay_rows([update_row], existing)

    assert calls == [
        {
            "method": "POST",
            "table": "guest_stays",
            "payload": [{**update_row, "id": "row-bill"}],
            "query": {"on_conflict": "id"},
            "prefer": "resolution=merge-duplicates,return=minimal",
        }
    ]


def test_import_loads_existing_rows_by_incoming_folio_and_bill(monkeypatch):
    calls = []

    def fake_supabase_request(method, table, payload=None, query=None, prefer=None):
        calls.append(query)
        if query and query.get("bill_no") == "in.(BN35634)":
            return [{"id": "row-bill", "folio_no": "FN35634", "bill_no": "BN35634"}]
        return []

    monkeypatch.setattr(hotel_app, "supabase_configured", lambda: True)
    monkeypatch.setattr(hotel_app, "supabase_request", fake_supabase_request)
    incoming = stay("GUEST", "305", "100.00", 1, folio_no="FN35634", bill_no="BN35634")

    rows = hotel_app.load_matching_stays_from_supabase([incoming])

    assert rows == [{"id": "row-bill", "folio_no": "FN35634", "bill_no": "BN35634"}]
    assert {"select": "*", "folio_no": "in.(FN35634)", "limit": "1000", "offset": "0"} in calls
    assert {"select": "*", "bill_no": "in.(BN35634)", "limit": "1000", "offset": "0"} in calls
def test_import_dedupe_keeps_latest_row_for_same_folio():
    first = stay_record(stay("OLD GUEST", "501", "100.00", 1, folio_no="FN777", bill_no="BN777"))
    second = stay_record(stay("UPDATED GUEST", "501", "120.00", 1, folio_no="fn777", bill_no="bn777"))
    rows, reviews = dedupe_import_records([first, second])
    assert len(rows) == 1
    assert rows[0]["guest_name"] == "UPDATED GUEST"
    assert rows[0]["total_amount"] == 120.0
    assert len(reviews) == 1
    assert reviews[0]["reason"] == "DUPLICATE_IN_UPLOAD"


def test_existing_match_uses_folio_or_bill_and_prefers_newest_row():
    incoming = stay_record(stay("UPDATED", "601", "150.00", 1, folio_no="FN888", bill_no="BN888"))
    existing = [
        {"id": "old", "folio_no": "FN888", "bill_no": "BN000", "updated_at": "2026-01-01T00:00:00Z"},
        {"id": "new", "folio_no": "FN000", "bill_no": "BN888", "updated_at": "2026-02-01T00:00:00Z"},
    ]
    matches = find_existing_matches(incoming, existing)
    assert [row["id"] for row in matches] == ["new", "old"]
