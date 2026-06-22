# KELESTARIAN KL Guest Hotel

This Vercel-ready Flask app generates two PDF reports from the guest check-in Excel export:

- Lampiran B: Laporan Penyata Kutipan Fi Kelestarian Negeri Selangor (Bulanan)
- Lampiran C: Laporan Transaksi Penggunaan Bilik (Harian)

## How it works

1. Upload the guest check-in `.xls` or `.xlsx` file.
2. Fill in the hotel and licence details.
3. Confirm the fee rate. The default is RM5.00 per room-night.
4. Download the ZIP containing both completed PDF files.

The parser expects the guest check-in export layout used in the provided sample:

- Column A: Voucher number
- Column B: Room number
- Column C: Room type
- Column D: Room rate
- Column E: Booking source (`WLKIN` = Walk-in, `XPDIA` = Online Booking)
- Column F: Guest name
- Column K: Check-in date
- Column L: Check-in time
- Column M: Check-out date
- Column N: Number of nights

The website ledger expands every stay across all occupied nights. One-night stays remain uncoloured. For multi-night stays, the check-in/payment row is yellow; subsequent already-paid nights are blue and show `PAID` instead of repeating the amounts. These audit colours and labels are website-only. Generated PDFs retain the official layouts: Lampiran C records each booking once on its check-in date, with the full number of nights and corresponding fee, so totals are not duplicated.

## Deploy to Vercel

Connect this repository to Vercel, or deploy from the Vercel CLI:

```bash
vercel deploy
```

Vercel detects the Flask app from `app.py` and installs dependencies from `requirements.txt`.
