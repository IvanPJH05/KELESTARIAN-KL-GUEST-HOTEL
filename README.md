# KELESTARIAN KL Guest Hotel

This Vercel-ready Flask app generates two DOCX reports from the guest check-in Excel export:

- Lampiran B: Laporan Penyata Kutipan Fi Kelestarian Negeri Selangor (Bulanan)
- Lampiran C: Laporan Transaksi Penggunaan Bilik (Harian)

## How it works

1. Upload the guest check-in `.xls` or `.xlsx` file.
2. Fill in the hotel and licence details.
3. Confirm the fee rate. The default is RM5.00 per room-night.
4. Download the ZIP containing both completed DOCX files.

The parser expects the guest check-in export layout used in the provided sample:

- Column A: Voucher number
- Column B: Room number
- Column C: Room type
- Column D: Room rate
- Column F: Guest name
- Column K: Check-in date
- Column L: Check-in time
- Column M: Check-out date
- Column N: Number of nights

## Deploy to Vercel

Connect this repository to Vercel, or deploy from the Vercel CLI:

```bash
vercel deploy
```

Vercel detects the Flask app from `app.py` and installs dependencies from `requirements.txt`.
