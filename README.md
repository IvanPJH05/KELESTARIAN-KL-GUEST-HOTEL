# KELESTARIAN KL Guest Hotel

Vercel-ready Flask app for KL Guest Hotel sales tracking and Lampiran PDF generation.

## What It Does

1. Upload the hotel `Sales Bill Register` Excel export.
2. Extract guest stays, rooms, nights, payment method, tax, and sales totals.
3. Save imported guest stays to Supabase when configured.
   - Folio Number is the unique guest-stay identity.
   - Bill Number verifies that an existing Folio is the same record.
   - Mismatched Folio/Bill pairs are sent to the Manual Review page.
4. Display a dashboard grouped by date and room.
5. Separate collections by payment method.
6. Generate Lampiran B and Lampiran C PDF reports.

## Supabase

The Supabase schema is in [supabase/schema.sql](supabase/schema.sql).

The base schema has been applied to project:

```text
akvvxxaufvlprqusxxzq
```

Add these Vercel environment variables:

```text
SUPABASE_URL=https://akvvxxaufvlprqusxxzq.supabase.co
SUPABASE_SERVICE_ROLE_KEY=your-service-role-key
```

Use the service role key only on the server. Do not expose it in browser code.

Before deploying the Folio/Bill verification feature, run
[`supabase/migrations/20260624_folio_bill_verification.sql`](supabase/migrations/20260624_folio_bill_verification.sql)
once in the Supabase SQL Editor. It removes existing duplicate identifiers (keeping the most recently updated row), sends those conflicts to the review queue, and adds the Folio uniqueness constraint.

## Deploy To Vercel

Connect this repository to Vercel. Vercel detects the Flask app from `app.py` and installs dependencies from `requirements.txt`.

After changing environment variables, redeploy the latest `main` branch.
