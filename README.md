# KELESTARIAN KL Guest Hotel

Vercel-ready Flask app for KL Guest Hotel sales tracking and Lampiran PDF generation.

## What It Does

1. Upload the hotel `Sales Bill Register` Excel export.
2. Extract guest stays, rooms, nights, payment method, tax, and sales totals.
3. Save imported guest stays to Supabase when configured.
4. Display a dashboard grouped by date and room.
5. Separate collections by payment method.
6. Generate Lampiran B and Lampiran C PDF reports.

## Supabase

The Supabase schema is in [supabase/schema.sql](supabase/schema.sql).

The schema has already been applied to project:

```text
joaoirpegnkexmktylop
```

Add these Vercel environment variables:

```text
SUPABASE_URL=https://joaoirpegnkexmktylop.supabase.co
SUPABASE_SERVICE_ROLE_KEY=your-service-role-key
```

Use the service role key only on the server. Do not expose it in browser code.

## Deploy To Vercel

Connect this repository to Vercel. Vercel detects the Flask app from `app.py` and installs dependencies from `requirements.txt`.

After changing environment variables, redeploy the latest `main` branch.
