create extension if not exists pgcrypto;

create table if not exists public.hotel_import_batches (
  id uuid primary key default gen_random_uuid(),
  source_filename text not null,
  imported_at timestamptz not null default now(),
  stay_count integer not null default 0,
  sale_row_count integer not null default 0,
  total_sales numeric(12,2) not null default 0,
  total_kelestarian numeric(12,2) not null default 0
);

create table if not exists public.guest_stays (
  id uuid primary key default gen_random_uuid(),
  import_batch_id uuid references public.hotel_import_batches(id) on delete set null,
  bill_no text unique,
  registration_no text,
  folio_no text,
  guest_name text,
  room_no text,
  departure_date date,
  check_in_date date,
  check_out_date date,
  number_of_pax integer,
  number_of_nights integer,
  rate_type text,
  tariff_before_tax numeric(12,2) not null default 0,
  tax_amount numeric(12,2) not null default 0,
  food_amount numeric(12,2) not null default 0,
  bar_amount numeric(12,2) not null default 0,
  laundry_amount numeric(12,2) not null default 0,
  call_amount numeric(12,2) not null default 0,
  misc_amount numeric(12,2) not null default 0,
  total_other_charges numeric(12,2) not null default 0,
  room_revenue numeric(12,2) not null default 0,
  total_amount numeric(12,2) not null default 0,
  amount_paid numeric(12,2) not null default 0,
  payment_method text,
  flags text[] not null default '{}',
  updated_at timestamptz not null default now()
);

create index if not exists guest_stays_check_in_date_idx on public.guest_stays(check_in_date);
create index if not exists guest_stays_import_batch_id_idx on public.guest_stays(import_batch_id);
create index if not exists guest_stays_room_no_idx on public.guest_stays(room_no);
create index if not exists guest_stays_payment_method_idx on public.guest_stays(payment_method);
create index if not exists guest_stays_guest_name_idx on public.guest_stays(guest_name);

alter table public.hotel_import_batches enable row level security;
alter table public.guest_stays enable row level security;

-- No public policies are created. The Flask app writes through SUPABASE_SERVICE_ROLE_KEY on the server only.
