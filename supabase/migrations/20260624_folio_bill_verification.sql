create extension if not exists pgcrypto;

create table if not exists public.guest_stay_review_queue (
  id uuid primary key default gen_random_uuid(),
  review_key text not null unique,
  import_batch_id uuid references public.hotel_import_batches(id) on delete set null,
  source_filename text,
  folio_no text,
  incoming_bill_no text,
  existing_folio_no text,
  existing_bill_no text,
  reason text not null,
  incoming_record jsonb not null,
  existing_record jsonb,
  status text not null default 'pending' check (status in ('pending', 'reviewed')),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

with ranked_folios as (
  select g.*, first_value(id) over (partition by upper(folio_no) order by updated_at desc, id desc) as keep_id,
         row_number() over (partition by upper(folio_no) order by updated_at desc, id desc) as row_no
  from public.guest_stays g
  where folio_no is not null and btrim(folio_no) <> ''
)
insert into public.guest_stay_review_queue
  (review_key, folio_no, incoming_bill_no, existing_folio_no, existing_bill_no, reason, incoming_record, existing_record)
select 'existing-folio|' || duplicate.id::text, upper(duplicate.folio_no), upper(duplicate.bill_no),
       upper(kept.folio_no), upper(kept.bill_no), 'EXISTING_DUPLICATE_FOLIO', to_jsonb(duplicate), to_jsonb(kept)
from ranked_folios duplicate
join public.guest_stays kept on kept.id = duplicate.keep_id
where duplicate.row_no > 1
on conflict (review_key) do nothing;

with ranked_folios as (
  select id, row_number() over (partition by upper(folio_no) order by updated_at desc, id desc) as row_no
  from public.guest_stays
  where folio_no is not null and btrim(folio_no) <> ''
)
delete from public.guest_stays
where id in (select id from ranked_folios where row_no > 1);

with ranked_bills as (
  select g.*, first_value(id) over (partition by upper(bill_no) order by updated_at desc, id desc) as keep_id,
         row_number() over (partition by upper(bill_no) order by updated_at desc, id desc) as row_no
  from public.guest_stays g
  where bill_no is not null and btrim(bill_no) <> ''
)
insert into public.guest_stay_review_queue
  (review_key, folio_no, incoming_bill_no, existing_folio_no, existing_bill_no, reason, incoming_record, existing_record)
select 'existing-bill|' || duplicate.id::text, upper(duplicate.folio_no), upper(duplicate.bill_no),
       upper(kept.folio_no), upper(kept.bill_no), 'EXISTING_DUPLICATE_BILL', to_jsonb(duplicate), to_jsonb(kept)
from ranked_bills duplicate
join public.guest_stays kept on kept.id = duplicate.keep_id
where duplicate.row_no > 1
on conflict (review_key) do nothing;

with ranked_bills as (
  select id, row_number() over (partition by upper(bill_no) order by updated_at desc, id desc) as row_no
  from public.guest_stays
  where bill_no is not null and btrim(bill_no) <> ''
)
delete from public.guest_stays
where id in (select id from ranked_bills where row_no > 1);

update public.guest_stays
set folio_no = upper(btrim(folio_no)), bill_no = upper(btrim(bill_no));

create unique index if not exists guest_stays_folio_no_unique_idx on public.guest_stays(folio_no);
create index if not exists guest_stay_review_queue_status_idx on public.guest_stay_review_queue(status, updated_at desc);

alter table public.guest_stay_review_queue enable row level security;
