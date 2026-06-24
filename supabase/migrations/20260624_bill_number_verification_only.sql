-- Folio Number is the unique stay identity.
-- Bill Number is kept for verification and review matching, but it should not block updates.
alter table if exists public.guest_stays
  drop constraint if exists guest_stays_bill_no_key;

drop index if exists public.guest_stays_bill_no_key;
create index if not exists guest_stays_bill_no_idx on public.guest_stays(bill_no);