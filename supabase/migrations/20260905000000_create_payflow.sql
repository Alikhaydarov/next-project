begin;

create extension if not exists pgcrypto;
create schema if not exists payflow;

do $migration$
begin
  create type payflow.account_kind as enum ('OUR', 'COMPANY');
exception
  when duplicate_object then null;
end
$migration$;

create table if not exists payflow.companies (
  id uuid primary key default gen_random_uuid(),
  name text not null unique,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists payflow.accounts (
  id uuid primary key default gen_random_uuid(),
  kind payflow.account_kind not null,
  company_id uuid references payflow.companies(id) on delete cascade,
  label text not null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  constraint accounts_kind_company_check check (
    (kind = 'OUR' and company_id is null)
    or (kind = 'COMPANY' and company_id is not null)
  )
);

create table if not exists payflow.app_settings (
  singleton boolean primary key default true check (singleton),
  access_key_sha256 text not null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists payflow.payments (
  id uuid primary key default gen_random_uuid(),
  company_id uuid not null references payflow.companies(id) on delete restrict,
  our_account_id uuid references payflow.accounts(id) on delete set null,
  company_account_id uuid references payflow.accounts(id) on delete set null,
  amount numeric not null check (amount > 0),
  paid_at timestamptz not null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  company_payment_method text check (company_payment_method in ('cash', 'card')),
  our_payment_method text not null check (our_payment_method in ('cash', 'card')),
  company_card_account_text text,
  our_card_account_text text,
  receipt_path text,
  description text,
  receipt_paths text[] not null default '{}',
  constraint payments_method_fields_check check (
    (
      our_payment_method = 'cash'
      and our_account_id is null
      and company_account_id is null
      and company_payment_method is null
      and nullif(trim(coalesce(our_card_account_text, '')), '') is null
      and nullif(trim(coalesce(company_card_account_text, '')), '') is null
    )
    or
    (
      our_payment_method = 'card'
      and company_payment_method is not null
      and nullif(trim(coalesce(our_card_account_text, '')), '') is not null
      and nullif(trim(coalesce(company_card_account_text, '')), '') is not null
    )
  )
);

create index if not exists accounts_company_id_idx
  on payflow.accounts(company_id);
create index if not exists accounts_kind_idx
  on payflow.accounts(kind);
create index if not exists payments_company_account_id_idx
  on payflow.payments(company_account_id);
create index if not exists payments_company_id_idx
  on payflow.payments(company_id);
create index if not exists payments_created_at_idx
  on payflow.payments(created_at desc);
create index if not exists payments_our_account_id_idx
  on payflow.payments(our_account_id);
create index if not exists payments_paid_at_idx
  on payflow.payments(paid_at desc);

alter table payflow.companies enable row level security;
alter table payflow.accounts enable row level security;
alter table payflow.app_settings enable row level security;
alter table payflow.payments enable row level security;

revoke all on schema payflow from public, anon, authenticated;
revoke all on all tables in schema payflow from public, anon, authenticated;
revoke all on all sequences in schema payflow from public, anon, authenticated;
grant usage on schema payflow to service_role;
grant select, insert, update, delete on all tables in schema payflow to service_role;
grant usage, select on all sequences in schema payflow to service_role;

alter default privileges in schema payflow
  revoke all on tables from public, anon, authenticated;
alter default privileges in schema payflow
  grant select, insert, update, delete on tables to service_role;
alter default privileges in schema payflow
  grant usage, select on sequences to service_role;

insert into storage.buckets (
  id,
  name,
  public,
  file_size_limit,
  allowed_mime_types
)
values (
  'payflow-receipts',
  'payflow-receipts',
  false,
  5242880,
  array['image/jpeg', 'image/png', 'image/webp']
)
on conflict (id) do update set
  name = excluded.name,
  public = excluded.public,
  file_size_limit = excluded.file_size_limit,
  allowed_mime_types = excluded.allowed_mime_types;

commit;
