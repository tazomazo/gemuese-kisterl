-- ============================================================
-- Andreas' Gemüsekisterl – Supabase Setup (v2)
-- Dieses Script im Supabase SQL Editor ausführen
-- ============================================================

-- Bestehende Tabellen löschen (falls vorhanden)
drop table if exists public.orders cascade;
drop table if exists public.products cascade;
drop table if exists public.users cascade;

-- 1. Tabelle: products
-- ============================================================
create table public.products (
  id         bigint generated always as identity primary key,
  name       text    not null,
  price      numeric not null,
  unit       text    not null default '',
  category   text    not null default 'Alle Produkte',
  special    boolean not null default false,
  created_at timestamptz not null default now()
);

-- 2. Tabelle: users
-- ============================================================
create table public.users (
  id           bigint generated always as identity primary key,
  name         text    not null unique,
  password     text    not null default '',
  created_at   timestamptz not null default now()
);

-- 3. Tabelle: orders
-- ============================================================
create table public.orders (
  id          bigint generated always as identity primary key,
  user_id     bigint  not null references public.users(id) on delete cascade,
  cart        jsonb   not null default '{}',
  updated_at  timestamptz not null default now(),
  unique(user_id)
);

-- 4. Row Level Security aktivieren
-- ============================================================
alter table public.products enable row level security;
alter table public.users    enable row level security;
alter table public.orders   enable row level security;

-- 5. Policies: products
-- ============================================================
create policy "products_select" on public.products for select using (true);
create policy "products_insert" on public.products for insert with check (true);
create policy "products_delete" on public.products for delete using (true);

-- 6. Policies: users
-- ============================================================
create policy "users_select" on public.users for select using (true);
create policy "users_insert" on public.users for insert with check (true);

-- 7. Policies: orders
-- ============================================================
create policy "orders_select" on public.orders for select using (true);
create policy "orders_insert" on public.orders for insert with check (true);
create policy "orders_update" on public.orders for update using (true);

-- ============================================================
-- Fertig!
-- ============================================================
