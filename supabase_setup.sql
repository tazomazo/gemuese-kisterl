-- ============================================================
-- Andreas' Gemüsekisterl – Supabase Setup (v3)
-- Dieses Script im Supabase SQL Editor ausführen
-- ============================================================

-- ⚠ MIGRATION (bestehende Datenbank, kein Datenverlust):
--
--   alter table public.products
--     add column category_id bigint references public.categories(id) on delete set null;
--   update public.products p set category_id = c.id
--     from public.categories c where c.name = p.category;
--   alter table public.products drop column category;
--   create policy "products_update" on public.products for update using (true);
--
-- Erst danach dieses Script NICHT mehr ausführen (würde Daten löschen).
-- ============================================================

-- Bestehende Tabellen löschen (Reihenfolge beachten!)
drop table if exists public.orders    cascade;
drop table if exists public.products  cascade;
drop table if exists public.categories cascade;
drop table if exists public.users     cascade;

-- 1. Tabelle: categories (muss vor products angelegt werden)
-- ============================================================
create table public.categories (
  id         bigint generated always as identity primary key,
  name       text    not null unique,
  sort_order int     not null default 0
);

-- 2. Tabelle: products (referenziert categories)
-- ============================================================
create table public.products (
  id          bigint generated always as identity primary key,
  name        text      not null,
  price       numeric   not null,
  unit        text      not null default '',
  category_id bigint    references public.categories(id) on delete set null,
  special     boolean   not null default false,
  created_at  timestamptz not null default now()
);

-- 3. Tabelle: users
-- ============================================================
create table public.users (
  id           bigint generated always as identity primary key,
  name         text    not null unique,
  password     text    not null default '',
  created_at   timestamptz not null default now()
);

-- 4. Tabelle: orders
-- ============================================================
create table public.orders (
  id          bigint generated always as identity primary key,
  user_id     bigint  not null references public.users(id) on delete cascade,
  cart        jsonb   not null default '{}',
  updated_at  timestamptz not null default now(),
  unique(user_id)
);

-- 5. Row Level Security aktivieren
-- ============================================================
alter table public.categories enable row level security;
alter table public.products   enable row level security;
alter table public.users      enable row level security;
alter table public.orders     enable row level security;

-- 6. Policies: categories
-- ============================================================
create policy "categories_select" on public.categories for select using (true);
create policy "categories_insert" on public.categories for insert with check (true);
create policy "categories_update" on public.categories for update using (true);
create policy "categories_delete" on public.categories for delete using (true);

-- 7. Policies: products
-- ============================================================
create policy "products_select" on public.products for select using (true);
create policy "products_insert" on public.products for insert with check (true);
create policy "products_update" on public.products for update using (true);
create policy "products_delete" on public.products for delete using (true);

-- 8. Policies: users
-- ============================================================
create policy "users_select" on public.users for select using (true);
create policy "users_insert" on public.users for insert with check (true);

-- 9. Policies: orders
-- ============================================================
create policy "orders_select" on public.orders for select using (true);
create policy "orders_insert" on public.orders for insert with check (true);
create policy "orders_update" on public.orders for update using (true);

-- ============================================================
-- Fertig!
-- ============================================================
