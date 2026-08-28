-- ============================================================
-- EVIDENCIJA RADA — Supabase shema (verzija 2)
-- Zaposlenik vidi SAMO ono što je sam unio; admin vidi SVE.
-- Zalijepi CIJELI sadržaj u Supabase → SQL Editor → Run.
-- Sigurno za ponovno pokretanje (može se izvršiti i preko stare verzije).
-- ============================================================

-- Profili korisnika
create table if not exists profiles (
  id uuid primary key references auth.users on delete cascade,
  name text not null default '',
  role text not null default 'employee' check (role in ('admin','employee')),
  created_at timestamptz default now()
);

create or replace function public.handle_new_user() returns trigger
language plpgsql security definer set search_path = public as $$
begin
  insert into public.profiles (id, name)
  values (new.id, coalesce(new.raw_user_meta_data->>'name', split_part(new.email,'@',1)))
  on conflict (id) do nothing;
  return new;
end $$;

drop trigger if exists on_auth_user_created on auth.users;
create trigger on_auth_user_created
  after insert on auth.users
  for each row execute function public.handle_new_user();

create or replace function public.is_admin() returns boolean
language sql stable security definer set search_path = public as $$
  select exists(select 1 from profiles where id = auth.uid() and role = 'admin');
$$;

-- Tablice
create table if not exists objects (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  created_by uuid references profiles,
  created_at timestamptz default now(),
  deleted_at timestamptz
);
alter table objects add column if not exists created_by uuid references profiles;

create table if not exists object_billing (
  object_id uuid primary key references objects on delete cascade,
  bill_rate numeric not null default 0
);

create table if not exists workers (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  phone text default '',
  base_rate numeric default 0,
  object_id uuid references objects,
  note text default '',
  archived boolean default false,
  archived_date date,
  created_by uuid references profiles,
  created_at timestamptz default now(),
  deleted_at timestamptz
);
alter table workers add column if not exists created_by uuid references profiles;

create table if not exists work_logs (
  id uuid primary key default gen_random_uuid(),
  worker_id uuid not null references workers on delete cascade,
  object_id uuid references objects,
  work_date date not null,
  from_t text default '',
  to_t text default '',
  hours numeric not null,
  monthly boolean default false,
  note text default '',
  created_by uuid references profiles,
  created_at timestamptz default now(),
  deleted_at timestamptz
);

create table if not exists payments (
  id uuid primary key default gen_random_uuid(),
  worker_id uuid references workers on delete cascade,
  pay_date date not null,
  type text not null check (type in ('avans','bonus','gorivo','ostalo')),
  amount numeric not null,
  note text default '',
  deduct boolean default true,
  created_by uuid references profiles,
  created_at timestamptz default now(),
  deleted_at timestamptz
);

-- Troškovi vezani za objekt (bez radnika): dozvoli prazan worker_id + dodaj object_id
alter table payments alter column worker_id drop not null;
alter table payments add column if not exists object_id uuid references objects on delete cascade;

-- Valuta isplate: EUR ili CZK (bez preračuna — vode se odvojeno)
alter table payments add column if not exists currency text default 'EUR';

-- Država objekta: HR ili CZ (za odvojene preglede)
alter table objects add column if not exists country text default 'HR';

-- Valuta satnice radnika i naplate objekta (EUR ili CZK)
alter table workers add column if not exists rate_currency text default 'EUR';
alter table object_billing add column if not exists bill_currency text default 'EUR';
alter table payouts add column if not exists amount_czk numeric default 0;

create table if not exists assignments (
  id uuid primary key default gen_random_uuid(),
  worker_id uuid not null references workers on delete cascade,
  object_id uuid references objects,
  from_date date not null,
  created_by uuid references profiles,
  created_at timestamptz default now(),
  deleted_at timestamptz
);

create table if not exists rate_changes (
  id uuid primary key default gen_random_uuid(),
  worker_id uuid not null references workers on delete cascade,
  rate numeric not null,
  from_date date not null,
  created_by uuid references profiles,
  created_at timestamptz default now(),
  deleted_at timestamptz
);

create table if not exists audit_log (
  id uuid primary key default gen_random_uuid(),
  at timestamptz default now(),
  user_id uuid references profiles,
  user_name text default '',
  action text not null
);

-- Dodjela objekata zaposlenicima (admin odredi tko vidi koji objekt)
create table if not exists object_members (
  object_id uuid references objects on delete cascade,
  user_id uuid references profiles on delete cascade,
  primary key (object_id, user_id)
);

-- Isplaćeni mjeseci (zaključavanje obračuna po radniku)
create table if not exists payouts (
  id uuid primary key default gen_random_uuid(),
  worker_id uuid not null references workers on delete cascade,
  month text not null,
  amount numeric not null default 0,
  paid_at date default current_date,
  created_by uuid references profiles,
  created_at timestamptz default now(),
  deleted_at timestamptz
);

-- Uplate od objekata/hotela (praćenje naplate) — samo admin
create table if not exists invoice_payments (
  id uuid primary key default gen_random_uuid(),
  object_id uuid not null references objects on delete cascade,
  month text not null,
  amount numeric not null,
  pay_date date default current_date,
  note text default '',
  created_at timestamptz default now(),
  deleted_at timestamptz
);

-- Podaci firme (za fakture/specifikacije) — samo admin
create table if not exists settings (
  id int primary key default 1 check (id = 1),
  company_name text default '',
  address text default '',
  oib text default '',
  iban text default ''
);

-- Istek dokumenata radnika
alter table workers add column if not exists permit_expiry date;
alter table workers add column if not exists contract_expiry date;

-- ============================================================
-- SIGURNOST: zaposlenik vidi samo SVOJE, admin vidi SVE
-- ============================================================
alter table profiles enable row level security;
alter table objects enable row level security;
alter table object_billing enable row level security;
alter table workers enable row level security;
alter table work_logs enable row level security;
alter table payments enable row level security;
alter table assignments enable row level security;
alter table rate_changes enable row level security;
alter table audit_log enable row level security;
alter table object_members enable row level security;
alter table payouts enable row level security;
alter table invoice_payments enable row level security;
alter table settings enable row level security;

-- Obriši stara pravila da ne dođe do sukoba (bezopasno ako ne postoje)
do $$ declare p record;
begin
  for p in
    select policyname, tablename from pg_policies
    where schemaname = 'public'
      and tablename in ('profiles','objects','object_billing','workers','work_logs','payments','assignments','rate_changes','audit_log','object_members','payouts','invoice_payments','settings')
  loop
    execute format('drop policy if exists %I on public.%I', p.policyname, p.tablename);
  end loop;
end $$;

-- Profili: vidiš sebe, admin vidi sve i mijenja uloge
create policy "profiles_select" on profiles for select to authenticated
  using (is_admin() or id = auth.uid());
create policy "profiles_admin_update" on profiles for update to authenticated
  using (is_admin());

-- Objekti: vidiš svoje + one koje ti je admin dodijelio; admin vidi sve
create policy "objects_select" on objects for select to authenticated
  using (
    is_admin()
    or created_by = auth.uid()
    or exists (select 1 from object_members m where m.object_id = objects.id and m.user_id = auth.uid())
  );
create policy "objects_insert" on objects for insert to authenticated
  with check (created_by = auth.uid());
create policy "objects_update" on objects for update to authenticated
  using (is_admin() or created_by = auth.uid());
create policy "objects_delete_admin" on objects for delete to authenticated
  using (is_admin());

-- Dodjele objekata: vidiš svoje dodjele, dodjeljuje i miče samo admin
create policy "om_select" on object_members for select to authenticated
  using (is_admin() or user_id = auth.uid());
create policy "om_insert_admin" on object_members for insert to authenticated
  with check (is_admin());
create policy "om_delete_admin" on object_members for delete to authenticated
  using (is_admin());

-- Isplaćeni mjeseci: vidiš svoje, admin sve; otključava (mijenja) samo admin
create policy "po_select" on payouts for select to authenticated
  using (is_admin() or created_by = auth.uid());
create policy "po_insert" on payouts for insert to authenticated
  with check (created_by = auth.uid());
create policy "po_update_admin" on payouts for update to authenticated
  using (is_admin());
create policy "po_delete_admin" on payouts for delete to authenticated
  using (is_admin());

-- Uplate od hotela: SAMO ADMIN
create policy "ip_admin_all" on invoice_payments for all to authenticated
  using (is_admin()) with check (is_admin());

-- Podaci firme: SAMO ADMIN
create policy "set_admin_all" on settings for all to authenticated
  using (is_admin()) with check (is_admin());

-- Naplata objekta: SAMO ADMIN
create policy "billing_admin_all" on object_billing for all to authenticated
  using (is_admin()) with check (is_admin());

-- Radnici: zaposlenik vidi svoje + sve radnike na objektima koji su mu dodijeljeni; admin sve
create policy "workers_select" on workers for select to authenticated
  using (
    is_admin()
    or created_by = auth.uid()
    or object_id in (select object_id from object_members where user_id = auth.uid())
    or id in (
      select worker_id from work_logs
      where deleted_at is null
        and object_id in (select object_id from object_members where user_id = auth.uid())
    )
  );
create policy "workers_insert" on workers for insert to authenticated
  with check (created_by = auth.uid());
create policy "workers_update" on workers for update to authenticated
  using (is_admin() or created_by = auth.uid());
create policy "workers_delete_admin" on workers for delete to authenticated
  using (is_admin());

-- Sati: zaposlenik vidi svoje + SVE upise na dodijeljenim objektima; admin sve
create policy "logs_select" on work_logs for select to authenticated
  using (
    is_admin()
    or created_by = auth.uid()
    or object_id in (select object_id from object_members where user_id = auth.uid())
  );
create policy "logs_insert" on work_logs for insert to authenticated
  with check (created_by = auth.uid());
create policy "logs_update" on work_logs for update to authenticated
  using (is_admin() or created_by = auth.uid());
create policy "logs_delete_admin" on work_logs for delete to authenticated
  using (is_admin());

-- Isplate
create policy "pay_select" on payments for select to authenticated
  using (is_admin() or created_by = auth.uid());
create policy "pay_insert" on payments for insert to authenticated
  with check (created_by = auth.uid());
create policy "pay_update" on payments for update to authenticated
  using (is_admin() or created_by = auth.uid());
create policy "pay_delete_admin" on payments for delete to authenticated
  using (is_admin());

-- Premještaji
create policy "asg_select" on assignments for select to authenticated
  using (is_admin() or created_by = auth.uid());
create policy "asg_insert" on assignments for insert to authenticated
  with check (created_by = auth.uid());
create policy "asg_update" on assignments for update to authenticated
  using (is_admin() or created_by = auth.uid());
create policy "asg_delete_admin" on assignments for delete to authenticated
  using (is_admin());

-- Promjene satnice
create policy "rc_select" on rate_changes for select to authenticated
  using (is_admin() or created_by = auth.uid());
create policy "rc_insert" on rate_changes for insert to authenticated
  with check (created_by = auth.uid());
create policy "rc_update" on rate_changes for update to authenticated
  using (is_admin() or created_by = auth.uid());
create policy "rc_delete_admin" on rate_changes for delete to authenticated
  using (is_admin());

-- Aktivnost: svatko upisuje svoje radnje, čita SAMO admin
create policy "audit_insert" on audit_log for insert to authenticated
  with check (user_id = auth.uid());
create policy "audit_select_admin" on audit_log for select to authenticated
  using (is_admin());

-- ============================================================
-- DOKUMENTI: zaposlenik vidi dokumente SAMO svojih radnika
-- ============================================================
insert into storage.buckets (id, name, public) values ('docs','docs', false)
  on conflict (id) do nothing;

drop policy if exists "docs_read" on storage.objects;
drop policy if exists "docs_write" on storage.objects;
drop policy if exists "docs_delete_admin" on storage.objects;

create policy "docs_read" on storage.objects for select to authenticated
  using (
    bucket_id = 'docs' and (
      is_admin() or exists (
        select 1 from public.workers w
        where w.id::text = (storage.foldername(name))[1] and w.created_by = auth.uid()
      )
    )
  );
create policy "docs_write" on storage.objects for insert to authenticated
  with check (
    bucket_id = 'docs' and (
      is_admin() or exists (
        select 1 from public.workers w
        where w.id::text = (storage.foldername(name))[1] and w.created_by = auth.uid()
      )
    )
  );
create policy "docs_delete_admin" on storage.objects for delete to authenticated
  using (bucket_id = 'docs' and is_admin());

-- ============================================================
-- MZDY NAPOMENE: automatski izračunata dodatna stavka iz češke platne liste
-- ============================================================
create table if not exists payroll_notes (
  id uuid primary key default gen_random_uuid(),
  worker_id uuid not null references workers on delete cascade,
  month text not null,
  amount numeric not null default 0,
  currency text not null default 'CZK',
  source_czk numeric not null default 0,
  note text default '',
  created_by uuid references profiles,
  created_at timestamptz default now(),
  deleted_at timestamptz
);
alter table payroll_notes enable row level security;
drop policy if exists "pn_admin_all" on payroll_notes;
create policy "pn_admin_all" on payroll_notes for all to authenticated
  using (is_admin()) with check (is_admin());

-- Tečaj CZK→EUR, koristi se SAMO za automatski izračun mzdy napomene iznad
alter table settings add column if not exists czk_rate numeric default 25;

-- ============================================================
-- GENERIRANE FAKTURE ZA OBJEKTE (izlazne fakture prema hotelima)
-- ============================================================
create table if not exists object_invoices (
  id uuid primary key default gen_random_uuid(),
  object_id uuid not null references objects on delete cascade,
  number text not null,
  period text not null,
  issue_date date not null default current_date,
  hours numeric not null default 0,
  rate numeric not null default 0,
  amount numeric not null default 0,
  currency text not null default 'EUR',
  note text default '',
  created_by uuid references profiles,
  created_at timestamptz default now(),
  deleted_at timestamptz
);
alter table object_invoices enable row level security;
drop policy if exists "oi_admin_all" on object_invoices;
create policy "oi_admin_all" on object_invoices for all to authenticated
  using (is_admin()) with check (is_admin());

-- ============================================================
-- FAKTURE OBJEKATA: vidi i uređuje samo admin
-- ============================================================
insert into storage.buckets (id, name, public) values ('invoices','invoices', false)
  on conflict (id) do nothing;

drop policy if exists "invoices_admin_all" on storage.objects;
create policy "invoices_admin_all" on storage.objects for all to authenticated
  using (bucket_id = 'invoices' and is_admin())
  with check (bucket_id = 'invoices' and is_admin());
