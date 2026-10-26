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

-- Pozicija radnika na objektu (npr. HSK, kuhinja, bar...)
alter table workers add column if not exists position text default '';

-- Isplata na račun (banka) — nova vrsta isplate, ostatak ide na kovertu (gotovina)
alter table payments drop constraint if exists payments_type_check;
alter table payments add constraint payments_type_check check (type in ('avans','bonus','gorivo','ostalo','racun'));

-- Naplata objekta po poziciji radnika (npr. HSK, kuhinja, bar plaćaju se drugačije)
-- mode='fixed' → objekt plaća fiksnu cijenu (rate) po satu za tu poziciju
-- mode='markup' → objekt plaća radnikovu satnicu + markup po satu za tu poziciju
create table if not exists position_billing (
  object_id uuid not null references objects on delete cascade,
  position text not null,
  mode text not null default 'fixed' check (mode in ('fixed','markup')),
  rate numeric not null default 0,
  markup numeric not null default 0,
  currency text not null default 'EUR',
  created_by uuid references profiles,
  created_at timestamptz default now(),
  primary key (object_id, position)
);
alter table position_billing enable row level security;
drop policy if exists "posb_admin_all" on position_billing;
create policy "posb_admin_all" on position_billing for all to authenticated
  using (is_admin()) with check (is_admin());

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

-- Dodatak na satnicu vikendom / praznikom, u % (npr. 25 = +25%). 0 = bez dodatka.
alter table settings add column if not exists weekend_pct numeric default 0;
alter table settings add column if not exists holiday_pct numeric default 0;

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

-- ============================================================
-- ZATVARANJE OBJEKATA (arhiva, kao kod radnika) — bez brisanja povijesti
-- ============================================================
alter table objects add column if not exists archived boolean default false;
alter table objects add column if not exists archived_date date;

-- Koliko dana unaprijed upozoravamo na istek dokumenata radnika (podesivo)
alter table settings add column if not exists expiry_warn_days numeric default 30;

-- ============================================================
-- PODSJETNICI: slobodni tekstualni podsjetnici s datumom — samo admin
-- ============================================================
create table if not exists reminders (
  id uuid primary key default gen_random_uuid(),
  text text not null,
  due_date date,
  done boolean default false,
  created_by uuid references profiles,
  created_at timestamptz default now(),
  deleted_at timestamptz
);
alter table reminders enable row level security;
drop policy if exists "rem_admin_all" on reminders;
create policy "rem_admin_all" on reminders for all to authenticated
  using (is_admin()) with check (is_admin());

-- ============================================================
-- MJESEČNI PRORAČUN PO OBJEKTU (upozorenje kad se troškovi približe/premaše)
-- ============================================================
alter table object_billing add column if not exists monthly_budget numeric default 0;

-- ============================================================
-- ULOGA "manager" — voditelj objekta (dodatno uz admin/employee)
-- vidi svoje dodijeljene objekte kao i employee (postojeća pravila), ali smije
-- i uređivati naplatu (object_billing, position_billing) SAMO za te objekte.
-- ============================================================
alter table profiles drop constraint if exists profiles_role_check;
alter table profiles add constraint profiles_role_check check (role in ('admin','employee','manager'));

drop policy if exists "billing_manager" on object_billing;
create policy "billing_manager" on object_billing for all to authenticated
  using (exists (
    select 1 from object_members m join profiles p on p.id = auth.uid()
    where m.object_id = object_billing.object_id and m.user_id = auth.uid() and p.role = 'manager'
  ))
  with check (exists (
    select 1 from object_members m join profiles p on p.id = auth.uid()
    where m.object_id = object_billing.object_id and m.user_id = auth.uid() and p.role = 'manager'
  ));

drop policy if exists "posb_manager" on position_billing;
create policy "posb_manager" on position_billing for all to authenticated
  using (exists (
    select 1 from object_members m join profiles p on p.id = auth.uid()
    where m.object_id = position_billing.object_id and m.user_id = auth.uid() and p.role = 'manager'
  ))
  with check (exists (
    select 1 from object_members m join profiles p on p.id = auth.uid()
    where m.object_id = position_billing.object_id and m.user_id = auth.uid() and p.role = 'manager'
  ));

-- ============================================================
-- ODOBRAVANJE ISPLATA U DVA KORAKA: "predloženo" pa tek onda "odobreno"
-- ============================================================
alter table payouts add column if not exists approved boolean default false;
alter table payouts add column if not exists approved_by uuid references profiles;
alter table payouts add column if not exists approved_at timestamptz;

-- ============================================================
-- POVIJEST PRIJAVA — jednostavan sigurnosni log (tko/kad se prijavio)
-- ============================================================
create table if not exists login_log (
  id uuid primary key default gen_random_uuid(),
  user_id uuid references profiles,
  user_name text default '',
  at timestamptz default now(),
  user_agent text default ''
);
alter table login_log enable row level security;
drop policy if exists "login_insert_self" on login_log;
create policy "login_insert_self" on login_log for insert to authenticated
  with check (user_id = auth.uid());
drop policy if exists "login_select_admin" on login_log;
create policy "login_select_admin" on login_log for select to authenticated
  using (is_admin());

-- ============================================================
-- GREŠKE APLIKACIJE — osnovno bilježenje (bez vanjskog servisa)
-- ============================================================
create table if not exists error_log (
  id uuid primary key default gen_random_uuid(),
  user_id uuid references profiles,
  user_name text default '',
  message text not null,
  stack text default '',
  url text default '',
  at timestamptz default now()
);
alter table error_log enable row level security;
drop policy if exists "errlog_insert_self" on error_log;
create policy "errlog_insert_self" on error_log for insert to authenticated
  with check (user_id = auth.uid());
drop policy if exists "errlog_select_admin" on error_log;
create policy "errlog_select_admin" on error_log for select to authenticated
  using (is_admin());

-- ============================================================
-- PROVIZIJA ZAPOSLENIKA PO SATU (npr. Ivan dobiva 0.5 €/h za svaki sat koji
-- su radnici odradili na objektima koji su mu dodijeljeni preko object_members)
-- ============================================================
alter table profiles add column if not exists hourly_commission numeric default 0;
alter table profiles add column if not exists commission_currency text default 'EUR';

-- Posebna provizija po satu za pojedini objekt (nadjačava zadanu gore, samo za taj objekt)
create table if not exists commission_rates (
  object_id uuid not null references objects on delete cascade,
  user_id uuid not null references profiles on delete cascade,
  rate numeric not null default 0,
  currency text not null default 'EUR',
  created_by uuid references profiles,
  created_at timestamptz default now(),
  primary key (object_id, user_id)
);
alter table commission_rates enable row level security;
drop policy if exists "commrate_admin_all" on commission_rates;
create policy "commrate_admin_all" on commission_rates for all to authenticated
  using (is_admin()) with check (is_admin());

-- Isplate provizije zaposleniku (koliko mu je stvarno dano) — oduzima se od zarađene provizije
create table if not exists commission_payouts (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null references profiles on delete cascade,
  pay_date date not null default current_date,
  amount numeric not null default 0,
  currency text not null default 'EUR',
  note text default '',
  created_by uuid references profiles,
  created_at timestamptz default now(),
  deleted_at timestamptz
);
alter table commission_payouts enable row level security;
drop policy if exists "commpay_admin_all" on commission_payouts;
create policy "commpay_admin_all" on commission_payouts for all to authenticated
  using (is_admin()) with check (is_admin());

-- ============================================================
-- ČIŠĆENJE PODATAKA: blokiraj prazno ime radnika ubuduće
-- (NOT VALID = ne provjerava postojeće retke, samo nove upise/izmjene —
-- sigurno i ako već postoji stari radnik s praznim imenom)
-- ============================================================
alter table workers drop constraint if exists workers_name_not_blank;
alter table workers add constraint workers_name_not_blank check (btrim(name) <> '') not valid;

-- Dan u mjesecu kad se isplaćuje plaća + koliko dana prije podsjetiti
alter table settings add column if not exists payday_day numeric default 5;
alter table settings add column if not exists payday_reminder_days numeric default 3;

-- ============================================================
-- IME/PREZIME: novi redoslijed je "Prezime Ime" (umjesto "Ime Prezime").
-- Ovaj flag pamti je li već napravljena jednokratna zamjena redoslijeda
-- postojećih radnika (gumb "Zamijeni svima" u Radnici tabu) — nakon toga
-- se ta ponuda više ne prikazuje (da se netko slučajno ne zamijeni dvaput).
-- ============================================================
alter table settings add column if not exists name_order_migrated boolean default false;

-- ============================================================
-- BRIGADA (ČEŠKA): kad na hotel pošalješ dodatnu osobu "na dan" i odmah je
-- platiš u cashu — bez imena, samo datum/objekt/sati/iznos. Ulazi i u naplatu
-- objekta (sati se naplaćuju hotelu po zadanoj cijeni objekta) i u trošak
-- (isplaćeni cash se oduzima od dobiti) — samo admin.
-- ============================================================
create table if not exists brigade_costs (
  id uuid primary key default gen_random_uuid(),
  object_id uuid not null references objects on delete cascade,
  work_date date not null default current_date,
  hours numeric not null default 0,
  amount numeric not null default 0,
  currency text not null default 'CZK',
  note text default '',
  created_by uuid references profiles,
  created_at timestamptz default now(),
  deleted_at timestamptz
);
alter table brigade_costs enable row level security;
drop policy if exists "brigade_admin_all" on brigade_costs;
create policy "brigade_admin_all" on brigade_costs for all to authenticated
  using (is_admin()) with check (is_admin());

-- Posebna cijena naplate za baš tu brigadu (nadjačava zadanu cijenu objekta gore) — 0 = koristi zadanu
alter table brigade_costs add column if not exists bill_rate numeric default 0;
alter table brigade_costs add column if not exists bill_currency text default 'EUR';

-- Brigada: broj ljudi, imena (opcionalno) i tko je platio — null = admin, inače koordinator
-- (profil) koji je platio iz gotovine koju mu je admin dao (vidi cash_advances)
alter table brigade_costs add column if not exists people integer default 1;
alter table brigade_costs add column if not exists names text default '';
alter table brigade_costs add column if not exists paid_by uuid references profiles;

-- GOTOVINA KOD KOORDINATORA: admin mu da novac (given) ili mu ovaj vrati višak (returned).
-- Saldo = dano − vraćeno − brigade koje je on platio (brigade_costs.paid_by)
create table if not exists cash_advances (
  id uuid primary key default gen_random_uuid(),
  profile_id uuid not null references profiles on delete cascade,
  kind text not null default 'given' check (kind in ('given', 'returned')),
  amount numeric not null default 0,
  currency text not null default 'CZK',
  pay_date date not null default current_date,
  note text default '',
  created_by uuid references profiles,
  created_at timestamptz default now(),
  deleted_at timestamptz
);
alter table cash_advances enable row level security;
drop policy if exists "cash_advances_admin_all" on cash_advances;
create policy "cash_advances_admin_all" on cash_advances for all to authenticated
  using (is_admin()) with check (is_admin());

-- ============================================================
-- PDV NA IZLAZNIM FAKTURAMA (hotelima) + PDF faktura spremljen uz zapis
-- vat_amount se upisuje ručno (onoliko koliko admin vidi na stvarnoj fakturi/PDF-u)
-- ============================================================
alter table object_invoices add column if not exists vat_amount numeric default 0;
alter table object_invoices add column if not exists vat_currency text default 'EUR';
alter table object_invoices add column if not exists file_path text default '';

-- Naziv kupca kako stoji na PDF fakturi ("Odběratel") — pamti se po objektu da se
-- kod sljedećeg uvoza PDF faktura automatski prepozna kojem objektu pripada
alter table objects add column if not exists invoice_customer_name text default '';

-- ============================================================
-- FINANCIJE (samo admin): bankovni izvodi, kategorije, pravila za
-- automatsko razvrstavanje. Iz toga se računa mjesečni pregled i
-- procjena poreza na dobit / dividende za českú s.r.o.
-- ============================================================
create table if not exists bank_imports (
  id uuid primary key default gen_random_uuid(),
  file_name text default '',
  date_from date,
  date_to date,
  row_count int default 0,
  created_by uuid references profiles,
  created_at timestamptz default now()
);

create table if not exists bank_transactions (
  id uuid primary key default gen_random_uuid(),
  import_id uuid references bank_imports on delete cascade, -- null = ručno upisano (npr. gotovina)
  tx_date date not null,
  amount numeric not null,                 -- + uplata, − isplata
  currency text not null default 'CZK',
  counterparty text default '',
  counter_account text default '',
  vs text default '',
  message text default '',
  category text default '',                -- '' = nerazvrstano
  category_auto boolean default false,     -- true = predložila pravila, admin još nije potvrdio
  vat_rate numeric,                        -- null = zadana stopa kategorije
  note text default '',
  hash text not null unique,               -- sprječava dvostruki uvoz istog izvoda
  created_by uuid references profiles,
  created_at timestamptz default now()
);
create index if not exists bank_transactions_date_idx on bank_transactions (tx_date);

create table if not exists bank_rules (
  id uuid primary key default gen_random_uuid(),
  pattern text not null,                   -- dio imena partnera, broja računa ili poruke
  category text not null,
  created_by uuid references profiles,
  created_at timestamptz default now()
);

alter table bank_imports enable row level security;
alter table bank_transactions enable row level security;
alter table bank_rules enable row level security;
drop policy if exists "bimp_admin_all" on bank_imports;
create policy "bimp_admin_all" on bank_imports for all to authenticated
  using (is_admin()) with check (is_admin());
drop policy if exists "btx_admin_all" on bank_transactions;
create policy "btx_admin_all" on bank_transactions for all to authenticated
  using (is_admin()) with check (is_admin());
drop policy if exists "brule_admin_all" on bank_rules;
create policy "brule_admin_all" on bank_rules for all to authenticated
  using (is_admin()) with check (is_admin());

-- ============================================================
-- KANDIDATI — javna prijavnica /prijava
-- Javni unos SAMO kroz funkciju submit_application (provjera + limit protiv spama).
-- Čitaju i mijenjaju admin i regruteri (profiles.can_recruit), briše samo admin.
-- ============================================================
create table if not exists public.candidates (
  id uuid primary key default gen_random_uuid(),
  created_at timestamptz not null default now(),
  full_name text not null check (char_length(full_name) between 2 and 120),
  phone text not null check (char_length(phone) between 6 and 30),
  messenger text not null default '' check (char_length(messenger) <= 20),
  positions text[] not null default '{}' check (cardinality(positions) <= 12),
  experience text not null default '' check (char_length(experience) <= 1000),
  residence text not null default '' check (char_length(residence) <= 40),
  city text not null default '' check (char_length(city) <= 80),
  start_from text not null default '' check (char_length(start_from) <= 40),
  heard_from text not null default '' check (char_length(heard_from) <= 40),
  referrer text not null default '' check (char_length(referrer) <= 120),
  source_tag text not null default '' check (char_length(source_tag) <= 60),
  lang text not null default 'UK' check (char_length(lang) <= 5),
  status text not null default 'novi' check (status in ('novi','kontaktiran','razgovor','poceo','nije_prikladan','nedostupan')),
  object_id uuid references public.objects on delete set null,
  admin_note text not null default '' check (char_length(admin_note) <= 2000),
  status_changed_at timestamptz,
  updated_by uuid references public.profiles on delete set null
);
create index if not exists candidates_created_idx on public.candidates (created_at desc);
create index if not exists candidates_phone_idx on public.candidates (phone);
alter table public.candidates enable row level security;
-- regruter (profiles.can_recruit) vidi i obrađuje prijave; briše samo admin
alter table public.profiles add column if not exists can_recruit boolean not null default false;
create or replace function public.can_recruit() returns boolean
language sql stable security definer set search_path = public as $$
  select exists(select 1 from profiles where id = auth.uid() and (role = 'admin' or can_recruit));
$$;
drop policy if exists "cand_admin_select" on public.candidates;
drop policy if exists "cand_admin_update" on public.candidates;
drop policy if exists "cand_recruit_select" on public.candidates;
drop policy if exists "cand_recruit_update" on public.candidates;
drop policy if exists "cand_admin_delete" on public.candidates;
create policy "cand_recruit_select" on public.candidates for select to authenticated using (public.can_recruit());
create policy "cand_recruit_update" on public.candidates for update to authenticated using (public.can_recruit()) with check (public.can_recruit());
create policy "cand_admin_delete" on public.candidates for delete to authenticated using (public.is_admin());
revoke all on public.candidates from anon;

create or replace function public.submit_application(
  p_full_name text, p_phone text, p_messenger text, p_positions text[],
  p_experience text, p_residence text, p_city text, p_start_from text,
  p_heard_from text, p_referrer text, p_source_tag text, p_lang text
) returns boolean
language plpgsql security definer set search_path = public as $$
declare
  v_phone text := regexp_replace(coalesce(p_phone,''), '[^0-9+]', '', 'g');
  v_name text := btrim(coalesce(p_full_name,''));
begin
  if char_length(v_name) < 2 or char_length(v_name) > 120 then raise exception 'invalid_name'; end if;
  if char_length(v_phone) < 6 or char_length(v_phone) > 20 then raise exception 'invalid_phone'; end if;
  if (select count(*) from candidates where phone = v_phone and created_at > now() - interval '24 hours') >= 3 then
    return true;
  end if;
  if (select count(*) from candidates where created_at > now() - interval '1 hour') >= 60 then
    raise exception 'too_many';
  end if;
  insert into candidates (full_name, phone, messenger, positions, experience, residence, city, start_from, heard_from, referrer, source_tag, lang)
  values (
    v_name, v_phone,
    left(btrim(coalesce(p_messenger,'')), 20),
    coalesce((select array_agg(left(x, 40)) from (select unnest(p_positions) x limit 12) s), '{}'),
    left(btrim(coalesce(p_experience,'')), 1000),
    left(btrim(coalesce(p_residence,'')), 40),
    left(btrim(coalesce(p_city,'')), 80),
    left(btrim(coalesce(p_start_from,'')), 40),
    left(btrim(coalesce(p_heard_from,'')), 40),
    left(btrim(coalesce(p_referrer,'')), 120),
    left(btrim(coalesce(p_source_tag,'')), 60),
    left(coalesce(p_lang,'UK'), 5)
  );
  return true;
end $$;
revoke all on function public.submit_application(text,text,text,text[],text,text,text,text,text,text,text,text) from public;
grant execute on function public.submit_application(text,text,text,text[],text,text,text,text,text,text,text,text) to anon, authenticated;
