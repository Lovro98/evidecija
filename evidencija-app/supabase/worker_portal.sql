-- Radnički portal (/moje): radnik se prijavi s ID + PIN i vidi svoje isplaćene obračune.
-- Radnici NEMAJU Supabase račun: prijavu provjerava edge funkcija "worker-portal" (service_role)
-- preko portal_login. PIN je u bazi samo kao bcrypt hash.

create table if not exists public.worker_portal (
  worker_id uuid primary key references public.workers(id) on delete cascade,
  code text not null unique,
  pin_hash text not null,
  failed int not null default 0,
  locked_until timestamptz,
  last_login timestamptz,
  issued_by uuid,
  issued_at timestamptz not null default now()
);
alter table public.worker_portal enable row level security;
create policy wp_select_admin on public.worker_portal for select to authenticated using ((select is_admin()));
revoke insert, update, delete on public.worker_portal from anon, authenticated;

-- obračun kakav je bio kad je isplata odobrena (HTML, oznake u ⟦…⟧ prevodi stranica /moje)
create table if not exists public.payslip_snapshots (
  worker_id uuid not null references public.workers(id) on delete cascade,
  month text not null,
  html text not null,
  net text,
  updated_by uuid,
  updated_at timestamptz not null default now(),
  primary key (worker_id, month)
);
alter table public.payslip_snapshots enable row level security;
create policy ps_select_admin on public.payslip_snapshots for select to authenticated using ((select is_admin()));
create policy ps_insert_admin on public.payslip_snapshots for insert to authenticated with check ((select is_admin()));
create policy ps_update_admin on public.payslip_snapshots for update to authenticated using ((select is_admin())) with check ((select is_admin()));
create policy ps_delete_admin on public.payslip_snapshots for delete to authenticated using ((select is_admin()));

-- admin izdaje pristup (ID + novi PIN). PIN se vraća samo ovaj put.
-- p_reset = false: radnik koji već ima pristup dobije samo svoj ID (pin = null)
create or replace function public.portal_issue(p_ids uuid[], p_reset boolean default false)
returns table (worker_id uuid, code text, pin text)
language plpgsql security definer set search_path = public, extensions as $$
#variable_conflict use_column
declare
  wid uuid; c text; p text; i int; ex record;
  alphabet constant text := 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
begin
  if not is_admin() then raise exception 'not_admin'; end if;
  foreach wid in array p_ids loop
    if not exists (select 1 from workers w where w.id = wid) then continue; end if;
    select * into ex from worker_portal w where w.worker_id = wid;
    if found and not p_reset then
      worker_id := wid; code := ex.code; pin := null; return next; continue;
    end if;
    c := case when found then ex.code else null end;
    if c is null then
      loop
        c := '';
        for i in 1..6 loop c := c || substr(alphabet, 1 + (get_byte(gen_random_bytes(1), 0) % 32), 1); end loop;
        c := substr(c, 1, 3) || '-' || substr(c, 4, 3);
        exit when not exists (select 1 from worker_portal w where w.code = c);
      end loop;
    end if;
    p := lpad(((('x' || encode(gen_random_bytes(4), 'hex'))::bit(32)::bigint) % 1000000)::text, 6, '0');
    insert into worker_portal as wp (worker_id, code, pin_hash, failed, locked_until, issued_by, issued_at)
      values (wid, c, crypt(p, gen_salt('bf', 8)), 0, null, auth.uid(), now())
      on conflict (worker_id) do update set pin_hash = excluded.pin_hash, failed = 0, locked_until = null,
        issued_by = excluded.issued_by, issued_at = now();
    worker_id := wid; code := c; pin := p; return next;
  end loop;
end $$;
revoke all on function public.portal_issue(uuid[], boolean) from public, anon;
grant execute on function public.portal_issue(uuid[], boolean) to authenticated;

create or replace function public.portal_revoke(p_id uuid)
returns void language plpgsql security definer set search_path = public as $$
begin
  if not is_admin() then raise exception 'not_admin'; end if;
  delete from worker_portal where worker_id = p_id;
end $$;
revoke all on function public.portal_revoke(uuid) from public, anon;
grant execute on function public.portal_revoke(uuid) to authenticated;

-- provjera prijave — SAMO za edge funkciju (service_role).
-- 5 krivih PIN-ova → 15 min zaključano, 10+ → 24 h
create or replace function public.portal_login(p_code text, p_pin text)
returns jsonb language plpgsql security definer set search_path = public, extensions as $$
declare r record;
begin
  select * into r from worker_portal where code = upper(trim(coalesce(p_code, '')));
  if not found then return jsonb_build_object('ok', false, 'reason', 'bad'); end if;
  if r.locked_until is not null and r.locked_until > now() then
    return jsonb_build_object('ok', false, 'reason', 'locked', 'until', r.locked_until);
  end if;
  if crypt(coalesce(p_pin, ''), r.pin_hash) = r.pin_hash then
    update worker_portal set failed = 0, locked_until = null, last_login = now() where worker_id = r.worker_id;
    return jsonb_build_object('ok', true, 'worker_id', r.worker_id);
  end if;
  update worker_portal set failed = r.failed + 1,
    locked_until = case when r.failed + 1 >= 10 then now() + interval '24 hours'
                        when r.failed + 1 >= 5 then now() + interval '15 minutes' else null end
    where worker_id = r.worker_id;
  return jsonb_build_object('ok', false, 'reason', 'bad');
end $$;
revoke all on function public.portal_login(text, text) from public, anon, authenticated;
grant execute on function public.portal_login(text, text) to service_role;
