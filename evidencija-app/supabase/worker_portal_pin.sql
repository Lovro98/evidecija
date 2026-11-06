-- Radnički portal: PIN se sprema i čitljivo (vidi ga samo admin, RLS wp_select_admin),
-- da se ID + PIN mogu ispisati na obračunu radnika. Prijava i dalje provjerava bcrypt hash.
alter table public.worker_portal add column if not exists pin text;

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
    if found and not p_reset and ex.pin is not null then
      worker_id := wid; code := ex.code; pin := ex.pin; return next; continue;
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
    insert into worker_portal as wp (worker_id, code, pin_hash, pin, failed, locked_until, issued_by, issued_at)
      values (wid, c, crypt(p, gen_salt('bf', 8)), p, 0, null, auth.uid(), now())
      on conflict (worker_id) do update set pin_hash = excluded.pin_hash, pin = excluded.pin, failed = 0, locked_until = null,
        issued_by = excluded.issued_by, issued_at = now();
    worker_id := wid; code := c; pin := p; return next;
  end loop;
end $$;
revoke all on function public.portal_issue(uuid[], boolean) from public, anon;
grant execute on function public.portal_issue(uuid[], boolean) to authenticated;
