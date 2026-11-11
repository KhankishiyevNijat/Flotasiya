-- Shared shift log for the Flotation Calculator.
-- Run this once in the Supabase SQL editor. Replace CHANGE-ME with your own shared code.

create table if not exists public.shift_log (
  id text primary key,
  data jsonb not null,
  updated_at timestamptz not null default now()
);
create table if not exists public.shift_secret (
  id int primary key default 1,
  code text not null
);

-- No direct access to the tables: everything goes through the two functions below.
alter table public.shift_log enable row level security;
alter table public.shift_secret enable row level security;

insert into public.shift_secret (id, code) values (1, 'CHANGE-ME')
on conflict (id) do update set code = excluded.code;

create or replace function public.shift_pull(p_code text)
returns setof public.shift_log
language plpgsql security definer set search_path = public as $$
begin
  if not exists (select 1 from shift_secret where code = p_code) then
    raise exception 'bad code';
  end if;
  return query select * from shift_log;
end $$;

create or replace function public.shift_push(p_code text, p_entries jsonb, p_deleted text[])
returns void
language plpgsql security definer set search_path = public as $$
begin
  if not exists (select 1 from shift_secret where code = p_code) then
    raise exception 'bad code';
  end if;
  insert into shift_log (id, data)
  select e->>'id', e->'data'
  from jsonb_array_elements(coalesce(p_entries, '[]'::jsonb)) e
  where e->>'id' is not null
  on conflict (id) do update set data = excluded.data, updated_at = now();
  delete from shift_log where id = any(coalesce(p_deleted, '{}'::text[]));
end $$;

revoke all on function public.shift_pull(text) from public;
revoke all on function public.shift_push(text, jsonb, text[]) from public;
grant execute on function public.shift_pull(text) to anon, authenticated;
grant execute on function public.shift_push(text, jsonb, text[]) to anon, authenticated;
