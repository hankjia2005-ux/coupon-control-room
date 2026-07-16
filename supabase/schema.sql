-- Shared current snapshot for the coupon-control dashboard.
-- Run this once in Supabase SQL Editor.

create table if not exists public.coupon_current (
  id bigint primary key check (id = 1),
  rows jsonb not null default '[]'::jsonb,
  source_name text not null default 'initial-data',
  updated_at timestamptz not null default now(),
  updated_by text
);

alter table public.coupon_current enable row level security;

drop policy if exists "coupon public read" on public.coupon_current;
drop policy if exists "coupon authenticated insert" on public.coupon_current;
drop policy if exists "coupon authenticated update" on public.coupon_current;

create policy "coupon public read"
on public.coupon_current for select
to anon, authenticated
using (true);

create policy "coupon authenticated insert"
on public.coupon_current for insert
to authenticated
with check (true);

create policy "coupon authenticated update"
on public.coupon_current for update
to authenticated
using (true)
with check (true);

grant select on public.coupon_current to anon, authenticated;
grant insert, update on public.coupon_current to authenticated;

create or replace function public.update_coupon_status(
  p_id text,
  p_usable boolean,
  p_updated_by text
)
returns jsonb
language plpgsql
security invoker
set search_path = public
as $$
declare
  result jsonb;
begin
  if auth.uid() is null then
    raise exception 'authentication required';
  end if;

  update public.coupon_current
  set rows = coalesce((
    select jsonb_agg(
      case
        when value->>'id' = p_id then value || jsonb_build_object(
          'usable', p_usable,
          'statusSource', 'manual',
          'updatedBy', p_updated_by,
          'updatedAt', now()::text
        )
        else value
      end
    )
    from jsonb_array_elements(rows) as item(value)
  ), '[]'::jsonb),
  updated_at = now(),
  updated_by = p_updated_by
  where id = 1
  returning jsonb_build_object(
    'id', id,
    'rows', rows,
    'source_name', source_name,
    'updated_at', updated_at
  ) into result;

  return coalesce(result, '{}'::jsonb);
end;
$$;

grant execute on function public.update_coupon_status(text, boolean, text) to authenticated;

alter publication supabase_realtime add table public.coupon_current;
