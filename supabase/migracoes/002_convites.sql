-- SmarTree · convites de colaboradores (rode uma vez no SQL Editor do Supabase)
create table if not exists public.convites (
  email         text primary key check (email = lower(email)),
  publisher_id  text not null references public.publishers(id) on delete cascade,
  convidado_por uuid default auth.uid() references public.perfis(id) on delete set null,
  criado_em     timestamptz not null default now()
);
alter table public.convites enable row level security;
drop policy if exists convites_admin on public.convites;
create policy convites_admin on public.convites for all using (public.eh_admin()) with check (public.eh_admin());
grant select, insert, update, delete on public.convites to authenticated;

-- perfil criado automaticamente no primeiro login (e acesso liberado se houver convite)
create or replace function public.criar_perfil() returns trigger
language plpgsql security definer set search_path = public as $$
declare v_pub text;
begin
  insert into public.perfis (id, email) values (new.id, new.email) on conflict (id) do nothing;
  select publisher_id into v_pub from public.convites where email = lower(new.email);
  if v_pub is not null then
    insert into public.acessos (user_id, publisher_id) values (new.id, v_pub)
      on conflict (user_id) do update set publisher_id = excluded.publisher_id, concedido_em = now();
    delete from public.convites where email = lower(new.email);
  end if;
  return new;
end; $$;
