-- =====================================================================
-- SmarTree · Planejador de Propostas — estrutura do banco (Supabase)
-- Rode este arquivo inteiro no Supabase: SQL Editor › New query › Run.
-- Pode ser executado de novo sem perder dados.
-- =====================================================================

create extension if not exists pgcrypto;

-- ---------- tabelas ----------
create table if not exists public.perfis (
  id         uuid primary key references auth.users(id) on delete cascade,
  email      text,
  nome       text,
  empresa    text,
  admin      boolean not null default false,
  criado_em  timestamptz not null default now()
);

create table if not exists public.publishers (
  id            text primary key check (id ~ '^[a-z0-9-]{1,64}$'),
  dados         jsonb not null,           -- formatos, CPMs, segmentos, inventário do GAM, marca
  atualizado_em timestamptz not null default now()
);

create table if not exists public.acessos (
  user_id      uuid primary key references public.perfis(id) on delete cascade,
  publisher_id text not null references public.publishers(id) on delete cascade,
  concedido_em timestamptz not null default now()
);

create table if not exists public.propostas (
  id           uuid primary key default gen_random_uuid(),
  publisher_id text not null references public.publishers(id) on delete cascade,
  criado_por   uuid default auth.uid() references public.perfis(id) on delete set null,
  dados        jsonb not null,
  criado_em    timestamptz not null default now()
);
create index if not exists propostas_publisher_idx on public.propostas(publisher_id);

-- ---------- funções auxiliares ----------
create or replace function public.eh_admin() returns boolean
language sql stable security definer set search_path = public as $$
  select coalesce((select admin from public.perfis where id = auth.uid()), false);
$$;

create or replace function public.meu_publisher() returns text
language sql stable security definer set search_path = public as $$
  select publisher_id from public.acessos where user_id = auth.uid();
$$;

-- perfil criado automaticamente no primeiro login
create or replace function public.criar_perfil() returns trigger
language plpgsql security definer set search_path = public as $$
begin
  insert into public.perfis (id, email) values (new.id, new.email) on conflict (id) do nothing;
  return new;
end; $$;
drop trigger if exists ao_criar_usuario on auth.users;
create trigger ao_criar_usuario after insert on auth.users
  for each row execute function public.criar_perfil();

-- ninguém se torna admin sozinho (o SQL Editor, sem usuário logado, pode)
create or replace function public.proteger_admin() returns trigger
language plpgsql security definer set search_path = public as $$
begin
  if new.admin is distinct from old.admin and auth.uid() is not null and not public.eh_admin() then
    raise exception 'Só administradores podem alterar quem é admin';
  end if;
  return new;
end; $$;
drop trigger if exists proteger_admin on public.perfis;
create trigger proteger_admin before update on public.perfis
  for each row execute function public.proteger_admin();

-- ---------- regras de acesso (cada publisher só enxerga o que é dele) ----------
alter table public.perfis     enable row level security;
alter table public.publishers enable row level security;
alter table public.acessos    enable row level security;
alter table public.propostas  enable row level security;

drop policy if exists perfis_ler    on public.perfis;
drop policy if exists perfis_editar on public.perfis;
create policy perfis_ler    on public.perfis for select using (id = auth.uid() or public.eh_admin());
create policy perfis_editar on public.perfis for update using (id = auth.uid() or public.eh_admin())
                                                with check (id = auth.uid() or public.eh_admin());

drop policy if exists publishers_ler   on public.publishers;
drop policy if exists publishers_admin on public.publishers;
create policy publishers_ler   on public.publishers for select using (public.eh_admin() or id = public.meu_publisher());
create policy publishers_admin on public.publishers for all    using (public.eh_admin()) with check (public.eh_admin());

drop policy if exists acessos_ler   on public.acessos;
drop policy if exists acessos_admin on public.acessos;
create policy acessos_ler   on public.acessos for select using (user_id = auth.uid() or public.eh_admin());
create policy acessos_admin on public.acessos for all    using (public.eh_admin()) with check (public.eh_admin());

drop policy if exists propostas_ler     on public.propostas;
drop policy if exists propostas_criar   on public.propostas;
drop policy if exists propostas_editar  on public.propostas;
drop policy if exists propostas_excluir on public.propostas;
create policy propostas_ler     on public.propostas for select using (public.eh_admin() or publisher_id = public.meu_publisher());
create policy propostas_criar   on public.propostas for insert with check (public.eh_admin() or publisher_id = public.meu_publisher());
create policy propostas_editar  on public.propostas for update using (public.eh_admin() or publisher_id = public.meu_publisher())
                                                      with check (public.eh_admin() or publisher_id = public.meu_publisher());
create policy propostas_excluir on public.propostas for delete using (public.eh_admin() or publisher_id = public.meu_publisher());

-- ---------- logos dos publishers (leitura pública, envio só por admin) ----------
insert into storage.buckets (id, name, public) values ('logos', 'logos', true) on conflict (id) do nothing;
drop policy if exists logos_ler     on storage.objects;
drop policy if exists logos_inserir on storage.objects;
drop policy if exists logos_alterar on storage.objects;
drop policy if exists logos_excluir on storage.objects;
create policy logos_ler     on storage.objects for select using (bucket_id = 'logos');
create policy logos_inserir on storage.objects for insert with check (bucket_id = 'logos' and public.eh_admin());
create policy logos_alterar on storage.objects for update using (bucket_id = 'logos' and public.eh_admin());
create policy logos_excluir on storage.objects for delete using (bucket_id = 'logos' and public.eh_admin());

-- ---------- primeiro administrador ----------
-- Depois de entrar no site uma vez com o seu e-mail, rode (troque o e-mail):
-- update public.perfis set admin = true where email = 'seu@email.com';
