-- SmarTree · marca do publisher na tela de login (rode uma vez no SQL Editor do Supabase)
-- Devolve só nome e identidade visual; formatos, CPMs e inventário continuam protegidos.
create or replace function public.marca_publica(p_id text) returns jsonb
language sql stable security definer set search_path = public as $$
  select jsonb_build_object('nome', dados->>'nome', 'marca', coalesce(dados->'marca', '{}'::jsonb))
  from public.publishers where id = p_id;
$$;
grant execute on function public.marca_publica(text) to anon, authenticated;
