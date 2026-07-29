-- ============================================================
-- GeoAcervo — Schema do banco de dados (Supabase / PostgreSQL)
-- Versão com autenticação real via Supabase Auth
-- ============================================================
-- Como usar:
-- 1. Crie um projeto gratuito em https://supabase.com
-- 2. Abra "SQL Editor" no painel do seu projeto
-- 3. Cole todo este arquivo e clique em "Run"
-- 4. Veja a seção "PRIMEIRO ADMINISTRADOR" no final — é um passo manual
--    obrigatório, feito uma única vez.
-- ============================================================

-- ---------------------------------------------------------------
-- Tabela de dados do acervo (minerais, rochas, histórico)
-- ---------------------------------------------------------------
create table if not exists public.kv_store (
  key text primary key,
  value text not null,
  updated_at timestamptz not null default now()
);

create index if not exists kv_store_key_prefix_idx on public.kv_store (key text_pattern_ops);

-- ---------------------------------------------------------------
-- Tabela de perfis de usuário (ligada ao Supabase Auth)
-- ---------------------------------------------------------------
create table if not exists public.profiles (
  id uuid primary key references auth.users (id) on delete cascade,
  nome text not null,
  email text not null,
  role text not null default 'Usuário' check (role in ('Usuário', 'Administrador')),
  status text not null default 'Pendente' check (status in ('Pendente', 'Aprovado', 'Recusado')),
  created_at timestamptz not null default now()
);

-- Função auxiliar: o usuário logado é um Administrador aprovado?
create or replace function public.is_admin()
returns boolean
language sql
security definer
set search_path = public
as $$
  select exists (
    select 1 from public.profiles
    where id = auth.uid() and role = 'Administrador' and status = 'Aprovado'
  );
$$;

-- Função auxiliar: o usuário logado está aprovado (qualquer perfil)?
create or replace function public.is_aprovado()
returns boolean
language sql
security definer
set search_path = public
as $$
  select exists (
    select 1 from public.profiles
    where id = auth.uid() and status = 'Aprovado'
  );
$$;

-- ---------------------------------------------------------------
-- Gatilho: cria automaticamente a linha em "profiles" sempre que
-- alguém se cadastra pelo Supabase Auth (roda com privilégio elevado,
-- então funciona mesmo com confirmação de e-mail ativada).
-- ---------------------------------------------------------------
create or replace function public.handle_new_user()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
begin
  insert into public.profiles (id, nome, email, role, status)
  values (
    new.id,
    coalesce(new.raw_user_meta_data->>'nome', new.email),
    new.email,
    'Usuário',
    'Pendente'
  );
  return new;
end;
$$;

drop trigger if exists on_auth_user_created on auth.users;
create trigger on_auth_user_created
  after insert on auth.users
  for each row execute procedure public.handle_new_user();

-- ---------------------------------------------------------------
-- RLS: profiles
-- ---------------------------------------------------------------
alter table public.profiles enable row level security;

create policy "profiles_select" on public.profiles
  for select using (auth.uid() = id or public.is_admin());

create policy "profiles_insert_self" on public.profiles
  for insert with check (auth.uid() = id);

create policy "profiles_update_admin" on public.profiles
  for update using (public.is_admin());

-- ---------------------------------------------------------------
-- RLS: kv_store (dados do acervo)
-- ---------------------------------------------------------------
alter table public.kv_store enable row level security;

create policy "kv_select_aprovados" on public.kv_store
  for select using (public.is_aprovado());

create policy "kv_insert_admin" on public.kv_store
  for insert with check (public.is_admin());

create policy "kv_update_admin" on public.kv_store
  for update using (public.is_admin());

create policy "kv_delete_admin" on public.kv_store
  for delete using (public.is_admin());

-- ============================================================
-- PRIMEIRO ADMINISTRADOR (passo manual, uma única vez)
-- ============================================================
-- 1. Rode este script inteiro.
-- 2. No seu app (rodando localmente ou já publicado), clique em
--    "Criar conta" e cadastre-se normalmente com seu e-mail e senha.
--    Sua conta vai ficar com status "Pendente" (ainda não dá pra entrar).
-- 3. Volte aqui no SQL Editor e rode o comando abaixo, trocando o
--    e-mail pelo que você acabou de cadastrar:
--
--    update public.profiles
--    set role = 'Administrador', status = 'Aprovado'
--    where email = 'seuemail@exemplo.com';
--
-- 4. Pronto — agora você consegue entrar como Administrador e aprovar
--    os próximos usuários direto pela tela "Usuários" do sistema.
-- ============================================================
