# GeoAcervo — do zero ao APK

Este projeto é a versão "para produção" do sistema: banco de dados real na
nuvem (Supabase), login com senha criptografada de verdade (Supabase Auth)
e app instalável no Android (via PWABuilder). Nenhuma parte exige
conhecimento avançado de programação — é seguir os passos abaixo, em ordem.

---

## Parte 1 — Criar o banco de dados (Supabase, grátis)

1. Acesse **https://supabase.com** → "Start your project" → crie uma conta grátis.
2. Clique em **New Project**. Escolha um nome (ex: `geoacervo`) e uma senha
   para o banco (guarde essa senha em local seguro).
3. Aguarde ~2 minutos até o projeto ficar pronto.
4. No menu lateral, clique em **SQL Editor** → **New query**.
5. Abra o arquivo `supabase-schema.sql` (está nesta pasta), copie todo o
   conteúdo, cole no editor e clique em **Run**. Isso cria:
   - a tabela `profiles` (usuários: nome, e-mail, perfil, status);
   - a tabela `kv_store` (dados do acervo: minerais, rochas, histórico);
   - as políticas de segurança (RLS) de cada uma;
   - um gatilho que cria automaticamente o perfil de todo novo usuário.
6. (Recomendado para testar mais rápido) Vá em **Authentication → Providers
   → Email** e desative a opção **"Confirm email"**. Assim, ao criar uma
   conta pelo app, ela já pode ser aprovada e usada na hora, sem precisar
   clicar em um link de confirmação por e-mail antes. Se preferir manter a
   confirmação ativada (mais seguro), tudo bem — só avise os usuários que
   eles precisam confirmar o e-mail antes do primeiro login.
7. No menu lateral, vá em **Project Settings → API**. Você vai precisar de
   dois valores nessa tela:
   - **Project URL**
   - **anon public key**

---

## Parte 2 — Configurar o projeto

1. Instale o [Node.js](https://nodejs.org) (versão 18 ou mais recente), se
   ainda não tiver.
2. Abra um terminal dentro desta pasta e rode:
   ```
   npm install
   ```
3. Copie o arquivo `.env.example` para um novo arquivo chamado `.env`:
   ```
   cp .env.example .env
   ```
4. Abra o `.env` e cole os dois valores que você pegou no Supabase:
   ```
   VITE_SUPABASE_URL=https://seuprojetoaqui.supabase.co
   VITE_SUPABASE_ANON_KEY=eyJhbGciOi... (chave longa)
   ```
5. Teste localmente:
   ```
   npm run dev
   ```
   Abra o link que aparecer no terminal (geralmente `http://localhost:5173`).

---

## Parte 2.5 — Criar o primeiro Administrador (obrigatório, uma vez só)

O sistema não vem mais com login/senha padrão (isso era só da versão de
demonstração) — agora as contas são reais, com senha gerenciada com
segurança pelo Supabase. Por isso, o primeiro administrador precisa ser
promovido manualmente, uma única vez:

1. Com o app rodando (`npm run dev`), clique em **Criar conta** e cadastre
   seu próprio e-mail e senha. Sua conta nasce como perfil "Usuário" e
   status "Pendente" — ainda não dá pra entrar.
2. No Supabase, vá em **SQL Editor** e rode (trocando pelo seu e-mail):
   ```sql
   update public.profiles
   set role = 'Administrador', status = 'Aprovado'
   where email = 'seuemail@exemplo.com';
   ```
3. Volte no app e faça login normalmente. A partir de agora, você aprova
   os próximos usuários direto pela tela **Usuários** do sistema (um
   clique em "Aprovar"), sem precisar mexer em SQL de novo.

---

## Parte 3 — Colocar o site no ar (necessário para gerar o APK)

O PWABuilder (próxima etapa) precisa de um **link público (https://)** do
seu sistema. A forma mais simples e gratuita é o **Vercel**:

1. Rode `npm run build` (gera a pasta `dist`).
2. Acesse **https://vercel.com** → crie conta grátis (pode entrar com GitHub).
3. Clique em **Add New → Project**.
   - Se você subiu este projeto para o GitHub: selecione o repositório.
   - Se preferir sem GitHub: use o **Vercel CLI** (`npm i -g vercel` e depois
     `vercel` dentro da pasta do projeto) ou arraste a pasta `dist` em
     **vercel.com/new** (opção de deploy manual).
4. Em **Environment Variables**, adicione `VITE_SUPABASE_URL` e
   `VITE_SUPABASE_ANON_KEY` com os mesmos valores do seu `.env`.
5. Clique em **Deploy**. Em ~1 minuto você recebe um link público, tipo
   `https://geoacervo.vercel.app`.

(Netlify funciona do mesmo jeito, se preferir.)

---

## Parte 4 — Compartilhar com o pessoal (Android e iPhone)

Não é preciso gerar nenhum arquivo pra instalar — é só mandar o link
(`https://geoacervo.vercel.app` ou o domínio que você configurar). O app
já está preparado como PWA, então funciona nos dois sistemas:

**No Android (Chrome):**
1. A pessoa abre o link.
2. Aparece um banner (ou o menu ⋮ → "Instalar app" / "Adicionar à tela
   inicial"). Um toque e pronto — ícone na tela inicial, abre em tela
   cheia como um app normal, sem nenhum aviso de segurança.

**No iPhone/iPad (Safari — precisa ser o Safari, não funciona pelo Chrome no iOS):**
1. A pessoa abre o link.
2. Como o iOS não mostra um botão automático de instalar, o próprio
   sistema já exibe um aviso na tela ensinando o passo a passo: tocar no
   ícone de **Compartilhar** (a seta para cima, na barra do Safari) e
   depois em **"Adicionar à Tela de Início"**.
3. Pronto — ícone na tela inicial, abre como app, sem barra de endereço.

Nenhum dos dois casos passa pela Google Play ou App Store, então não
tem taxa, não tem revisão, e funciona imediatamente assim que você
publica o link.

**Se no futuro você ainda quiser um arquivo `.apk` de verdade** (por
exemplo para distribuir por fora, sem depender de internet no momento da
instalação), o caminho é colar esse mesmo link em **pwabuilder.com** →
"Package for Stores" → Android. Só vale lembrar que, nesse caso, quem for
instalar o `.apk` vai ver um aviso do Android pedindo permissão para
"instalar de fontes desconhecidas" — isso é uma trava do próprio sistema
para qualquer app que não venha da Play Store, não tem como remover sem
publicar oficialmente lá.

---

## Como funciona o controle de acesso agora

- **Login/senha**: gerenciados pelo **Supabase Auth** — senhas ficam
  criptografadas de verdade, nada de texto simples salvo em lugar nenhum.
  Existe também recuperação de senha ("Esqueci minha senha" na tela de
  login), usando o envio de e-mail padrão do Supabase — não precisa
  configurar nada além do projeto já criado (o Supabase free tier envia
  um número limitado de e-mails por hora; para uso maior, dá pra plugar
  um provedor próprio de e-mail nas configurações do projeto).
- **Cadastro**: qualquer pessoa pode criar conta pela tela "Criar conta",
  mas nasce com status **Pendente**.
- **Aprovação**: só o Administrador vê e aprova/recusa solicitações, na
  aba **Usuários**. Lá também é possível promover alguém a Administrador
  ou revogar o acesso de alguém (a conta deixa de conseguir entrar, mas
  não é apagada do Supabase Auth — isso exigiria uma função de servidor
  com chave "service role", que é um possível próximo passo).
- **Dados do acervo** (minerais, rochas, histórico): só usuários com
  status "Aprovado" conseguem ler; só Administradores aprovados
  conseguem gravar — essas regras já estão reforçadas no próprio banco
  (RLS), não só na tela.

---

## Resumo do que já está pronto neste projeto

- ✅ Autenticação real com Supabase Auth (e-mail + senha).
- ✅ Recuperação de senha por e-mail ("Esqueci minha senha").
- ✅ Tabela `profiles` com aprovação de cadastro e perfis (Usuário/Administrador).
- ✅ Regras de segurança (RLS) no próprio banco, não só na interface.
- ✅ Conexão do acervo (minerais/rochas/histórico) com o Postgres do
  Supabase via `src/lib/supabaseStorageAdapter.js`.
- ✅ Manifesto PWA e ícones em todos os tamanhos exigidos por Android e
  iOS (`vite.config.js`, `public/icons/`, meta tags do `index.html`).
- ✅ Banner automático ensinando o passo a passo de instalação no iPhone
  (`src/IOSInstallBanner.jsx`) — o Safari não mostra isso sozinho.
- ✅ Schema SQL completo, pronto para colar no Supabase (`supabase-schema.sql`).

## Próximos passos possíveis (me chame se quiser evoluir)

- Excluir de fato a conta no Supabase Auth (hoje só revoga o acesso),
  usando uma Edge Function com a chave "service role".
- Migrar minerais/rochas de um `kv_store` genérico para tabelas
  relacionais próprias (melhor para relatórios e para 10.000+ registros).
- Publicar oficialmente na Google Play (exige conta de desenvolvedor
  Google, taxa única de US$25).
