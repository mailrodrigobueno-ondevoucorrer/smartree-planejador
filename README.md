# SmarTree · Planejador de Propostas

Plataforma para publishers montarem propostas comerciais a partir do próprio inventário do Google Ad Manager (GAM), com a marca de cada publisher no sistema, no Excel e na apresentação em PowerPoint.

## O que a aplicação faz

- **Proposta:** anunciante, período, verba, contexto, praças, segmentos, desconto e formatos. As impressões ficam limitadas ao que o inventário entrega no período.
- **Exportação:** Excel editável (com fórmulas) e apresentação de defesa da campanha em PowerPoint, as duas com o logo e as cores do publisher.
- **Propostas salvas:** salvar, abrir, aprovar e reaproveitar a proposta aprovada do mesmo anunciante.
- **Admin:** cadastro de publishers (formatos, CPM tabela, segmentos, identidade visual), importação do relatório do GAM e liberação de usuários.
- **Publisher:** cada usuário vê só o publisher liberado para ele. O isolamento é garantido no banco de dados, não só na tela.

## Como funciona

| Parte | Serviço | Papel |
|---|---|---|
| Site | GitHub Pages | Publica as telas a cada atualização do repositório |
| Login, banco e logos | Supabase | Usuários (login com e-mail e senha), publishers, propostas e arquivos de logo |
| Inventário | Google Ad Manager | Relatório importado em CSV/Excel (atualização automática via Google Apps Script) |

Estrutura dos arquivos:

```
index.html            telas
css/app.css           estilos
js/app.js             aplicação (cálculo, Excel, PowerPoint, telas e camada de dados)
config.js             endereço e chave pública do Supabase
supabase/schema.sql   tabelas e regras de acesso do banco
assets/               ícone
```

## Instalação

### 1. Banco de dados (Supabase)

1. Crie uma conta em [supabase.com](https://supabase.com) e um projeto novo (região **South America (São Paulo)**).
2. Abra **SQL Editor › New query**, cole o conteúdo de [`supabase/schema.sql`](supabase/schema.sql) e clique em **Run**.
3. Em **Authentication › Sign In / Providers › Email**, deixe o provedor ativo e **desligue "Confirm email"**. Assim a conta funciona logo depois de criada, sem depender do envio de e-mail. Religue quando tiver SMTP próprio (veja "Antes de vender").
4. Em **Authentication › URL Configuration**:
   - **Site URL:** o endereço do site (passo 3, por exemplo `https://mailrodrigobueno-ondevoucorrer.github.io/smartree-planejador/`).
   - **Redirect URLs:** o mesmo endereço. Quando tiver domínio próprio, acrescente-o aqui também.
5. Em **Project Settings › API**, copie o **Project URL** e a chave **anon / publishable**.

> A chave **service_role / secret** nunca vai para o site nem para o repositório. Ela dá acesso total ao banco.

### 2. Configuração do site

Edite o arquivo [`config.js`](config.js) com os dois valores do passo anterior e salve (commit) no GitHub:

```js
window.SMARTREE_CONFIG = {
  supabaseUrl: "https://xxxxxxxx.supabase.co",
  supabaseAnonKey: "chave-publica-anon"
};
```

A chave pública pode ficar no repositório: o que protege os dados são as regras de acesso do `schema.sql`.

### 3. Publicação (GitHub Pages)

1. No repositório, abra **Settings › Pages**.
2. Em **Source**, escolha **Deploy from a branch**, branch **main**, pasta **/ (root)** e salve.
3. Em um ou dois minutos o site fica disponível em `https://mailrodrigobueno-ondevoucorrer.github.io/smartree-planejador/`.

Cada novo commit na branch `main` atualiza o site sozinho.

### 4. Primeiro administrador

1. Abra o site, clique em **Primeiro acesso? Criar conta** e crie a conta com o seu e-mail e uma senha.
2. No Supabase, em **SQL Editor**, rode (com o seu e-mail):
   ```sql
   update public.perfis set admin = true where email = 'seu@email.com';
   ```
3. Recarregue o site: as abas **Configurar publisher** e **Usuários** aparecem.

### 5. Primeiro publisher

- Crie o publisher em **Configurar publisher › + Novo publisher**, ou carregue os dados prontos que você recebeu em arquivo `.seed.sql` (rode no SQL Editor).
- Envie o logo e escolha a cor em **Identidade visual**.
- Importe o relatório do GAM em **Inventário**.

### 6. Liberar usuários de um publisher

**Convidar colaborador (recomendado):** em **Usuários › Convidar colaborador**, informe o e-mail e o publisher e clique em **Registrar convite**. Copie a mensagem gerada e envie por WhatsApp ou e-mail. A pessoa cria a conta com esse e-mail e entra direto no publisher.

**Pedido de acesso:** quem entra sem convite informa a empresa e clica em **Pedir acesso**. Em **Usuários › Aguardando acesso**, escolha o publisher e clique em **Liberar acesso**.

> Projetos criados antes dos convites: rode uma vez [`supabase/migracoes/002_convites.sql`](supabase/migracoes/002_convites.sql) no SQL Editor.

### 7. Domínio próprio

1. Em **Settings › Pages › Custom domain**, informe o domínio, por exemplo `app.smartree.com.br`, e ative **Enforce HTTPS**.
2. No provedor do domínio, crie um registro **CNAME** de `app` apontando para `mailrodrigobueno-ondevoucorrer.github.io`.
3. No Supabase, atualize **Site URL** e **Redirect URLs** com o novo endereço.

## Antes de vender: pontos de produção

- **E-mails do sistema:** a recuperação de senha usa o envio padrão do Supabase, com limite baixo por hora. Configure um SMTP próprio em **Authentication › Emails › SMTP Settings** (por exemplo Resend, Brevo ou Amazon SES) com remetente do seu domínio e, depois, religue **Confirm email**: com a confirmação desligada, alguém que conheça o e-mail de um convite pendente poderia criar a conta antes da pessoa convidada.
- **Dados de clientes:** este repositório é público. Nunca grave aqui dados de publishers (inventário, planilhas, logos de clientes); eles ficam só no Supabase.
- **Backup:** o plano gratuito do Supabase não tem backup diário automático; avalie o plano Pro quando houver clientes pagantes.
- **Atualização automática do GAM:** hoje o script do Google Apps Script grava o relatório numa planilha e o inventário é importado em CSV. A próxima etapa é uma função no Supabase que busca o relatório e atualiza o inventário sozinha.

## Desenvolvimento local

O site é estático. Para testar no computador, rode um servidor simples na pasta do projeto, por exemplo `python3 -m http.server 8000`, e acesse `http://localhost:8000`. Inclua `http://localhost:8000` nas **Redirect URLs** do Supabase para o login funcionar localmente.
