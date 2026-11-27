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
| Site | Cloudflare Pages (ou GitHub Pages para testes) | Publica as telas a cada atualização do repositório, com um subdomínio por publisher |
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

### 3. Publicação de teste (GitHub Pages)

1. No repositório, abra **Settings › Pages**.
2. Em **Source**, escolha **Deploy from a branch**, branch **main**, pasta **/ (root)** e salve.
3. Em um ou dois minutos o site fica disponível em `https://mailrodrigobueno-ondevoucorrer.github.io/smartree-planejador/`.

Cada novo commit na branch `main` atualiza o site sozinho. Para testar a marca de um publisher nesse endereço, acrescente `?p=<código>` (ex.: `?p=marketup`). Para domínio próprio e subdomínios, veja o passo 7.

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

### 7. Domínio e subdomínios por publisher (Cloudflare Pages)

Cada publisher abre no próprio subdomínio, com o código do publisher: `marketup.smartree.com.br`. O GitHub Pages aceita um único domínio por site; por isso a publicação com domínio próprio usa o **Cloudflare Pages** (gratuito, até 100 domínios por projeto), ligado a este mesmo repositório.

1. **Cloudflare:** crie uma conta em [cloudflare.com](https://cloudflare.com), clique em **Add a domain**, informe `smartree.com.br` e escolha o plano **Free**.
2. **Registro.br:** em **Domínios › smartree.com.br › DNS › Alterar servidores DNS**, troque pelos dois servidores que a Cloudflare indicar. A troca leva de minutos a algumas horas.
3. **Projeto:** na Cloudflare, abra **Workers & Pages › Create › Pages › Connect to Git**, escolha este repositório, branch `main`, **Framework preset: None**, **Build command:** vazio, **Build output directory:** `/`. Clique em **Save and Deploy**.
4. **Domínios:** no projeto, abra **Custom domains › Set up a custom domain** e adicione `app.smartree.com.br` (endereço do admin) e `marketup.smartree.com.br`. Para cada novo cliente, adicione `<código-do-publisher>.smartree.com.br`.
5. **Supabase › Authentication › URL Configuration:** **Site URL** `https://app.smartree.com.br` e, em **Redirect URLs**, `https://*.smartree.com.br/**`.

O `config.js` já traz `dominioBase: "smartree.com.br"`: o site reconhece o publisher pelo subdomínio, mostra o logo e as cores dele na tela de login e gera os convites com o endereço certo. Rode uma vez [`supabase/migracoes/003_marca_publica.sql`](supabase/migracoes/003_marca_publica.sql) em projetos criados antes desta função.

## Antes de vender: pontos de produção

- **E-mails do sistema:** a recuperação de senha usa o envio padrão do Supabase, com limite baixo por hora. Configure um SMTP próprio em **Authentication › Emails › SMTP Settings** (por exemplo Resend, Brevo ou Amazon SES) com remetente do seu domínio e, depois, religue **Confirm email**: com a confirmação desligada, alguém que conheça o e-mail de um convite pendente poderia criar a conta antes da pessoa convidada.
- **Dados de clientes:** este repositório é público. Nunca grave aqui dados de publishers (inventário, planilhas, logos de clientes); eles ficam só no Supabase.
- **Backup:** o plano gratuito do Supabase não tem backup diário automático; avalie o plano Pro quando houver clientes pagantes.
- **Atualização automática do GAM:** hoje o script do Google Apps Script grava o relatório numa planilha e o inventário é importado em CSV. A próxima etapa é uma função no Supabase que busca o relatório e atualiza o inventário sozinha.

## Desenvolvimento local

O site é estático. Para testar no computador, rode um servidor simples na pasta do projeto, por exemplo `python3 -m http.server 8000`, e acesse `http://localhost:8000`. Inclua `http://localhost:8000` nas **Redirect URLs** do Supabase para o login funcionar localmente.
