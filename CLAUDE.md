# Painel do Forno — memória do projeto

PWA (HTML/CSS/JS puro, sem build) + **Firebase Realtime Database** para a usina de
**secagem de areia / consumo de óleo** da Kalfix. Tudo vive em `painel-forno-COMPLETO-SYNC/`.
Branch de trabalho: `claude/siga-control-reconfig-b2ryff`.

## Arquivos servidos
- `index.html` — **app do operador** (registrar turno por lote, recebimento de óleo). PWA.
- `painel-644bba82.html` — **painel admin** (edição, gráficos, export CSV, parâmetros). Link secreto.
- `supervisor-5143484a.html` — **painel supervisor** (só leitura, mesmos gráficos). Link secreto.
- `import-turnos-50e86ba5.html` — importador de turnos do histórico (upsert idempotente).
- `import.html` — importador de recebimentos de óleo (+ botão "preencher em branco").
- `sw.js` — service worker (`painel-forno-v7`, network-first). Subir a versão a cada mudança limpa o cache.
- **PWA instalável (3 apps distintos)**: cada página tem seu manifest e ícone próprios:
  - operador → `manifest.json` + `icon-192.png`/`icon-512.png` (secador Kalfix).
  - admin → `manifest-admin.json` + `icon-admin-*.png` (secador + faixa **ADMIN** vermelha), `start_url:/painel-644bba82.html`.
  - supervisor → `manifest-supervisor.json` + `icon-supervisor-*.png` (secador + faixa **SUPERVISOR** grafite), `start_url:/supervisor-5143484a.html`.
  - **Ícones = PNG same-origin** referenciados por caminho relativo (`/icon-*.png`) + `apple-touch-icon`. NÃO usar
    URL externa (raw.githubusercontent) nem JPEG → o navegador **não renderiza** o ícone de PWA assim (fica "V" cinza).
  - Atalho já instalado **não** troca o ícone sozinho: apagar e reinstalar (e, se preciso, limpar cache do Chrome).
- Todas as páginas têm `<meta name="robots" content="noindex, nofollow">` (acesso por link secreto).

## Firebase (config pública de cliente — ok versionar; segurança fica nas regras)
`{apiKey:"AIzaSyB2SIxkvxmB_bqbIbzYSo9Sb5IQxwUTGR0", databaseURL:"https://painel-forno-default-rtdb.firebaseio.com", projectId:"painel-forno"}`
Coleções: `turnos`, `recebimentos`, `parametros`.

## Modelo de dados
- **turno**: `{data(YYYY-MM-DD), turno(Manhã|Tarde|Noite), operador, fornecedor_areia(ASSIS|SKAL),
  silo, lotes:[{ini,fim,conchas,tempo_min}], conchas, m3_umida, tempo_secagem_min,
  oleo_consumido?, obs?, origem(app|historico-planilha)}`
- **recebimento**: `{data, hora, antes, depois, fornecedor, recebedor, preco_litro?, teste_combustao?('Aprovado'|'Reprovado'), teste_agua?}`
- **parametros**: `{umidade(%), densidade(t/m³), meta(L/ton), tolerancia(%), concha_m3}` — default `{8, 1.5, 8, 15, 2.1}`.

## Cálculos (idênticos nos 3 arquivos — manter em sincronia)
- Concha → m³: **2,1 m³/concha** (caçamba Komatsu WA200-5), parâmetro `concha_m3`.
- `tonOf(t)`: m³ úmida (de `m3_umida` ou `conchas*concha_m3`) → `ton = m3u*(1-umidade/100)*densidade`.
- `mediaLton()` = Σ litros recebidos ÷ Σ ton seca (histórico).
- `compute(t)` → `{m3u,m3s,ton,oleo,lton,estimado}`. **Regra do L/ton** (corrigida): usa consumo real
  se `medido && oleo>0`; **senão estima pela média** (`else if(!lton&&ton>0.01)`, com `~` = estimado).
  Nunca deixar L/ton em branco quando há dados.
- `tempoMinLote(ini,fim)` = minutos fim−início (soma +1440 se virar a meia-noite). `tempoTurno(t)` soma os lotes.
- `normOp(o)` normaliza operador (ignora acento/maiúsculas; junta Lívio/lívio/Livio, Anailson/Anaílson/Costa/Cotia).
- Custo: **R$/ton = L/ton × preço médio do litro** (`precoMedio()` = média ponderada de `preco_litro`).

## Deploy — publicação automática pelo GitHub (git-connected)  ⚠️ MUDOU (ago/2026)
Agora **É git-connected**. O projeto Vercel `painel-forno-areia-oleo` está ligado ao repositório
**`jorgeguill/AREIA-OLEO`**, **branch de produção `main`**, **Root Directory = `painel-forno-COMPLETO-SYNC`**
(Framework: Other, sem build). **Publicar = mesclar no `main`:**
1. Editar os arquivos reais em `painel-forno-COMPLETO-SYNC/`, validar (JS + JSON), commitar na branch de trabalho.
2. Abrir PR e **mesclar no `main`** → a Vercel **publica sozinha** (o robô comenta no PR com status **Ready**).
   - O commit precisa **tocar `painel-forno-COMPLETO-SYNC/`** (o toggle "Ignorar implantações sem alteração no
     Root Directory" está ligado).
   - Se o deploy do merge ficar como **Preview** e não virar produção, promover: no deploy →
     `...` → **Force Promote to Production**.
3. **OBSOLETO:** o antigo padrão de **"loaders"** (páginas minúsculas com `document.write` do raw@SHA) e o
   deploy manual via `mcp__Vercel__deploy_to_vercel`. A Vercel agora serve os arquivos **reais** direto da subpasta.

**⚠️ Repositório certo = `AREIA-OLEO`.** Já aconteceu de o projeto Vercel ficar conectado ao
`jorgeguill/KALFIX-SIGA` (que só tem um `.zip` — não serve site) → **nada publicava** e a produção
ficava congelada num deploy manual antigo. Se os deploys não aparecerem/atualizarem: Vercel → projeto →
Settings → Git → conferir que o repo ligado é o **AREIA-OLEO** (e Settings → Build → Root Directory correto).

## Restrições do ambiente (importantes)
- O **proxy bloqueia** Firebase, Vercel e raw/jsDelivr a partir do Bash (403). → **não dá para escrever no
  Firebase daqui**; importações e migrações de dados precisam ser rodadas **pelo usuário no navegador**.
- `git push`/`fetch` por HTTPS **funciona** — é assim que se publica (push na branch → PR → merge no `main`).
- **A MCP da Vercel dá `403` para o time `siga-kalfix`**: não dá para deployar, listar deploys nem ler o site
  por ela; `web_fetch_vercel_url` também falha ("Unable to create shareable URL"). **Verificar o deploy** pelo
  **comentário do robô da Vercel no PR** (status Ready + inspector/preview URL) e pedindo ao **usuário** abrir a
  URL de produção (ex.: `/icon-192.png` deve carregar a imagem, não 404).
- **WebFetch/Bash bloqueiam** `*.vercel.app` (egress) → não dá para buscar o site publicado daqui. **WebFetch**
  alcança `raw.githubusercontent.com` e a **MCP do GitHub** lê o repo (usar para conferir a fonte no `main`).
- **LibreOffice não funciona** aqui ("source file could not be loaded") → não converter pptx→pdf nem gerar
  prévia de slides. Para PDF projetado use reportlab (guia.py) — instalar `reportlab` via pip.
- `pptxgenjs` não vem instalado no scratchpad — `npm install pptxgenjs` antes de gerar deck.
- O **classificador de segurança do deploy** às vezes fica indisponível → só reintentar.

## Validação (sempre antes de mesclar)
Extrair os `<script>` inline e checar sintaxe com `node`+`vm.Script` (padrão em `scripts/check.js`).
Validar também o **JSON dos manifests**. Depois do merge: confirmar pelo **comentário do robô da Vercel no
PR** (Ready) e pedir ao usuário abrir a URL de produção.

## Repo / Vercel
- GitHub: **`jorgeguill/AREIA-OLEO`** (público). Sub-pasta `painel-forno-COMPLETO-SYNC/` (= Root Directory na Vercel).
- Vercel: projeto `painel-forno-areia-oleo` (`prj_Q2Nb93E70F24CJCzezcXMpxAgMTA`), team **`siga-kalfix`**
  (`team_GcUWH8OVBZpk4n8dSDgPhhrt`). **Git-connected ao AREIA-OLEO, branch de produção `main`.**
- Produção: **`https://painel-forno-areia-oleo.vercel.app/`** (sem o sufixo `-siga-kalfix`).

## Roadmap (pendências combinadas com o usuário)
- **Leitura de tanque** → L/ton real (hoje é estimado quando o óleo não é informado).
- **PIN por perfil** (hoje o acesso é só por link secreto).
