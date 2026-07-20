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
- `manifest.json`, `sw.js` — PWA (service worker `painel-forno-v5`, network-first).
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

## Deploy — LER ANTES DE PUBLICAR
O deploy segue um padrão específico. **Não** é git-connected e **não** se manda os HTML grandes inline.
Ver o skill do projeto: **`.claude/skills/deploy-painel-forno/SKILL.md`**. Resumo:
1. Editar os arquivos reais em `painel-forno-COMPLETO-SYNC/`, validar o JS, commitar e **push** para o GitHub.
2. A Vercel serve **páginas "loader" minúsculas** que buscam o arquivo real do
   `raw.githubusercontent.com` **fixado no SHA do commit** (fallback jsDelivr) e renderizam via `document.write`.
3. Publicar = redeployar os 5 loaders **apontando para o novo SHA** + `manifest.json`/`sw.js` reais (base64).
   Cada deploy **substitui todos** os arquivos — sempre mandar o conjunto completo.

## Restrições do ambiente (importantes)
- O **proxy bloqueia** Firebase, Vercel e raw/jsDelivr a partir do Bash (403). → **não dá para escrever no
  Firebase daqui**; importações e migrações de dados precisam ser rodadas **pelo usuário no navegador**.
- `git push`/`fetch` por HTTPS **funciona**. A **ferramenta MCP da Vercel** alcança a Vercel.
  **WebFetch** alcança `raw.githubusercontent.com` (usar para verificar o conteúdo publicado).
- **LibreOffice não funciona** aqui ("source file could not be loaded") → não converter pptx→pdf nem gerar
  prévia de slides. Para PDF projetado use reportlab (guia.py) — instalar `reportlab` via pip.
- `pptxgenjs` não vem instalado no scratchpad — `npm install pptxgenjs` antes de gerar deck.
- O **classificador de segurança do deploy** às vezes fica indisponível → só reintentar.

## Validação (sempre antes do deploy)
Extrair os `<script>` inline e checar sintaxe com `node`+`vm.Script` (padrão em `scratchpad/check.js`).
Depois do deploy: `web_fetch_vercel_url` no loader (conferir o SHA) e `WebFetch` no raw (conferir strings
da feature). `manifest.json` correto tem etag/md5 `627c78e76a96e40293eb4c8c87e5634d`.

## Repo / Vercel
- GitHub: `jorgeguill/AREIA-OLEO` (público). Sub-pasta `painel-forno-COMPLETO-SYNC/`.
- Vercel: projeto `painel-forno-areia-oleo`, team **`siga-kalfix`**. Produção:
  `https://painel-forno-areia-oleo-siga-kalfix.vercel.app/`.

## Roadmap (pendências combinadas com o usuário)
- **Leitura de tanque** → L/ton real (hoje é estimado quando o óleo não é informado).
- **PIN por perfil** (hoje o acesso é só por link secreto).
