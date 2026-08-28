---
name: deploy-painel-forno
description: Publica o Painel do Forno na Vercel. Use sempre que precisar fazer deploy/publicar/atualizar em produção as páginas em painel-forno-COMPLETO-SYNC/ (index.html, painel-644bba82.html, supervisor-5143484a.html, import-turnos-50e86ba5.html, import.html), os manifests ou os ícones. Cobre o fluxo git-connected (mesclar no main), a promoção manual e a verificação.
---

# Deploy do Painel do Forno — publicação automática pelo GitHub

**⚠️ MUDOU (ago/2026). Agora é git-connected.** O projeto Vercel `painel-forno-areia-oleo` está
ligado ao repositório **`jorgeguill/AREIA-OLEO`**, **branch de produção `main`**, com
**Root Directory = `painel-forno-COMPLETO-SYNC`** (Framework: Other, sem build). A Vercel serve os
arquivos **reais** direto da subpasta. **Publicar = mesclar no `main`.**

> **OBSOLETO:** o antigo padrão de "loaders" (páginas minúsculas com `document.write` do raw@SHA) e o
> deploy manual via `mcp__Vercel__deploy_to_vercel`. **Não use mais.** A MCP da Vercel dá `403` para o
> time `siga-kalfix` (não deploya, não lista, não lê o site).

## Pré-requisitos
- Acesso `git push`/PR ao `jorgeguill/AREIA-OLEO` (funciona por HTTPS).
- MCP do GitHub (para abrir/mesclar PR e conferir o comentário do robô da Vercel).

## Passo a passo
1. **Editar** os arquivos reais em `painel-forno-COMPLETO-SYNC/`.
2. **Validar** antes de mesclar (nunca pular):
   - JS inline com `vm.Script` (padrão em `scripts/check.js`): `node scripts/check.js painel-forno-COMPLETO-SYNC/*.html`
   - JSON dos manifests: `node -e "JSON.parse(require('fs').readFileSync('painel-forno-COMPLETO-SYNC/manifest.json','utf8'))"`
     (idem `manifest-admin.json`, `manifest-supervisor.json`).
3. **Commit na branch de trabalho + push** (`claude/siga-control-reconfig-b2ryff`). O commit precisa
   **tocar `painel-forno-COMPLETO-SYNC/`** (o toggle "Ignorar implantações sem alteração no Root Directory"
   está ligado; ex.: subir a versão do `sw.js` já garante).
4. **Abrir PR e mesclar no `main`** (MCP `create_pull_request` + `merge_pull_request`). O merge dispara a
   Vercel automaticamente.
5. **Verificar**:
   - O **robô da Vercel comenta no PR** com status **Ready** (+ inspector/preview URL). É a confirmação de build.
   - Se o deploy do merge entrar como **Preview** e não virar produção sozinho, promover: no deploy →
     `...` → **Force Promote to Production** (feito pelo usuário no painel; a MCP não consegue).
   - Pedir ao **usuário** abrir a URL de produção — ex.: `https://painel-forno-areia-oleo.vercel.app/icon-192.png`
     deve **carregar a imagem** (não 404). WebFetch/Bash/`web_fetch_vercel_url` não alcançam o `*.vercel.app` daqui.

## PWA / ícones (cuidados)
- Ícones de PWA = **PNG same-origin**, referenciados por caminho relativo (`/icon-*.png`) + `apple-touch-icon`.
  **Nunca** referenciar ícone por URL externa (raw.githubusercontent) nem em JPEG → o navegador não renderiza
  e o atalho fica com "V" cinza.
- 3 apps instaláveis, cada um com manifest e `start_url` próprios: `manifest.json` (operador, `/`),
  `manifest-admin.json` (`/painel-644bba82.html`), `manifest-supervisor.json` (`/supervisor-5143484a.html`).
- Atalho já instalado não troca o ícone sozinho → apagar e reinstalar (e, se preciso, limpar cache do Chrome).
- Ao mexer em cache/PWA, subir a versão do `sw.js` (`painel-forno-vN`).

## Erros comuns / notas
- **Repositório errado conectado**: sintoma clássico é "nada que eu mesclo aparece / produção congelada numa
  data antiga / `/arquivo-novo` dá 404". Conferir Vercel → projeto → Settings → Git: tem que ser o
  **`AREIA-OLEO`**. Já houve engano de ligar o `jorgeguill/KALFIX-SIGA` (que só tem um `.zip` — não serve site).
- **Root Directory**: se `/` der 404 em tudo, provavelmente o Root Directory não está como `painel-forno-COMPLETO-SYNC`.
- **Firebase**: não dá para escrever no Firebase daqui (proxy 403) — importações/migrações rodam pelo usuário no navegador.
