---
name: deploy-painel-forno
description: Publica o Painel do Forno na Vercel. Use sempre que precisar fazer deploy/publicar/atualizar em produção as páginas em painel-forno-COMPLETO-SYNC/ (index.html, painel-644bba82.html, supervisor-5143484a.html, import-turnos-50e86ba5.html, import.html). Cobre o padrão de loaders apontando para o SHA do commit, a lista completa de arquivos, e a verificação pós-deploy.
---

# Deploy do Painel do Forno

A Vercel **não** é git-connected e os HTML reais são grandes demais para mandar inline de forma
confiável pela ferramenta de deploy. Por isso a produção serve **loaders minúsculos** que buscam o
arquivo real do GitHub (via `raw.githubusercontent.com`, fallback `jsDelivr`) **fixado no SHA do commit**
e o renderizam com `document.write`. Publicar = apontar os loaders para o novo SHA.

## Pré-requisitos
- Vercel MCP tool disponível (`mcp__Vercel__deploy_to_vercel`, `mcp__Vercel__web_fetch_vercel_url`).
- Projeto Vercel: `name: painel-forno-areia-oleo`, `teamId: siga-kalfix`, `target: production`.
- Produção: `https://painel-forno-areia-oleo-siga-kalfix.vercel.app/`.

## Passo a passo

1. **Editar** os arquivos reais em `painel-forno-COMPLETO-SYNC/`.

2. **Validar** o JS inline antes de qualquer deploy (nunca publicar sem isso). Extrair cada `<script>`
   e checar com `vm.Script`:
   ```js
   // scripts/check.js  (padrão do projeto)
   const fs=require('fs'),vm=require('vm');
   for(const f of process.argv.slice(2)){const h=fs.readFileSync(f,'utf8');let m,n=0,bad=0;
     const re=/<script>([\s\S]*?)<\/script>/g;
     while((m=re.exec(h))){n++;try{new vm.Script(m[1]);}catch(e){bad++;console.log('FAIL',f,e.message);}}
     console.log(`${f}: ${n} script(s), ${bad} erro(s)`);}
   ```
   `node scripts/check.js painel-forno-COMPLETO-SYNC/*.html`

3. **Commit + push** para `jorgeguill/AREIA-OLEO` na branch de trabalho. Pegar o **SHA completo**:
   `git rev-parse HEAD`.

4. **Deploy** com `mcp__Vercel__deploy_to_vercel` (`target: production`, `name: painel-forno-areia-oleo`,
   `teamId: siga-kalfix`). Mandar **SEMPRE o conjunto completo** (o deploy substitui todos os arquivos):
   - 5 **loaders** (utf-8): `index.html`, `painel-644bba82.html`, `supervisor-5143484a.html`,
     `import-turnos-50e86ba5.html`, `import.html` — cada um é o template abaixo com `S`=novo SHA e `F`=nome do arquivo.
   - `manifest.json` e `sw.js` **reais** (encoding `base64`): gerar com `base64 -w0 painel-forno-COMPLETO-SYNC/manifest.json`
     (e o mesmo para `sw.js`). O `manifest.json` correto tem md5 `627c78e76a96e40293eb4c8c87e5634d`.
   - Os loaders usam **aspas simples e uma linha só** (sem `"` e sem quebras) → transmitem como utf-8 sem
     problema de escape. NUNCA usar placeholder/texto encurtado — sempre o conteúdo real.

   **Template do loader** (trocar `SHA` e `FILE`):
   ```html
   <!doctype html><html lang='pt-BR'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width, initial-scale=1, maximum-scale=1, viewport-fit=cover'><meta name='theme-color' content='#12151A'><meta name='robots' content='noindex, nofollow'><link rel='manifest' href='/manifest.json'><title>Painel Forno</title><style>html,body{margin:0;height:100%;background:#12151A;color:#8D96A1;font-family:system-ui,sans-serif}#ld{display:flex;height:100vh;align-items:center;justify-content:center;text-align:center;padding:24px;font-size:14px}</style></head><body><div id='ld'>Carregando...</div><script>(function(){var S='SHA',F='FILE',U=['https://raw.githubusercontent.com/jorgeguill/AREIA-OLEO/'+S+'/painel-forno-COMPLETO-SYNC/'+F,'https://cdn.jsdelivr.net/gh/jorgeguill/AREIA-OLEO@'+S+'/painel-forno-COMPLETO-SYNC/'+F];function go(i){if(i>=U.length){document.getElementById('ld').textContent='Nao foi possivel carregar. Verifique a conexao e recarregue a pagina.';return;}fetch(U[i]).then(function(r){if(!r.ok)throw 0;return r.text();}).then(function(h){document.open();document.write(h);document.close();}).catch(function(){go(i+1);});}go(0);})();</script></body></html>
   ```

5. **Verificar** (o proxy do Bash bloqueia raw/jsDelivr/Vercel, então usar as ferramentas certas):
   - `mcp__Vercel__web_fetch_vercel_url` em `.../painel-644bba82.html` → confere que o loader tem o **novo SHA**.
   - `WebFetch` em `https://raw.githubusercontent.com/jorgeguill/AREIA-OLEO/<SHA>/painel-forno-COMPLETO-SYNC/<arquivo>`
     → confere strings da feature nova.
   - Conferir `manifest.json` na Vercel (etag deve bater com o md5 acima).

## Erros comuns / notas
- **Classificador de segurança do deploy indisponível** → só reintentar depois de alguns instantes.
- **Nunca** encurtar/placeholderizar arquivos no deploy — já quebrou produção antes (404). Conteúdo real sempre.
- Se ao retypar o `manifest.json` base64 der divergência (ex.: "produtivity"), gerar de novo do arquivo e
  conferir o md5 antes de publicar.
- **Não** dá para escrever no Firebase daqui — migrações/importações de dados são executadas pelo usuário
  no navegador (as páginas `import-*.html` são idempotentes: upsert por `data|turno` e `data|hora|fornecedor`).
- Adicionar `CLAUDE.md`/skills ao repo **não** exige redeploy (não muda o conteúdo servido).
