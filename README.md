[README.md](https://github.com/user-attachments/files/29808146/README.md)
# Painel do Forno — Aplicativo PWA

**Secagem de areia com conversão automática m³ úmida → seca**

## 📦 Arquivos

- `index.html` — Aplicativo (HTML + CSS + JS em um arquivo)
- `manifest.json` — Configuração PWA (nome, ícone, cores)
- `sw.js` — Service Worker (cache offline)
- `.gitignore` — Ignorar arquivos ao fazer commit

## 🚀 Publicar em 5 minutos

### Opção A: Vercel (recomendado)

1. Acesse https://vercel.com e faça login com GitHub (ou cria conta)
2. Clique **"New Project"**
3. Escolha **"Import Git Repository"** ou **"Add GitHub App"**
4. Se não tiver repositório GitHub:
   - Crie um em https://github.com/new
   - Nome: `painel-forno`
   - Faça upload dos 4 arquivos (index.html, manifest.json, sw.js, .gitignore)
   - Faça commit
5. De volta ao Vercel, escolha o repositório `painel-forno`
6. Clique **"Deploy"** e aguarde (2-3 minutos)
7. Recebe um link tipo: **`https://painel-forno.vercel.app`**

### Opção B: Netlify

1. Acesse https://netlify.com
2. Faça drag-and-drop dos 4 arquivos
3. Site fica pronto em segundos com um link gerado automaticamente

### Opção C: GitHub Pages

1. Crie repositório em https://github.com/new com nome `painel-forno`
2. Coloque os 4 arquivos na pasta `docs/`
3. Acesse Settings → Pages → Deploy from branch: `main`, folder `docs`
4. Link fica: `seuuser.github.io/painel-forno`

## 📱 Usar no celular

**Android (Chrome/Firefox):**
- Abra o link
- Menu → **"Instalar aplicativo"**

**iPhone (Safari):**
- Abra o link
- Compartilhar → **"Adicionar à tela de início"**

## ⚙️ Primeira execução

1. Abra o app
2. Toque em ⚙️ (engrenagem)
3. Configure:
   - **Umidade bruta:** % de água na areia (padrão: 8%)
   - **Densidade seca:** ton/m³ (padrão: 1.5)
   - **Meta L/ton:** sua meta (padrão: 8)
   - **Tolerância:** % acima da meta (padrão: 15%)
4. Clique "Salvar"

## 📊 Como usar

1. **Registrar Turno:**
   - Data, turno, operador
   - m³ úmida que passou no elevador
   - Óleo: litros início e fim
   - App calcula automaticamente L/ton

2. **Recebimento:** registra descarga de óleo (opcional)

3. **Histórico:** últimos turnos aparecem na lista

## 🔄 Atualizar depois

Se precisar mudar algo:
1. Edite `index.html` (ou outro arquivo)
2. Faça commit e push no GitHub
3. Vercel/Netlify detecta e faz deploy automático
4. App atualiza quando usuários abrem de novo

## 🌐 Dados

- Dados salvos **localmente** no celular (localStorage)
- Cada device tem seus próprios dados
- Para **equipe usar junto** com dados sincronizados, integrar API de backend (fora do escopo)

## ✅ Checklist antes de compartilhar

- [ ] App publicado online (link ativo)
- [ ] Link abre no celular
- [ ] Consegue instalar como app
- [ ] Configurações salvas
- [ ] Registrou um turno de teste
- [ ] Dados aparecem no histórico

Pronto para compartilhar o link com a equipe! 🎉

---

**Dúvidas?** Fale com Jorge (Radar Assessoria Empresarial)
