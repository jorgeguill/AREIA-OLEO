# Painel do Forno — Módulo (fase de teste)

Versão **plugável** do Painel do Forno, preparada para rodar **à parte** na fase
de teste e, depois, ser **agregada ao SIGA como mais uma ferramenta** — sem
dependências cruzadas com o resto do sistema.

O app original (arquivo único) continua em `../painel-forno-COMPLETO-SYNC/`.
Este módulo é uma reorganização desse app, **sem mudar o que ele faz**: mesma
conta de m³ úmida → seca, mesmo consumo L/ton, mesmos registros de turno e
recebimento, mesmo Firebase opcional.

## 📦 Arquivos

| Arquivo            | Papel                                                        |
|--------------------|-------------------------------------------------------------|
| `painel-forno.js`  | O módulo em si. Expõe `PainelForno.mount(container, opcoes)`.|
| `painel-forno.css` | Estilos **isolados** sob `.pf-app` (não vazam para o host).  |
| `index.html`       | Host **standalone** (fase de teste). Monta o módulo sozinho. |
| `manifest.json`    | Config PWA (caminhos relativos).                            |
| `sw.js`            | Service Worker offline (caminhos relativos ao escopo).       |
| `module.json`      | Descritor para o SIGA registrar a ferramenta.               |

## 🧪 Testar à parte

Não precisa de build. Basta servir a pasta por HTTP (o Service Worker e o
`fetch` não funcionam em `file://`):

```bash
cd modulo-painel-forno
python3 -m http.server 8080
# abra http://localhost:8080
```

Publicação (Vercel/Netlify/GitHub Pages): aponte o deploy para esta pasta.
Como tudo usa caminhos relativos, funciona tanto na raiz quanto numa subpasta.

## 🔌 Como o SIGA agrega depois (contrato do módulo)

O módulo não toca em `window`/`body` do host além de expor `PainelForno`. Para
plugar no SIGA, basta carregar o CSS + JS e montar num container:

```html
<link rel="stylesheet" href="/ferramentas/painel-forno/painel-forno.css">
<script src="/ferramentas/painel-forno/painel-forno.js"></script>

<div id="area-da-ferramenta"></div>
<script>
  const painel = PainelForno.mount('#area-da-ferramenta', {
    storagePrefix: 'siga:painel-forno:',   // isola os dados dentro do SIGA
    firebasePath: 'painel-forno/'          // isola os nós no Realtime Database
  });

  // ao sair da tela/ferramenta:
  // painel.unmount();
</script>
```

### Fronteiras garantidas (sem dependências cruzadas)

- **CSS isolado:** todas as regras são escopadas sob `.pf-app`. Nada de
  `body`, `header` ou `.card` globais colidindo com o SIGA.
- **DOM isolado:** o módulo só consulta elementos dentro do container recebido
  (nada de `getElementById` global).
- **Storage com namespace:** as chaves do `localStorage` levam prefixo
  configurável (`storagePrefix`), então convivem com outras ferramentas.
- **Firebase isolado e opcional:** app Firebase nomeado por instância e nós
  configuráveis via `firebasePath`; se o SDK não estiver presente, cai para
  `localStorage` automaticamente.
- **Ciclo de vida:** `mount()` devolve `{ unmount, refresh, state }` para o
  SIGA controlar quando a ferramenta entra e sai da tela.

## ⚙️ Opções de `mount(container, opcoes)`

| Opção           | Default            | Para quê                                             |
|-----------------|--------------------|------------------------------------------------------|
| `storagePrefix` | `painel-forno:`    | Prefixo das chaves no `localStorage`.                |
| `firebasePath`  | `""`               | Prefixo dos nós no Firebase Realtime Database.       |

## ✅ Checklist da fase de teste

- [ ] Servir a pasta e abrir no navegador
- [ ] Registrar um turno e ver o L/ton no medidor
- [ ] Registrar um recebimento
- [ ] Ajustar parâmetros da areia em ⚙️ e conferir o recálculo
- [ ] (Opcional) Configurar Firebase e ver "sincronizado"
- [ ] Instalar como PWA no celular

---

**Dúvidas?** Fale com Jorge (Radar Assessoria Empresarial)
