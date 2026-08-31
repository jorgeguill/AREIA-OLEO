# SKAL Industrial Cost Intelligence System

**Digital Cost Twin** da operação industrial da SKAL / Grupo Kalfix — um modelo
econômico funcional, modular, auditável e parametrizado, entregue como workbook
Excel gerado por código (fonte da verdade: `build_skal_cost_system.py`).

> Não é "planilha de custo por saco". É um sistema que responde, para qualquer SKU:
> quanto custou, quanto deveria ter custado, onde/por que desviou, quanto é
> controlável e quanto vale eliminá-lo.

## Como reconstruir o arquivo

```bash
pip install openpyxl
python3 build_skal_cost_system.py       # gera SKAL_Industrial_Cost_Intelligence_System.xlsx
```

`build_skal_cost_system.py` chama automaticamente `ingest_compras.py`, que lê todas as
extrações reais em `data/SKAL_Compras_*.xlsx` e popula as abas `R0`–`R5`. Para adicionar
um novo mês, basta soltar o arquivo de extração em `data/` e rodar de novo.

O arquivo `.xlsx` (59 abas) é versionado junto para uso imediato.

## 1. O que foi construído

Sistema com **53 abas** organizadas em 18 módulos, do cadastro à inteligência
executiva. Núcleo econômico com **fórmulas vivas e reconciliadas**:

- **Master data** (produtos, materiais, fornecedores) com chaves únicas e conversões
  automáticas de embalagem — inclusive a exceção da linha **Extras** (20 kg/saco,
  75 sacos/palete) vs. padrão (15 kg/saco, 100/palete).
- **Landed Net Cost** por material: bruto − descontos − rebates + frete/seguro/
  descarga/armazenagem + tributos não-recuperáveis − créditos recuperáveis.
- **Tax engine** parametrizado: ICMS com benefício de 80% calculado como
  `nominal × %devido` (= 4,5%, **nunca hardcoded**); PIS/COFINS Lucro Real;
  IRPJ/CSLL/SUDENE (75% só sobre IRPJ); campos de vigência para a reforma
  (CBS/IBS 2026–2033).
- **BOM técnica e econômica** com controle de fechamento (Σ kg/t = 1.000).
- **Standard Cost, Actual Cost e Variance Engine**: PPV, Usage, decomposição
  Mix/Yield e variação de conversão.
- **Capacity & Constraint** (unused capacity cost visível, MC/hora do gargalo).
- **Cost Gap** (5 níveis: Actual → Current Std → Best Demonstrated → Entitlement →
  Theoretical) e **Savings Engine** (funil de 8 estágios, regra de "Realizado").
- **Cost to Serve, Pricing (7 níveis) e Profitability** cliente × SKU com
  Price Waterfall → Pocket Price → Pocket Margin e detecção de destruidores de valor.
- **Reconciliações** física, industrial (Std + Var = Actual) e contábil (ponte CPV).
- **5 dashboards** (executivo, produção, compras, comercial, redução de custos) +
  **CONTROLS_ALERTS** com testes automáticos + **DATA_DICTIONARY**, **DATA_GAPS**
  e **LEGACY_RECON**.

## 2. Quais dados foram utilizados

**Nenhum dado real da SKAL foi inventado.** O sistema tem duas camadas claramente
separadas:

**A) Camada REAL (abas `R0`–`R5`)** — extraída das notas de compra reais fornecidas
(`data/SKAL_Compras_2026-01_Extracao.xlsx` e `..._2026-08_Extracao.xlsx`), via
`ingest_compras.py`. Itens de OCR sem quantidade/valor legíveis foram **descartados,
nunca inferidos**. Contém: 15+ fornecedores reais, 154 documentos no ledger, itens de
produto legíveis (cimento, calcário, energia), 31 CT-e de frete e o custo **landed real**.
Agosto é limpo; Janeiro contribui sobretudo com o ledger documento-a-documento e
fornecedores (itens de OCR ruidosos foram filtrados).

**B) Camada DEMO (motor de custeio)** — como ainda **não há dados reais de BOM,
produção e consumo**, o motor de Standard/Actual/Variance usa um pequeno dataset de
**DEMONSTRAÇÃO explicitamente marcado** (3 produtos, 5 materiais) só para exercitar e
provar as fórmulas de reconciliação. Ligar/desligar em `01_CONTROL_PANEL!MODO_DEMO`.

As planilhas legadas citadas no briefing (`Planilha_Precificação.xlsx`,
`CustoFormula2026.1.xlsb`) **não foram fornecidas** — lacuna nº 1 em `DATA_GAPS`.

## 3. Lacunas que ainda precisam ser preenchidas

Ver a aba **`DATA_GAPS`** (15 itens priorizados). Prioridade ALTA:
arquivos legados; landed cost real por MP/fornecedor/período; preços históricos
(PPV); perdas/reprocesso e tempos reais por OF; base de vendas cliente × SKU;
razão contábil/CPV. Cada lacuna já tem **o campo de entrada criado** no sistema.

## 4. Controles implementados

Aba **`CONTROLS_ALERTS`** (validação matemática, todos verdes no dataset demo,
exceto o alerta legítimo que sinaliza o destruidor de valor da demonstração):

- BOM fecha (Σ kg/t = 1.000) por produto.
- **Standard + Variances = Actual** (resíduo por linha → 0; resíduo por produto → 0).
- Sem custo real negativo; toda MP usada tem custo rastreável.
- Mass balance sem perda não-explicada; ICMS efetivo consistente (sem hardcode).
- Detecção de produtos/clientes destruidores de valor.

Validação independente: recálculo do workbook com o motor `formulas` →
**0 erros de fórmula em 885 células**; `PPV + Usage + Conversão = Cost Gap` exato.

## 5. Inconsistências do modelo legado

**Não avaliáveis ainda** — as planilhas legadas não foram fornecidas. A aba
`LEGACY_RECON` já está pronta (produto a produto: preço/custo atual vs. novo custo
industrial/econômico, diferença abs./%, causa) para ser preenchida assim que o
arquivo de referência chegar. A análise será **estritamente técnica** (divergência
de metodologia / premissa sem suporte / dado não rastreável), nunca pessoal.

## 6. Principais oportunidades econômicas identificadas (dados reais)

Já com os dados reais de compra (abas `R0`–`R5`):

- **Frete inbound do cimento é um custo de primeira ordem.** A NF do Cimento CP-V ARI
  (Apodi) é **R$ 505/t**, mas o **landed cost real é ~R$ 835/t** — o frete de
  Quixeré/CE → Teresina/PI adiciona **~R$ 330/t (≈ 65%)**. Precificar/custear pela NF
  subestima o cimento em dois terços. Oportunidades: renegociar frete, revisar origem/
  modal, avaliar fornecedor mais próximo (Should Cost / Supplier TCO).
- **Concentração de gasto:** Cimento Apodi (~R$ 353 mil) + seu frete F G Inteligência
  (~R$ 344 mil) somam **~52% de todo o gasto mapeado** (R$ 1,34 mi nos dois meses).
  Aqui está a maior alavanca de redução de custo de materiais.
- **Cimento é ICMS-ST (CST 500):** sem crédito de ICMS na entrada — o custo econômico
  não deve descontar crédito de ICMS deste item (tratado corretamente no landed).

O motor de valor (Cost Gap 5 níveis, savings priority score, unused capacity, pocket
margin) está pronto para receber a BOM e a produção reais e traduzir isso em R$/t por SKU.

## Governança

INPUT (amarelo) / FORMULA (branco) / LINKED (azul) / CONTROL (verde) / ALERT
(vermelho) / CHAVE (cinza). Premissas globais centralizadas e referenciadas por
**nome** em `01_CONTROL_PANEL` (23 parâmetros). Papéis de dono do dado, dono da
metodologia, aprovador e usuário documentados no briefing e refletidos nas colunas
de governança das tabelas mestre.
