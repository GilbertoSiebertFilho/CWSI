# CWSI — Planilha Mestre e Guia de Monitores Agrícolas

Planilha para cálculo do **CWSI (Crop Water Stress Index)** a partir de dados de estação
meteorológica + termometria de dossel, combinada com um **guia de transferência de arquivos
entre o escritório e o monitor da máquina** (prescrições, linhas de guia AB, dados de trabalho).

**Entregável:** [`CWSI_Planilha_Mestre.xlsx`](CWSI_Planilha_Mestre.xlsx)

---

## Dois entregáveis neste repositório

| Pasta | O que é |
|---|---|
| raiz + `tools/` | A **planilha CWSI** e o guia de monitores (este documento) |
| [`platform/`](platform/) | A **AB Line Platform** — aplicação web que responde *como colocar arquivos no monitor e como tirar os dados dele*, por marca, monitor e versão de software |

A plataforma leva adiante as abas `09_Passo_a_Passo` e `10_Base_Procedimentos`
da planilha e as transforma em produto: você escolhe tipo de equipamento →
monitor → **versão do software** → o que quer fazer → como os dados viajam, e
recebe o formato de arquivo, a pasta exata no pen drive, os cliques numerados,
como conferir se deu certo e o que costuma dar errado — com ícone do terminal e
pronto para imprimir. São 264 procedimentos em 23 monitores, cobrindo pen drive, nuvem de 14 plataformas e software de escritório. Cada procedimento vira link compartilhável, e cada monitor tem um manual completo para impressão.

Gerar linhas AB continua disponível, mas é o papel de apoio.

A interface da plataforma é em inglês. Os ícones dos monitores são gerados por
`tools/gerar_icones.py` e servem aos dois entregáveis (com legenda na planilha,
sem legenda na web).

```bash
cd platform && pip install -r requirements.txt && python3 run.py --seed && python3 run.py
```

Documentação completa: [`platform/README.md`](platform/README.md).

---

## O que a planilha faz

### Parte 1 — Cálculo do CWSI

Dois métodos, selecionáveis em `01_Config`:

| Método | Fórmula | Quando usar |
|---|---|---|
| **Empírico (Idso, 1981)** | `CWSI = (dT − dT_LL) / (dT_UL − dT_LL)`<br>`dT_LL = a + b·VPD`  ·  `dT_UL = a + b·VPG`<br>`VPG = es(Ta) − es(Ta + a)` | Padrão. Simples, mas exige baseline (a, b) calibrada localmente. |
| **Teórico (Jackson, 1981)** | Balanço de energia com `ra`, `Rn`, `Δ`, `γ*` | Quando há medida de Rn e vento confiáveis. |

Toda a cadeia intermediária fica visível coluna a coluna em `04_Calculo` — `es`, `ea`, `VPD`,
`Δ`, `γ`, `u2`, `dT`, `VPG`, os dois limites, `ra`, `γ*` — para que qualquer resultado possa ser
auditado sem abrir fórmula.

### Parte 2 — Guia de monitores

`09_Passo_a_Passo` monta a sequência de cliques a partir de seis escolhas encadeadas:

```
marca → tipo de equipamento → monitor → versão do software → objetivo → origem/destino
```

e devolve: **formato de arquivo**, **pasta exata no pen drive**, **sistema de arquivos**,
**até 10 passos numerados**, **cuidados**, **erros comuns**, **grau de confiança** e **fonte**.

As listas de monitor e de versão são dependentes (mudam conforme a marca / o monitor escolhido).
Quando não existe procedimento para a versão específica, a planilha cai automaticamente na
entrada `Todas as versoes`.

---

## Abas

| Aba | Função |
|---|---|
| `00_Inicio` | Capa, índice e roteiro de uso |
| `01_Config` | **Comece aqui.** Local, sensores, unidades, cultura, filtros, limiares |
| `02_Estacao` | Dados brutos da estação meteorológica (colar) |
| `03_Dossel` | Temperatura de dossel do infravermelho (colar) |
| `04_Calculo` | Motor de cálculo, etapa por etapa. Só fórmulas |
| `05_QC_Diagnostico` | Painel automático + checklist dos 22 erros clássicos |
| `06_Resultados` | CWSI diário e decisão de irrigação |
| `07_Baselines` | Coeficientes a e b por cultura + roteiro de calibração local |
| `08_Equipamentos` | Catálogo ilustrado de 14 monitores/terminais |
| `09_Passo_a_Passo` | Gerador do passo a passo |
| `10_Base_Procedimentos` | Banco de procedimentos — **é aqui que o guia cresce** |
| `11_Formatos_Arquivos` | Shapefile, ISOXML, .agsetup, .agdata, AgData, Rx, TASKDATA… |
| `12_Fontes` | Referências, premissas e limitações |

Convenção de cores: **amarelo + azul** = você digita · **cinza** = fórmula · **laranja** = linha
de exemplo (apagar) · **vermelho** = alerta · **verde** = verificação aprovada.

---

## Como ampliar o guia

O gerador da aba 09 é só uma consulta à aba `10_Base_Procedimentos`. Para cobrir uma combinação
nova, acrescente uma linha lá (há 20 linhas em branco já formatadas, com listas suspensas) —
ela passa a aparecer na aba 09 sozinha, sem mexer em fórmula.

A coluna **Confiança** classifica cada linha:

- **Verificado** — formato e caminho de pasta confirmados na fonte citada.
- **Confirmar** — a estrutura está correta, mas o nome exato do menu muda entre versões.
  Confirme na máquina e atualize a linha.

---

## Reconstruir os arquivos

```bash
pip install openpyxl pillow
python3 tools/gerar_icones.py          # gera assets/icons/*.png
python3 tools/construir_planilha.py    # gera CWSI_Planilha_Mestre.xlsx
```

Para uma versão menor durante testes: `LINHAS_CWSI=50 python3 tools/construir_planilha.py`
(o padrão são 1000 linhas de dados).

Depois de gerar, recalcule com o LibreOffice para gravar os valores em cache:

```bash
soffice --headless --convert-to xlsx --outdir . CWSI_Planilha_Mestre.xlsx
```

> Requer o pacote `libreoffice-calc` — só o `libreoffice-core` não abre planilhas.

### Arquivos

```
CWSI_Planilha_Mestre.xlsx      entregável
tools/construir_planilha.py    monta o .xlsx (layout, fórmulas, validações)
tools/conteudo.py              conteúdo: baselines, catálogo, procedimentos, formatos
tools/gerar_icones.py          desenha os ícones e as figuras de fluxo
assets/icons/                  18 PNGs gerados
```

---

## Sobre as figuras

Os ícones dos monitores são **desenhos esquemáticos próprios**, gerados por script, que
representam o formato físico do terminal (proporção de tela, teclas físicas, encoder rotativo).
Não reproduzem fotografias, logotipos nem marcas figurativas dos fabricantes; os nomes dos
modelos aparecem apenas como referência textual.

---

## Limitações assumidas

- Os coeficientes de baseline da aba 07 são de **Idso (1982)**, levantados em clima árido
  (Phoenix/AZ). Em clima úmido tendem a errar — **calibração local não é opcional** para uso
  operacional. O roteiro está na própria aba 07.
- Quando a estação não mede radiação líquida, `Rn` é estimado como
  `(1 − albedo)·Rs − onda longa` com a onda longa fixada em `01_Config`. É aproximação para
  meio-dia com céu limpo.
- O pareamento entre estação e dossel é por **carimbo de tempo exato**, sem tolerância.
  Registros sem par são marcados e excluídos — nunca interpolados em silêncio.
- O CWSI pressupõe que o sensor enxerga apenas dossel. Abaixo de ~60% de cobertura do solo o
  índice perde validade; a planilha alerta, mas não corrige.
