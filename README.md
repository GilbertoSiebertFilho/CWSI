# CWSI — Master Spreadsheet and Agricultural Monitor Guide

A spreadsheet for calculating **CWSI (Crop Water Stress Index)** from weather
station data plus canopy thermometry, combined with a **guide to moving files
between the office and the machine's monitor** (prescriptions, AB guidance
lines, work data).

**Deliverable:** [`CWSI_Planilha_Mestre.xlsx`](CWSI_Planilha_Mestre.xlsx)

---

## Two deliverables in this repository

| Folder | What it is |
|---|---|
| root + `tools/` | The **CWSI spreadsheet** and its monitor guide (this document) |
| [`platform/`](platform/) | The **AB Line Platform** — a web application answering *how do I get files into the monitor and how do I get the data back out*, by brand, display and software version |

The platform takes the spreadsheet's `09_Passo_a_Passo` and
`10_Base_Procedimentos` tabs and turns them into a product: pick equipment type
→ display → **software version** → what you want to do → how the data travels,
and get the file format, the exact folder on the USB stick, numbered clicks, how
to check it worked, and what usually goes wrong — with the terminal's icon and
ready to print.

**264 procedures across 23 displays**, covering USB, fourteen cloud platforms
and desktop software. Every procedure is a shareable link, every display has a
full printable handbook, and every card carries a button for sending back a
correction from the machine.

Generating AB lines is still there, but it is the supporting act.

```bash
cd platform && pip install -r requirements.txt && python3 run.py --seed && python3 run.py
```

Full documentation: [`platform/README.md`](platform/README.md).

> **Language note.** The platform is entirely in English. The spreadsheet
> itself and the two scripts that build it (`tools/conteudo.py`,
> `tools/construir_planilha.py`) are still in Portuguese, along with the tab
> names above. The icon generator and everything under `platform/` have been
> translated. Say the word and the spreadsheet follows.

---

## What the spreadsheet does

### Part 1 — CWSI calculation

Two methods, selectable in `01_Config`:

| Method | Formula | When to use it |
|---|---|---|
| **Empirical (Idso, 1981)** | `CWSI = (dT − dT_LL) / (dT_UL − dT_LL)`<br>`dT_LL = a + b·VPD`  ·  `dT_UL = a + b·VPG`<br>`VPG = es(Ta) − es(Ta + a)` | The default. Simple, but needs a locally calibrated baseline (a, b). |
| **Theoretical (Jackson, 1981)** | Energy balance with `ra`, `Rn`, `Δ`, `γ*` | When you have reliable net radiation and wind measurements. |

Every intermediate step stays visible column by column in `04_Calculo` — `es`,
`ea`, `VPD`, `Δ`, `γ`, `u2`, `dT`, `VPG`, both limits, `ra`, `γ*` — so any
result can be audited without opening a formula.

### Part 2 — Monitor guide

`09_Passo_a_Passo` builds the click sequence from six chained choices:

```
brand → equipment type → monitor → software version → objective → source/destination
```

and returns: **file format**, **exact folder on the USB stick**, **filesystem**,
**up to 10 numbered steps**, **cautions**, **common errors**, **confidence
level** and **source**.

The monitor and version lists are dependent (they change with the brand and
monitor chosen). When no procedure exists for the specific version, the
spreadsheet falls back to the `Todas as versoes` entry automatically.

> This is the part the platform supersedes. The spreadsheet holds 24 procedures,
> all tagged "all versions"; the platform holds 264 with a real version
> dimension. Use the spreadsheet for CWSI, the platform for monitors.

---

## Tabs

| Tab | Purpose |
|---|---|
| `00_Inicio` | Cover, index and usage guide |
| `01_Config` | **Start here.** Location, sensors, units, crop, filters, thresholds |
| `02_Estacao` | Raw weather station data (paste) |
| `03_Dossel` | Canopy temperature from the infrared sensor (paste) |
| `04_Calculo` | Calculation engine, step by step. Formulas only |
| `05_QC_Diagnostico` | Automatic panel plus a checklist of the 22 classic errors |
| `06_Resultados` | Daily CWSI and irrigation decision |
| `07_Baselines` | Coefficients a and b per crop, plus a local calibration routine |
| `08_Equipamentos` | Illustrated catalog of 14 monitors and terminals |
| `09_Passo_a_Passo` | Step-by-step generator |
| `10_Base_Procedimentos` | Procedure database — **this is where the guide grows** |
| `11_Formatos_Arquivos` | Shapefile, ISOXML, .agsetup, .agdata, AgData, Rx, TASKDATA… |
| `12_Fontes` | References, assumptions and limitations |

Colour convention: **yellow + blue** = you type · **grey** = formula · **orange**
= example row (delete it) · **red** = warning · **green** = check passed.

---

## Extending the guide

The generator on tab 09 is only a lookup into `10_Base_Procedimentos`. To cover
a new combination, add a row there (there are 20 blank pre-formatted rows with
dropdown lists already) — it appears on tab 09 by itself, with no formula
changes.

The **Confiança** column classifies each row:

- **Verificado** — format and folder path confirmed against the cited source.
- **Confirmar** — the structure is right, but the exact menu name changes
  between versions. Confirm it on the machine and update the row.

---

## Rebuilding the files

```bash
pip install openpyxl pillow
python3 tools/gerar_icones.py          # generates assets/icons/*.png
python3 tools/construir_planilha.py    # generates CWSI_Planilha_Mestre.xlsx
```

For a smaller version while testing: `LINHAS_CWSI=50 python3 tools/construir_planilha.py`
(the default is 1000 data rows).

After generating, recalculate with LibreOffice to write the cached values:

```bash
soffice --headless --convert-to xlsx --outdir . CWSI_Planilha_Mestre.xlsx
```

> Requires the `libreoffice-calc` package — `libreoffice-core` alone will not
> open spreadsheets.

### Files

```
CWSI_Planilha_Mestre.xlsx      the deliverable
tools/construir_planilha.py    builds the .xlsx (layout, formulas, validation)
tools/conteudo.py              content: baselines, catalog, procedures, formats
tools/gerar_icones.py          draws the icons and the flow figures
assets/icons/                  25 terminal icons in two variants, plus 4 figures
platform/                      the AB Line Platform (see platform/README.md)
```

---

## About the figures

The monitor icons are **our own schematic drawings**, generated by script, that
represent the physical shape of the terminal (screen proportion, physical keys,
rotary encoder). They do not reproduce photographs, logos or figurative marks of
the manufacturers; model names appear as text reference only.

Two variants are generated from one source: **captioned** for the spreadsheet,
and **caption-free** under `assets/icons/ui/` for the web platform, where the
model name is already real text beside the picture.

---

## Assumed limitations

- The baseline coefficients on tab 07 come from **Idso (1982)**, measured in an
  arid climate (Phoenix, AZ). In humid climates they tend to be wrong — **local
  calibration is not optional** for operational use. The routine is on tab 07
  itself.
- When the station does not measure net radiation, `Rn` is estimated as
  `(1 − albedo)·Rs − longwave`, with the longwave term fixed in `01_Config`.
  This is an approximation for midday under a clear sky.
- Station and canopy records are paired by **exact timestamp**, with no
  tolerance. Unpaired records are flagged and excluded — never silently
  interpolated.
- CWSI assumes the sensor sees canopy only. Below roughly 60% ground cover the
  index loses validity; the spreadsheet warns but does not correct for it.
