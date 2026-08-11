#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Constroi a planilha CWSI + Guia de Monitores (.xlsx).

Uso:  python3 tools/construir_planilha.py
Saida: CWSI_Planilha_Mestre.xlsx  (na raiz do repositorio)
"""

from __future__ import annotations

import os

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import ColorScaleRule, FormulaRule
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.drawing.image import Image as XLImage
from openpyxl.comments import Comment

import conteudo as C

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ICONES = os.path.join(RAIZ, "assets", "icons")
SAIDA = os.path.join(RAIZ, "CWSI_Planilha_Mestre.xlsx")

VERSAO = "1.0"
# Nº de linhas de dados pré-formatadas. Pode ser ajustado por variável de
# ambiente para testes rápidos:  LINHAS_CWSI=60 python3 construir_planilha.py
LINHAS = int(os.environ.get("LINHAS_CWSI", "500"))
L0 = 6                 # primeira linha de dados
L1 = L0 + LINHAS - 1   # ultima linha de dados (1005)

# ------------------------------------------------------------------ estilo --
FONTE = "Arial"
F_TITULO   = Font(name=FONTE, size=16, bold=True, color="FF1F3864")
F_SUB      = Font(name=FONTE, size=10, italic=True, color="FF595959")
F_SECAO    = Font(name=FONTE, size=11, bold=True, color="FFFFFFFF")
F_CAB      = Font(name=FONTE, size=10, bold=True, color="FFFFFFFF")
F_ROT      = Font(name=FONTE, size=10, bold=False, color="FF212529")
F_ROT_B    = Font(name=FONTE, size=10, bold=True, color="FF212529")
F_ENTRADA  = Font(name=FONTE, size=10, bold=True, color="FF0000FF")   # azul = digitar
F_FORMULA  = Font(name=FONTE, size=10, color="FF000000")              # preto = formula
F_LINK     = Font(name=FONTE, size=10, color="FF008000")              # verde = outra aba
F_NOTA     = Font(name=FONTE, size=9, italic=True, color="FF6C757D")
F_TXT      = Font(name=FONTE, size=10)
F_TXT_B    = Font(name=FONTE, size=10, bold=True)
F_PASSO    = Font(name=FONTE, size=10)
F_ALERTA   = Font(name=FONTE, size=10, bold=True, color="FF9C0006")

P_SECAO    = PatternFill("solid", fgColor="FF1F3864")
P_SECAO2   = PatternFill("solid", fgColor="FF2E5F8A")
P_CAB      = PatternFill("solid", fgColor="FF404E5C")
P_ENTRADA  = PatternFill("solid", fgColor="FFFFF2CC")   # amarelo = preencher
P_CALC     = PatternFill("solid", fgColor="FFF2F2F2")
P_EXEMPLO  = PatternFill("solid", fgColor="FFFCE4D6")   # laranja = linha exemplo
P_OK       = PatternFill("solid", fgColor="FFE2EFDA")
P_ALERTA   = PatternFill("solid", fgColor="FFFFC7CE")
P_DESTAQUE = PatternFill("solid", fgColor="FFDDEBF7")

_thin = Side(style="thin", color="FFBFBFBF")
B_CAIXA = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)

# --------------------------------------------------------------- unidades --
U_T   = ["°C", "°F"]
U_UR  = ["% (0-100)", "fração (0-1)"]
U_RS  = ["W/m²", "kW/m²", "MJ/m²/h"]
U_VEL = ["m/s", "km/h"]
U_P   = ["kPa", "hPa (mbar)", "Pa"]
METODOS = ["Empírico (Idso)", "Teórico (Jackson)"]
STATUS_QC = ["A verificar", "OK", "Corrigir", "Não se aplica"]

CULT_PERS = "PERSONALIZADO (calibracao local)"
VER_TODAS = "Todas as versoes"


# ------------------------------------------------------------- utilitarios --
def titulo(ws, texto, sub=None, largura=10):
    ws["A1"] = texto
    ws["A1"].font = F_TITULO
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=largura)
    if sub:
        ws["A2"] = sub
        ws["A2"].font = F_SUB
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=largura)
    ws.row_dimensions[1].height = 22


def secao(ws, linha, texto, largura=10, cor=P_SECAO):
    c = ws.cell(row=linha, column=1, value=texto)
    c.font = F_SECAO
    c.fill = cor
    c.alignment = Alignment(vertical="center", indent=1)
    for col in range(2, largura + 1):
        ws.cell(row=linha, column=col).fill = cor
    ws.merge_cells(start_row=linha, start_column=1, end_row=linha, end_column=largura)
    ws.row_dimensions[linha].height = 19


def rotulo(ws, linha, texto, col=1, nota=None):
    c = ws.cell(row=linha, column=col, value=texto)
    c.font = F_ROT
    c.alignment = Alignment(vertical="center", wrap_text=False)
    if nota:
        c.comment = Comment(nota, "Planilha CWSI")
    return c


def entrada(ws, linha, col, valor=None, fmt=None, nome=None, wb=None, aba=None):
    c = ws.cell(row=linha, column=col, value=valor)
    c.font = F_ENTRADA
    c.fill = P_ENTRADA
    c.border = B_CAIXA
    c.alignment = Alignment(horizontal="center", vertical="center")
    if fmt:
        c.number_format = fmt
    if nome and wb is not None and aba is not None:
        ref = f"'{aba}'!${get_column_letter(col)}${linha}"
        wb.defined_names.add(DefinedName(nome, attr_text=ref))
    return c


def calculada(ws, linha, col, formula, fmt=None, nome=None, wb=None, aba=None):
    c = ws.cell(row=linha, column=col, value=formula)
    c.font = F_FORMULA
    c.fill = P_CALC
    c.border = B_CAIXA
    c.alignment = Alignment(horizontal="center", vertical="center")
    if fmt:
        c.number_format = fmt
    if nome and wb is not None and aba is not None:
        ref = f"'{aba}'!${get_column_letter(col)}${linha}"
        wb.defined_names.add(DefinedName(nome, attr_text=ref))
    return c


def nota(ws, linha, texto, col=1, largura=10):
    c = ws.cell(row=linha, column=col, value=texto)
    c.font = F_NOTA
    c.alignment = Alignment(vertical="center", wrap_text=True)
    ws.merge_cells(start_row=linha, start_column=col, end_row=linha, end_column=largura)
    return c


def cabecalhos(ws, linha, textos, larguras=None, altura=30):
    for i, t in enumerate(textos, start=1):
        c = ws.cell(row=linha, column=i, value=t)
        c.font = F_CAB
        c.fill = P_CAB
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = B_CAIXA
    ws.row_dimensions[linha].height = altura
    if larguras:
        for i, w in enumerate(larguras, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w


def nome_global(wb, nome, aba, ref):
    wb.defined_names.add(DefinedName(nome, attr_text=f"'{aba}'!{ref}"))


def impressao(ws, paisagem=True, repetir_topo=None):
    """Ajusta a aba para caber na largura de uma página ao imprimir/gerar PDF."""
    ws.page_setup.orientation = "landscape" if paisagem else "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.5
    if repetir_topo:
        ws.print_title_rows = repetir_topo


# =========================================================================== #
def construir():
    wb = Workbook()
    wb.remove(wb.active)

    ws_inicio  = wb.create_sheet("00_Inicio")
    ws_cfg     = wb.create_sheet("01_Config")
    ws_est     = wb.create_sheet("02_Estacao")
    ws_dos     = wb.create_sheet("03_Dossel")
    ws_calc    = wb.create_sheet("04_Calculo")
    ws_qc      = wb.create_sheet("05_QC_Diagnostico")
    ws_res     = wb.create_sheet("06_Resultados")
    ws_bl      = wb.create_sheet("07_Baselines")
    ws_eq      = wb.create_sheet("08_Equipamentos")
    ws_passo   = wb.create_sheet("09_Passo_a_Passo")
    ws_db      = wb.create_sheet("10_Base_Procedimentos")
    ws_fmt     = wb.create_sheet("11_Formatos_Arquivos")
    ws_fontes  = wb.create_sheet("12_Fontes")
    ws_list    = wb.create_sheet("_Listas")

    montar_baselines(wb, ws_bl)
    montar_listas(wb, ws_list)
    montar_config(wb, ws_cfg)
    montar_estacao(wb, ws_est)
    montar_dossel(wb, ws_dos)
    montar_calculo(wb, ws_calc)
    montar_qc(wb, ws_qc)
    montar_resultados(wb, ws_res)
    montar_db(wb, ws_db)
    montar_equipamentos(wb, ws_eq)
    montar_passo(wb, ws_passo)
    montar_formatos(wb, ws_fmt)
    montar_fontes(wb, ws_fontes)
    montar_inicio(wb, ws_inicio)

    # Abas de leitura/impressão: cabem na largura de uma página A4
    impressao(ws_inicio, paisagem=False)
    impressao(ws_passo)
    impressao(ws_eq)
    impressao(ws_fmt, repetir_topo="4:4")
    impressao(ws_bl, paisagem=False, repetir_topo="6:6")
    impressao(ws_qc, repetir_topo="4:4")
    impressao(ws_res, repetir_topo="7:7")
    impressao(ws_fontes, paisagem=False)
    impressao(ws_cfg, paisagem=False)

    ws_list.sheet_state = "hidden"
    wb.active = 0
    wb.save(SAIDA)
    print(f"Gerado: {SAIDA}")


# --------------------------------------------------------- 07_Baselines ----
def montar_baselines(wb, ws):
    titulo(ws, "07 · Baselines não-estressadas (linha de base inferior)",
           "Tc − Ta = a + b × VPD   |   VPD em kPa, dT em °C, b negativo. "
           "Selecione a cultura na aba 01_Config.", largura=5)

    ws["A4"] = ("ATENÇÃO: os valores abaixo são de LITERATURA (Idso, 1982), obtidos em clima árido "
                "(Phoenix/AZ, EUA). Servem como ponto de partida. Para uso operacional a calibração "
                "LOCAL é obrigatória — meça Tc−Ta e VPD em uma parcela bem irrigada, ao meio-dia solar, "
                "em céu limpo, e ajuste a reta.")
    ws["A4"].font = F_ALERTA
    ws["A4"].fill = P_ALERTA
    ws["A4"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.merge_cells("A4:E4")
    ws.row_dimensions[4].height = 48

    cabecalhos(ws, 6, ["Cultura", "a  (°C)", "b  (°C/kPa)", "Fonte / observação", ""],
               larguras=[36, 12, 14, 46, 3])

    r = 7
    for cult, a, b, obs in C.BASELINES:
        ws.cell(row=r, column=1, value=cult).font = F_TXT
        if a is None:
            ws.cell(row=r, column=2, value="—").font = F_TXT
            ws.cell(row=r, column=3, value="—").font = F_TXT
        else:
            ws.cell(row=r, column=2, value=a).font = F_TXT
            ws.cell(row=r, column=3, value=b).font = F_TXT
        ws.cell(row=r, column=2).number_format = "0.00"
        ws.cell(row=r, column=3).number_format = "0.00"
        ws.cell(row=r, column=4, value=obs).font = F_NOTA
        for col in range(1, 5):
            ws.cell(row=r, column=col).border = B_CAIXA
            ws.cell(row=r, column=col).alignment = Alignment(
                horizontal="center" if col in (2, 3) else "left", vertical="center")
        r += 1
    ult = r - 1

    nome_global(wb, "BL_CULTURA", "07_Baselines", f"$A$7:$A${ult}")
    nome_global(wb, "BL_A", "07_Baselines", f"$B$7:$B${ult}")
    nome_global(wb, "BL_B", "07_Baselines", f"$C$7:$C${ult}")

    r += 2
    ws.cell(row=r, column=1, value="Como calibrar a baseline localmente (resumo):").font = F_TXT_B
    passos_cal = [
        "1. Escolha uma parcela SEM restrição hídrica (irrigada 24-48 h antes) e com cobertura total do solo.",
        "2. Meça Tc (termômetro infravermelho) e Ta/UR da estação, entre 12h e 15h solar, em céu limpo.",
        "3. Repita em vários dias, cobrindo uma faixa ampla de VPD (dias secos e úmidos).",
        "4. Plote dT = Tc − Ta (eixo Y) contra VPD (eixo X) e ajuste uma reta por mínimos quadrados.",
        "5. O intercepto é 'a' e a inclinação é 'b' (deve dar negativo). Use R² ≥ 0,80 como critério mínimo.",
        "6. Lance a e b nos campos de calibração local em 01_Config — eles têm prioridade sobre a literatura.",
    ]
    for p in passos_cal:
        r += 1
        ws.cell(row=r, column=1, value=p).font = F_TXT
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)

    ws.freeze_panes = "A7"


# ------------------------------------------------------------- _Listas ----
def montar_listas(wb, ws):
    ws["A1"] = "LISTAS DE APOIO — aba oculta. Alterar aqui muda as listas suspensas da planilha."
    ws["A1"].font = F_TXT_B

    marcas = []
    for m in C.MONITORES:
        if m[1] not in marcas:
            marcas.append(m[1])
    monitores_flat = [m[2] for m in C.MONITORES]

    def col_lista(col, cabecalho, itens, nome):
        ws.cell(row=3, column=col, value=cabecalho).font = F_TXT_B
        for i, v in enumerate(itens):
            ws.cell(row=4 + i, column=col, value=v).font = F_TXT
        if itens:
            nome_global(wb, nome, "_Listas",
                        f"${get_column_letter(col)}$4:${get_column_letter(col)}${3 + len(itens)}")
        ws.column_dimensions[get_column_letter(col)].width = 34

    col_lista(1,  "MARCAS",        marcas,            "LST_MARCAS")
    col_lista(2,  "MONITORES",     monitores_flat,    "LST_MONITORES")
    col_lista(3,  "OBJETIVOS",     C.OBJETIVOS,       "LST_OBJETIVOS")
    col_lista(4,  "ORIGENS",       C.ORIGENS,         "LST_ORIGENS")
    col_lista(5,  "TIPOS EQUIP.",  C.TIPOS_EQUIPAMENTO, "LST_TIPOS")
    col_lista(6,  "UNID. TEMP.",   U_T,               "LST_UT")
    col_lista(7,  "UNID. UR",      U_UR,              "LST_UUR")
    col_lista(8,  "UNID. RAD.",    U_RS,              "LST_URS")
    col_lista(9,  "UNID. VENTO",   U_VEL,             "LST_UVEL")
    col_lista(10, "UNID. PRESSÃO", U_P,               "LST_UP")
    col_lista(11, "MÉTODO CWSI",   METODOS,           "LST_METODO")
    col_lista(12, "STATUS QC",     STATUS_QC,         "LST_STATUS")

    # MON_<i> : monitores da marca i (i = posicao em LST_MARCAS)
    col = 15
    for i, marca in enumerate(marcas, start=1):
        itens = [m[2] for m in C.MONITORES if m[1] == marca]
        ws.cell(row=3, column=col, value=f"MON_{i} · {marca}").font = F_NOTA
        for k, v in enumerate(itens):
            ws.cell(row=4 + k, column=col, value=v).font = F_TXT
        nome_global(wb, f"MON_{i}", "_Listas",
                    f"${get_column_letter(col)}$4:${get_column_letter(col)}${3 + len(itens)}")
        ws.column_dimensions[get_column_letter(col)].width = 30
        col += 1

    # VER_<j> : versoes do monitor j (j = posicao em LST_MONITORES)
    col += 1
    for j, mon in enumerate(monitores_flat, start=1):
        itens = C.VERSOES.get(mon, [VER_TODAS])
        ws.cell(row=3, column=col, value=f"VER_{j}").font = F_NOTA
        for k, v in enumerate(itens):
            ws.cell(row=4 + k, column=col, value=v).font = F_TXT
        nome_global(wb, f"VER_{j}", "_Listas",
                    f"${get_column_letter(col)}$4:${get_column_letter(col)}${3 + len(itens)}")
        ws.column_dimensions[get_column_letter(col)].width = 26
        col += 1


# ------------------------------------------------------------ 01_Config ----
def montar_config(wb, ws):
    A = "01_Config"
    titulo(ws, "01 · Configuração do cálculo",
           "Preencha SOMENTE as células amarelas. As cinzas são calculadas — não digite nelas.",
           largura=6)
    ws.column_dimensions["A"].width = 46
    ws.column_dimensions["B"].width = 3
    ws.column_dimensions["C"].width = 26
    ws.column_dimensions["D"].width = 3
    ws.column_dimensions["E"].width = 62

    def linha_cfg(r, rot, valor=None, fmt=None, nome=None, calc=None, ajuda=None):
        rotulo(ws, r, rot)
        if calc is not None:
            calculada(ws, r, 3, calc, fmt=fmt, nome=nome, wb=wb, aba=A)
        else:
            entrada(ws, r, 3, valor, fmt=fmt, nome=nome, wb=wb, aba=A)
        if ajuda:
            c = ws.cell(row=r, column=5, value=ajuda)
            c.font = F_NOTA
            c.alignment = Alignment(vertical="center", wrap_text=True)

    # A — identificacao
    secao(ws, 4, "A · IDENTIFICAÇÃO", largura=6)
    linha_cfg(5, "Fazenda", "", ajuda="Texto livre.")
    linha_cfg(6, "Talhão / parcela", "")
    linha_cfg(7, "Cultura", "Milho", nome="CFG_CULTURA",
              ajuda="Define a baseline (a, b) puxada da aba 07_Baselines.")
    linha_cfg(8, "Safra / ano agrícola", "")
    linha_cfg(9, "Responsável técnico", "")
    linha_cfg(10, "Data do relatório", None, fmt="dd/mm/yyyy")

    # B — local
    secao(ws, 12, "B · LOCALIZAÇÃO", largura=6, cor=P_SECAO2)
    linha_cfg(13, "Latitude (graus decimais, sul negativo)", None, fmt="0.0000", nome="CFG_LAT")
    linha_cfg(14, "Longitude (graus decimais, oeste negativo)", None, fmt="0.0000", nome="CFG_LON")
    linha_cfg(15, "Altitude (m)", 500, fmt="0", nome="CFG_ALT",
              ajuda="Usada para estimar a pressão atmosférica quando a estação não mede pressão.")
    linha_cfg(16, "Pressão atmosférica padrão (kPa)", nome="CFG_P", fmt="0.00",
              calc="=101.3*((293-0.0065*CFG_ALT)/293)^5.26",
              ajuda="FAO-56 eq. 7. Só é usada nas linhas em que a coluna de pressão da estação estiver vazia.")

    # C — estacao
    secao(ws, 18, "C · ESTAÇÃO METEOROLÓGICA", largura=6, cor=P_SECAO2)
    linha_cfg(19, "Marca / modelo da estação", "")
    linha_cfg(20, "Altura do anemômetro — zm (m)", 2.0, fmt="0.00", nome="CFG_ZANEM",
              ajuda="Altura REAL do sensor de vento. Erro clássico: usar vento de 10 m como se fosse de 2 m.")
    linha_cfg(21, "Altura do sensor de T/UR — zh (m)", 2.0, fmt="0.00", nome="CFG_ZAR")
    linha_cfg(22, "Intervalo de registro (min)", 15, fmt="0")
    linha_cfg(23, "Nº de série / identificação", "")
    linha_cfg(24, "Data da última calibração", None, fmt="dd/mm/yyyy")
    linha_cfg(25, "Distância da estação até o talhão (m)", None, fmt="0",
              ajuda="Acima de ~1 km em relevo/manejo diferente, o dado deixa de representar o talhão.")

    # D — sensor de dossel
    secao(ws, 27, "D · SENSOR DE TEMPERATURA DE DOSSEL (IRT / termal)", largura=6, cor=P_SECAO2)
    linha_cfg(28, "Marca / modelo do sensor", "")
    linha_cfg(29, "Emissividade configurada no sensor", 0.98, fmt="0.000",
              ajuda="Dossel vegetal: 0,97–0,99. Emissividade errada desloca Tc em até 1–2 °C.")
    linha_cfg(30, "Ângulo de visada em relação à horizontal (°)", 45, fmt="0",
              ajuda="Visada oblíqua (~45°) reduz a captação de solo exposto.")
    linha_cfg(31, "Altura de instalação (m)", 1.5, fmt="0.00")
    linha_cfg(32, "Campo de visão — FOV (°)", 20, fmt="0")
    linha_cfg(33, "Cobertura do solo pela cultura (%)", None, fmt="0",
              ajuda="Abaixo de ~60% de cobertura o sensor 'enxerga' solo e o CWSI perde validade.")
    linha_cfg(34, "Data da última calibração", None, fmt="dd/mm/yyyy")

    # E — unidades
    secao(ws, 36, "E · UNIDADES DOS DADOS DE ENTRADA (abas 02 e 03)", largura=6, cor=P_SECAO2)
    linha_cfg(37, "Temperatura do ar", U_T[0], nome="CFG_UT")
    linha_cfg(38, "Umidade relativa", U_UR[0], nome="CFG_UUR")
    linha_cfg(39, "Radiação solar", U_RS[0], nome="CFG_URS")
    linha_cfg(40, "Velocidade do vento", U_VEL[0], nome="CFG_UVEL")
    linha_cfg(41, "Pressão atmosférica", U_P[0], nome="CFG_UP")
    linha_cfg(42, "Temperatura de dossel", U_T[0], nome="CFG_UTC")
    nota(ws, 43, "A conversão é automática. Confira a unidade real exportada pela sua estação — "
                 "misturar hPa com kPa ou % com fração é a origem mais comum de CWSI absurdo.",
         col=1, largura=6)

    # F — metodo e baseline
    secao(ws, 45, "F · MÉTODO E BASELINE", largura=6, cor=P_SECAO2)
    linha_cfg(46, "Método do CWSI", METODOS[0], nome="CFG_METODO",
              ajuda="Empírico (Idso): simples, precisa de baseline calibrada. "
                    "Teórico (Jackson): usa balanço de energia, precisa de Rn, vento e resistências.")
    linha_cfg(47, "a — literatura (cultura selecionada)", fmt="0.00", nome="CFG_A_LIT",
              calc='=IFERROR(IF(ISNUMBER(INDEX(BL_A,MATCH(CFG_CULTURA,BL_CULTURA,0))),'
                   'INDEX(BL_A,MATCH(CFG_CULTURA,BL_CULTURA,0)),""),"")')
    linha_cfg(48, "b — literatura (cultura selecionada)", fmt="0.00", nome="CFG_B_LIT",
              calc='=IFERROR(IF(ISNUMBER(INDEX(BL_B,MATCH(CFG_CULTURA,BL_CULTURA,0))),'
                   'INDEX(BL_B,MATCH(CFG_CULTURA,BL_CULTURA,0)),""),"")')
    linha_cfg(49, "a — CALIBRAÇÃO LOCAL (preencha para sobrepor)", None, fmt="0.00", nome="CFG_A_MAN",
              ajuda="Se preenchido, tem prioridade sobre a literatura. É o caminho recomendado.")
    linha_cfg(50, "b — CALIBRAÇÃO LOCAL (preencha para sobrepor)", None, fmt="0.00", nome="CFG_B_MAN")
    linha_cfg(51, "a EM USO", fmt="0.00", nome="CFG_A",
              calc='=IF(CFG_A_MAN<>"",CFG_A_MAN,CFG_A_LIT)')
    linha_cfg(52, "b EM USO", fmt="0.00", nome="CFG_B",
              calc='=IF(CFG_B_MAN<>"",CFG_B_MAN,CFG_B_LIT)')
    linha_cfg(53, "R² da calibração local", None, fmt="0.00",
              ajuda="Registro de qualidade. Abaixo de 0,80 a baseline não é confiável.")

    # G — filtros
    secao(ws, 55, "G · FILTROS DE QUALIDADE (definem quais registros entram no CWSI)", largura=6, cor=P_SECAO2)
    linha_cfg(56, "Hora inicial da janela (h local)", 13, fmt="0", nome="CFG_HINI",
              ajuda="O CWSI só é válido perto do meio-dia solar. Ajuste ao seu fuso e longitude.")
    linha_cfg(57, "Hora final da janela (h local)", 15, fmt="0", nome="CFG_HFIM")
    linha_cfg(58, "Radiação solar mínima (W/m²)", 400, fmt="0", nome="CFG_RSMIN",
              ajuda="Corta nuvem. Abaixo disso a baseline de Idso não se aplica.")
    linha_cfg(59, "Vento mínimo (m/s)", 0.5, fmt="0.0", nome="CFG_UMIN",
              ajuda="Calmaria desacopla o dossel do ar e infla o dT.")
    linha_cfg(60, "Vento máximo (m/s)", 6.0, fmt="0.0", nome="CFG_UMAX",
              ajuda="Vento forte comprime o dT e mascara o estresse.")
    linha_cfg(61, "VPD mínimo (kPa)", 0.5, fmt="0.0", nome="CFG_VPDMIN",
              ajuda="Com VPD muito baixo o denominador do CWSI colapsa.")
    linha_cfg(62, "Denominador mínimo |dT_UL − dT_LL| (°C)", 0.5, fmt="0.0", nome="CFG_DENOMMIN",
              ajuda="Trava de segurança: evita divisão por número quase zero, que gera CWSI gigante.")

    # H — parametros fisicos
    secao(ws, 64, "H · PARÂMETROS FÍSICOS (usados no método teórico de Jackson)", largura=6, cor=P_SECAO2)
    linha_cfg(65, "Altura do dossel — h (m)", 1.00, fmt="0.00", nome="CFG_HDOSSEL",
              ajuda="Precisa ser MENOR que a altura do anemômetro; caso contrário ra não tem solução.")
    linha_cfg(66, "Albedo do dossel", 0.23, fmt="0.00", nome="CFG_ALBEDO")
    linha_cfg(67, "Onda longa líquida estimada (W/m²)", 60, fmt="0", nome="CFG_RNL",
              ajuda="Usada só quando a estação NÃO mede Rn: Rn ≈ (1−albedo)·Rs − este valor. "
                    "Aproximação para meio-dia com céu limpo. Se você mede Rn, preencha a coluna G da aba 02.")
    linha_cfg(68, "Resistência do dossel — rc (s/m)", 50, fmt="0", nome="CFG_RC",
              ajuda="Resistência mínima com o dossel bem hidratado. Varia com a cultura (25–100 s/m).")
    linha_cfg(69, "ρ·cp — capacidade térmica do ar (J m⁻³ °C⁻¹)", 1200, fmt="0", nome="CFG_RHOCP")
    linha_cfg(70, "Constante de von Kármán — k", 0.41, fmt="0.00", nome="CFG_K")

    # I — decisao
    secao(ws, 72, "I · LIMIARES DE DECISÃO DE IRRIGAÇÃO", largura=6, cor=P_SECAO2)
    linha_cfg(73, "CWSI para IRRIGAR (limiar superior)", 0.35, fmt="0.00", nome="CFG_LIMIAR",
              ajuda="Valor de partida na literatura para grãos. Ajuste ao seu sistema e cultura.")
    linha_cfg(74, "CWSI para ATENÇÃO (limiar inferior)", 0.22, fmt="0.00", nome="CFG_LIMIAR_AT")

    # validacoes
    dvs = [
        ("BL_CULTURA", "C7"), ("LST_UT", "C37"), ("LST_UUR", "C38"), ("LST_URS", "C39"),
        ("LST_UVEL", "C40"), ("LST_UP", "C41"), ("LST_UT", "C42"), ("LST_METODO", "C46"),
    ]
    for nome_lista, celula in dvs:
        dv = DataValidation(type="list", formula1=f"={nome_lista}", allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(ws[celula])

    # legenda de cores
    secao(ws, 76, "LEGENDA DE CORES", largura=6, cor=P_SECAO)
    leg = [
        ("Amarelo + texto azul", "Célula de ENTRADA — é aqui que você digita.", P_ENTRADA, F_ENTRADA),
        ("Cinza + texto preto", "Célula CALCULADA — não digite, é fórmula.", P_CALC, F_FORMULA),
        ("Laranja", "Linha de EXEMPLO — apague antes de usar de verdade.", P_EXEMPLO, F_TXT),
        ("Vermelho", "Alerta / erro detectado.", P_ALERTA, F_ALERTA),
        ("Verde", "Verificação aprovada.", P_OK, F_TXT),
    ]
    r = 77
    for cor_nome, desc, fill, fnt in leg:
        c = ws.cell(row=r, column=1, value=cor_nome)
        c.fill = fill
        c.font = fnt
        c.border = B_CAIXA
        c.alignment = Alignment(horizontal="center")
        d = ws.cell(row=r, column=3, value=desc)
        d.font = F_TXT
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)
        r += 1

    ws.freeze_panes = "A4"


# ----------------------------------------------------------- 02_Estacao ----
def montar_estacao(wb, ws):
    titulo(ws, "02 · Dados da estação meteorológica (COLE AQUI)",
           "Uma linha por registro. A coluna A (data/hora) é a CHAVE que liga esta aba à aba 03_Dossel — "
           "os carimbos de tempo precisam ser idênticos.", largura=9)

    cabecalhos(ws, 4, [
        "Data e hora", "Temperatura do ar", "Umidade relativa", "Radiação solar",
        "Velocidade do vento", "Pressão atmosférica\n(opcional)",
        "Radiação líquida Rn\n(opcional)", "Precipitação\n(opcional)", "Observações",
    ], larguras=[19, 15, 15, 15, 15, 16, 16, 14, 30], altura=40)

    unid = ["dd/mm/aaaa hh:mm", "conforme 01_Config", "conforme 01_Config", "conforme 01_Config",
            "conforme 01_Config", "conforme 01_Config", "W/m²", "mm", ""]
    for i, u in enumerate(unid, start=1):
        c = ws.cell(row=5, column=i, value=u)
        c.font = F_NOTA
        c.fill = P_DESTAQUE
        c.alignment = Alignment(horizontal="center")
        c.border = B_CAIXA

    exemplo = ["2026-01-15 13:30", 31.5, 48, 820, 2.4, None, None, 0, "LINHA DE EXEMPLO — APAGUE"]
    from datetime import datetime
    exemplo[0] = datetime(2026, 1, 15, 13, 30)
    for i, v in enumerate(exemplo, start=1):
        c = ws.cell(row=L0, column=i, value=v)
        c.fill = P_EXEMPLO
        c.font = F_TXT
        c.border = B_CAIXA
        c.alignment = Alignment(horizontal="center")
    ws.cell(row=L0, column=1).number_format = "dd/mm/yyyy hh:mm"

    for r in range(L0 + 1, L1 + 1):
        for col in range(1, 10):
            c = ws.cell(row=r, column=col)
            c.font = F_TXT
            c.border = B_CAIXA
            c.alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=1).number_format = "dd/mm/yyyy hh:mm"

    ws.freeze_panes = "B6"


# ------------------------------------------------------------ 03_Dossel ----
def montar_dossel(wb, ws):
    titulo(ws, "03 · Temperatura de dossel — sensor infravermelho (COLE AQUI)",
           "A data/hora precisa bater EXATAMENTE com a da aba 02_Estacao. Se os equipamentos gravam em "
           "intervalos diferentes, agregue antes de colar — a aba 05 avisa quando não encontra o par.",
           largura=7)

    cabecalhos(ws, 4, [
        "Data e hora", "Tc sensor 1", "Tc sensor 2", "Tc sensor 3", "Tc sensor 4",
        "Tc média (°C)\n[calculada]", "Observações",
    ], larguras=[19, 13, 13, 13, 13, 17, 34], altura=40)

    unid = ["dd/mm/aaaa hh:mm", "conforme 01_Config", "conforme 01_Config",
            "conforme 01_Config", "conforme 01_Config", "convertida para °C", ""]
    for i, u in enumerate(unid, start=1):
        c = ws.cell(row=5, column=i, value=u)
        c.font = F_NOTA
        c.fill = P_DESTAQUE
        c.alignment = Alignment(horizontal="center")
        c.border = B_CAIXA

    from datetime import datetime
    ws.cell(row=L0, column=1, value=datetime(2026, 1, 15, 13, 30))
    ws.cell(row=L0, column=2, value=33.1)
    ws.cell(row=L0, column=7, value="LINHA DE EXEMPLO — APAGUE")

    for r in range(L0, L1 + 1):
        for col in range(1, 8):
            c = ws.cell(row=r, column=col)
            c.font = F_TXT
            c.border = B_CAIXA
            c.alignment = Alignment(horizontal="center")
            if r == L0:
                c.fill = P_EXEMPLO
        ws.cell(row=r, column=1).number_format = "dd/mm/yyyy hh:mm"
        f = ws.cell(row=r, column=6)
        f.value = (f'=IF(COUNT($B{r}:$E{r})=0,"",'
                   f'IF(CFG_UTC="°F",(AVERAGE($B{r}:$E{r})-32)*5/9,AVERAGE($B{r}:$E{r})))')
        f.font = F_FORMULA
        f.number_format = "0.00"
        if r != L0:
            f.fill = P_CALC

    ws.freeze_panes = "B6"


# ----------------------------------------------------------- 04_Calculo ----
COLS_CALC = [
    ("Data e hora", 18, "dd/mm/yyyy hh:mm"),
    ("Ta\n(°C)", 9, "0.00"),
    ("UR\n(%)", 9, "0.0"),
    ("Rs\n(W/m²)", 10, "0"),
    ("u\n(m/s)", 9, "0.00"),
    ("P\n(kPa)", 9, "0.00"),
    ("Rn\n(W/m²)", 10, "0"),
    ("Tc\n(°C)", 9, "0.00"),
    ("Par\ndossel", 10, None),
    ("es(Ta)\n(kPa)", 10, "0.000"),
    ("ea\n(kPa)", 10, "0.000"),
    ("VPD\n(kPa)", 10, "0.000"),
    ("Δ\n(kPa/°C)", 11, "0.000"),
    ("γ\n(kPa/°C)", 11, "0.0000"),
    ("u2\n(m/s)", 9, "0.00"),
    ("dT medido\n(°C)", 11, "0.00"),
    ("a\n(°C)", 9, "0.00"),
    ("b\n(°C/kPa)", 10, "0.00"),
    ("VPG\n(kPa)", 10, "0.000"),
    ("dT_LL emp.\n(°C)", 12, "0.00"),
    ("dT_UL emp.\n(°C)", 12, "0.00"),
    ("CWSI\nempírico", 11, "0.000"),
    ("ra\n(s/m)", 10, "0.0"),
    ("γ*\n(kPa/°C)", 11, "0.000"),
    ("dT_LL teór.\n(°C)", 12, "0.00"),
    ("dT_UL teór.\n(°C)", 12, "0.00"),
    ("CWSI\nteórico", 11, "0.000"),
    ("F. janela", 9, "0"),
    ("F. radiação", 10, "0"),
    ("F. vento", 9, "0"),
    ("F. VPD", 9, "0"),
    ("F. denom.", 10, "0"),
    ("F. dados", 9, "0"),
    ("VÁLIDO", 9, "0"),
    ("CWSI bruto", 11, "0.000"),
    ("CWSI [0–1]", 11, "0.000"),
    ("Status do registro", 34, None),
]


def montar_calculo(wb, ws):
    titulo(ws, "04 · Motor de cálculo do CWSI",
           "Aba 100% de fórmulas — NÃO DIGITE AQUI. Cada etapa intermediária fica visível de propósito, "
           "para que qualquer resultado possa ser auditado passo a passo.", largura=12)

    grupos = [
        (1, 1, "ENTRADA", "FF6C757D"),
        (2, 9, "DADOS CONVERTIDOS E PAREADOS", "FF2E5F8A"),
        (10, 15, "PSICROMETRIA", "FF1F6E5C"),
        (16, 22, "MÉTODO EMPÍRICO (IDSO)", "FF8A5A2B"),
        (23, 27, "MÉTODO TEÓRICO (JACKSON)", "FF5B3E90"),
        (28, 34, "FILTROS DE QUALIDADE", "FF9C4A4A"),
        (35, 37, "RESULTADO", "FF1F3864"),
    ]
    for c0, c1, nome, cor in grupos:
        cel = ws.cell(row=3, column=c0, value=nome)
        cel.font = Font(name=FONTE, size=9, bold=True, color="FFFFFFFF")
        cel.fill = PatternFill("solid", fgColor=cor)
        cel.alignment = Alignment(horizontal="center", vertical="center")
        for cc in range(c0, c1 + 1):
            ws.cell(row=3, column=cc).fill = PatternFill("solid", fgColor=cor)
        if c1 > c0:
            ws.merge_cells(start_row=3, start_column=c0, end_row=3, end_column=c1)

    cabecalhos(ws, 4, [c[0] for c in COLS_CALC],
               larguras=[c[1] for c in COLS_CALC], altura=34)

    ref = [
        "← aba 02_Estacao", "conversão", "conversão", "conversão", "conversão",
        "estação ou altitude", "medido ou estimado", "PROCV em 03_Dossel", "controle",
        "0,6108·e^(17,27T/(T+237,3))", "es·UR/100", "es−ea", "4098·es/(T+237,3)²",
        "0,000665·P", "FAO-56 eq.47", "Tc−Ta", "01_Config", "01_Config",
        "es(Ta)−es(Ta+a)", "a+b·VPD", "a+b·VPG", "(dT−LL)/(UL−LL)",
        "FAO-56", "γ(1+rc/ra)", "Jackson (1981)", "ra·Rn/(ρcp)", "(dT−LL)/(UL−LL)",
        "hora", "Rs", "vento", "VPD", "|UL−LL|", "completos", "E(todos)",
        "método escolhido", "limitado a [0;1]", "motivo",
    ]
    for i, t in enumerate(ref, start=1):
        c = ws.cell(row=5, column=i, value=t)
        c.font = F_NOTA
        c.fill = P_DESTAQUE
        c.alignment = Alignment(horizontal="center", wrap_text=True)
        c.border = B_CAIXA
    ws.row_dimensions[5].height = 26

    est = "'02_Estacao'"
    dosA = f"'03_Dossel'!$A${L0}:$A${L1}"
    dosF = f"'03_Dossel'!$F${L0}:$F${L1}"
    idx = f"INDEX({dosF},MATCH($A{{r}},{dosA},0))"

    for r in range(L0, L1 + 1):
        f = {}
        f[1]  = f"=IF({est}!A{r}=\"\",\"\",{est}!A{r})"
        f[2]  = (f"=IF(OR($A{r}=\"\",{est}!B{r}=\"\"),\"\","
                 f"IF(CFG_UT=\"{U_T[1]}\",({est}!B{r}-32)*5/9,{est}!B{r}))")
        f[3]  = (f"=IF(OR($A{r}=\"\",{est}!C{r}=\"\"),\"\","
                 f"IF(CFG_UUR=\"{U_UR[1]}\",{est}!C{r}*100,{est}!C{r}))")
        f[4]  = (f"=IF(OR($A{r}=\"\",{est}!D{r}=\"\"),\"\","
                 f"IF(CFG_URS=\"{U_RS[2]}\",{est}!D{r}*277.7778,"
                 f"IF(CFG_URS=\"{U_RS[1]}\",{est}!D{r}*1000,{est}!D{r})))")
        f[5]  = (f"=IF(OR($A{r}=\"\",{est}!E{r}=\"\"),\"\","
                 f"IF(CFG_UVEL=\"{U_VEL[1]}\",{est}!E{r}/3.6,{est}!E{r}))")
        f[6]  = (f"=IF($A{r}=\"\",\"\",IF({est}!F{r}=\"\",CFG_P,"
                 f"IF(CFG_UP=\"{U_P[1]}\",{est}!F{r}/10,"
                 f"IF(CFG_UP=\"{U_P[2]}\",{est}!F{r}/1000,{est}!F{r}))))")
        f[7]  = (f"=IF($A{r}=\"\",\"\",IF({est}!G{r}<>\"\",{est}!G{r},"
                 f"IF($D{r}=\"\",\"\",(1-CFG_ALBEDO)*$D{r}-CFG_RNL)))")
        f[8]  = (f"=IF($A{r}=\"\",\"\",IFERROR(IF(ISNUMBER({idx.format(r=r)}),"
                 f"{idx.format(r=r)},\"\"),\"\"))")
        f[9]  = f"=IF($A{r}=\"\",\"\",IF($H{r}=\"\",\"SEM PAR\",\"OK\"))"
        f[10] = f"=IF($B{r}=\"\",\"\",0.6108*EXP(17.27*$B{r}/($B{r}+237.3)))"
        f[11] = f"=IF(OR($J{r}=\"\",$C{r}=\"\"),\"\",$J{r}*$C{r}/100)"
        f[12] = f"=IF(OR($J{r}=\"\",$K{r}=\"\"),\"\",$J{r}-$K{r})"
        f[13] = f"=IF($J{r}=\"\",\"\",4098*$J{r}/($B{r}+237.3)^2)"
        f[14] = f"=IF($F{r}=\"\",\"\",0.000665*$F{r})"
        f[15] = f"=IF($E{r}=\"\",\"\",IFERROR($E{r}*4.87/LN(67.8*CFG_ZANEM-5.42),\"\"))"
        f[16] = f"=IF(OR($B{r}=\"\",$H{r}=\"\"),\"\",$H{r}-$B{r})"
        f[17] = f"=IF($A{r}=\"\",\"\",IF(ISNUMBER(CFG_A),CFG_A,\"\"))"
        f[18] = f"=IF($A{r}=\"\",\"\",IF(ISNUMBER(CFG_B),CFG_B,\"\"))"
        f[19] = (f"=IF(OR($B{r}=\"\",$Q{r}=\"\"),\"\","
                 f"0.6108*EXP(17.27*$B{r}/($B{r}+237.3))"
                 f"-0.6108*EXP(17.27*($B{r}+$Q{r})/($B{r}+$Q{r}+237.3)))")
        f[20] = f"=IF(OR($Q{r}=\"\",$R{r}=\"\",$L{r}=\"\"),\"\",$Q{r}+$R{r}*$L{r})"
        f[21] = f"=IF(OR($Q{r}=\"\",$R{r}=\"\",$S{r}=\"\"),\"\",$Q{r}+$R{r}*$S{r})"
        # IFs ANINHADOS de propósito: OR() avalia TODOS os argumentos, então
        # ABS(""-"") dispararia #VALOR! mesmo com a célula vazia detectada antes.
        f[22] = (f"=IF(OR($P{r}=\"\",$T{r}=\"\",$U{r}=\"\"),\"\","
                 f"IF(ABS($U{r}-$T{r})<CFG_DENOMMIN,\"\",($P{r}-$T{r})/($U{r}-$T{r})))")
        f[23] = (f"=IF($E{r}=\"\",\"\",IFERROR("
                 f"LN((CFG_ZANEM-0.67*CFG_HDOSSEL)/(0.123*CFG_HDOSSEL))"
                 f"*LN((CFG_ZAR-0.67*CFG_HDOSSEL)/(0.0123*CFG_HDOSSEL))"
                 f"/(CFG_K^2*$E{r}),\"\"))")
        f[24] = f"=IF(OR($N{r}=\"\",$W{r}=\"\",$W{r}=0),\"\",$N{r}*(1+CFG_RC/$W{r}))"
        f[25] = (f"=IF(OR($X{r}=\"\",$M{r}=\"\",$G{r}=\"\",$L{r}=\"\"),\"\","
                 f"($W{r}*$G{r}/CFG_RHOCP)*($X{r}/($M{r}+$X{r}))-$L{r}/($M{r}+$X{r}))")
        f[26] = f"=IF(OR($W{r}=\"\",$G{r}=\"\"),\"\",$W{r}*$G{r}/CFG_RHOCP)"
        f[27] = (f"=IF(OR($P{r}=\"\",$Y{r}=\"\",$Z{r}=\"\"),\"\","
                 f"IF(ABS($Z{r}-$Y{r})<CFG_DENOMMIN,\"\",($P{r}-$Y{r})/($Z{r}-$Y{r})))")
        f[28] = (f"=IF($A{r}=\"\",\"\",IF(AND(HOUR($A{r})>=CFG_HINI,HOUR($A{r})<CFG_HFIM),1,0))")
        f[29] = f"=IF($A{r}=\"\",\"\",IF($D{r}=\"\",0,IF($D{r}>=CFG_RSMIN,1,0)))"
        f[30] = (f"=IF($A{r}=\"\",\"\",IF($E{r}=\"\",0,"
                 f"IF(AND($E{r}>=CFG_UMIN,$E{r}<=CFG_UMAX),1,0)))")
        f[31] = f"=IF($A{r}=\"\",\"\",IF($L{r}=\"\",0,IF($L{r}>=CFG_VPDMIN,1,0)))"
        f[32] = (f"=IF($A{r}=\"\",\"\",IF(CFG_METODO=\"{METODOS[1]}\","
                 f"IF(OR($Y{r}=\"\",$Z{r}=\"\"),0,IF(ABS($Z{r}-$Y{r})>=CFG_DENOMMIN,1,0)),"
                 f"IF(OR($T{r}=\"\",$U{r}=\"\"),0,IF(ABS($U{r}-$T{r})>=CFG_DENOMMIN,1,0))))")
        f[33] = (f"=IF($A{r}=\"\",\"\",IF(AND($B{r}<>\"\",$C{r}<>\"\",$H{r}<>\"\"),1,0))")
        f[34] = (f"=IF($A{r}=\"\",\"\",IF(AND($AB{r}=1,$AC{r}=1,$AD{r}=1,$AE{r}=1,"
                 f"$AF{r}=1,$AG{r}=1),1,0))")
        f[35] = (f"=IF($A{r}=\"\",\"\",IF($AH{r}<>1,\"\","
                 f"IF(CFG_METODO=\"{METODOS[1]}\",$AA{r},$V{r})))")
        f[36] = f"=IF($AI{r}=\"\",\"\",MAX(0,MIN(1,$AI{r})))"
        f[37] = (f"=IF($A{r}=\"\",\"\","
                 f"IF($AG{r}=0,\"Dados incompletos\"&IF($I{r}=\"SEM PAR\","
                 f"\" — sem Tc no mesmo horário\",\"\"),"
                 f"IF($AB{r}=0,\"Fora da janela horária\","
                 f"IF($AC{r}=0,\"Radiação abaixo do mínimo\","
                 f"IF($AD{r}=0,\"Vento fora da faixa\","
                 f"IF($AE{r}=0,\"VPD abaixo do mínimo\","
                 f"IF($AF{r}=0,\"Denominador muito pequeno\",\"VÁLIDO\")))))))")

        for col, formula in f.items():
            c = ws.cell(row=r, column=col, value=formula)
            c.font = F_FORMULA
            c.border = B_CAIXA
            c.alignment = Alignment(horizontal="center")
            fmt = COLS_CALC[col - 1][2]
            if fmt:
                c.number_format = fmt
        ws.cell(row=r, column=37).alignment = Alignment(horizontal="left", indent=1)

    ws.conditional_formatting.add(
        f"AJ{L0}:AJ{L1}",
        ColorScaleRule(start_type="num", start_value=0, start_color="FF63BE7B",
                       mid_type="num", mid_value=0.5, mid_color="FFFFEB84",
                       end_type="num", end_value=1, end_color="FFF8696B"))
    ws.conditional_formatting.add(
        f"AK{L0}:AK{L1}",
        FormulaRule(formula=[f'$AK{L0}="VÁLIDO"'], fill=P_OK, font=Font(name=FONTE, size=10)))
    ws.conditional_formatting.add(
        f"AI{L0}:AI{L1}",
        FormulaRule(formula=[f'AND($AI{L0}<>"",OR($AI{L0}<0,$AI{L0}>1))'],
                    fill=P_ALERTA, font=F_ALERTA))

    ws.freeze_panes = "B6"


# ------------------------------------------------------ 05_QC_Diagnostico --
def montar_qc(wb, ws):
    titulo(ws, "05 · Diagnóstico e controle de qualidade",
           "Painel automático + checklist dos erros que mais aparecem em planilhas de CWSI. "
           "Rode este painel ANTES de olhar qualquer resultado.", largura=6)
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 52
    ws.column_dimensions["C"].width = 17
    ws.column_dimensions["D"].width = 58
    ws.column_dimensions["E"].width = 17
    ws.column_dimensions["F"].width = 30

    rngA  = f"'04_Calculo'!$A${L0}:$A${L1}"
    rngI  = f"'04_Calculo'!$I${L0}:$I${L1}"
    rngAH = f"'04_Calculo'!$AH${L0}:$AH${L1}"
    rngAI = f"'04_Calculo'!$AI${L0}:$AI${L1}"
    rngAJ = f"'04_Calculo'!$AJ${L0}:$AJ${L1}"
    rngL  = f"'04_Calculo'!$L${L0}:$L${L1}"
    rngP  = f"'04_Calculo'!$P${L0}:$P${L1}"

    secao(ws, 4, "PAINEL AUTOMÁTICO — números que revelam o estado dos dados", largura=6)
    painel = [
        ("Registros carregados (com data/hora)", f"=COUNT({rngA})", "0",
         "Total de linhas lidas da aba 02_Estacao."),
        ("Registros com Tc pareada", f'=COUNTIF({rngI},"OK")', "0",
         "Encontraram temperatura de dossel no MESMO carimbo de tempo."),
        ("Registros SEM par de dossel", f'=COUNTIF({rngI},"SEM PAR")', "0",
         "Se este número for alto, os relógios/intervalos das duas fontes não batem. "
         "É o problema nº 1 em planilhas de CWSI."),
        ("Registros VÁLIDOS (passaram em todos os filtros)", f"=COUNTIF({rngAH},1)", "0",
         "É a base real do seu CWSI."),
        ("Aproveitamento dos dados",
         f'=IFERROR(COUNTIF({rngAH},1)/COUNT({rngA}),"")', "0.0%",
         "Abaixo de ~10% revise a janela horária e o limite de radiação."),
        ("CWSI fora de [0; 1] ANTES do corte",
         f'=COUNTIF({rngAI},"<0")+COUNTIF({rngAI},">1")', "0",
         "Valores fora da faixa indicam baseline inadequada, unidade errada ou sensor descalibrado. "
         "NÃO basta cortar em [0;1] — investigue a causa."),
        ("CWSI médio do período (válidos)", f'=IFERROR(AVERAGE({rngAJ}),"")', "0.000",
         "Média simples de todos os registros válidos."),
        ("VPD médio dos registros válidos",
         f'=IFERROR(AVERAGEIF({rngAH},1,{rngL}),"")', "0.000",
         "kPa. VPD muito baixo torna o CWSI instável."),
        ("dT médio dos registros válidos",
         f'=IFERROR(AVERAGEIF({rngAH},1,{rngP}),"")', "0.00",
         "°C. dT persistentemente negativo em dia quente sugere Tc subestimada."),
        ("Baseline em uso: a", "=CFG_A", "0.00", "Vem de 01_Config (calibração local tem prioridade)."),
        ("Baseline em uso: b", "=CFG_B", "0.00", "Deve ser NEGATIVO. Positivo = baseline errada."),
    ]
    r = 5
    for rot, formula, fmt, desc in painel:
        ws.cell(row=r, column=2, value=rot).font = F_ROT
        c = ws.cell(row=r, column=3, value=formula)
        c.font = F_FORMULA
        c.fill = P_CALC
        c.border = B_CAIXA
        c.number_format = fmt
        c.alignment = Alignment(horizontal="center")
        d = ws.cell(row=r, column=4, value=desc)
        d.font = F_NOTA
        d.alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[r].height = 28
        r += 1

    r += 1
    ws.cell(row=r, column=2, value="Semáforo geral")
    ws.cell(row=r, column=2).font = F_ROT_B
    sem = ws.cell(row=r, column=3, value=(
        f'=IF(COUNT({rngA})=0,"SEM DADOS",'
        f'IF(COUNTIF({rngAH},1)=0,"NENHUM REGISTRO VÁLIDO",'
        f'IF(COUNTIF({rngAI},"<0")+COUNTIF({rngAI},">1")>COUNTIF({rngAH},1)*0.2,'
        f'"REVISAR — muitos CWSI fora da faixa","OK")))'))
    sem.font = F_TXT_B
    sem.border = B_CAIXA
    sem.alignment = Alignment(horizontal="center")
    ws.conditional_formatting.add(
        f"C{r}", FormulaRule(formula=[f'$C${r}="OK"'], fill=P_OK))
    ws.conditional_formatting.add(
        f"C{r}", FormulaRule(formula=[f'$C${r}<>"OK"'], fill=P_ALERTA, font=F_ALERTA))
    linha_semaforo = r

    # ------------------------------------------------ checklist de erros --
    r += 3
    secao(ws, r, "CHECKLIST — os 22 erros clássicos no cálculo do CWSI", largura=6)
    r += 1
    cab = ["#", "Erro clássico", "Verificação", "Por que quebra o CWSI / como corrigir",
           "Situação", "Anotações"]
    for i, t in enumerate(cab, start=1):
        c = ws.cell(row=r, column=i, value=t)
        c.font = F_CAB
        c.fill = P_CAB
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = B_CAIXA
    ws.row_dimensions[r].height = 30
    r0_check = r + 1

    checks = [
        ("Carimbos de tempo desalinhados entre estação e sensor de dossel",
         f'=IF(COUNTIF({rngI},"SEM PAR")=0,"OK","{{}} sem par")',
         "O CWSI compara Tc e Ta do MESMO instante. Fuso diferente, horário de verão ou intervalos "
         "de registro distintos (1 min x 15 min) destroem o pareamento. Agregue as duas séries ao "
         "mesmo intervalo antes de colar."),
        ("Umidade relativa em fração (0–1) tratada como porcentagem",
         "=IF(CFG_UUR=\"" + U_UR[1] + "\",\"Convertendo\",\"Verificar origem\")",
         "UR de 0,48 lida como 48% (ou 48 lido como 4800%) muda o VPD em uma ordem de grandeza. "
         "Declare a unidade real em 01_Config."),
        ("Pressão em hPa/mbar somada como kPa",
         "=IF(CFG_UP=\"kPa\",\"Verificar origem\",\"Convertendo\")",
         "1013 hPa = 101,3 kPa. Se entrar como 1013 kPa, γ fica 10× maior e o método teórico desaba."),
        ("Radiação em MJ/m²/h usada como W/m²",
         "=IF(CFG_URS=\"W/m²\",\"Verificar origem\",\"Convertendo\")",
         "1 MJ/m²/h = 277,8 W/m². Com a unidade errada o filtro de radiação corta tudo (ou nada)."),
        ("Vento de 10 m usado como se fosse de 2 m",
         "=IF(CFG_ZANEM=2,\"Altura = 2 m\",\"Corrigindo p/ 2 m\")",
         "A planilha corrige pela FAO-56, mas só se a altura real estiver em 01_Config. "
         "Vento errado distorce ra e todo o método teórico."),
        ("Usar média diária em vez do instante próximo ao meio-dia solar",
         "=\"Janela: \"&CFG_HINI&\"h–\"&CFG_HFIM&\"h\"",
         "A baseline de Idso vale para alta radiação perto do meio-dia. Média diária dilui o sinal "
         "e produz CWSI sem significado."),
        ("Não filtrar dias/horários nublados",
         "=\"Rs mínima: \"&CFG_RSMIN&\" W/m²\"",
         "Sob nuvem a transpiração cai por falta de energia, não por falta de água. O CWSI lê isso "
         "como 'sem estresse' e engana."),
        ("Baseline (a, b) de outra cultura, região ou fenologia",
         "=IF(CFG_A_MAN<>\"\",\"Calibração local\",\"Literatura — calibrar!\")",
         "Os coeficientes de Idso vieram do Arizona. Em clima úmido eles erram bastante. Calibre "
         "localmente (roteiro na aba 07)."),
        ("Coeficiente b positivo",
         "=IF(CFG_B=\"\",\"Sem valor\",IF(CFG_B<0,\"OK (negativo)\",\"ERRO: b positivo\"))",
         "A reta Tc−Ta x VPD é DECRESCENTE. b positivo inverte todo o índice."),
        ("Limite superior calculado com VPD em vez de VPG",
         "=\"Fórmula fixa na coluna U\"",
         "dT_UL = a + b·VPG, com VPG = es(Ta) − es(Ta + a). Usar VPD nos dois limites zera o "
         "denominador e o CWSI explode."),
        ("Divisão por denominador quase zero",
         f'=IF(COUNTIF({rngAI},"<0")+COUNTIF({rngAI},">1")=0,"OK","{{}} fora de [0;1]")',
         "Quando dT_UL ≈ dT_LL qualquer ruído vira CWSI gigante. A trava está em 01_Config "
         "(denominador mínimo)."),
        ("Cortar o CWSI em [0; 1] e esconder o problema",
         "=\"Coluna AI mostra o valor bruto\"",
         "Limitar é correto para reportar, mas o valor BRUTO precisa continuar visível. "
         "Muitos valores fora da faixa = erro de método, não de arredondamento."),
        ("es calculado com a temperatura do dossel em vez da do ar",
         "=\"Fórmula fixa na coluna J\"",
         "O VPD do ar usa es(Ta). Usar es(Tc) mistura causa e efeito e cria correlação artificial."),
        ("Emissividade errada no termômetro infravermelho",
         "=IF('01_Config'!$C$29=\"\",\"Não informada\",'01_Config'!$C$29)",
         "Dossel vegetal: 0,97–0,99. Sensor configurado em 1,00 ou 0,95 desloca Tc em 1–2 °C — "
         "o suficiente para mudar a decisão de irrigar."),
        ("Sensor enxergando solo exposto",
         "=IF('01_Config'!$C$33=\"\",\"Informar cobertura\","
         "IF('01_Config'!$C$33>=60,\"OK\",\"Cobertura baixa\"))",
         "Com cobertura abaixo de ~60% o solo quente entra no campo de visão e infla Tc. "
         "Use visada oblíqua e registre a cobertura."),
        ("Estação meteorológica distante ou em ambiente diferente do talhão",
         "=IF('01_Config'!$C$25=\"\",\"Informar distância\","
         "IF('01_Config'!$C$25<=1000,\"OK\",\"Estação distante\"))",
         "Estação sobre grama irrigada e talhão em sequeiro medem atmosferas diferentes."),
        ("Vento calmo (desacoplamento) ou muito forte",
         "=\"Faixa: \"&CFG_UMIN&\"–\"&CFG_UMAX&\" m/s\"",
         "Em calmaria o dossel desacopla do ar e o dT infla; com vento forte o dT comprime e o "
         "estresse some."),
        ("VPD muito baixo (manhã úmida, pós-chuva)",
         "=\"VPD mínimo: \"&CFG_VPDMIN&\" kPa\"",
         "Com VPD perto de zero os dois limites se encontram e o índice perde resolução."),
        ("Rn estimado tratado como medido (método teórico)",
         "=IF(COUNT('02_Estacao'!$G$" + str(L0) + ":$G$" + str(L1) + ")>0,"
         "\"Rn medido em uso\",\"Rn estimado\")",
         "Rn ≈ (1−albedo)·Rs − onda longa é aproximação grosseira. Para o método de Jackson, "
         "medir Rn muda bastante o resultado."),
        ("Altura do dossel maior que a altura do anemômetro",
         "=IF(CFG_HDOSSEL*0.67>=CFG_ZANEM,\"ERRO: h muito alto\",\"OK\")",
         "ra usa ln((zm − 0,67h)/z0m). Se 0,67h ≥ zm o logaritmo não existe e a coluna ra fica vazia."),
        ("Separador decimal trocado na importação (vírgula x ponto)",
         f'=IF(COUNT({rngA})=0,"Sem dados","Conferir manualmente")',
         "Dado colado como TEXTO não entra em nenhuma média. Se as colunas B a E da aba 02 "
         "alinham à esquerda, foram lidas como texto."),
        ("Misturar fusos: dado em UTC e janela horária em hora local",
         "=\"Janela em hora LOCAL\"",
         "Muitas estações gravam em UTC. No Brasil isso desloca a janela em 3 h e o filtro pega "
         "o fim da tarde em vez do meio-dia."),
    ]

    r = r0_check
    for i, (erro, verif, porque) in enumerate(checks, start=1):
        ws.cell(row=r, column=1, value=i).font = F_TXT
        ws.cell(row=r, column=1).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=r, column=2, value=erro).font = F_TXT_B
        ws.cell(row=r, column=2).alignment = Alignment(wrap_text=True, vertical="center")

        v = verif
        if "{}" in v:
            if "SEM PAR" in v:
                v = f'=IF(COUNTIF({rngI},"SEM PAR")=0,"OK",COUNTIF({rngI},"SEM PAR")&" sem par")'
            else:
                v = (f'=IF(COUNTIF({rngAI},"<0")+COUNTIF({rngAI},">1")=0,"OK",'
                     f'COUNTIF({rngAI},"<0")+COUNTIF({rngAI},">1")&" fora de [0;1]")')
        c = ws.cell(row=r, column=3, value=v)
        c.font = F_FORMULA
        c.fill = P_CALC
        c.border = B_CAIXA
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        d = ws.cell(row=r, column=4, value=porque)
        d.font = F_TXT
        d.alignment = Alignment(wrap_text=True, vertical="center")

        s = ws.cell(row=r, column=5, value=STATUS_QC[0])
        s.font = F_ENTRADA
        s.fill = P_ENTRADA
        s.border = B_CAIXA
        s.alignment = Alignment(horizontal="center", vertical="center")

        ws.cell(row=r, column=6).fill = P_ENTRADA
        ws.cell(row=r, column=6).border = B_CAIXA

        ws.row_dimensions[r].height = 44
        r += 1

    dv = DataValidation(type="list", formula1="=LST_STATUS", allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"E{r0_check}:E{r - 1}")

    ws.conditional_formatting.add(
        f"C{r0_check}:C{r - 1}",
        FormulaRule(formula=[f'ISNUMBER(SEARCH("ERRO",$C{r0_check}))'], fill=P_ALERTA, font=F_ALERTA))
    ws.conditional_formatting.add(
        f"C{r0_check}:C{r - 1}",
        FormulaRule(formula=[f'$C{r0_check}="OK"'], fill=P_OK))
    ws.conditional_formatting.add(
        f"E{r0_check}:E{r - 1}",
        FormulaRule(formula=[f'$E{r0_check}="Corrigir"'], fill=P_ALERTA, font=F_ALERTA))
    ws.conditional_formatting.add(
        f"E{r0_check}:E{r - 1}",
        FormulaRule(formula=[f'$E{r0_check}="OK"'], fill=P_OK))

    ws.freeze_panes = "A5"
    return linha_semaforo


# -------------------------------------------------------- 06_Resultados ----
def montar_resultados(wb, ws):
    titulo(ws, "06 · Resultados diários e decisão de irrigação",
           "Agrega os registros VÁLIDOS da aba 04 por dia. Nada aqui é digitado.", largura=9)

    rngA  = f"'04_Calculo'!$A${L0}:$A${L1}"
    rngAH = f"'04_Calculo'!$AH${L0}:$AH${L1}"
    rngAJ = f"'04_Calculo'!$AJ${L0}:$AJ${L1}"
    rngP  = f"'04_Calculo'!$P${L0}:$P${L1}"
    rngL  = f"'04_Calculo'!$L${L0}:$L${L1}"

    ws.cell(row=4, column=1, value="Primeira data com dados").font = F_ROT
    d0 = ws.cell(row=4, column=3, value=f'=IFERROR(INT(SMALL({rngA},1)),"")')
    ws.cell(row=5, column=1, value="Última data com dados").font = F_ROT
    d1 = ws.cell(row=5, column=3, value=f'=IFERROR(INT(LARGE({rngA},1)),"")')
    for c in (d0, d1):
        c.font = F_FORMULA
        c.fill = P_CALC
        c.border = B_CAIXA
        c.number_format = "dd/mm/yyyy"
        c.alignment = Alignment(horizontal="center")

    cabecalhos(ws, 7, [
        "Data", "Registros\nno dia", "Registros\nválidos", "CWSI médio", "CWSI mín.",
        "CWSI máx.", "dT médio\n(°C)", "VPD médio\n(kPa)", "Decisão",
    ], larguras=[13, 12, 12, 13, 12, 12, 13, 13, 22], altura=34)

    r0, n = 8, 90
    for i in range(n):
        r = r0 + i
        if i == 0:
            fa = '=IF($C$4="","",$C$4)'
        else:
            # aninhado: OR() avalia todos os argumentos e ""+1 daria #VALOR!
            fa = f'=IF($A{r-1}="","",IF($A{r-1}+1>$C$5,"",$A{r-1}+1))'
        crit = f'{rngA},">="&$A{r},{rngA},"<"&$A{r}+1'
        vals = {
            1: fa,
            2: f'=IF($A{r}="","",COUNTIFS({crit}))',
            3: f'=IF($A{r}="","",COUNTIFS({crit},{rngAH},1))',
            4: f'=IF(OR($A{r}="",$C{r}=0),"",AVERAGEIFS({rngAJ},{crit},{rngAH},1))',
            5: f'=IF(OR($A{r}="",$C{r}=0),"",_xlfn.MINIFS({rngAJ},{crit},{rngAH},1))',
            6: f'=IF(OR($A{r}="",$C{r}=0),"",_xlfn.MAXIFS({rngAJ},{crit},{rngAH},1))',
            7: f'=IF(OR($A{r}="",$C{r}=0),"",AVERAGEIFS({rngP},{crit},{rngAH},1))',
            8: f'=IF(OR($A{r}="",$C{r}=0),"",AVERAGEIFS({rngL},{crit},{rngAH},1))',
            9: (f'=IF($D{r}="","",IF($D{r}>=CFG_LIMIAR,"IRRIGAR",'
                f'IF($D{r}>=CFG_LIMIAR_AT,"ATENÇÃO","SEM ESTRESSE")))'),
        }
        fmts = {1: "dd/mm/yyyy", 2: "0", 3: "0", 4: "0.000", 5: "0.000",
                6: "0.000", 7: "0.00", 8: "0.000", 9: None}
        for col, formula in vals.items():
            c = ws.cell(row=r, column=col, value=formula)
            c.font = F_FORMULA
            c.border = B_CAIXA
            c.alignment = Alignment(horizontal="center")
            if fmts[col]:
                c.number_format = fmts[col]
    rN = r0 + n - 1

    ws.conditional_formatting.add(
        f"D{r0}:D{rN}",
        ColorScaleRule(start_type="num", start_value=0, start_color="FF63BE7B",
                       mid_type="num", mid_value=0.5, mid_color="FFFFEB84",
                       end_type="num", end_value=1, end_color="FFF8696B"))
    ws.conditional_formatting.add(
        f"I{r0}:I{rN}", FormulaRule(formula=[f'$I{r0}="IRRIGAR"'], fill=P_ALERTA, font=F_ALERTA))
    ws.conditional_formatting.add(
        f"I{r0}:I{rN}", FormulaRule(formula=[f'$I{r0}="SEM ESTRESSE"'], fill=P_OK))

    r = rN + 2
    ws.cell(row=r, column=1, value=(
        "Leitura: CWSI 0 = cultura transpirando sem restrição; CWSI 1 = transpiração interrompida. "
        "Os limiares vêm de 01_Config e devem ser ajustados à sua cultura, solo e sistema de irrigação. "
        "Uma decisão de irrigação nunca deve sair de um único dia isolado — olhe a tendência de 2 a 3 dias "
        "e cruze com umidade de solo e previsão de chuva.")).font = F_NOTA
    ws.cell(row=r, column=1).alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=r, start_column=1, end_row=r + 2, end_column=9)

    ws.freeze_panes = "A8"


# ------------------------------------------- 10_Base_Procedimentos ---------
DB_CAB = ["ID", "Marca", "Monitor / terminal", "Versão", "Objetivo", "Origem / destino",
          "Formato de arquivo", "Pasta / caminho", "Sistema de arquivos", "CHAVE"] + \
         [f"Passo {i:02d}" for i in range(1, 11)] + \
         ["Cuidados", "Erros comuns", "Confiança", "Fonte"]
DB_EXTRA = 20


def montar_db(wb, ws):
    titulo(ws, "10 · Base de procedimentos (o cérebro do passo a passo)",
           "Cada linha é um procedimento completo. A aba 09 apenas consulta esta tabela. "
           "Para ampliar o guia, acrescente linhas aqui — a aba 09 passa a encontrá-las automaticamente.",
           largura=12)

    larg = [6, 24, 30, 20, 30, 24, 34, 34, 18, 46] + [46] * 10 + [46, 46, 14, 40]
    cabecalhos(ws, 4, DB_CAB, larguras=larg, altura=32)

    r = L0
    for i, p in enumerate(C.PROCEDIMENTOS, start=1):
        ws.cell(row=r, column=1, value=i)
        ws.cell(row=r, column=2, value=p["marca"])
        ws.cell(row=r, column=3, value=p["monitor"])
        ws.cell(row=r, column=4, value=p["versao"])
        ws.cell(row=r, column=5, value=p["objetivo"])
        ws.cell(row=r, column=6, value=p["origem"])
        ws.cell(row=r, column=7, value=p["formato"])
        ws.cell(row=r, column=8, value=p["pasta"])
        ws.cell(row=r, column=9, value=p["fs"])
        for k in range(10):
            ws.cell(row=r, column=11 + k,
                    value=p["passos"][k] if k < len(p["passos"]) else "")
        ws.cell(row=r, column=21, value=p["cuidados"])
        ws.cell(row=r, column=22, value=p["erros"])
        ws.cell(row=r, column=23, value=p["confianca"])
        ws.cell(row=r, column=24, value=p["fonte"])
        r += 1
    ultimo_dado = r - 1
    ultimo = ultimo_dado + DB_EXTRA

    for rr in range(L0, ultimo + 1):
        ws.cell(row=rr, column=10, value=(
            f'=IF($B{rr}="","",$B{rr}&"|"&$C{rr}&"|"&$D{rr}&"|"&$E{rr}&"|"&$F{rr})'))
        for col in range(1, 25):
            c = ws.cell(row=rr, column=col)
            c.font = F_TXT if col != 10 else F_NOTA
            c.border = B_CAIXA
            c.alignment = Alignment(wrap_text=True, vertical="top")
            if rr > ultimo_dado:
                c.fill = P_ENTRADA
        ws.row_dimensions[rr].height = 46

    for col, nome in [(7, "DB_FORMATO"), (8, "DB_PASTA"), (9, "DB_FS"), (10, "DB_CHAVE"),
                      (21, "DB_CUIDADOS"), (22, "DB_ERROS"), (23, "DB_CONF"), (24, "DB_FONTE")]:
        L = get_column_letter(col)
        nome_global(wb, nome, "10_Base_Procedimentos", f"${L}${L0}:${L}${ultimo}")
    nome_global(wb, "DB_PASSOS", "10_Base_Procedimentos", f"$K${L0}:$T${ultimo}")

    for celula, lista in [("B", "LST_MARCAS"), ("C", "LST_MONITORES"),
                          ("E", "LST_OBJETIVOS"), ("F", "LST_ORIGENS")]:
        dv = DataValidation(type="list", formula1=f"={lista}", allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f"{celula}{ultimo_dado + 1}:{celula}{ultimo}")

    ws.freeze_panes = "C6"
    ws.auto_filter.ref = f"A4:X{ultimo}"


# ------------------------------------------------------ 08_Equipamentos ----
def montar_equipamentos(wb, ws):
    titulo(ws, "08 · Catálogo de monitores e terminais",
           "Figura esquemática do formato físico de cada terminal (proporção de tela, teclas, encoder), "
           "versões conhecidas e como cada um troca arquivos. Desenhos próprios — não são fotos nem "
           "logotipos dos fabricantes.", largura=8)

    # A=recuo · B=ícone · C=rótulo · D=valor · E=rótulo 2 · F..H=valor 2
    larguras = [3, 24, 21, 40, 18, 26, 22, 26]
    for i, w in enumerate(larguras, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    r = 4
    for m in C.MONITORES:
        (arq, marca, monitor, formato_fisico, tamanho, versoes,
         plataforma, telemetria, midia, formatos, obs) = m

        secao(ws, r, f"{marca}  ·  {monitor}", largura=8, cor=P_SECAO2)
        r += 1

        img_path = os.path.join(ICONES, arq)
        if os.path.exists(img_path):
            img = XLImage(img_path)
            img.width, img.height = 152, 120
            ws.add_image(img, f"B{r}")

        campos = [
            ("Formato físico", formato_fisico, "Tamanho de tela", tamanho),
            ("Versões conhecidas", versoes, None, None),
            ("Plataforma web", plataforma, "Telemetria", telemetria),
            ("Mídia aceita", midia, None, None),
            ("Formatos de arquivo", formatos, None, None),
            ("Observações", obs, None, None),
        ]
        for k, (r1, v1, r2, v2) in enumerate(campos):
            rr = r + k
            ws.row_dimensions[rr].height = 22
            c = ws.cell(row=rr, column=3, value=r1)
            c.font = F_ROT_B
            c.alignment = Alignment(vertical="center")
            if r2:
                v = ws.cell(row=rr, column=4, value=v1)
                v.font = F_TXT
                v.alignment = Alignment(wrap_text=True, vertical="center")
                c2 = ws.cell(row=rr, column=5, value=r2)
                c2.font = F_ROT_B
                c2.alignment = Alignment(vertical="center")
                v2c = ws.cell(row=rr, column=6, value=v2)
                v2c.font = F_TXT
                v2c.alignment = Alignment(wrap_text=True, vertical="center")
                ws.merge_cells(start_row=rr, start_column=6, end_row=rr, end_column=8)
            else:
                v = ws.cell(row=rr, column=4, value=v1)
                v.font = F_TXT
                v.alignment = Alignment(wrap_text=True, vertical="center")
                ws.merge_cells(start_row=rr, start_column=4, end_row=rr, end_column=8)
        r += len(campos) + 1

    ws.freeze_panes = "A4"


# ------------------------------------------------------ 09_Passo_a_Passo ---
def montar_passo(wb, ws):
    titulo(ws, "09 · Passo a passo — como colocar arquivos no monitor e tirar dados dele",
           "Selecione na ordem: marca → tipo → monitor → versão → o que quer fazer → por onde. "
           "A sequência de cliques aparece embaixo, montada a partir da aba 10.", largura=8)

    for col, w in [(1, 4), (2, 34), (3, 46), (4, 62), (5, 4), (6, 30), (7, 30), (8, 30)]:
        ws.column_dimensions[get_column_letter(col)].width = w

    # ---- bloco de selecao
    secao(ws, 4, "SELEÇÃO — preencha de cima para baixo, nesta ordem", largura=8)
    sel = [
        (5, "1 · Marca do equipamento", "LST_MARCAS",
         "Define quais monitores aparecem no passo 3."),
        (6, "2 · Tipo de equipamento", "LST_TIPOS",
         "Campo de registro. Não filtra os procedimentos — o que manda é o monitor."),
        (7, "3 · Monitor / terminal", None,
         "Lista dependente da marca. Se estiver vazia, escolha a marca primeiro."),
        (8, "4 · Versão do software", None,
         "Lista dependente do monitor. 'Todas as versoes' serve quando o procedimento não muda."),
        (9, "5 · O que você quer fazer", "LST_OBJETIVOS",
         "Levar arquivo para o monitor, ou tirar dados dele."),
        (10, "6 · Por onde", "LST_ORIGENS",
         "Pen drive ou plataforma do fabricante."),
    ]
    for linha, rot, lista, ajuda in sel:
        c = ws.cell(row=linha, column=2, value=rot)
        c.font = F_ROT_B
        c.alignment = Alignment(vertical="center")
        e = ws.cell(row=linha, column=3)
        e.font = F_ENTRADA
        e.fill = P_ENTRADA
        e.border = B_CAIXA
        e.alignment = Alignment(horizontal="center", vertical="center")
        a = ws.cell(row=linha, column=4, value=ajuda)
        a.font = F_NOTA
        a.alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[linha].height = 24

    dv1 = DataValidation(type="list", formula1="=LST_MARCAS", allow_blank=True)
    ws.add_data_validation(dv1); dv1.add(ws["C5"])
    dv2 = DataValidation(type="list", formula1="=LST_TIPOS", allow_blank=True)
    ws.add_data_validation(dv2); dv2.add(ws["C6"])
    dv3 = DataValidation(type="list", formula1='=INDIRECT("MON_"&MATCH($C$5,LST_MARCAS,0))',
                         allow_blank=True)
    ws.add_data_validation(dv3); dv3.add(ws["C7"])
    dv4 = DataValidation(type="list", formula1='=INDIRECT("VER_"&MATCH($C$7,LST_MONITORES,0))',
                         allow_blank=True)
    ws.add_data_validation(dv4); dv4.add(ws["C8"])
    dv5 = DataValidation(type="list", formula1="=LST_OBJETIVOS", allow_blank=True)
    ws.add_data_validation(dv5); dv5.add(ws["C9"])
    dv6 = DataValidation(type="list", formula1="=LST_ORIGENS", allow_blank=True)
    ws.add_data_validation(dv6); dv6.add(ws["C10"])

    # ---- celulas auxiliares
    ws.cell(row=4, column=10, value="AUXILIAR — não editar").font = F_NOTA
    ws["J5"] = '=$C$5&"|"&$C$7&"|"&$C$8&"|"&$C$9&"|"&$C$10'
    ws["J6"] = f'=$C$5&"|"&$C$7&"|{VER_TODAS}|"&$C$9&"|"&$C$10'
    ws["J7"] = '=IFERROR(MATCH($J$5,DB_CHAVE,0),IFERROR(MATCH($J$6,DB_CHAVE,0),""))'
    for cel in ("J5", "J6", "J7"):
        ws[cel].font = F_NOTA
    ws.column_dimensions["J"].width = 3

    # ---- mensagem de estado
    msg = ws.cell(row=12, column=2, value=(
        '=IF(COUNTA($C$5,$C$7,$C$9,$C$10)<4,'
        '"Complete a seleção acima (marca, monitor, objetivo e origem).",'
        'IF($J$7="",'
        '"Combinação ainda NÃO cadastrada. Abra a aba 10_Base_Procedimentos e acrescente uma linha '
        'para esta combinação — ela aparecerá aqui automaticamente.",'
        '"Procedimento encontrado. Siga a sequência abaixo."))'))
    msg.font = F_TXT_B
    msg.alignment = Alignment(vertical="center", wrap_text=True)
    ws.merge_cells("B12:H12")
    ws.row_dimensions[12].height = 30
    ws.conditional_formatting.add(
        "B12", FormulaRule(formula=['ISNUMBER(SEARCH("NÃO cadastrada",$B$12))'],
                           fill=P_ALERTA, font=F_ALERTA))
    ws.conditional_formatting.add(
        "B12", FormulaRule(formula=['ISNUMBER(SEARCH("encontrado",$B$12))'], fill=P_OK))

    # ---- resumo do arquivo
    secao(ws, 14, "O QUE PREPARAR", largura=8)
    resumo = [
        (15, "Formato do arquivo", "DB_FORMATO"),
        (16, "Pasta / caminho no pen drive", "DB_PASTA"),
        (17, "Sistema de arquivos do pen drive", "DB_FS"),
    ]
    for linha, rot, nome in resumo:
        c = ws.cell(row=linha, column=2, value=rot)
        c.font = F_ROT_B
        c.alignment = Alignment(vertical="center")
        v = ws.cell(row=linha, column=3, value=f'=IF($J$7="","",INDEX({nome},$J$7))')
        v.font = Font(name=FONTE, size=10, bold=True, color="FF1F3864")
        v.fill = P_DESTAQUE
        v.border = B_CAIXA
        v.alignment = Alignment(wrap_text=True, vertical="center")
        ws.merge_cells(start_row=linha, start_column=3, end_row=linha, end_column=8)
        ws.row_dimensions[linha].height = 30

    # ---- passos
    secao(ws, 19, "SEQUÊNCIA DE PASSOS — siga na ordem", largura=8)
    for k in range(10):
        linha = 20 + k
        n = ws.cell(row=linha, column=2, value=f"Passo {k + 1}")
        n.font = F_ROT_B
        n.alignment = Alignment(horizontal="center", vertical="center")
        n.fill = P_CALC
        n.border = B_CAIXA
        v = ws.cell(row=linha, column=3,
                    value=f'=IF($J$7="","",IF(INDEX(DB_PASSOS,$J$7,{k + 1})="","",'
                          f'INDEX(DB_PASSOS,$J$7,{k + 1})))')
        v.font = F_PASSO
        v.border = B_CAIXA
        v.alignment = Alignment(wrap_text=True, vertical="center", indent=1)
        ws.merge_cells(start_row=linha, start_column=3, end_row=linha, end_column=8)
        ws.row_dimensions[linha].height = 30

    # ---- cuidados / erros / fonte
    secao(ws, 31, "CUIDADOS, ERROS COMUNS E PROCEDÊNCIA", largura=8)
    blocos = [
        (32, "Cuidados", "DB_CUIDADOS", 44),
        (33, "Erros mais comuns", "DB_ERROS", 56),
        (34, "Confiança da informação", "DB_CONF", 22),
        (35, "Fonte", "DB_FONTE", 30),
    ]
    for linha, rot, nome, alt in blocos:
        c = ws.cell(row=linha, column=2, value=rot)
        c.font = F_ROT_B
        c.alignment = Alignment(vertical="center")
        v = ws.cell(row=linha, column=3, value=f'=IF($J$7="","",INDEX({nome},$J$7))')
        v.font = F_TXT
        v.border = B_CAIXA
        v.alignment = Alignment(wrap_text=True, vertical="center", indent=1)
        ws.merge_cells(start_row=linha, start_column=3, end_row=linha, end_column=8)
        ws.row_dimensions[linha].height = alt
    ws.conditional_formatting.add(
        "C34", FormulaRule(formula=['$C$34="Confirmar"'], fill=P_ENTRADA))
    ws.conditional_formatting.add(
        "C34", FormulaRule(formula=['$C$34="Verificado"'], fill=P_OK))
    nota(ws, 36,
         "«Verificado» = formato e caminho de pasta confirmados na fonte citada.  "
         "«Confirmar» = a estrutura está correta, mas o NOME do menu muda entre versões — "
         "confirme na máquina e registre o caminho real na aba 10.", col=2, largura=8)

    # ---- figuras
    secao(ws, 38, "OS TRÊS FLUXOS POSSÍVEIS", largura=8)
    figs = [
        ("fig_fluxo_usb_para_monitor.png", "B40", 620, 153),
        ("fig_fluxo_monitor_para_pc.png", "B50", 620, 153),
        ("fig_fluxo_nuvem.png", "B60", 620, 153),
        ("fig_estrutura_pastas.png", "B70", 620, 300),
    ]
    for arq, ancora, w, h in figs:
        p = os.path.join(ICONES, arq)
        if os.path.exists(p):
            img = XLImage(p)
            img.width, img.height = w, h
            ws.add_image(img, ancora)

    ws.freeze_panes = "A4"


# --------------------------------------------------- 11_Formatos_Arquivos --
def montar_formatos(wb, ws):
    titulo(ws, "11 · Formatos de arquivo — o que é cada um e onde erra",
           "Referência rápida. A maioria dos 'o monitor não lê meu arquivo' está nesta tabela.",
           largura=5)
    cabecalhos(ws, 4, ["Formato", "Extensões / estrutura", "O que carrega",
                       "Quem usa", "Cuidados e pegadinhas"],
               larguras=[26, 34, 40, 40, 62], altura=32)
    r = 5
    for fmt, ext, carrega, quem, cuidados in C.FORMATOS:
        ws.cell(row=r, column=1, value=fmt).font = F_TXT_B
        ws.cell(row=r, column=2, value=ext).font = F_TXT
        ws.cell(row=r, column=3, value=carrega).font = F_TXT
        ws.cell(row=r, column=4, value=quem).font = F_TXT
        ws.cell(row=r, column=5, value=cuidados).font = F_TXT
        for col in range(1, 6):
            ws.cell(row=r, column=col).border = B_CAIXA
            ws.cell(row=r, column=col).alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 62
        r += 1
    ws.freeze_panes = "A5"


# ------------------------------------------------------------ 12_Fontes ----
def montar_fontes(wb, ws):
    titulo(ws, "12 · Fontes, premissas e limitações", None, largura=4)
    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 96

    secao(ws, 4, "A · CÁLCULO DO CWSI", largura=4)
    itens_cwsi = [
        ("Idso, Jackson, Pinter, Reginato & Hatfield (1981)",
         "Normalizing the stress-degree-day parameter for environmental variability. "
         "Agricultural Meteorology 24:45-55. Origem do CWSI empírico."),
        ("Idso (1982)",
         "Non-water-stressed baselines: a key to measuring and interpreting plant water stress. "
         "Agricultural Meteorology 27:59-70. Origem da tabela de coeficientes a e b da aba 07."),
        ("Jackson, Idso, Reginato & Pinter (1981)",
         "Canopy temperature as a crop water stress indicator. Water Resources Research 17:1133-1138. "
         "Base do método teórico (balanço de energia)."),
        ("Jackson, Kustas & Choudhury (1988)",
         "A reexamination of the crop water stress index. Irrigation Science 9:309-317."),
        ("Allen, Pereira, Raes & Smith (1998) — FAO-56",
         "Crop evapotranspiration. Fonte das equações de es(T), Δ, γ, pressão pela altitude, "
         "correção do vento para 2 m e resistência aerodinâmica ra."),
        ("Premissa desta planilha — Rn estimado",
         "Quando a estação não mede radiação líquida, a planilha usa Rn ≈ (1−albedo)·Rs − onda longa "
         "líquida, com a onda longa fixada em 01_Config (padrão 60 W/m²). É uma aproximação para "
         "meio-dia com céu limpo; para trabalho de precisão, meça Rn."),
        ("Premissa desta planilha — pareamento por carimbo de tempo",
         "A aba 04 procura na aba 03 a data/hora EXATA da aba 02 (correspondência exata, sem "
         "tolerância). Registros sem par são marcados e excluídos, nunca interpolados em silêncio."),
        ("Limitação — cobertura do solo",
         "O CWSI pressupõe que o sensor enxerga apenas dossel. Abaixo de ~60% de cobertura o índice "
         "perde validade e a planilha apenas alerta; não corrige."),
        ("Limitação — os coeficientes são de clima árido",
         "A tabela de Idso (1982) foi levantada em Phoenix, Arizona. Em clima úmido tende a errar. "
         "A calibração local não é opcional para uso operacional."),
    ]
    r = 5
    for t, d in itens_cwsi:
        ws.cell(row=r, column=2, value=t).font = F_TXT_B
        ws.cell(row=r, column=3, value=d).font = F_TXT
        for col in (2, 3):
            ws.cell(row=r, column=col).alignment = Alignment(wrap_text=True, vertical="top")
            ws.cell(row=r, column=col).border = B_CAIXA
        ws.row_dimensions[r].height = 46
        r += 1

    r += 1
    secao(ws, r, "B · PROCEDIMENTOS DOS MONITORES", largura=4)
    r += 1
    itens_mon = [
        ("John Deere — Gen 4 / G5",
         "Deere StellarSupport, notas de versão do Generation 4 OS e ajuda do Gerenciador de Arquivos "
         "(displaysimulator.deere.com). Pasta 'Rx' na raiz para prescrições; exportação de Setup em "
         "'JD4600' e de dados de trabalho em 'JD-Data'."),
        ("John Deere — GreenStar 3 2630",
         "GS3 2630 User Guide. Estrutura GS3_2630 > <Perfil> > RCD para dados gravados; pasta 'Rx' na "
         "raiz para prescrições; pen drive em FAT/FAT32."),
        ("Case IH — AFS Pro 700 / AFS Pro 1200",
         "Quick Reference Card oficial de importação de shapefile do AFS Pro 700; AFS Pro 1200 Software "
         "Operating Manual (Importing Shapefile Data); material Case IH sobre AFS Connect."),
        ("Trimble — Precision-IQ (GFX-750, GFX-1060/1260) e TMX-2050",
         "Documentação Trimble Farm Data Compatibility e materiais de suporte. Prescrições shapefile em "
         "AgData\\Prescriptions; ISOXML em pasta TASKDATA."),
        ("Ag Leader — InCommand 800 / 1200",
         "InCommand User Guide (import .agsetup, export .agdata) e artigo 'AgSetup File Supported Uses' "
         "do portal Ag Leader; regra do .pat entre gerações."),
        ("Topcon X35 e terminais AGCO (Fendt / Valtra / Massey)",
         "Topcon X35 Operator's Manual (TASKDATA via USB em FAT32) e documentação ISOBUS / TaskDoc. "
         "Padrão ISO 11783-10 (ISOXML), perfis TC-BAS e TC-GEO."),
        ("Confiança das linhas marcadas como «Confirmar»",
         "Nessas linhas o formato e a estrutura de pastas estão corretos, mas o nome exato do menu "
         "muda entre versões de software. Confirme na máquina e atualize a aba 10 — é assim que este "
         "guia deve evoluir."),
    ]
    for t, d in itens_mon:
        ws.cell(row=r, column=2, value=t).font = F_TXT_B
        ws.cell(row=r, column=3, value=d).font = F_TXT
        for col in (2, 3):
            ws.cell(row=r, column=col).alignment = Alignment(wrap_text=True, vertical="top")
            ws.cell(row=r, column=col).border = B_CAIXA
        ws.row_dimensions[r].height = 52
        r += 1

    r += 1
    secao(ws, r, "C · SOBRE AS FIGURAS", largura=4)
    r += 1
    ws.cell(row=r, column=2, value="Ícones dos monitores").font = F_TXT_B
    ws.cell(row=r, column=3, value=(
        "São desenhos esquemáticos próprios, gerados por script (tools/gerar_icones.py), que "
        "representam o formato físico do terminal: proporção de tela, teclas físicas e encoder "
        "rotativo. Não reproduzem fotografias, logotipos nem marcas figurativas dos fabricantes. "
        "Os nomes dos modelos aparecem apenas como referência textual.")).font = F_TXT
    ws.cell(row=r, column=3).alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 62


# ------------------------------------------------------------ 00_Inicio ----
def montar_inicio(wb, ws):
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 88
    ws.column_dimensions["D"].width = 20

    ws["B2"] = "PLANILHA MESTRE — CWSI e Guia de Monitores Agrícolas"
    ws["B2"].font = Font(name=FONTE, size=20, bold=True, color="FF1F3864")
    ws.merge_cells("B2:D2")
    ws.row_dimensions[2].height = 30

    ws["B3"] = (f"Versão {VERSAO}  ·  Cálculo do Crop Water Stress Index a partir de estação "
                f"meteorológica + termometria de dossel, e guia de transferência de arquivos "
                f"entre escritório e monitor de máquina.")
    ws["B3"].font = F_SUB
    ws.merge_cells("B3:D3")

    secao(ws, 5, "COMO USAR — parte 1: calcular o CWSI", largura=4)
    p1 = [
        ("1", "Abra 01_Config", "Preencha as células AMARELAS: local, alturas dos sensores, unidades "
                                "dos seus dados, cultura e filtros. É o passo que mais evita erro."),
        ("2", "Cole os dados em 02_Estacao", "Uma linha por registro. A coluna A (data/hora) é a chave."),
        ("3", "Cole a temperatura de dossel em 03_Dossel",
              "Mesma data/hora da aba 02 — precisa bater exatamente."),
        ("4", "Confira 05_QC_Diagnostico",
              "Olhe o semáforo e o checklist ANTES de acreditar em qualquer número."),
        ("5", "Leia 06_Resultados", "CWSI por dia e a decisão de irrigação."),
        ("6", "Audite em 04_Calculo", "Cada etapa intermediária está visível, coluna por coluna."),
    ]
    r = 6
    for n, t, d in p1:
        ws.cell(row=r, column=1, value=n).font = F_TXT_B
        ws.cell(row=r, column=1).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=2, value=t).font = F_TXT_B
        ws.cell(row=r, column=3, value=d).font = F_TXT
        ws.cell(row=r, column=3).alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[r].height = 30
        r += 1

    r += 1
    secao(ws, r, "COMO USAR — parte 2: colocar arquivos no monitor / tirar dados dele", largura=4)
    r += 1
    p2 = [
        ("1", "Abra 09_Passo_a_Passo", "Selecione marca → tipo → monitor → versão → objetivo → origem."),
        ("2", "Leia o bloco «O QUE PREPARAR»", "Formato do arquivo, pasta e sistema de arquivos."),
        ("3", "Siga a sequência de passos", "Cada passo é uma ação na tela do monitor."),
        ("4", "Consulte 08_Equipamentos", "Figura do terminal, versões e formas de conexão."),
        ("5", "Consulte 11_Formatos_Arquivos", "O que é cada formato e onde ele costuma falhar."),
        ("6", "Amplie em 10_Base_Procedimentos",
              "Faltou uma combinação? Acrescente uma linha lá e ela aparece na aba 09 sozinha."),
    ]
    for n, t, d in p2:
        ws.cell(row=r, column=1, value=n).font = F_TXT_B
        ws.cell(row=r, column=1).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=2, value=t).font = F_TXT_B
        ws.cell(row=r, column=3, value=d).font = F_TXT
        ws.cell(row=r, column=3).alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[r].height = 30
        r += 1

    r += 1
    secao(ws, r, "ÍNDICE DAS ABAS", largura=4)
    r += 1
    idx = [
        ("01_Config", "Todos os parâmetros. Comece aqui."),
        ("02_Estacao", "Dados brutos da estação meteorológica (você cola aqui)."),
        ("03_Dossel", "Temperatura de dossel do infravermelho (você cola aqui)."),
        ("04_Calculo", "Motor de cálculo, etapa por etapa. Só fórmulas."),
        ("05_QC_Diagnostico", "Painel automático + checklist dos 22 erros clássicos."),
        ("06_Resultados", "CWSI diário e decisão de irrigação."),
        ("07_Baselines", "Coeficientes a e b por cultura + roteiro de calibração local."),
        ("08_Equipamentos", "Catálogo ilustrado de monitores e terminais."),
        ("09_Passo_a_Passo", "Gerador do passo a passo de transferência de arquivos."),
        ("10_Base_Procedimentos", "Banco de procedimentos. É aqui que o guia cresce."),
        ("11_Formatos_Arquivos", "Shapefile, ISOXML, .agsetup, .agdata, AgData, Rx, TASKDATA…"),
        ("12_Fontes", "Referências, premissas e limitações assumidas."),
    ]
    for aba, desc in idx:
        ws.cell(row=r, column=2, value=aba).font = Font(name=FONTE, size=10, bold=True,
                                                        color="FF1F3864")
        ws.cell(row=r, column=3, value=desc).font = F_TXT
        r += 1

    r += 1
    ws.cell(row=r, column=2, value="ANTES DE USAR DE VERDADE").font = F_ALERTA
    ws.cell(row=r, column=3, value=(
        "Apague a LINHA DE EXEMPLO (fundo laranja, linha 6) das abas 02_Estacao e 03_Dossel. "
        "Ela existe apenas para mostrar o formato esperado e para você ver a planilha funcionando.")
    ).font = F_ALERTA
    ws.cell(row=r, column=3).fill = P_ALERTA
    ws.cell(row=r, column=3).alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[r].height = 34


if __name__ == "__main__":
    construir()
