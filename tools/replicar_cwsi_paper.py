#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Replica o CWSI de BALANCO HIDRICO de Zulfiqar et al. (2026),
Sci Rep 16:25364, doi:10.1038/s41598-026-55297-z, secao "Crop water stress index",
usando os dados de estacao das planilhas CWSI_* (abas 02_Baseline_Input e
05_Season_Input, export ZL6/ATMOS 41 a cada 15 min).

Os tres passos, como o artigo os descreve:

  Passo 1  PET por Penman-Monteith, depois NORMALIZADA por escore-z:
           PET_norm,i = (PET_i - media(PET)) / desvio(PET)

  Passo 2  ETa por balanco hidrico:
           ETa_i = P_i + (S_{i-1} - S_i)
           S = armazenamento de agua no perfil, em mm.
           Depois normalizada pelo mesmo escore-z.

  Passo 3  CWSI = (PET_norm - ETa_norm) / PET_norm

ATENCAO: o Passo 3 ao pe da letra NAO produz valores em [0,1] — ver o bloco
DIAGNOSTICO na saida. O script calcula tambem 1 - ETa/PET com os fluxos
brutos, que e' a forma coerente com as faixas de classificacao do artigo
(definidas em percentual de ETa sobre PET).

Uso:
  python3.13 tools/replicar_cwsi_paper.py --xlsx ARQ.xlsx
  python3.13 tools/replicar_cwsi_paper.py --xlsx ARQ.xlsx --sheet 02_Baseline_Input
  python3.13 tools/replicar_cwsi_paper.py --xlsx ARQ.xlsx --csv saida.csv
"""

from __future__ import annotations

import argparse
import csv
import math
import sys
from collections import defaultdict
from datetime import datetime

SIGMA = 4.903e-9
ALBEDO = 0.23
G_SC = 0.0820

# Cabecalho na linha 4, dados a partir da linha 5 (export ZL6 colado na planilha)
LINHA_CAB = 4

# Dois layouts possiveis.
#
# "bruto"    abas 02_Baseline_Input e 05_Season_Input: o export ZL6/ATMOS 41
#            cru, 37 colunas, vento em km/h, perfil de umidade em 4
#            profundidades (15/30/45/60 cm).
#
# "derivado" abas 03_Baseline_Fit e 06_CWSI: a versao ja processada pela
#            planilha, vento em m/s, mas com UMA unica profundidade de
#            umidade (WC 15 cm). Traz tambem o flag de QC da planilha.
COL_BRUTO = {
    "ts": 1, "rs": 2, "precip": 3, "wind_kmh": 10, "tair": 12,
    "ea": 13, "patm": 14, "tmin": 16, "tmax": 17, "vpd": 20,
    "wc15": 21, "wc30": 23, "wc45": 25, "wc60": 27, "tc": 32,
}
COL_DERIVADO = {
    "ts": 2, "tair": 5, "tc": 6, "ea": 7, "vpd": 9, "rs": 10,
    "wind_ms": 15, "precip": 17, "wc15": 19, "dT": 20,
}

# Camadas de solo, espessuras em mm. Cada sensor representa a camada
# limitada pelos pontos medios entre ele e os vizinhos.
#
#   0-15    so o sensor de superficie, camada de 150 mm.
#   0-60    toda a coluna instrumentada: os 4 sensores, do topo ate a
#           profundidade do sensor mais fundo (60 cm = 600 mm).
#   0-67.5  idem, mas estendendo a camada do sensor de 60 cm ate 67,5 cm
#           pela convencao do ponto medio (675 mm). Extrapola 7,5 cm
#           abaixo do sensor mais fundo.
CAMADAS = {
    "0-15":   [("wc15", 150.0)],
    "0-60":   [("wc15", 225.0), ("wc30", 150.0), ("wc45", 150.0), ("wc60", 75.0)],
    "0-67.5": [("wc15", 225.0), ("wc30", 150.0), ("wc45", 150.0), ("wc60", 150.0)],
}

LAYOUT_DE = {
    "02_Baseline_Input": "bruto", "05_Season_Input": "bruto",
    "03_Baseline_Fit": "derivado", "06_CWSI": "derivado",
}


# --------------------------------------------------------------- fisica ----
def es(T):
    return 0.6108 * math.exp(17.27 * T / (T + 237.3))


def delta_svp(T):
    return 4098 * es(T) / (T + 237.3) ** 2


def ra_extraterrestre(lat_deg, doy):
    phi = math.radians(lat_deg)
    dr = 1 + 0.033 * math.cos(2 * math.pi * doy / 365)
    decl = 0.409 * math.sin(2 * math.pi * doy / 365 - 1.39)
    x = max(-1.0, min(1.0, -math.tan(phi) * math.tan(decl)))
    ws = math.acos(x)
    return (24 * 60 / math.pi) * G_SC * dr * (
        ws * math.sin(phi) * math.sin(decl)
        + math.cos(phi) * math.cos(decl) * math.sin(ws))


def rn_fao56(rs, tmax, tmin, ea, lat_deg, doy, z):
    ra = ra_extraterrestre(lat_deg, doy)
    rso = (0.75 + 2e-5 * z) * ra
    rns = (1 - ALBEDO) * rs
    frac = max(0.0, min(1.0, rs / rso)) if rso > 0 else 0.0
    rnl = (SIGMA * (((tmax + 273.16) ** 4 + (tmin + 273.16) ** 4) / 2)
           * (0.34 - 0.14 * math.sqrt(max(ea, 0.0)))
           * (1.35 * frac - 0.35))
    return rns - rnl


def pet_pm_diaria(tmean, tmax, tmin, ea, u2, rs_mj, patm, lat, z, doy):
    """ET0 de referencia FAO-56 (mm/dia)."""
    g_psi = 0.000665 * patm
    D = delta_svp(tmean)
    es_mean = (es(tmax) + es(tmin)) / 2
    rn = rn_fao56(rs_mj, tmax, tmin, ea, lat, doy, z)
    num = 0.408 * D * rn + g_psi * (900 / (tmean + 273)) * u2 * max(es_mean - ea, 0.0)
    den = D + g_psi * (1 + 0.34 * u2)
    return num / den, rn


# ----------------------------------------------------- passos do artigo ----
def normalizar(v):
    n = len(v)
    mu = sum(v) / n
    sd = math.sqrt(sum((x - mu) ** 2 for x in v) / (n - 1)) if n > 1 else 0.0
    if sd == 0:
        return [0.0] * n, mu, 0.0
    return [(x - mu) / sd for x in v], mu, sd


def classe(c):
    if c is None:
        return ""
    if c < 0.20: return "No Stress"
    if c < 0.40: return "Low Stress"
    if c < 0.60: return "Moderate Stress"
    if c < 0.80: return "High Stress"
    return "Extreme Stress"


# ------------------------------------------------------------- leitura ----
def num(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).strip().replace(",", "."))
    except ValueError:
        return None


def ler_15min(path, aba, solo="0-60"):
    """Le a aba e devolve registros com chaves uniformes, seja qual for o layout.

    As abas derivadas (03/06) so trazem WC 15 cm, entao qualquer escolha de
    solo que exija os sensores profundos cai de volta para 0-15 cm.
    """
    layout = LAYOUT_DE.get(aba, "bruto")
    cols = COL_BRUTO if layout == "bruto" else COL_DERIVADO
    camadas = CAMADAS[solo] if layout == "bruto" else CAMADAS["0-15"]

    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True, read_only=True)
    if aba not in wb.sheetnames:
        print(f"ERRO: aba '{aba}' nao existe. Abas: {wb.sheetnames}", file=sys.stderr)
        sys.exit(2)
    ws = wb[aba]
    ic = cols["ts"] - 1
    regs = []
    for r in ws.iter_rows(min_row=LINHA_CAB + 1, values_only=True):
        if not r or len(r) <= ic or r[ic] is None:
            continue
        ts = r[ic]
        if not isinstance(ts, datetime):
            try:
                ts = datetime.fromisoformat(str(ts).split(".")[0])
            except ValueError:
                continue
        d = {"ts": ts}
        for k, c in cols.items():
            if k == "ts":
                continue
            d[k] = num(r[c - 1]) if c - 1 < len(r) else None
        # vento sempre em m/s
        if d.get("wind_ms") is not None:
            d["wind"] = d["wind_ms"]
        elif d.get("wind_kmh") is not None:
            d["wind"] = d["wind_kmh"] / 3.6
        else:
            d["wind"] = None
        for k in ("patm", "tmax", "tmin", "wc30", "wc45", "wc60"):
            d.setdefault(k, None)
        regs.append(d)
    wb.close()
    return regs, layout, camadas


def ler_setup(path):
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb["01_Setup"]
    out = {}
    for r in ws.iter_rows(min_row=1, max_row=60, max_col=4, values_only=True):
        if not r or r[0] is None:
            continue
        rot = str(r[0]).strip().lower()
        val = num(r[1])
        if val is None:
            continue
        if rot.startswith("latitude (deg"):
            out["lat"] = val
        elif rot.startswith("elevation"):
            out["alt"] = val
    wb.close()
    return out


def agregar_diario(regs, camadas, alt, min_frac=0.80):
    """15 min -> diario. Descarta dias com menos de min_frac dos 96 registros.

    Quando a aba nao traz pressao atmosferica (layout derivado), usa a
    pressao padrao pela altitude (FAO-56 eq. 7). Quando nao traz Tmax/Tmin,
    tira do proprio Ta de 15 min, o que e' mais fiel ao dia do que as
    colunas agregadas do logger.
    """
    patm_padrao = 101.3 * ((293 - 0.0065 * alt) / 293) ** 5.26
    dias = defaultdict(list)
    for d in regs:
        dias[d["ts"].date()].append(d)
    saida, descartados = [], []
    for data in sorted(dias):
        g = dias[data]

        def col(k):
            return [x[k] for x in g if x.get(k) is not None]

        tair, rs, wind, ea = col("tair"), col("rs"), col("wind"), col("ea")
        patm, precip = col("patm"), col("precip")
        if len(g) < 96 * min_frac or not (tair and rs and wind and ea):
            descartados.append((data, len(g)))
            continue
        s_vals = []
        for x in g:
            tot, ok = 0.0, True
            for k, esp in camadas:
                if x.get(k) is None:
                    ok = False
                    break
                tot += (x[k] / 100.0) * esp
            if ok:
                s_vals.append(tot)
        if not s_vals:
            descartados.append((data, len(g)))
            continue
        tcs = col("tc")
        saida.append({
            "data": data,
            "n": len(g),
            "tmean": sum(tair) / len(tair),
            "tmax": max(col("tmax") or tair),
            "tmin": min(col("tmin") or tair),
            "ea": sum(ea) / len(ea),
            "u2": sum(wind) / len(wind),                   # ja em m/s
            "rs_mj": (sum(rs) / len(rs)) * 0.0864,         # W/m2 medio -> MJ/m2/dia
            "patm": (sum(patm) / len(patm)) if patm else patm_padrao,
            "precip": sum(precip),
            "S": sum(s_vals) / len(s_vals),                # armazenamento medio, mm
            "tc": (sum(tcs) / len(tcs)) if tcs else None,
        })
    return saida, descartados


# ---------------------------------------------------------------- main ----
def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--xlsx", required=True)
    ap.add_argument("--sheet", default="05_Season_Input")
    ap.add_argument("--csv")
    ap.add_argument("--lat", type=float)
    ap.add_argument("--alt", type=float)
    ap.add_argument("--mostrar", type=int, default=12)
    ap.add_argument("--solo", choices=sorted(CAMADAS), default="0-60",
                    help="profundidade de solo no balanco hidrico. 0-60 (padrao) usa "
                         "toda a coluna instrumentada. So vale nas abas brutas 02/05; "
                         "as derivadas 03/06 so tem WC 15 cm.")
    args = ap.parse_args()

    setup = ler_setup(args.xlsx)
    lat = args.lat if args.lat is not None else setup.get("lat", 51.7782)
    alt = args.alt if args.alt is not None else setup.get("alt", 995.0)

    regs, layout, camadas = ler_15min(args.xlsx, args.sheet, args.solo)
    dias, desc = agregar_diario(regs, camadas, alt)
    if len(dias) < 3:
        print("Dados diarios insuficientes.", file=sys.stderr)
        sys.exit(1)

    perfil_mm = sum(t for _, t in camadas)
    prof = f"{'0-15 cm' if len(camadas) == 1 else args.solo.replace('.', ',') + ' cm'}"
    if layout == "derivado" and args.solo != "0-15":
        prof += "  (a aba derivada so tem WC 15 cm)"
    print("=" * 92)
    print(f"ABA {args.sheet}  (layout {layout})  |  lat {lat:.4f}  alt {alt:.0f} m")
    print(f"umidade do solo: {prof}  ({perfil_mm:.0f} mm de perfil, {len(camadas)} sensor(es))")
    print(f"{len(regs)} registros de 15 min  ->  {len(dias)} dias completos"
          f"  ({len(desc)} dias descartados por cobertura < 80%)")
    print(f"periodo: {dias[0]['data']}  a  {dias[-1]['data']}")
    print("=" * 92)

    # ---- Passo 1
    pet, rn_l = [], []
    for d in dias:
        p, rn = pet_pm_diaria(d["tmean"], d["tmax"], d["tmin"], d["ea"], d["u2"],
                              d["rs_mj"], d["patm"], lat, alt,
                              d["data"].timetuple().tm_yday)
        pet.append(p)
        rn_l.append(rn)
    pet_n, mu_p, sd_p = normalizar(pet)

    # ---- Passo 2
    eta = [None]
    for i in range(1, len(dias)):
        eta.append(dias[i]["precip"] + (dias[i - 1]["S"] - dias[i]["S"]))
    idx = list(range(1, len(dias)))
    eta_n_vals, mu_e, sd_e = normalizar([eta[i] for i in idx])
    eta_n = [None] + eta_n_vals

    # ---- Passo 3
    c_art, c_flx = [None], [None]
    for i in idx:
        p, e = pet_n[i], eta_n[i]
        c_art.append(float("inf") if abs(p) < 1e-9 else (p - e) / p)
        c_flx.append(max(0.0, min(1.0, 1 - eta[i] / pet[i])) if pet[i] > 0 else None)

    print(f"\nPasso 1  PET  media {mu_p:7.3f} mm/dia   desvio {sd_p:6.3f}"
          f"   min {min(pet):.2f}  max {max(pet):.2f}")
    print(f"Passo 2  ETa  media {mu_e:7.3f} mm/dia   desvio {sd_e:6.3f}"
          f"   min {min(eta[i] for i in idx):.2f}  max {max(eta[i] for i in idx):.2f}")
    print(f"         armazenamento do perfil: {min(d['S'] for d in dias):.1f} a "
          f"{max(d['S'] for d in dias):.1f} mm")

    print("\n" + "-" * 92)
    print(f"{'data':>11}{'PET':>7}{'ETa':>8}{'PETn':>7}{'ETan':>7}"
          f"{'CWSI art.':>11}{'CWSI flux':>10}  {'classe (fluxos)':<17}{'chuva':>7}")
    print("-" * 92)
    for i in idx[:args.mostrar]:
        ca = c_art[i]
        ca_s = "      inf" if ca == float("inf") else f"{ca:9.2f}"
        print(f"{str(dias[i]['data']):>11}{pet[i]:7.2f}{eta[i]:8.2f}{pet_n[i]:7.2f}"
              f"{eta_n[i]:7.2f}{ca_s:>11}"
              f"{(f'{c_flx[i]:.3f}' if c_flx[i] is not None else '-'):>10}  "
              f"{classe(c_flx[i]):<17}{dias[i]['precip']:7.1f}")

    # ---- diagnostico
    vals = [c_art[i] for i in idx if c_art[i] != float("inf")]
    fora = [c for c in vals if c < 0 or c > 1]
    quase0 = sum(1 for i in idx if abs(pet_n[i]) < 0.25)
    print("\n" + "=" * 92)
    print("DIAGNOSTICO DO PASSO 3 LITERAL   CWSI = (PET_norm - ETa_norm) / PET_norm")
    print("=" * 92)
    print(f"  dias avaliados .................. {len(vals)}")
    print(f"  minimo ......................... {min(vals):12.2f}")
    print(f"  maximo ......................... {max(vals):12.2f}")
    print(f"  fora de [0, 1] ................. {len(fora)} de {len(vals)}"
          f"  ({100 * len(fora) / len(vals):.0f}%)")
    print(f"  |PET_norm| < 0,25 .............. {quase0} dias (denominador quase zero)")
    print('\n  O artigo afirma "The CWSI is bounded on [0, 1]".')
    print("  Com escore-z o denominador tem media zero: cruza zero, troca de sinal,")
    print("  e o quociente nao tem limite. A formula, como publicada, nao fecha.")

    v2 = [c_flx[i] for i in idx if c_flx[i] is not None]
    print("\n" + "=" * 92)
    print("VARIANTE COM FLUXOS BRUTOS   CWSI = 1 - ETa/PET, limitada a [0, 1]")
    print("=" * 92)
    print(f"  min {min(v2):.3f}   max {max(v2):.3f}   media {sum(v2) / len(v2):.3f}")
    cont = defaultdict(int)
    for c in v2:
        cont[classe(c)] += 1
    for k in ["No Stress", "Low Stress", "Moderate Stress", "High Stress", "Extreme Stress"]:
        if cont[k]:
            print(f"    {k:<17}{cont[k]:>4} dias  ({100 * cont[k] / len(v2):4.1f}%)")

    if args.csv:
        with open(args.csv, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["data", "n_registros", "Tmean_C", "Tmax_C", "Tmin_C", "ea_kPa",
                        "u2_m_s", "Rs_MJ_m2_dia", "Rn_MJ_m2_dia", "Patm_kPa",
                        "chuva_mm", "S_perfil_mm", "Tc_C",
                        "PET_mm", "ETa_mm", "PET_norm", "ETa_norm",
                        "CWSI_artigo_literal", "CWSI_fluxos_0_1", "classe_fluxos"])
            for i, d in enumerate(dias):
                w.writerow([d["data"], d["n"], f"{d['tmean']:.2f}", f"{d['tmax']:.2f}",
                            f"{d['tmin']:.2f}", f"{d['ea']:.3f}", f"{d['u2']:.2f}",
                            f"{d['rs_mj']:.2f}", f"{rn_l[i]:.2f}", f"{d['patm']:.2f}",
                            f"{d['precip']:.1f}", f"{d['S']:.1f}",
                            (f"{d['tc']:.2f}" if d["tc"] is not None else ""),
                            f"{pet[i]:.3f}",
                            (f"{eta[i]:.3f}" if eta[i] is not None else ""),
                            f"{pet_n[i]:.3f}",
                            (f"{eta_n[i]:.3f}" if eta_n[i] is not None else ""),
                            ("" if c_art[i] is None else
                             ("inf" if c_art[i] == float("inf") else f"{c_art[i]:.4f}")),
                            (f"{c_flx[i]:.4f}" if c_flx[i] is not None else ""),
                            classe(c_flx[i]) if c_flx[i] is not None else ""])
        print(f"\nCSV gravado: {args.csv}")


if __name__ == "__main__":
    main()
