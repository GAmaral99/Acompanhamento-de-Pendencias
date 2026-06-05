#!/usr/bin/env python3
"""
historico.py
------------
Acumula o histórico diário de baixas por funcionário em:
  data/historico/historico_baixas.xlsx

Chamado pelo encontrar_e_gerar.py após gerar o relatório HTML.

Uso:
  python scripts/historico.py \
      --placares-json '<json>' \
      --data-ref 2026-06-05 \
      [--apenas-sp]
"""

import os
import sys
import json
import argparse
from datetime import datetime, date
from pathlib import Path

import pandas as pd

HISTORICO_DIR  = Path("data/historico")
HISTORICO_FILE = HISTORICO_DIR / "historico_baixas.xlsx"
SHEET_NAME     = "Historico"

FILIAIS_SP = ["SP"]   # filiais consideradas "São Paulo"

# Colunas do arquivo histórico
COLUNAS = [
    "Data",           # date  YYYY-MM-DD
    "DiaSemana",      # str   Segunda, Terça...
    "Semana",         # str   2026-W23
    "Mes",            # str   2026-06
    "Filial",
    "Departamento",
    "Funcionario",
    "Coordenador",
    "BaixasDia",
    "BaixasSemana",
    "BaixasMes",
]

DIAS_PT = ["Segunda","Terça","Quarta","Quinta","Sexta","Sábado","Domingo"]


def semana_iso(d: date) -> str:
    return f"{d.isocalendar()[0]}-W{d.isocalendar()[1]:02d}"


def mes_str(d: date) -> str:
    return d.strftime("%Y-%m")


def carregar_historico() -> pd.DataFrame:
    if HISTORICO_FILE.exists():
        try:
            df = pd.read_excel(HISTORICO_FILE, sheet_name=SHEET_NAME, dtype=str)
            df.columns = [str(c).strip() for c in df.columns]
            # garantir colunas numéricas
            for col in ["BaixasDia","BaixasSemana","BaixasMes"]:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)
            return df
        except Exception as e:
            print(f"[!] Erro ao ler histórico existente: {e}")
    return pd.DataFrame(columns=COLUNAS)


def salvar_historico(df: pd.DataFrame):
    HISTORICO_DIR.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(HISTORICO_FILE, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=SHEET_NAME, index=False)
        # Ajusta largura das colunas
        ws = writer.sheets[SHEET_NAME]
        for col in ws.columns:
            max_len = max((len(str(c.value)) for c in col if c.value), default=8)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 40)
    print(f"[✓] Histórico salvo em {HISTORICO_FILE}")


def processar_placares(placares: dict, data_ref: date, apenas_sp: bool) -> list[dict]:
    """
    Converte o dict de placares (gerado por calcular_placares) em
    linhas prontas para o DataFrame histórico.
    """
    linhas = []

    dia_semana = DIAS_PT[data_ref.weekday()]
    semana     = semana_iso(data_ref)
    mes        = mes_str(data_ref)

    # Monta um índice (filial, dep, resp) → {dia, sem, mes}
    idx: dict[tuple, dict] = {}

    for periodo in ["dia", "sem", "mes"]:
        pd_data = placares.get(periodo, {})
        for r in pd_data.get("por_resp", []):
            filial = r.get("unidade", "")
            dep    = r.get("dep", "")
            resp   = r.get("resp", "")
            coord  = r.get("coord", "")
            n      = int(r.get("n", 0))

            if apenas_sp and filial not in FILIAIS_SP:
                continue

            chave = (filial, dep, resp, coord)
            if chave not in idx:
                idx[chave] = {"dia": 0, "sem": 0, "mes": 0}
            idx[chave][periodo] = n

    for (filial, dep, resp, coord), vals in idx.items():
        linhas.append({
            "Data":        str(data_ref),
            "DiaSemana":   dia_semana,
            "Semana":      semana,
            "Mes":         mes,
            "Filial":      filial,
            "Departamento":dep,
            "Funcionario": resp,
            "Coordenador": coord,
            "BaixasDia":   vals["dia"],
            "BaixasSemana":vals["sem"],
            "BaixasMes":   vals["mes"],
        })

    return linhas


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--placares-json", required=True,
                        help="JSON string com o dict de placares")
    parser.add_argument("--data-ref", default=None,
                        help="Data de referência YYYY-MM-DD (padrão: hoje)")
    parser.add_argument("--apenas-sp", action="store_true",
                        help="Processa apenas filiais de SP")
    args = parser.parse_args()

    # Data de referência
    if args.data_ref:
        data_ref = date.fromisoformat(args.data_ref)
    else:
        data_ref = date.today()

    print(f"[*] historico.py — data de referência: {data_ref}")

    # Parse do JSON de placares
    try:
        placares = json.loads(args.placares_json)
    except Exception as e:
        print(f"[!] Erro ao parsear placares JSON: {e}")
        sys.exit(1)

    # Carrega histórico existente
    df_hist = carregar_historico()
    print(f"[*] Histórico existente: {len(df_hist)} registros")

    # Remove registros do mesmo dia para evitar duplicatas
    data_str = str(data_ref)
    if not df_hist.empty and "Data" in df_hist.columns:
        antes = len(df_hist)
        df_hist = df_hist[df_hist["Data"] != data_str].copy()
        removidos = antes - len(df_hist)
        if removidos:
            print(f"[*] {removidos} registro(s) do dia {data_str} removidos (serão reescritos)")

    # Processa novos dados
    novas_linhas = processar_placares(placares, data_ref, args.apenas_sp)
    print(f"[*] Novos registros a inserir: {len(novas_linhas)}")

    if novas_linhas:
        df_novo = pd.DataFrame(novas_linhas, columns=COLUNAS)
        df_hist = pd.concat([df_hist, df_novo], ignore_index=True)

    # Ordena por Data + Filial + Dep + Funcionario
    if not df_hist.empty:
        df_hist = df_hist.sort_values(
            ["Data","Filial","Departamento","Funcionario"]
        ).reset_index(drop=True)

    salvar_historico(df_hist)
    print(f"[*] Total de registros no histórico: {len(df_hist)}")


if __name__ == "__main__":
    main()
