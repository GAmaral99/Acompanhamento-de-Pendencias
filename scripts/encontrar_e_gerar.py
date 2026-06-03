#!/usr/bin/env python3
"""
encontrar_e_gerar.py
--------------------
Executado pela GitHub Action. Responsável por:
  1. Listar todos os .xlsx na pasta data/
  2. Ordenar por data/hora extraída do nome (ou mtime)
  3. Identificar arquivos de referência:
       - ATUAL      = arquivo mais recente
       - BASE_DIA   = primeiro arquivo do mesmo dia calendário que ATUAL
       - BASE_SEM   = primeiro arquivo da semana (segunda-feira) que contém ATUAL
       - BASE_MES   = primeiro arquivo do mês calendário que contém ATUAL
  4. Chamar gerar_relatorio.py com todos os parâmetros
"""

import os
import sys
import re
import subprocess
from datetime import datetime, date, timedelta
from pathlib import Path

DATA_DIR = Path("data")
OUTPUT   = Path("index.html")
SCRIPT   = Path("scripts/gerar_relatorio.py")

# ── Helpers ──────────────────────────────────────────────────────────────────

def segunda_da_semana(d: date) -> date:
    return d - timedelta(days=d.weekday())

def extrair_dt_do_nome(nome: str) -> datetime | None:
    padroes = [
        r'(\d{2})-(\d{2})-(\d{2})_(\d{4})',   # DD-MM-YY_HHMM
        r'(\d{4})-(\d{2})-(\d{2})_(\d{4})',   # YYYY-MM-DD_HHMM
        r'(\d{2})-(\d{2})-(\d{4})_(\d{4})',   # DD-MM-YYYY_HHMM
        r'(\d{2})-(\d{2})-(\d{2})',            # DD-MM-YY
        r'(\d{4})-(\d{2})-(\d{2})',            # YYYY-MM-DD
    ]
    for p in padroes:
        m = re.search(p, nome)
        if m:
            g = m.groups()
            try:
                if len(g) == 4:
                    if len(g[2]) == 2:
                        return datetime(2000+int(g[2]), int(g[1]), int(g[0]),
                                        int(g[3][:2]), int(g[3][2:]))
                    elif len(g[0]) == 4:
                        return datetime(int(g[0]), int(g[1]), int(g[2]),
                                        int(g[3][:2]), int(g[3][2:]))
                    else:
                        return datetime(int(g[2]), int(g[1]), int(g[0]),
                                        int(g[3][:2]), int(g[3][2:]))
                elif len(g) == 3:
                    if len(g[2]) == 2:
                        return datetime(2000+int(g[2]), int(g[1]), int(g[0]))
                    elif len(g[0]) == 4:
                        return datetime(int(g[0]), int(g[1]), int(g[2]))
            except ValueError:
                continue
    return None

def mtime(p: Path) -> datetime:
    return datetime.fromtimestamp(p.stat().st_mtime)

def ordenar_arquivos(arquivos: list[Path]) -> list[tuple[datetime, Path]]:
    resultado = []
    for arq in arquivos:
        dt = extrair_dt_do_nome(arq.stem) or mtime(arq)
        resultado.append((dt, arq))
    return sorted(resultado, key=lambda x: x[0])

def primeiro_da_condicao(ordenados, condicao):
    """Retorna (dt, path) do primeiro arquivo que satisfaz a condição, ou None."""
    for dt, arq in ordenados:
        if condicao(dt):
            return dt, arq
    return None

# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    if not DATA_DIR.exists():
        print(f"[!] Pasta '{DATA_DIR}' não encontrada.")
        sys.exit(1)

    arquivos = list(DATA_DIR.glob("*.xlsx"))
    if not arquivos:
        print(f"[!] Nenhum .xlsx encontrado em '{DATA_DIR}'.")
        sys.exit(1)

    ordenados = ordenar_arquivos(arquivos)

    print(f"[*] {len(ordenados)} arquivo(s) encontrado(s) em '{DATA_DIR}':")
    for dt, arq in ordenados:
        print(f"    {dt.strftime('%d/%m/%Y %H:%M')}  →  {arq.name}")

    # ATUAL = mais recente
    dt_atual, arq_atual = ordenados[-1]
    d_atual = dt_atual.date()

    # BASE_DIA  = primeiro arquivo do mesmo dia
    res_dia = primeiro_da_condicao(ordenados, lambda dt: dt.date() == d_atual)
    dt_dia, arq_dia = res_dia if res_dia else (dt_atual, arq_atual)

    # BASE_SEM  = primeiro arquivo a partir da segunda-feira da semana atual
    segunda = segunda_da_semana(d_atual)
    res_sem = primeiro_da_condicao(ordenados, lambda dt: dt.date() >= segunda)
    dt_sem, arq_sem = res_sem if res_sem else (dt_atual, arq_atual)

    # BASE_MES  = primeiro arquivo do mês atual
    res_mes = primeiro_da_condicao(
        ordenados, lambda dt: dt.year == d_atual.year and dt.month == d_atual.month
    )
    dt_mes, arq_mes = res_mes if res_mes else (dt_atual, arq_atual)

    print(f"\n[*] ATUAL    : {arq_atual.name}  ({dt_atual.strftime('%d/%m/%Y %H:%M')})")
    print(f"[*] BASE DIA : {arq_dia.name}   ({dt_dia.strftime('%d/%m/%Y %H:%M')})")
    print(f"[*] BASE SEM : {arq_sem.name}   ({dt_sem.strftime('%d/%m/%Y %H:%M')})")
    print(f"[*] BASE MÊS : {arq_mes.name}   ({dt_mes.strftime('%d/%m/%Y %H:%M')})")
    
    COORD_XLSX = str(Path("data/coordenadores/Coordenadores (SP-RJ-Santos).xlsx"))   
    cmd = [
        sys.executable, str(SCRIPT),
        str(arq_mes),    # base principal (mês) — mantém compatibilidade
        str(arq_atual),
        "--base-dia", str(arq_dia),
        "--base-sem", str(arq_sem),
        "--base-mes", str(arq_mes),
        "--coordenadores", COORD_XLSX,
        "-o", str(OUTPUT),
    ]

    print(f"\n[*] Executando: {' '.join(cmd)}\n")
    result = subprocess.run(cmd)
    sys.exit(result.returncode)

if __name__ == "__main__":
    main()