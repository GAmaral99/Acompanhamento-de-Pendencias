#!/usr/bin/env python3
"""
encontrar_e_gerar.py
--------------------
Executado pela GitHub Action. Responsável por:
  1. Listar todos os .xlsx na pasta data/
  2. Ordenar por data de modificação (nome do arquivo contém timestamp)
  3. Selecionar:
       - BASE    = primeiro arquivo da semana atual (segunda-feira)
       - ATUAL   = arquivo mais recente
  4. Chamar gerar_relatorio.py para produzir o index.html
"""

import os
import sys
import re
import subprocess
from datetime import datetime, date, timedelta
from pathlib import Path

DATA_DIR   = Path("data")
OUTPUT     = Path("index.html")
SCRIPT     = Path("scripts/gerar_relatorio.py")

# ── Helpers ──────────────────────────────────────────────────────────────────

def segunda_da_semana(d: date) -> date:
    """Retorna a segunda-feira da semana que contém d."""
    return d - timedelta(days=d.weekday())

def extrair_dt_do_nome(nome: str) -> datetime | None:
    """
    Tenta extrair um datetime do nome do arquivo.
    Aceita padrões comuns:
      01-06-26_1009   →  01/06/2026 10:09
      2026-06-01      →  01/06/2026 00:00
      01-06-2026_1009 →  01/06/2026 10:09
    Caso não encontre, retorna None.
    """
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
                    if len(g[2]) == 2:                              # DD-MM-YY_HHMM
                        return datetime(2000+int(g[2]), int(g[1]), int(g[0]),
                                        int(g[3][:2]), int(g[3][2:]))
                    elif len(g[0]) == 4:                            # YYYY-MM-DD_HHMM
                        return datetime(int(g[0]), int(g[1]), int(g[2]),
                                        int(g[3][:2]), int(g[3][2:]))
                    else:                                           # DD-MM-YYYY_HHMM
                        return datetime(int(g[2]), int(g[1]), int(g[0]),
                                        int(g[3][:2]), int(g[3][2:]))
                elif len(g) == 3:
                    if len(g[2]) == 2:                              # DD-MM-YY
                        return datetime(2000+int(g[2]), int(g[1]), int(g[0]))
                    elif len(g[0]) == 4:                            # YYYY-MM-DD
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

# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    if not DATA_DIR.exists():
        print(f"[!] Pasta '{DATA_DIR}' não encontrada. Crie-a e adicione os arquivos Excel.")
        sys.exit(1)

    arquivos = list(DATA_DIR.glob("*.xlsx"))
    if not arquivos:
        print(f"[!] Nenhum .xlsx encontrado em '{DATA_DIR}'.")
        sys.exit(1)

    ordenados = ordenar_arquivos(arquivos)

    print(f"[*] {len(ordenados)} arquivo(s) encontrado(s) em '{DATA_DIR}':")
    for dt, arq in ordenados:
        print(f"    {dt.strftime('%d/%m/%Y %H:%M')}  →  {arq.name}")

    # Arquivo mais recente = ATUAL
    dt_atual, arq_atual = ordenados[-1]

    # Primeiro arquivo da mesma semana (seg a dom) = BASE
    segunda = segunda_da_semana(dt_atual.date())
    da_semana = [(dt, arq) for dt, arq in ordenados if dt.date() >= segunda]

    if da_semana:
        dt_base, arq_base = da_semana[0]
    else:
        # fallback: usa o mais antigo disponível
        dt_base, arq_base = ordenados[0]

    print(f"\n[*] BASE  : {arq_base.name}  ({dt_base.strftime('%d/%m/%Y %H:%M')})")
    print(f"[*] ATUAL : {arq_atual.name}  ({dt_atual.strftime('%d/%m/%Y %H:%M')})")

    if arq_base == arq_atual:
        print("[*] Apenas um arquivo na semana — modo visualização única.")
        cmd = [sys.executable, str(SCRIPT), str(arq_base), "-o", str(OUTPUT)]
    else:
        cmd = [sys.executable, str(SCRIPT), str(arq_base), str(arq_atual), "-o", str(OUTPUT)]

    print(f"\n[*] Executando: {' '.join(cmd)}\n")
    result = subprocess.run(cmd)
    sys.exit(result.returncode)

if __name__ == "__main__":
    main()
