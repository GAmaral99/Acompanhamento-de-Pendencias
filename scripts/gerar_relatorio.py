#!/usr/bin/env python3
"""
MG Contécnica – Gerador de Relatório HTML Consolidado
Versão 6.1
"""

import os, re, sys, json, argparse
import pandas as pd
from datetime import datetime

SHEET = "Pendencias CONSOLIDADO"

C_COD_CLIENTE  = "CodCliente"
C_RAZAO        = "RazaoSocial"
C_UNIDADE      = "Unidade"
C_DEPARTAMENTO = "Departamento"
C_TITULO       = "Titulo"
C_VENCIMENTO   = "DataVencimento"
C_PREVISAO     = "DataPrevisaoConclusao"
C_RESPONSAVEL  = "UsuarioResponsavel"
C_GRUPO        = "Grupo"
C_COMENTARIO   = "Comentario"

def normalizar_dep(dep):
    return re.sub(r'\s*-\s*\(Cliente Novo\)\s*$', '', str(dep).strip(), flags=re.IGNORECASE)

def formatar_data(valor):
    if pd.isna(valor) or str(valor).strip() in ("", "nan", "NaT", "None"):
        return ""
    try:
        v = pd.to_datetime(valor, dayfirst=True, errors="coerce")
        return v.strftime("%d/%m/%Y") if not pd.isna(v) else str(valor).strip()
    except:
        return str(valor).strip()

def extrair_data_comentario(texto):
    oc = re.findall(r'Data Coment[aá]rio[:\s]+(\d{2}/\d{2}/\d{4}(?:\s+\d{2}:\d{2})?)', str(texto), re.IGNORECASE)
    return oc[-1].strip() if oc else ""

def limpar(v):
    s = str(v).strip()
    return "" if s in ("nan", "None", "NaT") else s

def carregar_coordenadores(caminho):
    if not caminho or not os.path.exists(caminho):
        return {}, set()
    try:
        df = pd.read_excel(caminho, sheet_name=0, header=0)
        df.columns = [str(c).strip() for c in df.columns]
        if len(df.columns) < 4:
            return {}, set()
        col_exib  = df.columns[1]
        col_coord = df.columns[3]

        resp_to_coords = {}
        all_coords = set()

        for _, row in df.iterrows():
            nome      = limpar(str(row[col_exib]))
            coord_raw = limpar(str(row[col_coord]))
            if not nome or not coord_raw:
                continue
            coords = [c.strip() for c in coord_raw.split("|") if c.strip()]
            resp_to_coords[nome] = coords
            for c in coords:
                all_coords.add(c)

        for coord in list(all_coords):
            if coord not in resp_to_coords:
                resp_to_coords[coord] = [coord]

        return resp_to_coords, all_coords
    except Exception as e:
        print(f"[!] Erro ao carregar coordenadores: {e}")
        return {}, set()

def ler_relatorio(caminho, resp_to_coords=None):
    try:
        df_raw = pd.read_excel(caminho, sheet_name=SHEET, dtype=str)
    except Exception as e:
        raise ValueError(f"Não foi possível ler '{os.path.basename(caminho)}': {e}")
    registros = []
    for _, row in df_raw.iterrows():
        dep_raw = limpar(row.get(C_DEPARTAMENTO, ""))
        titulo  = limpar(row.get(C_TITULO, ""))
        if not dep_raw and not titulo:
            continue
        dep_norm     = normalizar_dep(dep_raw)
        cliente_novo = bool(re.search(r'\(Cliente Novo\)', dep_raw, re.IGNORECASE))
        comt_raw     = limpar(row.get(C_COMENTARIO, ""))
        responsavel  = limpar(row.get(C_RESPONSAVEL, ""))

        coord_list = []
        if resp_to_coords and responsavel:
            coord_list = resp_to_coords.get(responsavel, [])

        registros.append({
            "CodCliente":   limpar(row.get(C_COD_CLIENTE, "")),
            "RazaoSocial":  limpar(row.get(C_RAZAO, "")),
            "Unidade":      limpar(row.get(C_UNIDADE, "")),
            "Departamento": dep_norm,
            "ClienteNovo":  cliente_novo,
            "Titulo":       titulo,
            "Vencimento":   formatar_data(row.get(C_VENCIMENTO, "")),
            "Previsao":     formatar_data(row.get(C_PREVISAO, "")),
            "Responsavel":  responsavel,
            "Grupo":        limpar(row.get(C_GRUPO, "")),
            "Comentario":   comt_raw,
            "DataComt":     extrair_data_comentario(comt_raw),
            "Coordenador":  "|".join(coord_list),
            "chave":        f"{limpar(row.get(C_COD_CLIENTE,''))}||{limpar(row.get('Cod',''))}||{dep_norm}||{titulo}",
        })
    df = pd.DataFrame(registros)
    if df.empty:
        raise ValueError(f"Nenhuma tarefa válida em '{os.path.basename(caminho)}'.")
    return df

def comparar(df_base, df_atual):
    cb = set(df_base["chave"]); ca = set(df_atual["chave"])
    return (df_atual[df_atual["chave"].isin(ca)].copy(),
            df_base[df_base["chave"].isin(cb - ca)].copy(),
            df_atual[df_atual["chave"].isin(ca - cb)].copy())

def df_to_js(df):
    if df.empty: return "[]"
    rows = []
    for _, r in df.iterrows():
        rows.append({
            "unidade": r["Unidade"], "dep": r["Departamento"], "novo": r["ClienteNovo"],
            "tit": r["Titulo"], "venc": r["Vencimento"], "prev": r["Previsao"],
            "resp": r["Responsavel"], "grupo": r["Grupo"],
            "cod": r["CodCliente"], "razao": r["RazaoSocial"],
            "dataComt": r["DataComt"], "comt": r["Comentario"],
            "coord": r.get("Coordenador", ""),
        })
    return json.dumps(rows, ensure_ascii=False)

def calcular_placares(df_atual, df_dia=None, df_sem=None, df_mes=None):
    def baixas_entre(df_base, df_ref):
        if df_base is None or df_ref is None:
            return pd.DataFrame()
        cb = set(df_base["chave"])
        cr = set(df_ref["chave"])
        return df_base[df_base["chave"].isin(cb - cr)].copy()

    resultado = {}
    for periodo, df_b in [("dia", df_dia), ("sem", df_sem), ("mes", df_mes)]:
        baixas = baixas_entre(df_b, df_atual)
        if baixas.empty:
            resultado[periodo] = {"total": 0, "por_dep": [], "por_resp": []}
        else:
            agg_dep = (baixas.groupby(["Unidade", "Departamento"])
                         .size()
                         .reset_index(name="n")
                         .sort_values(["Unidade", "n"], ascending=[True, False]))
            agg_resp = (baixas.groupby(["Unidade", "Departamento", "Responsavel", "Coordenador"])
                          .size()
                          .reset_index(name="n")
                          .sort_values(["Unidade", "Departamento", "n"], ascending=[True, True, False]))
            resultado[periodo] = {
                "total": int(baixas.shape[0]),
                "por_dep": [
                    {"unidade": r["Unidade"], "dep": r["Departamento"], "n": int(r["n"])}
                    for _, r in agg_dep.iterrows()
                ],
                "por_resp": [
                    {
                        "unidade": r["Unidade"],
                        "dep": r["Departamento"],
                        "resp": r["Responsavel"],
                        "coord": r["Coordenador"],
                        "n": int(r["n"])
                    }
                    for _, r in agg_resp.iterrows()
                ],
            }
    return resultado

def gerar_html(arquivo_base, arquivo_atual, em_andamento, finalizadas, reabertas,
               placares=None, all_coords=None):
    nome_base  = os.path.basename(arquivo_base)
    nome_atual = os.path.basename(arquivo_atual) if arquivo_atual else nome_base
    gerado_em  = datetime.now().strftime("%d/%m/%Y às %H:%M")
    modo_comp  = arquivo_atual and arquivo_atual != arquivo_base

    js_man = df_to_js(em_andamento)
    js_fin = df_to_js(finalizadas)
    js_add = df_to_js(reabertas)
    total_man = len(em_andamento)
    total_fin = len(finalizadas)
    total_add = len(reabertas)

    js_placares = json.dumps(placares or {}, ensure_ascii=False)

    if all_coords:
        js_all_coords = json.dumps(sorted(all_coords), ensure_ascii=False)
    else:
        coords_set = set()
        for df in [em_andamento, finalizadas, reabertas]:
            if not df.empty and "Coordenador" in df.columns:
                for val in df["Coordenador"].dropna():
                    for c in str(val).split("|"):
                        c = c.strip()
                        if c:
                            coords_set.add(c)
        js_all_coords = json.dumps(sorted(coords_set), ensure_ascii=False)

    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>MG Contécnica · Relatório de Pendências</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js"></script>
<style>
:root{{
  --bg:#0a0a0a;--surface:#111111;--surface2:#181818;--border:#2a2a2a;
  --red:#E31E24;--red-dim:#7a1012;--red-glow:rgba(227,30,36,.18);
  --white:#ffffff;--off-white:#f0f0f0;--muted:#666;--muted2:#444;
  --green:#22c97a;--amber:#f5a623;--blue:#4da6ff;
  --font:'Inter',sans-serif;--mono:'JetBrains Mono',monospace;
}}
*,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
html{{scroll-behavior:smooth}}
body{{font-family:var(--font);background:var(--bg);color:var(--white);min-height:100vh;overflow-x:hidden;font-size:14px;line-height:1.5;-webkit-font-smoothing:antialiased}}

body::before{{content:'';position:fixed;inset:0;
  background-image:linear-gradient(rgba(227,30,36,.025) 1px,transparent 1px),linear-gradient(90deg,rgba(227,30,36,.025) 1px,transparent 1px);
  background-size:48px 48px;pointer-events:none;z-index:0}}

.screen{{position:relative;z-index:1;display:none;min-height:100vh;flex-direction:column}}
.screen.active{{display:flex}}

/* ── LOGO BAR ── */
.logo-bar{{display:flex;align-items:center;gap:14px;padding:24px 48px 0}}
.logo-img{{width:52px;height:52px;object-fit:contain;border-radius:50%;background:#fff;padding:4px}}
.logo-text{{font-size:11px;font-weight:600;letter-spacing:.14em;text-transform:uppercase;color:var(--muted)}}
.logo-text strong{{color:var(--white);display:block;font-size:13px;letter-spacing:.06em}}

/* ── SCREEN 1 FILIAL ── */
#screen-filial{{align-items:center;justify-content:center;padding:48px}}
.hero{{text-align:center;margin-bottom:56px}}
.hero-tag{{font-size:10px;font-weight:700;letter-spacing:.22em;text-transform:uppercase;color:var(--red);margin-bottom:14px}}
.hero-title{{font-size:clamp(28px,4vw,44px);font-weight:800;line-height:1.1;letter-spacing:-.02em;margin-bottom:10px}}
.hero-title em{{color:var(--red);font-style:normal}}
.hero-sub{{font-size:12px;color:var(--muted);font-family:var(--mono)}}

.filial-grid{{display:grid;grid-template-columns:repeat(2,1fr);gap:14px;max-width:640px;width:100%}}
.filial-card{{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:26px 28px;cursor:pointer;transition:all .22s;position:relative;overflow:hidden;display:flex;flex-direction:column;gap:6px}}
.filial-card::after{{content:'';position:absolute;bottom:0;left:0;right:0;height:2px;background:var(--red);transform:scaleX(0);transition:transform .22s;transform-origin:left}}
.filial-card:hover{{border-color:var(--red);transform:translateY(-2px);box-shadow:0 12px 40px var(--red-glow)}}
.filial-card:hover::after{{transform:scaleX(1)}}
.fc-num{{font-size:10px;font-weight:700;letter-spacing:.14em;text-transform:uppercase;color:var(--muted);font-family:var(--mono)}}
.fc-name{{font-size:20px;font-weight:700;letter-spacing:-.01em}}
.fc-badge{{font-size:11px;color:var(--red);font-family:var(--mono);margin-top:2px;display:flex;align-items:center;gap:5px}}
.fc-badge::before{{content:'';width:5px;height:5px;background:var(--red);border-radius:50%;display:inline-block}}

/* ── SCREEN 2 DEP ── */
#screen-dep{{padding:48px}}
.back-btn{{display:inline-flex;align-items:center;gap:8px;background:none;border:1px solid var(--border);color:var(--muted);font-family:var(--font);font-size:11px;font-weight:600;letter-spacing:.08em;text-transform:uppercase;padding:7px 14px;border-radius:7px;cursor:pointer;transition:all .18s;margin-bottom:36px}}
.back-btn:hover{{border-color:var(--red);color:var(--white)}}
.screen-header{{margin-bottom:36px}}
.screen-header .stag{{font-size:10px;font-weight:700;letter-spacing:.18em;text-transform:uppercase;color:var(--red);margin-bottom:8px}}
.screen-header h2{{font-size:clamp(22px,3vw,32px);font-weight:700;letter-spacing:-.02em}}
.screen-header p{{font-size:12px;color:var(--muted);font-family:var(--mono);margin-top:4px}}

.dep-list{{display:grid;grid-template-columns:repeat(auto-fill,minmax(260px,1fr));gap:10px;max-width:1100px}}
.dep-item{{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:14px 18px;cursor:pointer;transition:all .18s;display:flex;align-items:center;justify-content:space-between;gap:10px}}
.dep-item:hover{{border-color:var(--red);background:rgba(227,30,36,.05)}}
.dep-item-name{{font-size:13px;font-weight:600}}
.dep-pills{{display:flex;gap:6px;flex-shrink:0}}
.dp{{font-size:11px;font-weight:700;padding:2px 8px;border-radius:16px}}
.dp-man{{background:rgba(227,30,36,.15);color:var(--red)}}
.dp-fin{{background:rgba(34,201,122,.15);color:var(--green)}}
.dp-add{{background:rgba(245,166,35,.15);color:var(--amber)}}

/* ── SCREEN 3 RESULTS ── */
#screen-results{{padding:0}}
.results-header{{background:linear-gradient(135deg,#130303 0%,#0a0a0a 60%);border-bottom:1px solid var(--border);padding:28px 48px 22px;position:relative;overflow:hidden}}
.results-header::after{{content:'';position:absolute;top:-80px;right:-80px;width:320px;height:320px;background:radial-gradient(circle,var(--red-glow) 0%,transparent 65%);pointer-events:none}}
.rh-meta{{font-size:11px;color:var(--muted);font-family:var(--mono);margin-top:10px;display:flex;flex-direction:column;gap:2px}}
.rh-title{{font-size:clamp(18px,2.5vw,24px);font-weight:700;letter-spacing:-.02em;margin-top:6px}}
.rh-title em{{color:var(--red);font-style:normal}}
.rh-back{{display:inline-flex;align-items:center;gap:7px;background:none;border:1px solid var(--border);color:var(--muted);font-family:var(--font);font-size:11px;font-weight:600;letter-spacing:.08em;text-transform:uppercase;padding:6px 13px;border-radius:7px;cursor:pointer;transition:all .18s}}
.rh-back:hover{{border-color:var(--red);color:var(--white)}}

.results-body{{padding:28px 48px 60px;max-width:1400px}}

/* ── STAT CARDS ── */
.cards-row{{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:28px}}
.stat-card{{background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:18px 22px;position:relative;overflow:hidden}}
.stat-card::after{{content:'';position:absolute;bottom:0;left:0;right:0;height:2px}}
.stat-card.man::after{{background:var(--red)}}
.stat-card.fin::after{{background:var(--green)}}
.stat-card.add::after{{background:var(--amber)}}
.sc-label{{font-size:10px;font-weight:700;letter-spacing:.12em;text-transform:uppercase;color:var(--muted);margin-bottom:8px}}
.sc-num{{font-size:36px;font-weight:800;letter-spacing:-.03em;line-height:1}}
.stat-card.man .sc-num{{color:var(--red)}}
.stat-card.fin .sc-num{{color:var(--green)}}
.stat-card.add .sc-num{{color:var(--amber)}}
.sc-sub{{font-size:11px;color:var(--muted);margin-top:4px}}

/* ── FILTER BAR ── */
.filter-bar{{display:flex;align-items:flex-start;gap:12px;margin-bottom:20px;flex-wrap:wrap}}

/* ── GENERIC DROPDOWN ── */
.filter-wrap{{position:relative;min-width:200px;max-width:340px;flex:1}}
.filter-label{{font-size:10px;font-weight:700;letter-spacing:.12em;text-transform:uppercase;color:var(--muted);white-space:nowrap;margin-bottom:6px}}
.filter-select-btn{{width:100%;background:var(--surface);border:1px solid var(--border);color:var(--white);font-family:var(--font);font-size:13px;font-weight:500;padding:9px 36px 9px 14px;border-radius:8px;cursor:pointer;text-align:left;transition:border-color .18s;display:flex;align-items:center;justify-content:space-between;gap:8px}}
.filter-select-btn:hover,.filter-select-btn.open{{border-color:var(--red)}}
.filter-select-btn .arrow{{font-size:10px;color:var(--muted);transition:transform .18s;flex-shrink:0}}
.filter-select-btn.open .arrow{{transform:rotate(180deg)}}
.filter-dropdown{{position:absolute;top:calc(100% + 6px);left:0;right:0;background:var(--surface2);border:1px solid var(--border);border-radius:10px;z-index:200;overflow:hidden;box-shadow:0 12px 40px rgba(0,0,0,.6);display:none;flex-direction:column}}
.filter-dropdown.open{{display:flex}}
.filter-search{{padding:10px 12px;border-bottom:1px solid var(--border)}}
.filter-search input{{width:100%;background:var(--surface);border:1px solid var(--border);color:var(--white);font-family:var(--font);font-size:12px;padding:6px 10px;border-radius:6px;outline:none}}
.filter-search input:focus{{border-color:var(--red)}}
.filter-options{{max-height:240px;overflow-y:auto}}
.filter-options::-webkit-scrollbar{{width:4px}}
.filter-options::-webkit-scrollbar-track{{background:transparent}}
.filter-options::-webkit-scrollbar-thumb{{background:var(--border);border-radius:4px}}
.filter-opt{{display:flex;align-items:center;justify-content:space-between;padding:9px 14px;cursor:pointer;transition:background .15s;font-size:13px}}
.filter-opt:hover{{background:rgba(227,30,36,.08)}}
.filter-opt.sel{{background:rgba(227,30,36,.12);color:var(--red)}}
.filter-opt-name{{font-weight:500}}
.filter-opt-count{{font-size:11px;font-family:var(--mono);color:var(--muted);background:var(--surface);padding:1px 7px;border-radius:10px}}
.filter-opt.sel .filter-opt-count{{background:rgba(227,30,36,.2);color:var(--red)}}
.filter-clear{{padding:8px 14px;border-top:1px solid var(--border);font-size:11px;color:var(--muted);cursor:pointer;transition:color .15s;text-align:center;font-weight:600;letter-spacing:.06em;text-transform:uppercase}}
.filter-clear:hover{{color:var(--red)}}
.filter-disabled .filter-select-btn{{opacity:.4;cursor:not-allowed;pointer-events:none}}

/* ── TABS ── */
.tabs{{display:flex;gap:2px;border-bottom:1px solid var(--border);margin-bottom:0}}
.tab-btn{{background:none;border:none;border-bottom:2px solid transparent;padding:10px 20px;font-family:var(--font);font-size:13px;font-weight:600;color:var(--muted);cursor:pointer;transition:all .18s;margin-bottom:-1px;letter-spacing:.01em}}
.tab-btn:hover{{color:var(--white)}}
.tab-btn.ativo{{color:var(--white);border-bottom-color:var(--red)}}
.tab-btn.t-fin.ativo{{color:var(--green);border-bottom-color:var(--green)}}
.tab-btn.t-add.ativo{{color:var(--amber);border-bottom-color:var(--amber)}}
.badge{{display:inline-block;font-size:10px;font-weight:700;padding:1px 7px;border-radius:14px;margin-left:5px;color:#fff;vertical-align:middle}}
.tab-content{{display:none}}
.tab-content.ativo{{display:block}}

/* ── TABLE ── */
.tbl-wrap{{overflow-x:auto;margin-top:14px}}
table{{width:100%;border-collapse:collapse;font-size:13px;background:var(--surface);border:1px solid var(--border);border-radius:10px;overflow:hidden}}
thead tr{{background:var(--surface2)}}
thead th{{padding:10px 14px;text-align:left;font-size:10px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:var(--muted);border-bottom:1px solid var(--border);white-space:nowrap}}
tbody tr{{border-bottom:1px solid var(--border);transition:background .12s}}
tbody tr:last-child{{border-bottom:none}}
tbody tr:hover{{background:rgba(255,255,255,.02)}}
tbody td{{padding:10px 14px;color:var(--white);vertical-align:top}}
tr.dep-row td{{background:rgba(227,30,36,.06);color:var(--red);font-weight:700;font-size:10px;letter-spacing:.1em;text-transform:uppercase;padding:7px 14px;border-top:1px solid var(--border)}}
.cn-badge{{font-size:9px;background:rgba(245,166,35,.12);color:var(--amber);border:1px solid rgba(245,166,35,.25);padding:1px 6px;border-radius:10px;font-weight:700;margin-left:7px;vertical-align:middle}}
.client-tag{{display:inline-flex;align-items:center;gap:5px;font-size:11px;font-family:var(--mono);background:rgba(227,30,36,.1);border:1px solid rgba(227,30,36,.2);color:var(--red);padding:2px 8px;border-radius:6px;cursor:default;white-space:nowrap;position:relative}}
.client-tag .razao{{font-weight:600;max-width:140px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}}
.client-tag .cod{{opacity:.7}}
.comt-date{{font-size:11px;color:var(--amber);font-family:var(--mono);white-space:nowrap}}
.comt-btn{{background:rgba(227,30,36,.1);border:1px solid rgba(227,30,36,.2);color:var(--red);font-size:11px;font-family:var(--font);padding:2px 9px;border-radius:14px;cursor:pointer;transition:background .15s;white-space:nowrap;margin-top:3px;display:inline-block;font-weight:600}}
.comt-btn:hover{{background:rgba(227,30,36,.22)}}
.comt-full{{display:none;margin-top:6px;font-size:11px;color:var(--muted);background:rgba(0,0,0,.35);border-left:2px solid var(--red);padding:8px 12px;border-radius:0 6px 6px 0;line-height:1.7;white-space:pre-wrap;font-family:var(--mono)}}
.comt-full.open{{display:block}}
.no-data{{color:var(--muted);font-size:13px;padding:24px 16px;text-align:center;font-style:italic}}
.resp-name{{font-size:12px;font-weight:600}}

/* ── PLACARES DE BAIXAS ── */
.placares-section{{margin:32px 0 28px}}
.placares-title{{font-size:10px;font-weight:700;letter-spacing:.14em;text-transform:uppercase;color:var(--muted);margin-bottom:16px;display:flex;align-items:center;gap:10px}}
.placares-title::after{{content:'';flex:1;height:1px;background:var(--border)}}
.placares-export-row{{display:flex;gap:10px;margin-bottom:18px;flex-wrap:wrap}}
.placar-export-btn{{display:inline-flex;align-items:center;gap:7px;background:var(--surface);border:1px solid var(--border);color:var(--muted);font-family:var(--font);font-size:11px;font-weight:600;letter-spacing:.08em;text-transform:uppercase;padding:7px 14px;border-radius:7px;cursor:pointer;transition:all .18s}}
.placar-export-btn:hover{{border-color:var(--green);color:var(--green)}}
.placares-grid{{display:grid;grid-template-columns:repeat(3,1fr);gap:14px}}
.placar-card{{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:20px 22px;position:relative;overflow:hidden}}
.placar-card::after{{content:'';position:absolute;bottom:0;left:0;right:0;height:3px}}
.placar-card.pc-dia::after{{background:var(--red)}}
.placar-card.pc-sem::after{{background:var(--amber)}}
.placar-card.pc-mes::after{{background:var(--green)}}
.pc-label{{font-size:10px;font-weight:700;letter-spacing:.12em;text-transform:uppercase;color:var(--muted);margin-bottom:8px}}
.pc-num{{font-size:42px;font-weight:800;letter-spacing:-.03em;line-height:1}}
.placar-card.pc-dia .pc-num{{color:var(--red)}}
.placar-card.pc-sem .pc-num{{color:var(--amber)}}
.placar-card.pc-mes .pc-num{{color:var(--green)}}
.pc-sub{{font-size:11px;color:var(--muted);margin-top:4px}}
.pc-deps{{margin-top:14px;border-top:1px solid var(--border);padding-top:12px;display:flex;flex-direction:column;gap:6px;max-height:220px;overflow-y:auto}}
.pc-deps::-webkit-scrollbar{{width:3px}}
.pc-deps::-webkit-scrollbar-thumb{{background:var(--border);border-radius:3px}}
.pc-dep-row{{display:flex;align-items:center;justify-content:space-between;gap:8px}}
.pc-dep-name{{font-size:11px;color:var(--off-white);flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}}
.pc-dep-unit{{font-size:10px;color:var(--muted);font-family:var(--mono);flex-shrink:0}}
.pc-dep-n{{font-size:12px;font-weight:700;flex-shrink:0;min-width:24px;text-align:right}}
.placar-card.pc-dia .pc-dep-n{{color:var(--red)}}
.placar-card.pc-sem .pc-dep-n{{color:var(--amber)}}
.placar-card.pc-mes .pc-dep-n{{color:var(--green)}}
.pc-bar-wrap{{margin-top:3px}}
.pc-bar{{height:3px;border-radius:2px;transition:width .3s}}
.placar-card.pc-dia .pc-bar{{background:var(--red)}}
.placar-card.pc-sem .pc-bar{{background:var(--amber)}}
.placar-card.pc-mes .pc-bar{{background:var(--green)}}
.pc-empty{{font-size:12px;color:var(--muted);font-style:italic;padding:8px 0}}

/* ── BAIXAS POR FUNCIONÁRIO TABLE ── */
.baixas-func-section{{margin:32px 0 28px}}
.baixas-func-title{{font-size:10px;font-weight:700;letter-spacing:.14em;text-transform:uppercase;color:var(--muted);margin-bottom:16px;display:flex;align-items:center;gap:10px}}
.baixas-func-title::after{{content:'';flex:1;height:1px;background:var(--border)}}
.baixas-func-periodo-tabs{{display:flex;gap:8px;margin-bottom:16px;flex-wrap:wrap}}
.bf-tab{{background:var(--surface);border:1px solid var(--border);color:var(--muted);font-family:var(--font);font-size:11px;font-weight:600;letter-spacing:.06em;text-transform:uppercase;padding:6px 14px;border-radius:20px;cursor:pointer;transition:all .18s}}
.bf-tab:hover{{border-color:var(--green);color:var(--green)}}
.bf-tab.ativo{{background:rgba(34,201,122,.12);border-color:var(--green);color:var(--green)}}
.bf-table-wrap{{overflow-x:auto}}
.bf-table{{width:100%;border-collapse:collapse;font-size:13px;background:var(--surface);border:1px solid var(--border);border-radius:10px;overflow:hidden}}
.bf-table thead tr{{background:var(--surface2)}}
.bf-table thead th{{padding:10px 14px;text-align:left;font-size:10px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:var(--muted);border-bottom:1px solid var(--border);white-space:nowrap}}
.bf-table thead th:last-child{{text-align:right}}
.bf-table tbody tr{{border-bottom:1px solid var(--border);transition:background .12s}}
.bf-table tbody tr:last-child{{border-bottom:none}}
.bf-table tbody tr:hover{{background:rgba(255,255,255,.02)}}
.bf-table tbody td{{padding:9px 14px;vertical-align:middle}}
.bf-table tbody td:last-child{{text-align:right}}
.bf-dep-row td{{background:rgba(34,201,122,.05);color:var(--green);font-weight:700;font-size:10px;letter-spacing:.1em;text-transform:uppercase;padding:7px 14px;border-top:1px solid var(--border)}}
.bf-resp-name{{font-size:12px;font-weight:600;color:var(--white)}}
.bf-coord-badge{{font-size:10px;color:var(--blue);font-family:var(--mono);margin-left:8px}}
.bf-num{{font-size:14px;font-weight:700;color:var(--green);font-family:var(--mono)}}
.bf-bar-cell{{width:120px}}
.bf-bar-bg{{background:rgba(34,201,122,.1);border-radius:3px;height:6px;overflow:hidden}}
.bf-bar-fill{{height:100%;background:var(--green);border-radius:3px;transition:width .3s}}
.bf-empty{{color:var(--muted);font-size:13px;padding:24px 16px;text-align:center;font-style:italic}}

/* capture overlay */
.capture-overlay{{position:fixed;inset:0;background:rgba(0,0,0,.85);z-index:9999;display:flex;align-items:center;justify-content:center;flex-direction:column;gap:16px}}
.capture-spinner{{width:36px;height:36px;border:3px solid var(--border);border-top-color:var(--green);border-radius:50%;animation:spin .8s linear infinite}}
@keyframes spin{{to{{transform:rotate(360deg)}}}}

@media(max-width:700px){{
  .filial-grid{{grid-template-columns:1fr}}
  #screen-dep,.results-body,.results-header{{padding:20px}}
  .cards-row{{grid-template-columns:1fr}}
  .logo-bar{{padding:16px 20px 0}}
  .filter-bar{{flex-direction:column}}
  .filter-wrap{{max-width:100%}}
  .placares-grid{{grid-template-columns:1fr}}
}}
</style>
</head>
<body>

<!-- ══ SCREEN 1: FILIAL ══ -->
<div class="screen active" id="screen-filial">
  <div class="logo-bar">
    <img class="logo-img" src="data:image/png;base64,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" onerror="this.style.display='none'" alt="MG">
    <div class="logo-text"><strong>MG Contécnica</strong>Relatório de Pendências</div>
  </div>
  <div style="flex:1;display:flex;flex-direction:column;align-items:center;justify-content:center;padding:40px 48px">
    <div class="hero">
      <div class="hero-tag">Painel de Controle</div>
      <h1 class="hero-title">Pendências <em>Ativas</em></h1>
      <p class="hero-sub">Gerado em {gerado_em} &nbsp;·&nbsp; {nome_base}</p>
    </div>
    <div class="filial-grid" id="filial-grid"></div>
  </div>
</div>

<!-- ══ SCREEN 2: DEPARTAMENTO ══ -->
<div class="screen" id="screen-dep">
  <div class="logo-bar">
    <img class="logo-img" src="data:image/png;base64,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" onerror="this.style.display='none'" alt="MG">
    <div class="logo-text"><strong>MG Contécnica</strong>Selecione o Departamento</div>
  </div>
  <div style="flex:1;padding:40px 48px">
    <button class="back-btn" onclick="voltarFilial()">← Voltar</button>
    <div class="screen-header">
      <div class="stag">Filial selecionada</div>
      <h2 id="dep-screen-title">—</h2>
      <p id="dep-screen-sub"></p>
    </div>
    <div class="dep-list" id="dep-list"></div>
  </div>
</div>

<!-- ══ SCREEN 3: RESULTADOS ══ -->
<div class="screen" id="screen-results">
  <div class="results-header">
    <button class="rh-back" onclick="voltarDep()">← Departamentos</button>
    <div class="rh-meta">
      <span id="rh-meta-base"></span>
      <span id="rh-meta-atual"></span>
    </div>
    <div class="rh-title" id="rh-title">—</div>
  </div>
  <div class="results-body">

    <div class="cards-row">
      <div class="stat-card man"><div class="sc-label">Em Andamento</div><div class="sc-num" id="sc-man">0</div><div class="sc-sub">tarefas abertas</div></div>
      <div class="stat-card fin"><div class="sc-label">Baixadas</div><div class="sc-num" id="sc-fin">0</div><div class="sc-sub">finalizadas desde a base</div></div>
      <div class="stat-card add"><div class="sc-label">Reabertas</div><div class="sc-num" id="sc-add">0</div><div class="sc-sub">reabertas desde a base</div></div>
    </div>

    <!-- ══ FILTER BAR ══ -->
    <div class="filter-bar">
      <div>
        <div class="filter-label">Coordenador</div>
        <div class="filter-wrap" id="coord-wrap">
          <button class="filter-select-btn" id="coord-btn" onclick="toggleFilter('coord')">
            <span id="coord-btn-label">Todos os coordenadores</span>
            <span class="arrow">▼</span>
          </button>
          <div class="filter-dropdown" id="coord-dropdown">
            <div class="filter-search"><input type="text" placeholder="Buscar coordenador..." id="coord-search-input" oninput="filtrarOpcoes('coord')"></div>
            <div class="filter-options" id="coord-options"></div>
            <div class="filter-clear" onclick="limparFiltro('coord')">Limpar filtro</div>
          </div>
        </div>
      </div>
      <div>
        <div class="filter-label">Responsável</div>
        <div class="filter-wrap" id="resp-wrap">
          <button class="filter-select-btn" id="resp-btn" onclick="toggleFilter('resp')">
            <span id="resp-btn-label">Todos os responsáveis</span>
            <span class="arrow">▼</span>
          </button>
          <div class="filter-dropdown" id="resp-dropdown">
            <div class="filter-search"><input type="text" placeholder="Buscar responsável..." id="resp-search-input" oninput="filtrarOpcoes('resp')"></div>
            <div class="filter-options" id="resp-options"></div>
            <div class="filter-clear" onclick="limparFiltro('resp')">Limpar filtro</div>
          </div>
        </div>
      </div>
      <div>
        <div class="filter-label">Grupo <span style="color:var(--muted);font-weight:400;letter-spacing:0;text-transform:none;font-size:10px">(opcional)</span></div>
        <div class="filter-wrap" id="grupo-wrap">
          <button class="filter-select-btn" id="grupo-btn" onclick="toggleFilter('grupo')">
            <span id="grupo-btn-label">Todos os grupos</span>
            <span class="arrow">▼</span>
          </button>
          <div class="filter-dropdown" id="grupo-dropdown">
            <div class="filter-search"><input type="text" placeholder="Buscar grupo..." id="grupo-search-input" oninput="filtrarOpcoes('grupo')"></div>
            <div class="filter-options" id="grupo-options"></div>
            <div class="filter-clear" onclick="limparFiltro('grupo')">Limpar filtro</div>
          </div>
        </div>
      </div>
    </div>

    <!-- Placares de Baixas -->
    <div class="placares-section" id="placares-section" style="display:none">
      <div class="placares-title">Placar de Baixas</div>
      <div class="placares-export-row">
        <button class="placar-export-btn" onclick="exportarPlacares('geral')">⬇ Baixar imagem geral</button>
        <button class="placar-export-btn" onclick="exportarPlacares('dep')">⬇ Baixar por departamento</button>
      </div>
      <div class="placares-grid" id="placares-grid">
        <div class="placar-card pc-dia" id="pc-dia">
          <div class="pc-label">Baixas no Dia</div>
          <div class="pc-num" id="pc-dia-num">0</div>
          <div class="pc-sub">tarefas concluídas hoje</div>
          <div class="pc-deps" id="pc-dia-deps"></div>
        </div>
        <div class="placar-card pc-sem" id="pc-sem">
          <div class="pc-label">Baixas na Semana</div>
          <div class="pc-num" id="pc-sem-num">0</div>
          <div class="pc-sub">desde segunda-feira</div>
          <div class="pc-deps" id="pc-sem-deps"></div>
        </div>
        <div class="placar-card pc-mes" id="pc-mes">
          <div class="pc-label">Baixas no Mês</div>
          <div class="pc-num" id="pc-mes-num">0</div>
          <div class="pc-sub">acumulado do mês</div>
          <div class="pc-deps" id="pc-mes-deps"></div>
        </div>
      </div>
    </div>

    <!-- Baixas por Funcionário -->
    <div class="baixas-func-section" id="baixas-func-section" style="display:none">
      <div class="baixas-func-title">Baixas por Funcionário</div>
      <div class="baixas-func-periodo-tabs">
        <button class="bf-tab ativo" id="bf-tab-dia" onclick="mudarBfTab('dia',this)">Dia</button>
        <button class="bf-tab" id="bf-tab-sem" onclick="mudarBfTab('sem',this)">Semana</button>
        <button class="bf-tab" id="bf-tab-mes" onclick="mudarBfTab('mes',this)">Mês</button>
      </div>
      <div class="bf-table-wrap">
        <table class="bf-table">
          <thead>
            <tr>
              <th>Funcionário</th>
              <th>Coordenador</th>
              <th>Departamento</th>
              <th style="width:120px">Progresso</th>
              <th>Baixas</th>
            </tr>
          </thead>
          <tbody id="bf-tbody"></tbody>
        </table>
      </div>
    </div>

    <div class="tabs">
      <button class="tab-btn ativo"  id="tbtn-man" onclick="mudarTab('man',this)">Em Andamento <span class="badge" id="bdg-man" style="background:var(--red)">0</span></button>
      <button class="tab-btn t-fin"  id="tbtn-fin" onclick="mudarTab('fin',this)">Baixadas <span class="badge" id="bdg-fin" style="background:var(--green)">0</span></button>
      <button class="tab-btn t-add"  id="tbtn-add" onclick="mudarTab('add',this)">Reabertas <span class="badge" id="bdg-add" style="background:var(--amber)">0</span></button>
    </div>

    <div id="tab-man" class="tab-content ativo"><div class="tbl-wrap"><table><thead><tr><th>Cliente</th><th>Título da Tarefa</th><th>Vencimento</th><th>Responsável</th><th>Previsão</th><th>Comentário</th></tr></thead><tbody id="tbody-man"></tbody></table></div></div>
    <div id="tab-fin" class="tab-content"><div class="tbl-wrap"><table><thead><tr><th>Cliente</th><th>Título da Tarefa</th><th>Vencimento</th><th>Responsável</th><th>Previsão</th><th>Comentário</th></tr></thead><tbody id="tbody-fin"></tbody></table></div></div>
    <div id="tab-add" class="tab-content"><div class="tbl-wrap"><table><thead><tr><th>Cliente</th><th>Título da Tarefa</th><th>Vencimento</th><th>Responsável</th><th>Previsão</th><th>Comentário</th></tr></thead><tbody id="tbody-add"></tbody></table></div></div>
  </div>
</div>

<script>
const DATA = {{ man:{js_man}, fin:{js_fin}, add:{js_add} }};
const MODO_COMP = {'true' if modo_comp else 'false'};
const ARQ_BASE  = {json.dumps(nome_base)};
const ARQ_ATUAL = {json.dumps(nome_atual)};
const GERADO    = {json.dumps(gerado_em)};
const PLACARES  = {js_placares};
const ALL_COORDS = {js_all_coords};

let filialAtual=null, depAtual=null;
let filtroCoord=null, filtroResp=null, filtroGrupo=null;
let abaAtual='man', bfPeriodoAtual='dia';

// ── Labels ────────────────────────────────────────────────────────────────────
const LABELS = {{'SP':'São Paulo','Santos':'Santos','RJ':'Rio de Janeiro','GOIAS':'Goiás'}};
function labelFilial(f){{ return LABELS[f]||f; }}

// ── Data helpers ──────────────────────────────────────────────────────────────
function getFiliais(){{
  return [...new Set([...DATA.man,...DATA.fin,...DATA.add].map(r=>r.unidade))].sort();
}}
function getDeps(filial){{
  return [...new Set([...DATA.man,...DATA.fin,...DATA.add].filter(r=>r.unidade===filial).map(r=>r.dep))].sort();
}}
function countIn(arr,filial,dep){{
  return arr.filter(r=>r.unidade===filial&&r.dep===dep).length;
}}

function rowPassaCoord(r){{
  if(!filtroCoord) return true;
  if(!r.coord) return false;
  return r.coord.split('|').map(s=>s.trim()).includes(filtroCoord);
}}

function filtrar(arr){{
  return arr.filter(r =>
    r.unidade === filialAtual &&
    r.dep === depAtual &&
    rowPassaCoord(r) &&
    (!filtroResp  || r.resp  === filtroResp) &&
    (!filtroGrupo || r.grupo === filtroGrupo)
  );
}}

// ── Screens ───────────────────────────────────────────────────────────────────
function showScreen(id){{document.querySelectorAll('.screen').forEach(s=>s.classList.remove('active'));document.getElementById(id).classList.add('active');window.scrollTo(0,0);}}
function voltarFilial(){{showScreen('screen-filial');filialAtual=null;depAtual=null;}}
function voltarDep(){{showScreen('screen-dep');depAtual=null;filtroCoord=null;filtroResp=null;filtroGrupo=null;}}

// ── Screen 1 ──────────────────────────────────────────────────────────────────
function buildFiliais(){{
  const grid=document.getElementById('filial-grid');
  getFiliais().forEach((f,i)=>{{
    const total=DATA.man.filter(r=>r.unidade===f).length;
    const c=document.createElement('div');
    c.className='filial-card';
    c.innerHTML=`<div class="fc-num">FILIAL ${{String(i+1).padStart(2,'0')}}</div><div class="fc-name">${{labelFilial(f)}}</div><div class="fc-badge">${{total}} pendência${{total!==1?'s':''}}</div>`;
    c.onclick=()=>selecionarFilial(f);
    grid.appendChild(c);
  }});
}}

// ── Screen 2 ──────────────────────────────────────────────────────────────────
function selecionarFilial(filial){{
  filialAtual=filial; filtroCoord=null; filtroResp=null; filtroGrupo=null;
  document.getElementById('dep-screen-title').textContent=labelFilial(filial);
  const deps=getDeps(filial);
  document.getElementById('dep-screen-sub').textContent=`${{deps.length}} departamentos`;
  const list=document.getElementById('dep-list');
  list.innerHTML='';
  deps.forEach(dep=>{{
    const m=countIn(DATA.man,filial,dep);
    const f=countIn(DATA.fin,filial,dep);
    const a=countIn(DATA.add,filial,dep);
    const el=document.createElement('div');
    el.className='dep-item';
    el.innerHTML=`<span class="dep-item-name">${{dep}}</span><div class="dep-pills"><span class="dp dp-man">${{m}}</span>${{MODO_COMP?`<span class="dp dp-fin">${{f}}</span><span class="dp dp-add">${{a}}</span>`:''}}</div>`;    el.onclick=()=>selecionarDep(dep);
    list.appendChild(el);
  }});
  showScreen('screen-dep');
}}

// ── Screen 3 ──────────────────────────────────────────────────────────────────
function selecionarDep(dep){{
  depAtual=dep; filtroCoord=null; filtroResp=null; filtroGrupo=null;
  abaAtual='man'; bfPeriodoAtual='dia';
  document.getElementById('rh-title').innerHTML=`<em>${{labelFilial(filialAtual)}}</em> · ${{dep}}`;
  document.getElementById('rh-meta-base').textContent=`Base: ${{ARQ_BASE}}`;
  document.getElementById('rh-meta-atual').textContent=MODO_COMP?`Atual: ${{ARQ_ATUAL}} · Gerado em: ${{GERADO}}`:`Gerado em: ${{GERADO}}`;
  document.querySelectorAll('.tab-btn').forEach(b=>b.classList.remove('ativo'));
  document.getElementById('tbtn-man').classList.add('ativo');
  document.querySelectorAll('.tab-content').forEach(t=>t.classList.remove('ativo'));
  document.getElementById('tab-man').classList.add('ativo');
  document.querySelectorAll('.bf-tab').forEach(b=>b.classList.remove('ativo'));
  document.getElementById('bf-tab-dia').classList.add('ativo');

  buildAllDropdowns();
  renderAll();
  renderPlacares();
  renderBaixasFunc();
  showScreen('screen-results');
}}

// ══════════════════════════════════════════════════════════════════════════════
// ── DROPDOWN SYSTEM ───────────────────────────────────────────────────────────
// ══════════════════════════════════════════════════════════════════════════════

function getOpcoesFiltro(tipo){{
  const all = [...DATA.man,...DATA.fin,...DATA.add].filter(r => r.unidade===filialAtual && r.dep===depAtual);

  if(tipo==='coord'){{
    const cnt={{}};
    all.forEach(r=>{{
      const coords = r.coord ? r.coord.split('|').map(s=>s.trim()).filter(Boolean) : [];
      if(!coords.length){{ cnt['']=(cnt['']||0)+1; return; }}
      coords.forEach(c=>{{ cnt[c]=(cnt[c]||0)+1; }});
    }});
    return Object.entries(cnt)
      .filter(([k])=>k)
      .sort((a,b)=>b[1]-a[1])
      .map(([k,v])=>{{return{{valor:k,label:k,count:v}}}});
  }}

  if(tipo==='resp'){{
    const filtrados = filtroCoord
      ? all.filter(r=>r.coord && r.coord.split('|').map(s=>s.trim()).includes(filtroCoord))
      : all;
    const cnt={{}};
    filtrados.forEach(r=>{{ if(r.resp) cnt[r.resp]=(cnt[r.resp]||0)+1; }});
    return Object.entries(cnt)
      .sort((a,b)=>b[1]-a[1])
      .map(([k,v])=>{{return{{valor:k,label:k,count:v}}}});
  }}

  if(tipo==='grupo'){{
    const filtrados = all.filter(r=>
      rowPassaCoord(r) &&
      (!filtroResp || r.resp===filtroResp)
    );
    const cnt={{}};
    filtrados.forEach(r=>{{ if(r.grupo) cnt[r.grupo]=(cnt[r.grupo]||0)+1; }});
    return Object.entries(cnt)
      .sort((a,b)=>b[1]-a[1])
      .map(([k,v])=>{{return{{valor:k,label:k,count:v}}}});
  }}
  return [];
}}

function buildAllDropdowns(){{
  renderDropdownOpcoes('coord','');
  renderDropdownOpcoes('resp','');
  renderDropdownOpcoes('grupo','');
  atualizarBotaoFiltro('coord', filtroCoord, 'Todos os coordenadores');
  atualizarBotaoFiltro('resp',  filtroResp,  'Todos os responsáveis');
  atualizarBotaoFiltro('grupo', filtroGrupo, 'Todos os grupos');
}}

function renderDropdownOpcoes(tipo, busca){{
  const el=document.getElementById(`${{tipo}}-options`);
  const opcoes=getOpcoesFiltro(tipo);
  const valorAtual = tipo==='coord'?filtroCoord : tipo==='resp'?filtroResp : filtroGrupo;
  const filtradas = busca ? opcoes.filter(o=>o.label.toLowerCase().includes(busca.toLowerCase())) : opcoes;
  if(!filtradas.length){{
    el.innerHTML='<div class="filter-opt" style="color:var(--muted);justify-content:center;cursor:default">Nenhuma opção</div>';
    return;
  }}
  el.innerHTML='';
  filtradas.forEach(o=>{{
    const opt=document.createElement('div');
    opt.className='filter-opt'+(valorAtual===o.valor?' sel':'');
    opt.innerHTML=`<span class="filter-opt-name">${{o.label}}</span><span class="filter-opt-count">${{o.count}}</span>`;
    opt.onclick=()=>selecionarFiltro(tipo,o.valor,o.label);
    el.appendChild(opt);
  }});
}}

function filtrarOpcoes(tipo){{
  const busca=document.getElementById(`${{tipo}}-search-input`).value;
  renderDropdownOpcoes(tipo,busca);
}}

function toggleFilter(tipo){{
  const btn=document.getElementById(`${{tipo}}-btn`);
  const dd=document.getElementById(`${{tipo}}-dropdown`);
  const open=dd.classList.toggle('open');
  btn.classList.toggle('open',open);
  if(open){{
    document.getElementById(`${{tipo}}-search-input`).value='';
    renderDropdownOpcoes(tipo,'');
    setTimeout(()=>document.getElementById(`${{tipo}}-search-input`).focus(),50);
  }}
}}

function fecharTodosDropdowns(){{
  ['coord','resp','grupo'].forEach(t=>{{
    document.getElementById(`${{t}}-dropdown`).classList.remove('open');
    document.getElementById(`${{t}}-btn`).classList.remove('open');
  }});
}}

function atualizarBotaoFiltro(tipo, valor, placeholder){{
  document.getElementById(`${{tipo}}-btn-label`).textContent = valor || placeholder;
}}

function selecionarFiltro(tipo, valor, label){{
  if(tipo==='coord'){{
    filtroCoord=valor;
    filtroResp=null;
    atualizarBotaoFiltro('coord', valor, 'Todos os coordenadores');
    atualizarBotaoFiltro('resp', null, 'Todos os responsáveis');
    renderDropdownOpcoes('resp','');
    renderDropdownOpcoes('grupo','');
  }} else if(tipo==='resp'){{
    filtroResp=valor;
    atualizarBotaoFiltro('resp', valor, 'Todos os responsáveis');
    renderDropdownOpcoes('grupo','');
  }} else {{
    filtroGrupo=valor;
    atualizarBotaoFiltro('grupo', valor, 'Todos os grupos');
  }}
  fecharTodosDropdowns();
  renderAll(); renderPlacares(); renderBaixasFunc();
}}

function limparFiltro(tipo){{
  if(tipo==='coord'){{
    filtroCoord=null; filtroResp=null; filtroGrupo=null;
    atualizarBotaoFiltro('coord', null, 'Todos os coordenadores');
    atualizarBotaoFiltro('resp',  null, 'Todos os responsáveis');
    atualizarBotaoFiltro('grupo', null, 'Todos os grupos');
    renderDropdownOpcoes('resp','');
    renderDropdownOpcoes('grupo','');
  }} else if(tipo==='resp'){{
    filtroResp=null; filtroGrupo=null;
    atualizarBotaoFiltro('resp',  null, 'Todos os responsáveis');
    atualizarBotaoFiltro('grupo', null, 'Todos os grupos');
    renderDropdownOpcoes('grupo','');
  }} else {{
    filtroGrupo=null;
    atualizarBotaoFiltro('grupo', null, 'Todos os grupos');
  }}
  fecharTodosDropdowns();
  renderAll(); renderPlacares(); renderBaixasFunc();
}}

document.addEventListener('click',e=>{{
  const wraps=['coord-wrap','resp-wrap','grupo-wrap'];
  const inside=wraps.some(id=>{{const el=document.getElementById(id);return el&&el.contains(e.target);}});
  if(!inside) fecharTodosDropdowns();
}});

// ── Render Tables ─────────────────────────────────────────────────────────────
function renderAll(){{
  const m=filtrar(DATA.man), f=filtrar(DATA.fin), a=filtrar(DATA.add);
  ['man','fin','add'].forEach((k,i)=>{{
    const n=[m,f,a][i].length;
    document.getElementById('sc-'+k).textContent=n;
    document.getElementById('bdg-'+k).textContent=n;
  }});
  renderTabela('man',m); renderTabela('fin',f); renderTabela('add',a);
}}

function renderTabela(aba,rows){{
  const tbody=document.getElementById('tbody-'+aba);
  if(!rows.length){{tbody.innerHTML='<tr><td colspan="6" class="no-data">Nenhuma tarefa encontrada.</td></tr>';return;}}
  const grupos={{}};
  rows.forEach(r=>{{
    const key=r.dep+(r.novo?' (Cliente Novo)':'');
    if(!grupos[key])grupos[key]={{dep:r.dep,novo:r.novo,rows:[]}};
    grupos[key].rows.push(r);
  }});
  let html='';
  Object.keys(grupos).sort().forEach(gkey=>{{
    const g=grupos[gkey];
    const nb=g.novo?'<span class="cn-badge">CLIENTE NOVO</span>':'';
    html+=`<tr class="dep-row"><td colspan="6">${{g.dep}}${{nb}}</td></tr>`;
    g.rows.forEach(r=>{{
      const razaoShort=(r.razao||r.grupo||'—').substring(0,22)+(((r.razao||'').length>22)?'…':'');
      const clientTag=`<span class="client-tag" title="${{r.razao||''}}"><span class="razao">${{razaoShort}}</span><span class="cod">#${{r.cod}}</span></span>`;
      let comtHtml='';
      if(r.dataComt)comtHtml+=`<div class="comt-date">📅 ${{r.dataComt}}</div>`;
      if(r.comt&&r.comt.trim()){{
        comtHtml+=`<button class="comt-btn" onclick="toggleComt(this)">Ver comentário</button>`;
        comtHtml+=`<div class="comt-full">${{r.comt.replace(/</g,'&lt;').replace(/>/g,'&gt;')}}</div>`;
      }}else if(!r.dataComt)comtHtml='<span style="color:var(--muted)">—</span>';
      html+=`<tr><td>${{clientTag}}</td><td>${{r.tit}}</td><td style="white-space:nowrap">${{r.venc}}</td><td class="resp-name">${{r.resp||'—'}}</td><td style="white-space:nowrap">${{r.prev}}</td><td>${{comtHtml}}</td></tr>`;
    }});
  }});
  tbody.innerHTML=html;
}}

function mudarTab(aba,btn){{
  document.querySelectorAll('.tab-content').forEach(t=>t.classList.remove('ativo'));
  document.querySelectorAll('.tab-btn').forEach(b=>b.classList.remove('ativo'));
  document.getElementById('tab-'+aba).classList.add('ativo');
  btn.classList.add('ativo'); abaAtual=aba;
}}
function toggleComt(btn){{
  const open=btn.nextElementSibling.classList.toggle('open');
  btn.textContent=open?'Ocultar':'Ver comentário';
}}

// ── Placares de Baixas ────────────────────────────────────────────────────────
function renderPlacares(){{
  const sec=document.getElementById('placares-section');
  if(!PLACARES||Object.keys(PLACARES).length===0){{sec.style.display='none';return;}}

  // Verifica se há qualquer baixa em qualquer período para este dep
  const temDados = ['dia','sem','mes'].some(p=>{{
    const pd=PLACARES[p]||{{por_dep:[]}};
    return pd.por_dep.some(r=>r.unidade===filialAtual&&r.dep===depAtual);
  }});
  sec.style.display = temDados ? 'block' : 'none';

 // DEPOIS:
  ['dia','sem','mes'].forEach(p=>{{
    const pd=PLACARES[p]||{{total:0,por_dep:[]}};
    const filtrados_resp=pd.por_resp.filter(r=>
     r.unidade===filialAtual &&
      r.dep===depAtual &&
      (!filtroCoord || (r.coord||'').split('|').map(s=>s.trim()).includes(filtroCoord)) &&
    (!filtroResp  || r.resp===filtroResp)
    );
    const filtrados = filtrados_resp;
    const total_dep=filtrados.reduce((a,b)=>a+b.n,0);
    document.getElementById(`pc-${{p}}-num`).textContent=total_dep;
    const container=document.getElementById(`pc-${{p}}-deps`);
    if(!filtrados.length){{container.innerHTML='<div class="pc-empty">Nenhuma baixa neste período</div>';return;}}
    const max_n=Math.max(...filtrados.map(r=>r.n),1);
    container.innerHTML=filtrados.map(r=>{{
      const pct=Math.round(r.n/max_n*100);
      return `<div class="pc-dep-row"><span class="pc-dep-name" title="${{r.resp}}">${{r.resp}}</span><span class="pc-dep-n">${{r.n}}</span></div><div class="pc-bar-wrap"><div class="pc-bar" style="width:${{pct}}%"></div></div>`;
    }}).join('');
  }});
}}

// ── Baixas por Funcionário ─────────────────────────────────────────────────────
function renderBaixasFunc(){{
  const sec=document.getElementById('baixas-func-section');
  if(!PLACARES||Object.keys(PLACARES).length===0){{sec.style.display='none';return;}}

  sec.style.display = 'block';

  const pd=PLACARES[bfPeriodoAtual]||{{total:0,por_resp:[]}};
  const rows=pd.por_resp.filter(r=>
    r.unidade===filialAtual &&
    r.dep===depAtual &&
    (!filtroCoord || (r.coord||'').split('|').map(s=>s.trim()).includes(filtroCoord)) &&
    (!filtroResp  || r.resp===filtroResp)
  );

  // Se o período atual não tem dados, mostra mensagem mas mantém seção visível
  if(!rows.length){{
    document.getElementById('bf-tbody').innerHTML=
      '<tr><td colspan="5" class="bf-empty">Nenhuma baixa registrada neste período.</td></tr>';
    return;
  }}

  const maxN=Math.max(...rows.map(r=>r.n),1);
  const porDep={{}};
  rows.forEach(r=>{{
    if(!porDep[r.dep]) porDep[r.dep]=[];
    porDep[r.dep].push(r);
  }});

  let html='';
  Object.keys(porDep).sort().forEach(dep=>{{
    html+=`<tr class="bf-dep-row"><td colspan="5">${{dep}}</td></tr>`;
    porDep[dep].forEach(r=>{{
      const pct=Math.round(r.n/maxN*100);
      const coordLabel=r.coord?r.coord.split('|').map(s=>s.trim()).join(', '):'—';
      html+=`<tr>
        <td><span class="bf-resp-name">${{r.resp||'—'}}</span></td>
        <td><span class="bf-coord-badge">${{coordLabel}}</span></td>
        <td style="font-size:12px;color:var(--muted)">${{r.dep}}</td>
        <td class="bf-bar-cell"><div class="bf-bar-bg"><div class="bf-bar-fill" style="width:${{pct}}%"></div></div></td>
        <td><span class="bf-num">${{r.n}}</span></td>
      </tr>`;
    }});
  }});
  document.getElementById('bf-tbody').innerHTML=html;
}}

function mudarBfTab(periodo,btn){{
  bfPeriodoAtual=periodo;
  document.querySelectorAll('.bf-tab').forEach(b=>b.classList.remove('ativo'));
  btn.classList.add('ativo');
  renderBaixasFunc();
}}

// ── Export Placares ───────────────────────────────────────────────────────────
async function exportarPlacares(modo){{
  const overlay=document.createElement('div');
  overlay.className='capture-overlay';
  overlay.innerHTML='<div class="capture-spinner"></div><span style="color:#aaa;font-size:13px">Gerando imagem...</span>';
  document.body.appendChild(overlay);
  await new Promise(r=>setTimeout(r,100));
  try{{
    if(modo==='geral'){{
      const el=document.getElementById('placares-grid');
      const canvas=await html2canvas(el,{{backgroundColor:'#111111',scale:2,useCORS:true,logging:false}});
      _downloadCanvas(canvas,`placares_geral_${{depAtual||'todos'}}.png`);
    }} else {{
      const periodos=[{{id:'pc-dia',label:'Baixas no Dia'}},{{id:'pc-sem',label:'Baixas na Semana'}},{{id:'pc-mes',label:'Baixas no Mês'}}];
      const capturas=await Promise.all(periodos.map(p=>html2canvas(document.getElementById(p.id),{{backgroundColor:'#111111',scale:2,useCORS:true,logging:false}})));
      const pad=24,headerH=60;
      const totalW=capturas.reduce((a,c)=>a+c.width,0)+pad*(capturas.length+1);
      const maxH=Math.max(...capturas.map(c=>c.height));
      const tmpCanvas=document.createElement('canvas');
      tmpCanvas.width=totalW; tmpCanvas.height=maxH+pad*2+headerH;
      const ctx=tmpCanvas.getContext('2d');
      ctx.fillStyle='#0a0a0a'; ctx.fillRect(0,0,tmpCanvas.width,tmpCanvas.height);
      ctx.fillStyle='#ffffff'; ctx.font='bold 28px Inter,sans-serif';
      ctx.fillText('MG Contécnica · Placar de Baixas',pad,38);
      ctx.fillStyle='#666'; ctx.font='20px Inter,sans-serif';
      ctx.fillText(`${{labelFilial(filialAtual)}} · ${{depAtual}} · ${{GERADO}}`,pad,62);
      let x=pad;
      capturas.forEach(c=>{{ctx.drawImage(c,x,headerH+pad);x+=c.width+pad;}});
      _downloadCanvas(tmpCanvas,`placares_dep_${{depAtual||'todos'}}.png`);
    }}
  }}catch(e){{console.error('Erro ao exportar:',e);alert('Erro ao gerar imagem.');}}
  finally{{document.body.removeChild(overlay);}}
}}
function _downloadCanvas(canvas,filename){{
  const link=document.createElement('a');link.download=filename;link.href=canvas.toDataURL('image/png');link.click();
}}

// ── Init ──────────────────────────────────────────────────────────────────────
buildFiliais();
</script>
</body>
</html>"""

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("base")
    parser.add_argument("atual", nargs="?", default=None)
    parser.add_argument("-o","--output", default=None)
    parser.add_argument("--base-dia",  default=None, dest="base_dia")
    parser.add_argument("--base-sem",  default=None, dest="base_sem")
    parser.add_argument("--base-mes",  default=None, dest="base_mes")
    parser.add_argument("--coordenadores", default=None, dest="coordenadores",
                        help="Caminho para a planilha de coordenadores (ex: coordenadores.xlsx)")
    args = parser.parse_args()

    resp_to_coords, all_coords = carregar_coordenadores(args.coordenadores)
    if args.coordenadores:
        print(f"[*] Coordenadores carregados: {len(all_coords)} coordenadores, {len(resp_to_coords)} responsáveis mapeados.")
    else:
        print("[!] Planilha de coordenadores não informada. Filtro de coordenador estará vazio.")

    print(f"[*] Lendo base: {args.base}")
    df_base = ler_relatorio(args.base, resp_to_coords)
    print(f"    {len(df_base)} tarefas.")

    if args.atual:
        print(f"[*] Lendo atual: {args.atual}")
        df_atual = ler_relatorio(args.atual, resp_to_coords)
        print(f"    {len(df_atual)} tarefas.")
        em_andamento, finalizadas, reabertas = comparar(df_base, df_atual)
    else:
       df_atual     = df_base
       em_andamento = df_base.copy()
       finalizadas  = pd.DataFrame(columns=df_base.columns)
       reabertas    = pd.DataFrame(columns=df_base.columns)
       args.atual   = args.base

    print(f"[*] Man:{len(em_andamento)} Fin:{len(finalizadas)} Rea:{len(reabertas)}")

    def ler_ou_none(caminho):
        if not caminho:
            return None
        try:
            df = ler_relatorio(caminho, resp_to_coords)
            print(f"[*] Placar base '{os.path.basename(caminho)}': {len(df)} tarefas.")
            return df
        except Exception as e:
            print(f"[!] Não foi possível ler '{caminho}' para placar: {e}")
            return None

    df_b_dia = ler_ou_none(args.base_dia)
    df_b_sem = ler_ou_none(args.base_sem)
    df_b_mes = ler_ou_none(args.base_mes)

    placares = calcular_placares(df_atual, df_b_dia, df_b_sem, df_b_mes)
    print(f"[*] Placares — Dia:{placares['dia']['total']} Sem:{placares['sem']['total']} Mês:{placares['mes']['total']}")

    html = gerar_html(args.base, args.atual, em_andamento, finalizadas, reabertas,
                  placares=placares, all_coords=all_coords or None)

    out = args.output or f"Relatorio_Pendencias_{datetime.now().strftime('%d-%m-%y_%H%M%S')}.html"
    with open(out,"w",encoding="utf-8") as f:
        f.write(html)
    print(f"[✓] Salvo: {out}")

if __name__=="__main__":
    main()