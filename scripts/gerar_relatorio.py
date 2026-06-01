#!/usr/bin/env python3
"""
MG Contécnica – Gerador de Relatório HTML Consolidado
Versão 5.0
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
        v = pd.to_datetime(valor, errors="coerce")
        return v.strftime("%d/%m/%Y") if not pd.isna(v) else str(valor).strip()
    except:
        return str(valor).strip()

def extrair_data_comentario(texto):
    oc = re.findall(r'Data Coment[aá]rio[:\s]+(\d{2}/\d{2}/\d{4}(?:\s+\d{2}:\d{2})?)', str(texto), re.IGNORECASE)
    return oc[-1].strip() if oc else ""

def limpar(v):
    s = str(v).strip()
    return "" if s in ("nan", "None", "NaT") else s

def ler_relatorio(caminho):
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
        registros.append({
            "CodCliente":   limpar(row.get(C_COD_CLIENTE, "")),
            "RazaoSocial":  limpar(row.get(C_RAZAO, "")),
            "Unidade":      limpar(row.get(C_UNIDADE, "")),
            "Departamento": dep_norm,
            "ClienteNovo":  cliente_novo,
            "Titulo":       titulo,
            "Vencimento":   formatar_data(row.get(C_VENCIMENTO, "")),
            "Previsao":     formatar_data(row.get(C_PREVISAO, "")),
            "Responsavel":  limpar(row.get(C_RESPONSAVEL, "")),
            "Grupo":        limpar(row.get(C_GRUPO, "")),
            "Comentario":   comt_raw,
            "DataComt":     extrair_data_comentario(comt_raw),
            "chave":        f"{limpar(row.get(C_COD_CLIENTE,''))}||{limpar(row.get('Cod',''))}||{dep_norm}||{titulo}",
        })
    df = pd.DataFrame(registros)
    if df.empty:
        raise ValueError(f"Nenhuma tarefa válida em '{os.path.basename(caminho)}'.")
    return df

def comparar(df_base, df_atual):
    cb = set(df_base["chave"]); ca = set(df_atual["chave"])
    return (df_base[df_base["chave"].isin(cb & ca)].copy(),
            df_base[df_base["chave"].isin(cb - ca)].copy(),
            df_atual[df_atual["chave"].isin(ca - cb)].copy())

def df_to_js(df):
    if df.empty: return "[]"
    rows = []
    for _, r in df.iterrows():
        rows.append({"unidade": r["Unidade"], "dep": r["Departamento"], "novo": r["ClienteNovo"],
                     "tit": r["Titulo"], "venc": r["Vencimento"], "prev": r["Previsao"],
                     "resp": r["Responsavel"], "grupo": r["Grupo"],
                     "cod": r["CodCliente"], "razao": r["RazaoSocial"],
                     "dataComt": r["DataComt"], "comt": r["Comentario"]})
    return json.dumps(rows, ensure_ascii=False)

def gerar_html(arquivo_base, arquivo_atual, em_andamento, finalizadas, adicionadas):
    nome_base  = os.path.basename(arquivo_base)
    nome_atual = os.path.basename(arquivo_atual) if arquivo_atual else nome_base
    gerado_em  = datetime.now().strftime("%d/%m/%Y às %H:%M")
    modo_comp  = arquivo_atual and arquivo_atual != arquivo_base

    js_man = df_to_js(em_andamento)
    js_fin = df_to_js(finalizadas)
    js_add = df_to_js(adicionadas)
    total_man = len(em_andamento)
    total_fin = len(finalizadas)
    total_add = len(adicionadas)

    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>MG Contécnica · Relatório de Pendências</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
:root{{
  --bg:#0a0a0a;--surface:#111111;--surface2:#181818;--border:#2a2a2a;
  --red:#E31E24;--red-dim:#7a1012;--red-glow:rgba(227,30,36,.18);
  --white:#ffffff;--off-white:#f0f0f0;--muted:#666;--muted2:#444;
  --green:#22c97a;--amber:#f5a623;
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

/* ── GRUPO DROPDOWN ── */
.grupo-bar{{display:flex;align-items:center;gap:12px;margin-bottom:20px;position:relative}}
.grupo-label{{font-size:10px;font-weight:700;letter-spacing:.12em;text-transform:uppercase;color:var(--muted);white-space:nowrap}}
.grupo-select-wrap{{position:relative;flex:1;max-width:380px}}
.grupo-select-btn{{width:100%;background:var(--surface);border:1px solid var(--border);color:var(--white);font-family:var(--font);font-size:13px;font-weight:500;padding:9px 36px 9px 14px;border-radius:8px;cursor:pointer;text-align:left;transition:border-color .18s;display:flex;align-items:center;justify-content:space-between;gap:8px}}
.grupo-select-btn:hover,.grupo-select-btn.open{{border-color:var(--red)}}
.grupo-select-btn .arrow{{font-size:10px;color:var(--muted);transition:transform .18s;flex-shrink:0}}
.grupo-select-btn.open .arrow{{transform:rotate(180deg)}}
.grupo-dropdown{{position:absolute;top:calc(100% + 6px);left:0;right:0;background:var(--surface2);border:1px solid var(--border);border-radius:10px;z-index:100;overflow:hidden;box-shadow:0 12px 40px rgba(0,0,0,.6);display:none;flex-direction:column}}
.grupo-dropdown.open{{display:flex}}
.grupo-search{{padding:10px 12px;border-bottom:1px solid var(--border)}}
.grupo-search input{{width:100%;background:var(--surface);border:1px solid var(--border);color:var(--white);font-family:var(--font);font-size:12px;padding:6px 10px;border-radius:6px;outline:none}}
.grupo-search input:focus{{border-color:var(--red)}}
.grupo-options{{max-height:260px;overflow-y:auto}}
.grupo-options::-webkit-scrollbar{{width:4px}}
.grupo-options::-webkit-scrollbar-track{{background:transparent}}
.grupo-options::-webkit-scrollbar-thumb{{background:var(--border);border-radius:4px}}
.grupo-opt{{display:flex;align-items:center;justify-content:space-between;padding:9px 14px;cursor:pointer;transition:background .15s;font-size:13px}}
.grupo-opt:hover{{background:rgba(227,30,36,.08)}}
.grupo-opt.sel{{background:rgba(227,30,36,.12);color:var(--red)}}
.grupo-opt-name{{font-weight:500}}
.grupo-opt-count{{font-size:11px;font-family:var(--mono);color:var(--muted);background:var(--surface);padding:1px 7px;border-radius:10px}}
.grupo-opt.sel .grupo-opt-count{{background:rgba(227,30,36,.2);color:var(--red)}}
.grupo-clear{{padding:8px 14px;border-top:1px solid var(--border);font-size:11px;color:var(--muted);cursor:pointer;transition:color .15s;text-align:center;font-weight:600;letter-spacing:.06em;text-transform:uppercase}}
.grupo-clear:hover{{color:var(--red)}}

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

/* ── CHARTS SECTION ── */
.charts-section{{margin:32px 0 28px}}
.charts-title{{font-size:10px;font-weight:700;letter-spacing:.14em;text-transform:uppercase;color:var(--muted);margin-bottom:16px;display:flex;align-items:center;gap:10px}}
.charts-title::after{{content:'';flex:1;height:1px;background:var(--border)}}
.charts-grid{{display:grid;grid-template-columns:1fr 1fr;gap:16px}}
.chart-card{{background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:20px 24px}}
.chart-card-title{{font-size:11px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:var(--muted);margin-bottom:16px}}
.chart-wrap{{position:relative;height:240px}}

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

@media(max-width:700px){{
  .filial-grid{{grid-template-columns:1fr}}
  #screen-dep,.results-body,.results-header{{padding:20px}}
  .cards-row,.charts-grid{{grid-template-columns:1fr}}
  .logo-bar{{padding:16px 20px 0}}
}}
</style>
</head>
<body>

<!-- ══ SCREEN 1: FILIAL ══ -->
<div class="screen active" id="screen-filial">
  <div class="logo-bar">
    <img class="logo-img" src="https://i.imgur.com/0GqDpHS.png" onerror="this.style.display='none'" alt="MG">
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
    <img class="logo-img" src="https://i.imgur.com/0GqDpHS.png" onerror="this.style.display='none'" alt="MG">
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
      <div class="stat-card add"><div class="sc-label">Adicionadas</div><div class="sc-num" id="sc-add">0</div><div class="sc-sub">novas desde a base</div></div>
    </div>

    <!-- Grupo filter dropdown -->
    <div class="grupo-bar">
      <span class="grupo-label">Grupo:</span>
      <div class="grupo-select-wrap" id="grupo-wrap">
        <button class="grupo-select-btn" id="grupo-btn" onclick="toggleDropdown()">
          <span id="grupo-btn-label">Todos os grupos</span>
          <span class="arrow">▼</span>
        </button>
        <div class="grupo-dropdown" id="grupo-dropdown">
          <div class="grupo-search"><input type="text" placeholder="Buscar grupo..." id="grupo-search-input" oninput="filtrarOpcoes()"></div>
          <div class="grupo-options" id="grupo-options"></div>
          <div class="grupo-clear" onclick="limparGrupo()">Limpar filtro</div>
        </div>
      </div>
    </div>

    <!-- Charts -->
    <div class="charts-section" id="charts-section">
      <div class="charts-title">Evolução por Departamento</div>
      <div class="charts-grid">
        <div class="chart-card">
          <div class="chart-card-title">Em Andamento por Departamento</div>
          <div class="chart-wrap"><canvas id="chart-man"></canvas></div>
        </div>
        <div class="chart-card">
          <div class="chart-card-title">Baixadas vs Adicionadas</div>
          <div class="chart-wrap"><canvas id="chart-compare"></canvas></div>
        </div>
      </div>
    </div>

    <div class="tabs">
      <button class="tab-btn ativo"  id="tbtn-man" onclick="mudarTab('man',this)">Em Andamento <span class="badge" id="bdg-man" style="background:var(--red)">0</span></button>
      <button class="tab-btn t-fin"  id="tbtn-fin" onclick="mudarTab('fin',this)">Baixadas <span class="badge" id="bdg-fin" style="background:var(--green)">0</span></button>
      <button class="tab-btn t-add"  id="tbtn-add" onclick="mudarTab('add',this)">Adicionadas <span class="badge" id="bdg-add" style="background:var(--amber)">0</span></button>
    </div>

    <div id="tab-man" class="tab-content ativo"><div class="tbl-wrap"><table><thead><tr><th>Cliente</th><th>Título da Tarefa</th><th>Vencimento</th><th>Responsável</th><th>Previsão</th><th>Comentário</th></tr></thead><tbody id="tbody-man"></tbody></table></div></div>
    <div id="tab-fin" class="tab-content"><div class="tbl-wrap"><table><thead><tr><th>Cliente</th><th>Título da Tarefa</th><th>Vencimento</th><th>Responsável</th><th>Previsão</th><th>Comentário</th></tr></thead><tbody id="tbody-fin"></tbody></table></div></div>
    <div id="tab-add" class="tab-content"><div class="tbl-wrap"><table><thead><tr><th>Cliente</th><th>Título da Tarefa</th><th>Vencimento</th><th>Responsável</th><th>Previsão</th><th>Comentário</th></tr></thead><tbody id="tbody-add"></tbody></table></div></div>
  </div>
</div>

<script>
const DATA = {{ man:{js_man}, fin:{js_fin}, add:{js_add} }};
const MODO_COMP = {'true' if modo_comp else 'false'};
const ARQ_BASE = {json.dumps(nome_base)};
const ARQ_ATUAL = {json.dumps(nome_atual)};
const GERADO = {json.dumps(gerado_em)};

let filialAtual=null, depAtual=null, grupoAtual=null, abaAtual='man';
let chartMan=null, chartComp=null;

// ── Helpers ─────────────────────────────────────────────────────────────────
const LABELS = {{'SP':'São Paulo','Santos':'Santos','RJ':'Rio de Janeiro','GOIAS':'Goiás'}};
function labelFilial(f){{ return LABELS[f]||f; }}

function getFiliais(){{
  return [...new Set([...DATA.man,...DATA.fin,...DATA.add].map(r=>r.unidade))].sort();
}}
function getDeps(filial){{
  return [...new Set([...DATA.man,...DATA.fin,...DATA.add].filter(r=>r.unidade===filial).map(r=>r.dep))].sort();
}}
function getGruposOrdenados(filial, dep){{
  const rows = [...DATA.man,...DATA.fin,...DATA.add].filter(r=>r.unidade===filial&&r.dep===dep&&r.grupo);
  const cnt={{}};
  rows.forEach(r=>{{ cnt[r.grupo]=(cnt[r.grupo]||0)+1; }});
  return Object.entries(cnt).sort((a,b)=>b[1]-a[1]).map(([g,c])=>{{return{{grupo:g,count:c}}}});
}}
function countIn(arr,filial,dep,grupo){{
  return arr.filter(r=>r.unidade===filial&&r.dep===dep&&(!grupo||r.grupo===grupo)).length;
}}
function filtrar(arr){{
  return arr.filter(r=>r.unidade===filialAtual&&r.dep===depAtual&&(!grupoAtual||r.grupo===grupoAtual));
}}

// ── Screens ──────────────────────────────────────────────────────────────────
function showScreen(id){{document.querySelectorAll('.screen').forEach(s=>s.classList.remove('active'));document.getElementById(id).classList.add('active');window.scrollTo(0,0);}}
function voltarFilial(){{showScreen('screen-filial');filialAtual=null;depAtual=null;}}
function voltarDep(){{showScreen('screen-dep');depAtual=null;grupoAtual=null;}}

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
  filialAtual=filial; grupoAtual=null;
  document.getElementById('dep-screen-title').textContent=labelFilial(filial);
  const deps=getDeps(filial);
  document.getElementById('dep-screen-sub').textContent=`${{deps.length}} departamentos`;
  const list=document.getElementById('dep-list');
  list.innerHTML='';
  deps.forEach(dep=>{{
    const m=countIn(DATA.man,filial,dep,null);
    const f=countIn(DATA.fin,filial,dep,null);
    const a=countIn(DATA.add,filial,dep,null);
    const el=document.createElement('div');
    el.className='dep-item';
    el.innerHTML=`<span class="dep-item-name">${{dep}}</span><div class="dep-pills"><span class="dp dp-man">${{m}}</span>${{MODO_COMP?`<span class="dp dp-fin">${{f}}</span><span class="dp dp-add">${{a}}</span>`:''}}</div>`;
    el.onclick=()=>selecionarDep(dep);
    list.appendChild(el);
  }});
  showScreen('screen-dep');
}}

// ── Screen 3 ──────────────────────────────────────────────────────────────────
function selecionarDep(dep){{
  depAtual=dep; grupoAtual=null;
  abaAtual='man';
  document.getElementById('rh-title').innerHTML=`<em>${{labelFilial(filialAtual)}}</em> · ${{dep}}`;
  document.getElementById('rh-meta-base').textContent=`Base: ${{ARQ_BASE}}`;
  document.getElementById('rh-meta-atual').textContent=MODO_COMP?`Atual: ${{ARQ_ATUAL}} · Gerado em: ${{GERADO}}`:`Gerado em: ${{GERADO}}`;
  document.querySelectorAll('.tab-btn').forEach(b=>b.classList.remove('ativo'));
  document.getElementById('tbtn-man').classList.add('ativo');
  document.querySelectorAll('.tab-content').forEach(t=>t.classList.remove('ativo'));
  document.getElementById('tab-man').classList.add('ativo');
  buildDropdown();
  renderAll();
  renderCharts();
  showScreen('screen-results');
}}

// ── Grupo Dropdown ────────────────────────────────────────────────────────────
let gruposData=[];
function buildDropdown(){{
  gruposData=getGruposOrdenados(filialAtual,depAtual);
  renderOpcoes('');
}}
function renderOpcoes(busca){{
  const el=document.getElementById('grupo-options');
  el.innerHTML='';
  const filtrados=busca?gruposData.filter(g=>g.grupo.toLowerCase().includes(busca.toLowerCase())):gruposData;
  filtrados.forEach(g=>{{
    const opt=document.createElement('div');
    opt.className='grupo-opt'+(grupoAtual===g.grupo?' sel':'');
    opt.innerHTML=`<span class="grupo-opt-name">${{g.grupo}}</span><span class="grupo-opt-count">${{g.count}}</span>`;
    opt.onclick=()=>selecionarGrupo(g.grupo);
    el.appendChild(opt);
  }});
}}
function filtrarOpcoes(){{
  renderOpcoes(document.getElementById('grupo-search-input').value);
}}
function toggleDropdown(){{
  const btn=document.getElementById('grupo-btn');
  const dd=document.getElementById('grupo-dropdown');
  const open=dd.classList.toggle('open');
  btn.classList.toggle('open',open);
  if(open){{ document.getElementById('grupo-search-input').value=''; renderOpcoes(''); setTimeout(()=>document.getElementById('grupo-search-input').focus(),50); }}
}}
function selecionarGrupo(g){{
  grupoAtual=g;
  document.getElementById('grupo-btn-label').textContent=g;
  document.getElementById('grupo-dropdown').classList.remove('open');
  document.getElementById('grupo-btn').classList.remove('open');
  renderAll(); renderCharts();
}}
function limparGrupo(){{
  grupoAtual=null;
  document.getElementById('grupo-btn-label').textContent='Todos os grupos';
  document.getElementById('grupo-dropdown').classList.remove('open');
  document.getElementById('grupo-btn').classList.remove('open');
  renderAll(); renderCharts();
}}
document.addEventListener('click',e=>{{
  const wrap=document.getElementById('grupo-wrap');
  if(wrap&&!wrap.contains(e.target)){{
    document.getElementById('grupo-dropdown').classList.remove('open');
    document.getElementById('grupo-btn').classList.remove('open');
  }}
}});

// ── Render Table ──────────────────────────────────────────────────────────────
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

// ── Charts ────────────────────────────────────────────────────────────────────
Chart.defaults.color='#666';
Chart.defaults.font.family="'Inter',sans-serif";

function renderCharts(){{
  const m=filtrar(DATA.man), f=filtrar(DATA.fin), a=filtrar(DATA.add);
  const deps=[...new Set([...m,...f,...a].map(r=>r.dep))].sort();

  if(chartMan)chartMan.destroy();
  if(chartComp)chartComp.destroy();

  if(!deps.length){{document.getElementById('charts-section').style.display='none';return;}}
  document.getElementById('charts-section').style.display='block';

  const shortDep=d=>d.length>18?d.substring(0,16)+'…':d;
  const manCounts=deps.map(d=>m.filter(r=>r.dep===d).length);
  const finCounts=deps.map(d=>f.filter(r=>r.dep===d).length);
  const addCounts=deps.map(d=>a.filter(r=>r.dep===d).length);

  // Chart 1: Em Andamento por dep (horizontal bar)
  const ctx1=document.getElementById('chart-man').getContext('2d');
  chartMan=new Chart(ctx1,{{
    type:'bar',
    data:{{
      labels:deps.map(shortDep),
      datasets:[{{
        label:'Em Andamento',
        data:manCounts,
        backgroundColor:'rgba(227,30,36,.7)',
        borderColor:'#E31E24',
        borderWidth:1,
        borderRadius:4,
      }}]
    }},
    options:{{
      indexAxis:'y',
      responsive:true,maintainAspectRatio:false,
      plugins:{{legend:{{display:false}},tooltip:{{callbacks:{{title:i=>deps[i[0].dataIndex]}}}}}},
      scales:{{
        x:{{grid:{{color:'rgba(255,255,255,.05)'}},ticks:{{color:'#666',font:{{size:11}}}}}},
        y:{{grid:{{display:false}},ticks:{{color:'#aaa',font:{{size:11}}}}}}
      }}
    }}
  }});

  // Chart 2: Baixadas vs Adicionadas (grouped bar) — only in comparison mode
  const ctx2=document.getElementById('chart-compare').getContext('2d');
  if(MODO_COMP && (f.length||a.length)){{
    chartComp=new Chart(ctx2,{{
      type:'bar',
      data:{{
        labels:deps.map(shortDep),
        datasets:[
          {{label:'Baixadas',data:finCounts,backgroundColor:'rgba(34,201,122,.7)',borderColor:'#22c97a',borderWidth:1,borderRadius:4}},
          {{label:'Adicionadas',data:addCounts,backgroundColor:'rgba(245,166,35,.7)',borderColor:'#f5a623',borderWidth:1,borderRadius:4}}
        ]
      }},
      options:{{
        responsive:true,maintainAspectRatio:false,
        plugins:{{legend:{{labels:{{color:'#aaa',font:{{size:11}}}}}},tooltip:{{callbacks:{{title:i=>deps[i[0].dataIndex]}}}}}},
        scales:{{
          x:{{grid:{{color:'rgba(255,255,255,.05)'}},ticks:{{color:'#666',font:{{size:11}}}}}},
          y:{{grid:{{color:'rgba(255,255,255,.05)'}},ticks:{{color:'#666',font:{{size:11}}}}}}
        }}
      }}
    }});
  }} else {{
    // No comparison: show distribution donut
    const total=manCounts.reduce((a,b)=>a+b,0);
    if(total>0){{
      chartComp=new Chart(ctx2,{{
        type:'doughnut',
        data:{{
          labels:deps.map(shortDep),
          datasets:[{{
            data:manCounts,
            backgroundColor:deps.map((_,i)=>`hsl(${{(i*47+5)%360}},60%,50%)`),
            borderColor:'#111',borderWidth:2
          }}]
        }},
        options:{{
          responsive:true,maintainAspectRatio:false,
          plugins:{{
            legend:{{position:'right',labels:{{color:'#aaa',font:{{size:10}},boxWidth:12}}}},
            tooltip:{{callbacks:{{title:i=>deps[i[0].dataIndex],label:i=>`${{i.raw}} (${{Math.round(i.raw/total*100)}}%)`}}}}
          }}
        }}
      }});
    }}
  }}
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
    args = parser.parse_args()

    print(f"[*] Lendo base: {args.base}")
    df_base = ler_relatorio(args.base)
    print(f"    {len(df_base)} tarefas.")

    if args.atual:
        print(f"[*] Lendo atual: {args.atual}")
        df_atual = ler_relatorio(args.atual)
        print(f"    {len(df_atual)} tarefas.")
        em_andamento, finalizadas, adicionadas = comparar(df_base, df_atual)
    else:
        em_andamento = df_base.copy()
        finalizadas  = pd.DataFrame(columns=df_base.columns)
        adicionadas  = pd.DataFrame(columns=df_base.columns)
        args.atual   = args.base

    print(f"[*] Man:{len(em_andamento)} Fin:{len(finalizadas)} Add:{len(adicionadas)}")
    html = gerar_html(args.base, args.atual, em_andamento, finalizadas, adicionadas)

    out = args.output or f"Relatorio_Pendencias_{datetime.now().strftime('%d-%m-%y_%H%M%S')}.html"
    with open(out,"w",encoding="utf-8") as f:
        f.write(html)
    print(f"[✓] Salvo: {out}")

if __name__=="__main__":
    main()
