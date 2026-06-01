#!/usr/bin/env python3
"""
MG Contécnica – Gerador de Relatório HTML Consolidado
Versão 4.0 — suporte a planilha CONSOLIDADA com múltiplas filiais

Uso:
    python gerar_relatorio.py <arquivo_base.xlsx> [arquivo_atual.xlsx]

Se apenas um arquivo for passado, gera visualização única (sem comparação).
Se dois arquivos forem passados, compara base × atual.
O HTML gerado contém fluxo de seleção: Filial → Departamento → Resultados.
"""

import os
import re
import sys
import json
import argparse
import pandas as pd
from datetime import datetime

SHEET = "Pendencias CONSOLIDADO"

# ── Colunas da planilha ──────────────────────────────────────────────────────
C_COD_CLIENTE   = "CodCliente"
C_UNIDADE       = "Unidade"
C_DEPARTAMENTO  = "Departamento"
C_TITULO        = "Titulo"
C_VENCIMENTO    = "DataVencimento"
C_PREVISAO      = "DataPrevisaoConclusao"
C_RESPONSAVEL   = "UsuarioResponsavel"
C_GRUPO         = "Grupo"
C_COMENTARIO    = "Comentario"

# ── Utilitários ──────────────────────────────────────────────────────────────

def normalizar_dep(dep: str) -> str:
    """Remove sufixo '- (Cliente Novo)' para agrupar no mesmo filtro."""
    return re.sub(r'\s*-\s*\(Cliente Novo\)\s*$', '', str(dep).strip(), flags=re.IGNORECASE)

def formatar_data(valor) -> str:
    if pd.isna(valor) or str(valor).strip() in ("", "nan", "NaT", "None"):
        return ""
    try:
        v = pd.to_datetime(valor, errors="coerce")
        return v.strftime("%d/%m/%Y") if not pd.isna(v) else str(valor).strip()
    except Exception:
        return str(valor).strip()

def extrair_data_comentario(texto: str) -> str:
    ocorrencias = re.findall(
        r'Data Coment[aá]rio[:\s]+(\d{2}/\d{2}/\d{4}(?:\s+\d{2}:\d{2})?)',
        str(texto), re.IGNORECASE
    )
    return ocorrencias[-1].strip() if ocorrencias else ""

def limpar(v) -> str:
    s = str(v).strip()
    return "" if s in ("nan", "None", "NaT") else s

# ── Leitura ──────────────────────────────────────────────────────────────────

def ler_relatorio(caminho: str) -> pd.DataFrame:
    try:
        df_raw = pd.read_excel(caminho, sheet_name=SHEET, dtype=str)
    except Exception as e:
        raise ValueError(f"Não foi possível ler '{os.path.basename(caminho)}': {e}")

    registros = []
    for _, row in df_raw.iterrows():
        dep_raw   = limpar(row.get(C_DEPARTAMENTO, ""))
        titulo    = limpar(row.get(C_TITULO, ""))
        if not dep_raw and not titulo:
            continue

        dep_norm  = normalizar_dep(dep_raw)
        cliente_novo = bool(re.search(r'\(Cliente Novo\)', dep_raw, re.IGNORECASE))

        comt_raw  = limpar(row.get(C_COMENTARIO, ""))
        registros.append({
            "CodCliente":    limpar(row.get(C_COD_CLIENTE, "")),
            "Unidade":       limpar(row.get(C_UNIDADE, "")),
            "Departamento":  dep_norm,
            "ClienteNovo":   cliente_novo,
            "Titulo":        titulo,
            "Vencimento":    formatar_data(row.get(C_VENCIMENTO, "")),
            "Previsao":      formatar_data(row.get(C_PREVISAO, "")),
            "Responsavel":   limpar(row.get(C_RESPONSAVEL, "")),
            "Grupo":         limpar(row.get(C_GRUPO, "")),
            "Comentario":    comt_raw,
            "DataComt":      extrair_data_comentario(comt_raw),
            "chave":         f"{limpar(row.get(C_COD_CLIENTE,''))}||{limpar(row.get('Cod',''))}||{dep_norm}||{titulo}",
        })

    df = pd.DataFrame(registros)
    if df.empty:
        raise ValueError(f"Nenhuma tarefa válida em '{os.path.basename(caminho)}'.")
    return df

# ── Comparação ───────────────────────────────────────────────────────────────

def comparar(df_base: pd.DataFrame, df_atual: pd.DataFrame):
    cb = set(df_base["chave"])
    ca = set(df_atual["chave"])
    em_andamento = df_base[df_base["chave"].isin(cb & ca)].copy()
    finalizadas  = df_base[df_base["chave"].isin(cb - ca)].copy()
    adicionadas  = df_atual[df_atual["chave"].isin(ca - cb)].copy()
    return em_andamento, finalizadas, adicionadas

# ── Serialização para JS ──────────────────────────────────────────────────────

def df_to_js(df: pd.DataFrame) -> str:
    if df.empty:
        return "[]"
    rows = []
    for _, r in df.iterrows():
        rows.append({
            "unidade":  r["Unidade"],
            "dep":      r["Departamento"],
            "novo":     r["ClienteNovo"],
            "tit":      r["Titulo"],
            "venc":     r["Vencimento"],
            "prev":     r["Previsao"],
            "resp":     r["Responsavel"],
            "grupo":    r["Grupo"],
            "cod":      r["CodCliente"],
            "dataComt": r["DataComt"],
            "comt":     r["Comentario"],
        })
    return json.dumps(rows, ensure_ascii=False)

# ── HTML ─────────────────────────────────────────────────────────────────────

def gerar_html(arquivo_base: str, arquivo_atual: str,
               em_andamento: pd.DataFrame, finalizadas: pd.DataFrame, adicionadas: pd.DataFrame) -> str:

    nome_base   = os.path.basename(arquivo_base)
    nome_atual  = os.path.basename(arquivo_atual) if arquivo_atual else nome_base
    gerado_em   = datetime.now().strftime("%d/%m/%Y às %H:%M")
    modo_comp   = arquivo_atual and arquivo_atual != arquivo_base

    js_man = df_to_js(em_andamento)
    js_fin = df_to_js(finalizadas)
    js_add = df_to_js(adicionadas)

    total_man = len(em_andamento)
    total_fin = len(finalizadas)
    total_add = len(adicionadas)

    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>MG Contécnica · Relatório de Pendências</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Syne:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
:root{{
  --bg:#080b12;--surface:#0e1220;--surface2:#141929;--border:#1e2540;
  --accent:#5b8fff;--green:#1dd98a;--red:#ff4f6e;--amber:#ffb020;
  --text:#dce4f5;--muted:#5a6480;
  --font:'Syne',sans-serif;--mono:'JetBrains Mono',monospace;
}}
*,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
html{{scroll-behavior:smooth}}
body{{font-family:var(--font);background:var(--bg);color:var(--text);min-height:100vh;overflow-x:hidden}}

/* ── GRID BACKGROUND ── */
body::before{{content:'';position:fixed;inset:0;
  background-image:linear-gradient(rgba(91,143,255,.03) 1px,transparent 1px),linear-gradient(90deg,rgba(91,143,255,.03) 1px,transparent 1px);
  background-size:40px 40px;pointer-events:none;z-index:0}}

/* ── SCREENS ── */
.screen{{position:relative;z-index:1;display:none;min-height:100vh;flex-direction:column}}
.screen.active{{display:flex}}

/* ── HEADER ── */
.logo-bar{{display:flex;align-items:center;gap:16px;padding:28px 48px 0}}
.logo-mark{{width:36px;height:36px;background:linear-gradient(135deg,var(--accent),#3060d0);border-radius:8px;display:grid;place-items:center;font-size:14px;font-weight:800;color:#fff}}
.logo-text{{font-size:13px;font-weight:600;letter-spacing:.12em;text-transform:uppercase;color:var(--muted)}}
.logo-text span{{color:var(--text)}}

/* ── SCREEN 1: FILIAL ── */
#screen-filial{{align-items:center;justify-content:center;gap:0;padding:40px 48px}}
.hero{{text-align:center;margin-bottom:60px}}
.hero-tag{{font-size:11px;font-weight:600;letter-spacing:.2em;text-transform:uppercase;color:var(--accent);margin-bottom:16px}}
.hero-title{{font-size:clamp(32px,5vw,52px);font-weight:800;line-height:1.1;letter-spacing:-.03em;margin-bottom:12px}}
.hero-title em{{color:var(--accent);font-style:normal}}
.hero-sub{{font-size:14px;color:var(--muted);font-family:var(--mono)}}

.filial-grid{{display:grid;grid-template-columns:repeat(2,1fr);gap:16px;max-width:680px;width:100%}}
.filial-card{{background:var(--surface);border:1px solid var(--border);border-radius:16px;padding:28px 32px;cursor:pointer;transition:all .25s;position:relative;overflow:hidden;display:flex;flex-direction:column;gap:8px}}
.filial-card::before{{content:'';position:absolute;inset:0;background:linear-gradient(135deg,var(--accent),transparent);opacity:0;transition:opacity .25s}}
.filial-card:hover{{border-color:var(--accent);transform:translateY(-3px);box-shadow:0 16px 48px rgba(91,143,255,.2)}}
.filial-card:hover::before{{opacity:.05}}
.fc-num{{font-size:11px;font-weight:600;letter-spacing:.12em;text-transform:uppercase;color:var(--muted);font-family:var(--mono)}}
.fc-name{{font-size:22px;font-weight:800;letter-spacing:-.02em}}
.fc-badge{{display:inline-flex;align-items:center;gap:6px;font-size:11px;color:var(--accent);font-family:var(--mono);margin-top:4px}}
.fc-badge span{{width:6px;height:6px;background:var(--accent);border-radius:50%;display:inline-block}}

/* ── SCREEN 2: DEPARTAMENTO ── */
#screen-dep{{padding:48px}}
.back-btn{{display:inline-flex;align-items:center;gap:8px;background:none;border:1px solid var(--border);color:var(--muted);font-family:var(--font);font-size:12px;font-weight:600;letter-spacing:.06em;text-transform:uppercase;padding:8px 16px;border-radius:8px;cursor:pointer;transition:all .2s;margin-bottom:40px}}
.back-btn:hover{{border-color:var(--accent);color:var(--text)}}
.screen-header{{margin-bottom:40px}}
.screen-header .tag{{font-size:11px;font-weight:600;letter-spacing:.15em;text-transform:uppercase;color:var(--accent);margin-bottom:10px}}
.screen-header h2{{font-size:clamp(24px,3vw,36px);font-weight:800;letter-spacing:-.02em}}
.screen-header p{{font-size:13px;color:var(--muted);font-family:var(--mono);margin-top:6px}}

.dep-list{{display:grid;grid-template-columns:repeat(auto-fill,minmax(280px,1fr));gap:12px;max-width:1100px}}
.dep-item{{background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:16px 20px;cursor:pointer;transition:all .2s;display:flex;align-items:center;justify-content:space-between;gap:12px}}
.dep-item:hover{{border-color:var(--accent);background:rgba(91,143,255,.06)}}
.dep-item-name{{font-size:14px;font-weight:600}}
.dep-item-counts{{display:flex;gap:8px;flex-shrink:0}}
.dep-pill{{font-size:11px;font-weight:700;padding:3px 9px;border-radius:20px}}
.dp-man{{background:rgba(91,143,255,.15);color:var(--accent)}}
.dp-fin{{background:rgba(29,217,138,.15);color:var(--green)}}
.dp-add{{background:rgba(255,79,110,.15);color:var(--red)}}

/* ── SCREEN 3: RESULTS ── */
#screen-results{{padding:0}}
.results-header{{background:linear-gradient(135deg,#0e1525 0%,#080b12 60%);border-bottom:1px solid var(--border);padding:32px 48px 24px;position:relative;overflow:hidden}}
.results-header::after{{content:'';position:absolute;top:-100px;right:-100px;width:400px;height:400px;background:radial-gradient(circle,rgba(91,143,255,.12) 0%,transparent 60%);pointer-events:none}}
.rh-meta{{font-size:11px;color:var(--muted);font-family:var(--mono);margin-top:12px;display:flex;flex-direction:column;gap:3px}}
.rh-title{{font-size:clamp(20px,2.5vw,28px);font-weight:800;letter-spacing:-.02em;margin-top:8px}}
.rh-title em{{color:var(--accent);font-style:normal}}
.rh-back{{display:inline-flex;align-items:center;gap:8px;background:none;border:1px solid var(--border);color:var(--muted);font-family:var(--font);font-size:12px;font-weight:600;letter-spacing:.06em;text-transform:uppercase;padding:7px 14px;border-radius:8px;cursor:pointer;transition:all .2s}}
.rh-back:hover{{border-color:var(--accent);color:var(--text)}}

.results-body{{padding:32px 48px 60px;max-width:1400px}}

/* ── CARDS ── */
.cards-row{{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-bottom:32px}}
.stat-card{{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:20px 24px;position:relative;overflow:hidden}}
.stat-card::after{{content:'';position:absolute;bottom:0;left:0;right:0;height:2px}}
.stat-card.man::after{{background:var(--accent)}}
.stat-card.fin::after{{background:var(--green)}}
.stat-card.add::after{{background:var(--red)}}
.sc-label{{font-size:10px;font-weight:600;letter-spacing:.12em;text-transform:uppercase;color:var(--muted);margin-bottom:10px}}
.sc-num{{font-size:38px;font-weight:800;letter-spacing:-.04em;line-height:1}}
.stat-card.man .sc-num{{color:var(--accent)}}
.stat-card.fin .sc-num{{color:var(--green)}}
.stat-card.add .sc-num{{color:var(--red)}}
.sc-sub{{font-size:11px;color:var(--muted);margin-top:5px}}

/* ── GRUPO FILTER ── */
.filter-row{{display:flex;align-items:flex-start;gap:12px;margin-bottom:24px;flex-wrap:wrap}}
.filter-label{{font-size:11px;font-weight:600;letter-spacing:.1em;text-transform:uppercase;color:var(--muted);padding-top:7px;white-space:nowrap}}
.chips{{display:flex;flex-wrap:wrap;gap:6px;flex:1}}
.chip{{background:var(--surface2);border:1px solid var(--border);color:var(--muted);font-family:var(--font);font-size:12px;font-weight:600;padding:5px 14px;border-radius:20px;cursor:pointer;transition:all .18s;user-select:none}}
.chip:hover{{border-color:var(--accent);color:var(--text)}}
.chip.sel{{background:rgba(91,143,255,.18);border-color:var(--accent);color:var(--accent)}}

/* ── TABS ── */
.tabs{{display:flex;gap:4px;border-bottom:1px solid var(--border);margin-bottom:0}}
.tab-btn{{background:none;border:none;border-bottom:2px solid transparent;padding:10px 22px;font-family:var(--font);font-size:13px;font-weight:600;color:var(--muted);cursor:pointer;transition:all .2s;margin-bottom:-1px;letter-spacing:.02em}}
.tab-btn:hover{{color:var(--text)}}
.tab-btn.ativo{{color:var(--text);border-bottom-color:var(--accent)}}
.tab-btn.t-fin.ativo{{color:var(--green);border-bottom-color:var(--green)}}
.tab-btn.t-add.ativo{{color:var(--red);border-bottom-color:var(--red)}}
.badge{{display:inline-block;font-size:11px;font-weight:700;padding:2px 8px;border-radius:20px;margin-left:6px;color:#fff;vertical-align:middle}}
.tab-content{{display:none;padding-top:1px}}
.tab-content.ativo{{display:block}}

/* ── TABLE ── */
.tbl-wrap{{overflow-x:auto;margin-top:16px}}
table{{width:100%;border-collapse:collapse;font-size:13px;background:var(--surface);border:1px solid var(--border);border-radius:12px;overflow:hidden}}
thead tr{{background:var(--surface2)}}
thead th{{padding:12px 16px;text-align:left;font-size:10px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:var(--muted);border-bottom:1px solid var(--border);white-space:nowrap}}
tbody tr{{border-bottom:1px solid var(--border);transition:background .15s}}
tbody tr:last-child{{border-bottom:none}}
tbody tr:hover{{background:rgba(255,255,255,.025)}}
tbody td{{padding:10px 16px;color:var(--text);vertical-align:top}}
tr.dep-row td{{background:rgba(91,143,255,.07);color:var(--accent);font-weight:700;font-size:11px;letter-spacing:.08em;text-transform:uppercase;padding:8px 16px;border-top:1px solid var(--border)}}
tr.dep-row td .cn-badge{{font-size:9px;background:rgba(255,176,32,.15);color:var(--amber);border:1px solid rgba(255,176,32,.3);padding:1px 7px;border-radius:12px;font-weight:700;margin-left:8px;letter-spacing:.06em;vertical-align:middle}}
.cod-tag{{font-size:11px;font-family:var(--mono);background:rgba(91,143,255,.1);color:var(--accent);border:1px solid rgba(91,143,255,.2);padding:2px 8px;border-radius:6px;margin-right:4px;white-space:nowrap}}
.comt-date{{font-size:11px;color:var(--amber);font-family:var(--mono);white-space:nowrap}}
.comt-btn{{background:rgba(91,143,255,.12);border:1px solid rgba(91,143,255,.25);color:var(--accent);font-size:11px;font-family:var(--font);padding:3px 10px;border-radius:16px;cursor:pointer;transition:background .2s;white-space:nowrap;margin-top:4px;display:inline-block}}
.comt-btn:hover{{background:rgba(91,143,255,.25)}}
.comt-full{{display:none;margin-top:7px;font-size:12px;color:var(--muted);background:rgba(0,0,0,.3);border-left:2px solid var(--accent);padding:9px 13px;border-radius:0 6px 6px 0;line-height:1.7;white-space:pre-wrap;font-family:var(--mono)}}
.comt-full.open{{display:block}}
.no-data{{color:var(--muted);font-size:13px;padding:24px 16px;font-style:italic;text-align:center}}
.resp-name{{font-size:12.5px;font-weight:600}}

/* ── RESPONSIVE ── */
@media(max-width:700px){{
  .filial-grid{{grid-template-columns:1fr}}
  #screen-dep{{padding:24px}}
  .results-header{{padding:24px}}
  .results-body{{padding:24px 20px 48px}}
  .cards-row{{grid-template-columns:1fr;gap:10px}}
  .logo-bar{{padding:20px 24px 0}}
}}
</style>
</head>
<body>

<!-- ══ SCREEN 1: FILIAL ══════════════════════════════════════════════ -->
<div class="screen active" id="screen-filial">
  <div class="logo-bar">
    <div class="logo-mark">MG</div>
    <div class="logo-text"><span>MG Contécnica</span> · Relatório de Pendências</div>
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

<!-- ══ SCREEN 2: DEPARTAMENTO ══════════════════════════════════════ -->
<div class="screen" id="screen-dep">
  <div class="logo-bar">
    <div class="logo-mark">MG</div>
    <div class="logo-text"><span>MG Contécnica</span> · Selecione o Departamento</div>
  </div>
  <div style="flex:1;padding:40px 48px">
    <button class="back-btn" onclick="voltarFilial()">← Voltar</button>
    <div class="screen-header">
      <div class="tag">Filial selecionada</div>
      <h2 id="dep-screen-title">—</h2>
      <p id="dep-screen-sub">Escolha um departamento para visualizar as pendências</p>
    </div>
    <div class="dep-list" id="dep-list"></div>
  </div>
</div>

<!-- ══ SCREEN 3: RESULTADOS ══════════════════════════════════════════ -->
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

    <div class="filter-row">
      <span class="filter-label">Grupo:</span>
      <div class="chips" id="grupo-chips"></div>
    </div>

    <div class="tabs">
      <button class="tab-btn ativo"  id="tbtn-man" onclick="mudarTab('man',this)">Em Andamento <span class="badge" id="bdg-man" style="background:var(--accent)">0</span></button>
      <button class="tab-btn t-fin"  id="tbtn-fin" onclick="mudarTab('fin',this)">Baixadas <span class="badge" id="bdg-fin" style="background:var(--green)">0</span></button>
      <button class="tab-btn t-add"  id="tbtn-add" onclick="mudarTab('add',this)">Adicionadas <span class="badge" id="bdg-add" style="background:var(--red)">0</span></button>
    </div>

    <div id="tab-man" class="tab-content ativo"><div class="tbl-wrap"><table><thead><tr><th>Departamento</th><th>Título da Tarefa</th><th>Vencimento</th><th>Responsável</th><th>Previsão</th><th>Comentário</th></tr></thead><tbody id="tbody-man"></tbody></table></div></div>
    <div id="tab-fin" class="tab-content"><div class="tbl-wrap"><table><thead><tr><th>Departamento</th><th>Título da Tarefa</th><th>Vencimento</th><th>Responsável</th><th>Previsão</th><th>Comentário</th></tr></thead><tbody id="tbody-fin"></tbody></table></div></div>
    <div id="tab-add" class="tab-content"><div class="tbl-wrap"><table><thead><tr><th>Departamento</th><th>Título da Tarefa</th><th>Vencimento</th><th>Responsável</th><th>Previsão</th><th>Comentário</th></tr></thead><tbody id="tbody-add"></tbody></table></div></div>
  </div>
</div>

<script>
const DATA = {{
  man: {js_man},
  fin: {js_fin},
  add: {js_add},
}};
const MODO_COMPARACAO = {'true' if modo_comp else 'false'};
const ARQ_BASE = {json.dumps(nome_base)};
const ARQ_ATUAL = {json.dumps(nome_atual)};
const GERADO = {json.dumps(gerado_em)};

// ── Estado ──────────────────────────────────────────────────────────────────
let filialAtual = null;
let depAtual    = null;
let gruposSel   = new Set();
let abaAtual    = 'man';

// ── Listas calculadas ────────────────────────────────────────────────────────
function getFiliais() {{
  const s = new Set([...DATA.man, ...DATA.fin, ...DATA.add].map(r => r.unidade));
  return [...s].sort();
}}

function getDeps(filial) {{
  const all = [...DATA.man, ...DATA.fin, ...DATA.add].filter(r => r.unidade === filial);
  const s = new Set(all.map(r => r.dep));
  return [...s].sort();
}}

function getGrupos(filial, dep) {{
  const all = [...DATA.man, ...DATA.fin, ...DATA.add]
    .filter(r => r.unidade === filial && r.dep === dep && r.grupo);
  const s = new Set(all.map(r => r.grupo));
  return [...s].sort();
}}

function countDep(arr, filial, dep) {{
  return arr.filter(r => r.unidade === filial && r.dep === dep).length;
}}

// ── Screen transitions ───────────────────────────────────────────────────────
function showScreen(id) {{
  document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
  document.getElementById(id).classList.add('active');
  window.scrollTo(0,0);
}}

function voltarFilial() {{ showScreen('screen-filial'); filialAtual = null; depAtual = null; }}
function voltarDep()    {{ showScreen('screen-dep');    depAtual = null; }}

// ── Screen 1: build filial cards ─────────────────────────────────────────────
function buildFiliais() {{
  const grid = document.getElementById('filial-grid');
  const filiais = getFiliais();
  const labels = {{ 'SP': 'São Paulo', 'Santos': 'Santos', 'RJ': 'Rio de Janeiro', 'GOIAS': 'Goiás' }};
  grid.innerHTML = '';
  filiais.forEach((f, i) => {{
    const total = DATA.man.filter(r => r.unidade === f).length;
    const card = document.createElement('div');
    card.className = 'filial-card';
    card.innerHTML = `
      <div class="fc-num">FILIAL ${{String(i+1).padStart(2,'0')}}</div>
      <div class="fc-name">${{labels[f] || f}}</div>
      <div class="fc-badge"><span></span>${{total}} pendência${{total !== 1 ? 's' : ''}} em aberto</div>
    `;
    card.onclick = () => selecionarFilial(f);
    grid.appendChild(card);
  }});
}}

// ── Screen 2: select filial → show deps ──────────────────────────────────────
function selecionarFilial(filial) {{
  filialAtual = filial;
  gruposSel.clear();
  const labels = {{ 'SP': 'São Paulo', 'Santos': 'Santos', 'RJ': 'Rio de Janeiro', 'GOIAS': 'Goiás' }};
  document.getElementById('dep-screen-title').textContent = labels[filial] || filial;
  document.getElementById('dep-screen-sub').textContent =
    `${{getDeps(filial).length}} departamentos · selecione para ver as pendências`;

  const list = document.getElementById('dep-list');
  list.innerHTML = '';
  getDeps(filial).forEach(dep => {{
    const m = countDep(DATA.man, filial, dep);
    const f = countDep(DATA.fin, filial, dep);
    const a = countDep(DATA.add, filial, dep);
    const el = document.createElement('div');
    el.className = 'dep-item';
    el.innerHTML = `
      <span class="dep-item-name">${{dep}}</span>
      <div class="dep-item-counts">
        <span class="dep-pill dp-man">${{m}}</span>
        ${{MODO_COMPARACAO ? `<span class="dep-pill dp-fin">${{f}}</span><span class="dep-pill dp-add">${{a}}</span>` : ''}}
      </div>
    `;
    el.onclick = () => selecionarDep(dep);
    list.appendChild(el);
  }});

  showScreen('screen-dep');
}}

// ── Screen 3: select dep → show results ──────────────────────────────────────
function selecionarDep(dep) {{
  depAtual = dep;
  gruposSel.clear();
  abaAtual = 'man';

  const labels = {{ 'SP': 'São Paulo', 'Santos': 'Santos', 'RJ': 'Rio de Janeiro', 'GOIAS': 'Goiás' }};
  document.getElementById('rh-title').innerHTML =
    `<em>${{labels[filialAtual] || filialAtual}}</em> · ${{dep}}`;
  document.getElementById('rh-meta-base').textContent  = `Base: ${{ARQ_BASE}}`;
  document.getElementById('rh-meta-atual').textContent = MODO_COMPARACAO
    ? `Atual: ${{ARQ_ATUAL}} · Gerado em: ${{GERADO}}`
    : `Gerado em: ${{GERADO}}`;

  // Build grupo chips
  const grupos = getGrupos(filialAtual, dep);
  const chips = document.getElementById('grupo-chips');
  chips.innerHTML = '';
  const all = document.createElement('span');
  all.className = 'chip sel'; all.textContent = 'Todos';
  all.onclick = () => {{ gruposSel.clear(); chips.querySelectorAll('.chip').forEach(c => c.classList.remove('sel')); all.classList.add('sel'); renderAll(); }};
  chips.appendChild(all);
  grupos.forEach(g => {{
    const c = document.createElement('span');
    c.className = 'chip'; c.textContent = g;
    c.onclick = () => {{
      all.classList.remove('sel');
      c.classList.toggle('sel');
      if (c.classList.contains('sel')) gruposSel.add(g); else gruposSel.delete(g);
      if (gruposSel.size === 0) all.classList.add('sel');
      renderAll();
    }};
    chips.appendChild(c);
  }});

  // Reset tabs
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('ativo'));
  document.getElementById('tbtn-man').classList.add('ativo');
  document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('ativo'));
  document.getElementById('tab-man').classList.add('ativo');
  abaAtual = 'man';

  renderAll();
  showScreen('screen-results');
}}

// ── Render ────────────────────────────────────────────────────────────────────
function filtrar(arr) {{
  return arr.filter(r =>
    r.unidade === filialAtual &&
    r.dep === depAtual &&
    (gruposSel.size === 0 || gruposSel.has(r.grupo))
  );
}}

function renderAll() {{
  const m = filtrar(DATA.man);
  const f = filtrar(DATA.fin);
  const a = filtrar(DATA.add);

  document.getElementById('sc-man').textContent  = m.length;
  document.getElementById('sc-fin').textContent  = f.length;
  document.getElementById('sc-add').textContent  = a.length;
  document.getElementById('bdg-man').textContent = m.length;
  document.getElementById('bdg-fin').textContent = f.length;
  document.getElementById('bdg-add').textContent = a.length;

  renderTabela('man', m);
  renderTabela('fin', f);
  renderTabela('add', a);
}}

function renderTabela(aba, rows) {{
  const tbody = document.getElementById('tbody-' + aba);
  if (!rows.length) {{
    tbody.innerHTML = '<tr><td colspan="6" class="no-data">Nenhuma tarefa encontrada.</td></tr>';
    return;
  }}

  // Agrup por dep (aqui todos são o mesmo dep, mas mantemos para clareza)
  const grupos = {{}};
  rows.forEach(r => {{
    const key = r.dep + (r.novo ? ' (Cliente Novo)' : '');
    if (!grupos[key]) grupos[key] = {{ dep: r.dep, novo: r.novo, rows: [] }};
    grupos[key].rows.push(r);
  }});

  // Dentro de cada dep, agrupar por grupo
  let html = '';
  Object.keys(grupos).sort().forEach(gkey => {{
    const g = grupos[gkey];
    const novoBadge = g.novo ? '<span class="cn-badge">CLIENTE NOVO</span>' : '';
    html += `<tr class="dep-row"><td colspan="6">${{g.dep}}${{novoBadge}}</td></tr>`;

    // sub-agrup por grupo de cliente
    const subGrupos = {{}};
    g.rows.forEach(r => {{
      const sg = r.grupo || '—';
      if (!subGrupos[sg]) subGrupos[sg] = [];
      subGrupos[sg].push(r);
    }});

    Object.keys(subGrupos).sort().forEach(sg => {{
      subGrupos[sg].forEach((r, idx) => {{
        const showCod = gruposSel.size > 0 || true; // sempre mostra cod quando filtro ativo
        const codTag = `<span class="cod-tag">#${{r.cod}}</span>`;
        const comtHtml = (() => {{
          let out = '';
          if (r.dataComt) out += `<div class="comt-date">📅 ${{r.dataComt}}</div>`;
          if (r.comt && r.comt.trim()) {{
            out += `<button class="comt-btn" onclick="toggleComt(this)">Ver comentário</button>`;
            out += `<div class="comt-full">${{r.comt.replace(/</g,'&lt;').replace(/>/g,'&gt;')}}</div>`;
          }} else if (!r.dataComt) {{
            out = '<span style="color:var(--muted);font-size:12px">—</span>';
          }}
          return out;
        }})();
        // Show grupo label on first row of subgroup
        const grupoLabel = idx === 0
          ? `<div style="font-size:11px;color:var(--amber);font-family:var(--mono);font-weight:600;margin-bottom:3px;letter-spacing:.04em">${{sg}}</div>`
          : '';
        html += `<tr>
          <td>${{grupoLabel}}${{codTag}}</td>
          <td>${{r.tit}}</td>
          <td style="white-space:nowrap">${{r.venc}}</td>
          <td class="resp-name">${{r.resp || '—'}}</td>
          <td style="white-space:nowrap">${{r.prev}}</td>
          <td>${{comtHtml}}</td>
        </tr>`;
      }});
    }});
  }});
  tbody.innerHTML = html;
}}

function mudarTab(aba, btn) {{
  document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('ativo'));
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('ativo'));
  document.getElementById('tab-' + aba).classList.add('ativo');
  btn.classList.add('ativo');
  abaAtual = aba;
}}

function toggleComt(btn) {{
  const open = btn.nextElementSibling.classList.toggle('open');
  btn.textContent = open ? 'Ocultar' : 'Ver comentário';
}}

// ── Init ─────────────────────────────────────────────────────────────────────
buildFiliais();
</script>
</body>
</html>"""
    return html


# ── Entry point ───────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Gerador de Relatório HTML – MG Contécnica v4.0")
    parser.add_argument("base",    help="Arquivo Excel base (1º do dia / semana)")
    parser.add_argument("atual",   nargs="?", default=None, help="Arquivo Excel atual (opcional, para comparação)")
    parser.add_argument("-o", "--output", default=None, help="Caminho de saída do HTML")
    args = parser.parse_args()

    print(f"[*] Lendo base: {args.base}")
    df_base = ler_relatorio(args.base)
    print(f"    {len(df_base)} tarefas carregadas.")

    if args.atual:
        print(f"[*] Lendo atual: {args.atual}")
        df_atual = ler_relatorio(args.atual)
        print(f"    {len(df_atual)} tarefas carregadas.")
        em_andamento, finalizadas, adicionadas = comparar(df_base, df_atual)
    else:
        print("[*] Modo visualização única (sem comparação).")
        em_andamento = df_base.copy()
        finalizadas  = pd.DataFrame(columns=df_base.columns)
        adicionadas  = pd.DataFrame(columns=df_base.columns)
        args.atual   = args.base

    print(f"[*] Em andamento: {len(em_andamento)} | Baixadas: {len(finalizadas)} | Adicionadas: {len(adicionadas)}")

    html = gerar_html(args.base, args.atual, em_andamento, finalizadas, adicionadas)

    if args.output:
        out = args.output
    else:
        ts  = datetime.now().strftime("%d-%m-%y_%H%M%S")
        out = f"Relatorio_Pendencias_{ts}.html"

    with open(out, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"[✓] Relatório salvo: {out}")
    return out


if __name__ == "__main__":
    main()
