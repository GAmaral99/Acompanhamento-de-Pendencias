#!/usr/bin/env python3
"""
MG Contécnica – Gerador de Relatório HTML Consolidado
Versão 6.1
"""

import os, re, sys, json, argparse
import pandas as pd
import base64 as _b64
from datetime import datetime
from pathlib import Path

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
    from datetime import timezone, timedelta
    tz_brasilia = timezone(timedelta(hours=-3))
    gerado_em   = datetime.now(tz_brasilia).strftime("%d/%m/%Y às %H:%M")
    modo_comp  = arquivo_atual and arquivo_atual != arquivo_base

    js_man = df_to_js(em_andamento)
    js_fin = df_to_js(finalizadas)
    js_add = df_to_js(reabertas)
    total_man = len(em_andamento)
    total_fin = len(finalizadas)
    total_add = len(reabertas)

    js_placares = json.dumps(placares or {}, ensure_ascii=False)

    # Lê o histórico para embutir no HTML
    historico_b64 = ""
    historico_path = Path("data/historico/historico_baixas.xlsx")
    if historico_path.exists():
        with open(historico_path, "rb") as _f:
            historico_b64 = _b64.b64encode(_f.read()).decode()
        print(f"[*] Histórico embutido no HTML ({len(historico_b64)//1024} KB base64)")
    else:
        print("[!] historico_baixas.xlsx não encontrado. PDF de histórico indisponível.")

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
<link rel="stylesheet" href="static/app.css">
</head>
<body class="light">

<button id="theme-toggle" onclick="toggleTema()" title="Alternar tema"></button>

<!-- ══ SCREEN 1: FILIAL ══ -->
<div class="screen active" id="screen-filial">
  <div class="logo-bar">
    <img class="logo-img" src="static/logo_mg.png" alt="MG Contécnica">
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
    <img class="logo-img" src="static/logo_mg.png" alt="MG Contécnica">
    <div class="logo-text"><strong>MG Contécnica</strong>Selecione o Departamento</div>
  </div>
  <div style="flex:1;padding:40px 48px">
    <button class="back-btn" onclick="voltarFilial()">← Voltar</button>
    <div class="screen-header">
      <div class="stag">Filial selecionada</div>
      <h2 id="dep-screen-title">—</h2>
      <p id="dep-screen-sub"></p>
    </div>
    <div class="dep-legend">
      <span class="legend-item lg-man"><i class="lg-dot"></i>Pendentes</span>
      {'<span class="legend-item lg-fin"><i class="lg-dot"></i>Baixadas</span><span class="legend-item lg-add"><i class="lg-dot"></i>Reabertas</span>' if modo_comp else ''}
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
      <div class="stat-card man"><div class="sc-label">Pendentes</div><div class="sc-num" id="sc-man">0</div><div class="sc-sub">tarefas abertas</div></div>
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
      <div style="align-self:flex-end">
        <div class="filter-label">&nbsp;</div>
        <button id="btn-vencidas" onclick="toggleVencidas()"
           style="background:var(--surface);border:1px solid var(--border);color:var(--muted);font-family:var(--font);font-size:13px;font-weight:600;padding:9px 18px;border-radius:8px;cursor:pointer;transition:all .18s;white-space:nowrap">
          ⚠ Somente Vencidas
         </button>
        </div>
    </div>

    <!-- Placares de Baixas -->
    <div class="placares-section" id="placares-section" style="display:none">
      <div class="placares-title">Placar de Baixas</div>
      <div class="placares-export-row">
        <button class="placar-export-btn" onclick="abrirModalPDF()" style="border-color:var(--red);color:var(--red)">📄 Relatório PDF</button>
      </div>

      <div class="placares-grid" id="placares-grid">
        <div class="placar-card pc-dia">
          <div class="pc-label">Baixas no Dia</div>
          <div class="pc-num" id="pc-dia-num">0</div>
          <div class="pc-sub">tarefas finalizadas hoje</div>
          <div class="pc-deps" id="pc-dia-deps"></div>
        </div>
        <div class="placar-card pc-sem">
          <div class="pc-label">Baixas na Semana</div>
          <div class="pc-num" id="pc-sem-num">0</div>
          <div class="pc-sub">tarefas finalizadas esta semana</div>
          <div class="pc-deps" id="pc-sem-deps"></div>
        </div>
        <div class="placar-card pc-mes">
          <div class="pc-label">Baixas no Mês</div>
          <div class="pc-num" id="pc-mes-num">0</div>
          <div class="pc-sub">tarefas finalizadas este mês</div>
          <div class="pc-deps" id="pc-mes-deps"></div>
        </div>
      </div>

<!-- Modal PDF -->
<div class="pdf-modal-overlay" id="pdf-modal-overlay" onclick="fecharModalPDF(event)">
  <div class="pdf-modal">
    <div>
      <div class="pdf-modal-title">📄 Relatório de Baixas</div>
      <div class="pdf-modal-sub">Escolha o período do relatório</div>
    </div>
    <div class="pdf-periodo-grid">
      <button class="pdf-periodo-btn" onclick="gerarPDF('sem')">
        <span class="ppb-icon">📅</span>
        <div class="ppb-label">Semana Atual</div>
        <div class="ppb-sub">Seg → hoje</div>
      </button>
      <button class="pdf-periodo-btn" onclick="gerarPDF('mes')">
        <span class="ppb-icon">🗓️</span>
        <div class="ppb-label">Mês Atual</div>
        <div class="ppb-sub">01 → hoje</div>
      </button>
    </div>
    <button class="pdf-modal-cancel" onclick="fecharModalPDF()">Cancelar</button>
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
      <button class="tab-btn ativo"  id="tbtn-man" onclick="mudarTab('man',this)">Pendentes <span class="badge" id="bdg-man" style="background:var(--red)">0</span></button>
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
const HIST_B64 = {json.dumps(historico_b64 if 'historico_b64' in dir() else '')};
</script>
<script src="static/app.js"></script>
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
    print(f"[OK] Salvo: {out}")

if __name__=="__main__":
    main()