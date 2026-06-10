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
<body>

<button id="theme-toggle" onclick="toggleTema()" title="Alternar tema"></button>

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