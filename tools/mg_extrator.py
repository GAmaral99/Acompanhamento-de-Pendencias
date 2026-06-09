# ============================================================
# MG Contécnica – Extrator de Relatório de Pendências
# Versão 5.2
# - Terminal visual com Rich (instalação automática)
# - Sem "pressione ENTER" no final
# - Opção "Todas as Filiais" (multisselect no select2)
# ============================================================

import os
import sys
import time
import subprocess

from datetime import datetime
import re
import argparse

# ============================================================
# GARANTIR DEPENDÊNCIAS
# ============================================================

def _instalar(pacote):
    subprocess.check_call(
        [sys.executable, "-m", "pip", "install", pacote, "--quiet"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )

try:
    from rich.console import Console
    from rich.panel import Panel
    from rich.table import Table
    from rich.text import Text
    from rich import box
    from rich.prompt import Prompt
    from rich.progress import Progress, SpinnerColumn, TextColumn, TimeElapsedColumn
    from rich.align import Align
except ImportError:
    print("Instalando dependência 'rich'...")
    _instalar("rich")
    from rich.console import Console
    from rich.panel import Panel
    from rich.table import Table
    from rich.text import Text
    from rich import box
    from rich.prompt import Prompt
    from rich.progress import Progress, SpinnerColumn, TextColumn, TimeElapsedColumn
    from rich.align import Align

console = Console()

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from webdriver_manager.chrome import ChromeDriverManager

# ============================================================
# CONFIGURAÇÕES
# ============================================================

USUARIO = "gamaral"
SENHA   = "mg@2026*"

PASTA_BASE      = r"C:\Users\gamaral\Desktop\Relatorios Diarios"
PASTA_REPO_DATA = r"C:\Users\gamaral\Desktop\Python\repo\data"

TIMEOUT = 40

URL_LOGIN = "https://aplicativo.mgcontecnica.com.br/#/login"

FILIAIS = {
    "1": ("SP",     "SP",     "SP"),
    "2": ("Santos", "SANTOS", "Santos"),
    "3": ("RJ",     "RJ",     "RJ"),
    "4": ("Goiás",  "GOIAS",  "Goias"),
}

OPCAO_TODAS = "0"

# ============================================================
# UTILITÁRIOS
# ============================================================

def agora():
    return datetime.now().strftime("%d/%m/%Y %H:%M:%S")

# ============================================================
# AGUARDAR DOWNLOAD
# ============================================================

def snapshot_xlsx(pasta: str) -> set:
    """Retorna o conjunto de caminhos absolutos dos .xlsx existentes na pasta."""
    try:
        return {
            os.path.join(pasta, arq)
            for arq in os.listdir(pasta)
            if arq.lower().endswith(".xlsx")
        }
    except FileNotFoundError:
        return set()

def aguardar_download(pasta, antes: set, progress, task, timeout=600):
    """
    Aguarda um arquivo .xlsx NOVO aparecer em `pasta` (não estava em `antes`)
    e ficar estável em tamanho por 3 iterações consecutivas.
    """
    inicio         = time.time()
    ultimo_arquivo = None
    ultimo_tamanho = -1
    tempo_estavel  = 0

    while time.time() - inicio < timeout:

        # Arquivos temporários do Chrome ainda em download
        arquivos_temp = [
            arq for arq in os.listdir(pasta)
            if arq.endswith(".crdownload") or arq.endswith(".tmp")
        ]
        if arquivos_temp:
            progress.update(task, description="[cyan]Arquivo temporário detectado, aguardando...")
            time.sleep(1)
            continue

        # Apenas arquivos .xlsx que NÃO existiam antes do clique
        novos = [
            os.path.join(pasta, arq)
            for arq in os.listdir(pasta)
            if arq.lower().endswith(".xlsx")
            and os.path.join(pasta, arq) not in antes
        ]

        if not novos:
            progress.update(task, description="[cyan]Aguardando arquivo aparecer na pasta...")
            time.sleep(1)
            continue

        arquivo_recente = max(novos, key=os.path.getctime)

        try:
            tamanho_atual = os.path.getsize(arquivo_recente)

            if ultimo_arquivo == arquivo_recente and ultimo_tamanho == tamanho_atual:
                tempo_estavel += 1
                progress.update(task, description=f"[cyan]Verificando estabilidade ({tempo_estavel}/3)...")
            else:
                tempo_estavel = 0
                progress.update(task, description="[cyan]Arquivo crescendo, aguardando...")

            ultimo_arquivo = arquivo_recente
            ultimo_tamanho = tamanho_atual

            if tempo_estavel >= 3:
                return arquivo_recente

        except Exception:
            pass

        time.sleep(1)

    return None

# ============================================================
# MENU
# ============================================================

def perguntar_filial():
    console.print()
    console.rule("[bold cyan]Selecionar Filial[/bold cyan]")
    console.print()

    table = Table(box=box.ROUNDED, border_style="dim", show_header=False, padding=(0, 2))
    table.add_column("Key",    style="bold cyan",  width=5)
    table.add_column("Filial", style="bold white")

    # Opção especial "Todas"
    table.add_row(f"[{OPCAO_TODAS}]", "[bold magenta]Todas as Filiais[/bold magenta]")

    for k, (nome, _, _) in FILIAIS.items():
        table.add_row(f"[{k}]", nome)

    console.print(Align.center(table))
    console.print()

    opcoes_validas = set(FILIAIS.keys()) | {OPCAO_TODAS}

    while True:
        opcao = Prompt.ask("[bold cyan]  Digite a opção[/bold cyan]").strip()
        if opcao in opcoes_validas:
            if opcao == OPCAO_TODAS:
                console.print(f"  [green]✔ Selecionado:[/green] [bold magenta]Todas as Filiais[/bold magenta]")
                return OPCAO_TODAS, None, "Todas"
            console.print(f"  [green]✔ Filial selecionada:[/green] [bold]{FILIAIS[opcao][0]}[/bold]")
            return FILIAIS[opcao]
        console.print("  [red]Opção inválida. Tente novamente.[/red]")


def ultimo_dia_mes(ref: datetime = None) -> str:
    """Retorna o último dia do mês da data de referência (padrão: hoje) em DD/MM/YYYY."""
    import calendar
    ref = ref or datetime.now()
    ultimo = calendar.monthrange(ref.year, ref.month)[1]
    return datetime(ref.year, ref.month, ultimo).strftime("%d/%m/%Y")

def perguntar_data():
    padrao = ultimo_dia_mes()
    console.print()
    console.print(
        f"  [dim]Data final padrão: último dia do mês atual → [bold cyan]{padrao}[/bold cyan][/dim]"
    )
    data = Prompt.ask(
        "[bold cyan]  Data final[/bold cyan] [dim](Enter = último dia do mês)[/dim]",
        default=padrao,
    ).strip()
    return data if data else padrao

# ============================================================
# DRIVER (headless)
# ============================================================

def iniciar_driver(pasta_download):

    os.makedirs(pasta_download, exist_ok=True)

    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    prefs = {
        "download.default_directory":                                  os.path.abspath(pasta_download),
        "download.prompt_for_download":                                False,
        "download.directory_upgrade":                                  True,
        "safebrowsing.enabled":                                        True,
        "safebrowsing.disable_download_protection":                    True,
        "profile.default_content_setting_values.automatic_downloads":  1,
    }
    options.add_experimental_option("prefs", prefs)

    service = Service(ChromeDriverManager().install())
    driver  = webdriver.Chrome(service=service, options=options)

    driver.execute_cdp_cmd(
        "Page.setDownloadBehavior",
        {"behavior": "allow", "downloadPath": os.path.abspath(pasta_download)},
    )

    return driver

# ============================================================
# LOGIN
# ============================================================

def fazer_login(driver, progress, task):

    progress.update(task, description="[cyan]Abrindo página de login...")
    driver.get(URL_LOGIN)
    wait = WebDriverWait(driver, TIMEOUT)

    progress.update(task, description="[cyan]Preenchendo credenciais...")
    campo_usuario = wait.until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, "input[formcontrolname='usuario'], input[id='usuario']")
        )
    )
    campo_usuario.clear()
    campo_usuario.send_keys(USUARIO)

    campo_senha = wait.until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, "input#senha[formcontrolname='senha']")
        )
    )
    campo_senha.clear()
    campo_senha.send_keys(SENHA)

    driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()

    progress.update(task, description="[cyan]Aguardando autenticação...")
    wait.until(EC.url_contains("/home"))

# ============================================================
# NAVEGAÇÃO
# ============================================================

def navegar_para_relatorio(driver, progress, task):

    wait = WebDriverWait(driver, TIMEOUT)

    progress.update(task, description="[cyan]Abrindo MG Controle...")
    mg_controle = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//div[@title='http://intranetmg/Aplicativos/Geral/Controle/']")
        )
    )
    driver.execute_script("arguments[0].click();", mg_controle)
    time.sleep(4)

    if len(driver.window_handles) > 1:
        driver.switch_to.window(driver.window_handles[-1])

    progress.update(task, description="[cyan]Acessando relatórios...")
    btn_relatorios = wait.until(
        EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_RelatorioLinkButton"))
    )
    driver.execute_script("arguments[0].click();", btn_relatorios)
    time.sleep(2)

    progress.update(task, description="[cyan]Abrindo relatório de pendências...")
    btn_pendencias = wait.until(
        EC.element_to_be_clickable(
            (By.ID, "ContentPlaceHolder1_lnkRelatorioGeralPendenciasPorVencimento")
        )
    )
    driver.execute_script("arguments[0].click();", btn_pendencias)
    time.sleep(4)

    if len(driver.window_handles) > 1:
        driver.switch_to.window(driver.window_handles[-1])

# ============================================================
# SELECIONAR FILIAIS NO SELECT2
# ============================================================

def selecionar_filiais_select2(driver, lista_filiais_texto, progress, task):
    """
    Recebe uma lista de textos de filial (ex: ["SP", "SANTOS", "RJ", "GOIAS"])
    e seleciona cada uma no campo multisselect select2.
    """
    wait = WebDriverWait(driver, TIMEOUT)

    for filial_texto in lista_filiais_texto:
        progress.update(task, description=f"[cyan]Selecionando filial: {filial_texto}...")

        # Clica no container do select2 para abrir/manter aberto
        select2 = wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, ".select2-container"))
        )
        driver.execute_script("arguments[0].click();", select2)
        time.sleep(1)

        # Digita o nome da filial no campo de busca
        campo_busca = wait.until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, ".select2-input"))
        )
        campo_busca.clear()
        campo_busca.send_keys(filial_texto)
        time.sleep(2)

        # Confirma a seleção com ENTER
        campo_busca.send_keys(Keys.ENTER)
        time.sleep(1)

    progress.update(task, description="[cyan]Todas as filiais selecionadas no filtro.")
    time.sleep(1)

# ============================================================
# EXTRAÇÃO
# ============================================================

def preencher_e_extrair(driver, filiais_para_selecionar, data_str, pasta_download, progress, task):
    """
    filiais_para_selecionar: lista de strings de texto (ex: ["SP"] ou ["SP","SANTOS","RJ","GOIAS"])
    """
    wait = WebDriverWait(driver, TIMEOUT)

    # Seleciona as filiais no multisselect
    selecionar_filiais_select2(driver, filiais_para_selecionar, progress, task)

    progress.update(task, description=f"[cyan]Preenchendo data: {data_str}...")
    campo_data = wait.until(
        EC.presence_of_element_located((By.ID, "dtVencimentoFinal"))
    )
    driver.execute_script("arguments[0].removeAttribute('readonly');", campo_data)
    driver.execute_script(f"arguments[0].value='{data_str}';",         campo_data)
    time.sleep(2)

    # Snapshot ANTES de clicar — detecta só arquivos realmente novos
    os.makedirs(pasta_download, exist_ok=True)
    antes = snapshot_xlsx(pasta_download)

    progress.update(task, description="[cyan]Disparando extração...")
    btn_extrair = wait.until(
        EC.element_to_be_clickable((By.ID, "btnExtrairRelatorio"))
    )
    driver.execute_script("arguments[0].click();", btn_extrair)

    return aguardar_download(pasta_download, antes, progress, task)

# ============================================================
# ORGANIZAR POR SEMANA
# ============================================================

def pasta_da_semana(data_ref: datetime, pasta_base: str) -> str:
    """
    Retorna o caminho da pasta no formato:
    PASTA_BASE/2026-06 Semana 1/
    """
    ano  = data_ref.year
    mes  = data_ref.month
    dia  = data_ref.day

    # Número da semana dentro do mês (1 a 5)
    semana_no_mes = (dia - 1) // 7 + 1

    nome_mes = data_ref.strftime("%Y-%m")
    nome_pasta = f"{nome_mes} Semana {semana_no_mes}"
    return os.path.join(pasta_base, nome_pasta)

# ============================================================
# RENOMEAR ARQUIVO
# ============================================================

def renomear_arquivo(caminho_original, label_filial, pasta_download):
    timestamp    = datetime.now().strftime("%d-%m-%y %H%M")
    novo_nome    = f"{label_filial} - Relatorio de Pendencias - {timestamp}.xlsx"

    # Se for "Todas", organiza na subpasta da semana
    if label_filial == "Todas":
        pasta_destino = pasta_da_semana(datetime.now(), pasta_download)
        os.makedirs(pasta_destino, exist_ok=True)
    else:
        pasta_destino = pasta_download

    novo_caminho = os.path.join(pasta_destino, novo_nome)
    os.rename(caminho_original, novo_caminho)
    return novo_caminho, novo_nome, pasta_destino

# ============================================================
# MIGRAR ARQUIVOS EXISTENTES PARA PASTAS DE SEMANA
# ============================================================

def migrar_arquivos_existentes():
    """
    Organiza os .xlsx que estão soltos na pasta data/
    para as subpastas de semana correspondentes.
    Ignora subpastas já existentes (coordenadores, historico, semanas).
    """
    pastas_ignorar = {"coordenadores", "historico"}

    arquivos = [
        f for f in os.listdir(PASTA_REPO_DATA)
        if f.lower().endswith(".xlsx")
        and os.path.isfile(os.path.join(PASTA_REPO_DATA, f))
    ]

    if not arquivos:
        return 0

    movidos = 0
    console.print()
    console.rule("[bold yellow]Migração de arquivos existentes[/bold yellow]")
    console.print()

    for nome in arquivos:
        caminho = os.path.join(PASTA_REPO_DATA, nome)

        # Tenta extrair a data do nome do arquivo
        padroes = [
            r'(\d{2})-(\d{2})-(\d{2})\s(\d{4})',  # DD-MM-YY HHMM
            r'(\d{2})-(\d{2})-(\d{2})_(\d{4})',   # DD-MM-YY_HHMM
            r'(\d{2})-(\d{2})-(\d{4})\s(\d{4})',  # DD-MM-YYYY HHMM
            r'(\d{2})-(\d{2})-(\d{4})_(\d{4})',   # DD-MM-YYYY_HHMM
        ]

        data_arquivo = None
        for p in padroes:
            m = re.search(p, nome)
            if m:
                g = m.groups()
                try:
                    if len(g[2]) == 2:
                        data_arquivo = datetime(
                            2000 + int(g[2]), int(g[1]), int(g[0]),
                            int(g[3][:2]), int(g[3][2:])
                        )
                    else:
                        data_arquivo = datetime(
                            int(g[2]), int(g[1]), int(g[0]),
                            int(g[3][:2]), int(g[3][2:])
                        )
                    break
                except ValueError:
                    continue

        # Se não achou data no nome, usa a data de modificação do arquivo
        if not data_arquivo:
            data_arquivo = datetime.fromtimestamp(os.path.getmtime(caminho))

        # Calcula a pasta de destino
        pasta_destino = pasta_da_semana(data_arquivo, PASTA_REPO_DATA)
        os.makedirs(pasta_destino, exist_ok=True)

        destino_final = os.path.join(pasta_destino, nome)
        os.rename(caminho, destino_final)

        semana_label = os.path.basename(pasta_destino)
        console.print(f"  [green]✔[/green] [dim]{nome}[/dim] → [cyan]{semana_label}[/cyan]")
        movidos += 1

    if movidos:
        console.print()
        console.print(f"  [bold green]{movidos} arquivo(s) organizado(s) com sucesso![/bold green]")

    return movidos

# ============================================================
# EXECUÇÃO
# ============================================================

def executar(auto=False):

    console.print()
    console.print(Panel(
        Align.center(
            Text.from_markup(
                "[bold cyan]MG CONTÉCNICA[/bold cyan]\n"
                "[dim]Extrator de Relatório de Pendências[/dim]\n"
                "[bold white]v5.2[/bold white]"
            )
        ),
        border_style="cyan",
        padding=(1, 6),
    ))

    # Migra arquivos soltos para pastas de semana
    migrar_arquivos_existentes()

    if auto:
        resultado_filial = (OPCAO_TODAS, None, "Todas")
        data_str = ultimo_dia_mes()
        console.print(
            f"  [dim][--auto] Filial: [bold magenta]Todas[/bold magenta] | "
            f"Data: [bold cyan]{data_str}[/bold cyan][/dim]"
        )
    else:
        resultado_filial = perguntar_filial()
        data_str = perguntar_data()

    # Define quais textos de filial serão selecionados e o label do arquivo
    if resultado_filial[0] == OPCAO_TODAS:
        filiais_para_selecionar = [v[1] for v in FILIAIS.values()]  # ["SP","SANTOS","RJ","GOIAS"]
        label_filial            = "Todas"
        pasta_filial_final      = PASTA_REPO_DATA  # salva direto no repositorio local (data/)
    else:
        # perguntar_filial() retornou a tupla (nome, texto_select2, pasta_subdir)
        label_filial, filial_texto_unico, subdir = resultado_filial
        filiais_para_selecionar = [filial_texto_unico]
        pasta_filial_final      = os.path.join(PASTA_BASE, subdir)

    pasta_download = pasta_filial_final

    console.print()
    console.rule("[bold cyan]Executando[/bold cyan]")
    console.print()

    driver    = None
    novo_nome = None

    with Progress(
        SpinnerColumn(style="cyan"),
        TextColumn("[progress.description]{task.description}"),
        TimeElapsedColumn(),
        console=console,
        transient=True,
    ) as progress:
        task = progress.add_task("[cyan]Iniciando...", total=None)

        try:
            progress.update(task, description="[cyan]Iniciando Chrome (headless)...")
            driver = iniciar_driver(pasta_download)

            fazer_login(driver, progress, task)
            navegar_para_relatorio(driver, progress, task)

            arquivo_bruto = preencher_e_extrair(
                driver, filiais_para_selecionar, data_str, pasta_download, progress, task
            )

            if arquivo_bruto:
                progress.update(task, description="[cyan]Renomeando arquivo...")
                _, novo_nome, pasta_download = renomear_arquivo(arquivo_bruto, label_filial, pasta_download)

        except Exception as erro:
            progress.stop()
            console.print()
            console.print(Panel(
                f"[red]{erro}[/red]",
                title="[red]Erro durante a execução[/red]",
                border_style="red",
            ))
            console.print_exception()

        finally:
            if driver:
                driver.quit()

    # ── Resultado ────────────────────────────────────────────
    console.print()

    if novo_nome:
        console.rule("[bold green]Concluído[/bold green]")
        console.print()

        resultado = Table(box=box.ROUNDED, border_style="dim", show_header=False, padding=(0, 2))
        resultado.add_column("Label", style="dim",        width=18)
        resultado.add_column("Valor", style="bold white")
        resultado.add_row("Filial",        label_filial)
        resultado.add_row("Data extraída", data_str)
        resultado.add_row("Arquivo",       novo_nome)
        resultado.add_row("Pasta",         pasta_download)
        console.print(Align.center(resultado))
        console.print()
        console.print(Panel(
            f"[green]✔[/green] [bold white]{novo_nome}[/bold white]",
            title="[green]Arquivo salvo[/green]",
            border_style="green",
            padding=(0, 2),
        ))

        # ── Auto git push quando opção "Todas" ──────────────────
        if label_filial == "Todas":
            console.print()
            console.rule("[bold cyan]Publicando no GitHub[/bold cyan]")
            console.print()
            repo_dir = r"C:\Users\gamaral\Desktop\Python\repo"
            try:
                with Progress(
                    SpinnerColumn(style="cyan"),
                    TextColumn("[progress.description]{task.description}"),
                    console=console,
                    transient=True,
                ) as pg:
                    t2 = pg.add_task("[cyan]Sincronizando com GitHub...", total=None)
                    pg.update(t2, description="[cyan]Baixando atualizações remotas (pull)...")
                    subprocess.run(["git", "-C", repo_dir, "pull", "--rebase", "--autostash"],
                                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                    pg.update(t2, description="[cyan]Adicionando arquivo...")
                    subprocess.run(["git", "-C", repo_dir, "add", "data/"], check=True,
                                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                    pg.update(t2, description="[cyan]Criando commit...")
                    msg = f"data: {novo_nome}"
                    subprocess.run(["git", "-C", repo_dir, "commit", "-m", msg], check=True,
                                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                    pg.update(t2, description="[cyan]Fazendo push...")
                    subprocess.run(["git", "-C", repo_dir, "push"], check=True,
                                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

                console.print(Panel(
                    "[green]✔[/green] Arquivo enviado para o GitHub!\n"
                    "[dim]O relatório online será atualizado em ~1 minuto.[/dim]\n"
                    "[cyan]https://gamaral99.github.io/Acompanhamento-de-Pendencias/[/cyan]",
                    title="[green]GitHub Pages[/green]",
                    border_style="green",
                    padding=(0, 2),
                ))
            except subprocess.CalledProcessError as e:
                console.print(Panel(
                    f"[yellow]Download salvo, mas o push falhou:[/yellow]\n[red]{e}[/red]\n\n"
                    "[dim]Faça o push manualmente via terminal na pasta do repo.[/dim]",
                    title="[yellow]Aviso[/yellow]",
                    border_style="yellow",
                ))
    else:
        console.rule("[bold yellow]Concluído com aviso[/bold yellow]")
        console.print()
        console.print(Panel(
            "[yellow]Download não detectado dentro do tempo limite.[/yellow]",
            title="[yellow]Aviso[/yellow]",
            border_style="yellow",
        ))

    console.print()

# ============================================================
# MAIN
# ============================================================

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Extrator de Relatório de Pendências MG Contécnica"
    )
    parser.add_argument(
        "--auto",
        action="store_true",
        help="Execução automática: seleciona Todas as Filiais e usa a data padrão (último dia do mês) sem prompts.",
    )
    args = parser.parse_args()
    executar(auto=args.auto)