#!/usr/bin/env python3
"""
mg_extrator_auto.py
-------------------
Versão automática do extrator para rodar no GitHub Actions.
- Sem interface Rich (terminal simples)
- Credenciais via variáveis de ambiente
- Chrome headless configurado para Linux
- Salva na pasta de semana correta
"""

import os
import sys
import time
import subprocess
from datetime import datetime
from pathlib import Path
import re

# ============================================================
# CONFIGURAÇÕES
# ============================================================

USUARIO = os.environ.get("MG_USUARIO", "")
SENHA   = os.environ.get("MG_SENHA", "")

PASTA_REPO_DATA = Path("data")
TIMEOUT = 60
URL_LOGIN = "https://aplicativo.mgcontecnica.com.br/#/login"

FILIAIS_TEXTO = ["SP", "SANTOS", "RJ", "GOIAS"]

# ============================================================
# HELPERS
# ============================================================

def log(msg):
    agora = datetime.now().strftime("%H:%M:%S")
    print(f"[{agora}] {msg}", flush=True)

def pasta_da_semana(data_ref: datetime, pasta_base: Path) -> Path:
    semana_no_mes = (data_ref.day - 1) // 7 + 1
    nome_pasta = f"{data_ref.strftime('%Y-%m')} Semana {semana_no_mes}"
    return pasta_base / nome_pasta

def snapshot_xlsx(pasta: Path) -> set:
    try:
        return {str(p) for p in pasta.glob("*.xlsx")}
    except FileNotFoundError:
        return set()

def aguardar_download(pasta: Path, antes: set, timeout=120):
    inicio = time.time()
    ultimo_arquivo = None
    ultimo_tamanho = -1
    tempo_estavel  = 0

    while time.time() - inicio < timeout:
        # Arquivos temporários ainda em download
        temps = list(pasta.glob("*.crdownload")) + list(pasta.glob("*.tmp"))
        if temps:
            log("Arquivo temporário detectado, aguardando...")
            time.sleep(1)
            continue

        novos = [
            str(p) for p in pasta.glob("*.xlsx")
            if str(p) not in antes
        ]

        if not novos:
            time.sleep(1)
            continue

        arquivo = max(novos, key=os.path.getctime)

        try:
            tamanho = os.path.getsize(arquivo)
            if ultimo_arquivo == arquivo and ultimo_tamanho == tamanho:
                tempo_estavel += 1
                log(f"Verificando estabilidade ({tempo_estavel}/3)...")
            else:
                tempo_estavel = 0
                log("Arquivo crescendo, aguardando...")
            ultimo_arquivo = arquivo
            ultimo_tamanho = tamanho
            if tempo_estavel >= 3:
                return arquivo
        except Exception:
            pass

        time.sleep(1)

    return None

# ============================================================
# DRIVER
# ============================================================

def iniciar_driver(pasta_download: Path):
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.chrome.service import Service
    from webdriver_manager.chrome import ChromeDriverManager

    pasta_download.mkdir(parents=True, exist_ok=True)

    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-extensions")

    prefs = {
        "download.default_directory":             str(pasta_download.resolve()),
        "download.prompt_for_download":           False,
        "download.directory_upgrade":             True,
        "safebrowsing.enabled":                   True,
        "safebrowsing.disable_download_protection": True,
        "profile.default_content_setting_values.automatic_downloads": 1,
    }
    options.add_experimental_option("prefs", prefs)

    service = Service(ChromeDriverManager().install())
    driver  = webdriver.Chrome(service=service, options=options)

    driver.execute_cdp_cmd(
        "Page.setDownloadBehavior",
        {"behavior": "allow", "downloadPath": str(pasta_download.resolve())},
    )

    return driver

# ============================================================
# LOGIN
# ============================================================

def fazer_login(driver):
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    log("Abrindo página de login...")
    driver.get(URL_LOGIN)
    wait = WebDriverWait(driver, TIMEOUT)

    log("Preenchendo credenciais...")
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

    log("Aguardando autenticação...")
    wait.until(EC.url_contains("/home"))
    log("Login efetuado com sucesso!")

# ============================================================
# NAVEGAÇÃO
# ============================================================

def navegar_para_relatorio(driver):
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    wait = WebDriverWait(driver, TIMEOUT)

    log("Abrindo MG Controle...")
    mg_controle = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//div[@title='http://intranetmg/Aplicativos/Geral/Controle/']")
        )
    )
    driver.execute_script("arguments[0].click();", mg_controle)
    time.sleep(4)

    if len(driver.window_handles) > 1:
        driver.switch_to.window(driver.window_handles[-1])

    log("Acessando relatórios...")
    btn_relatorios = wait.until(
        EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_RelatorioLinkButton"))
    )
    driver.execute_script("arguments[0].click();", btn_relatorios)
    time.sleep(2)

    log("Abrindo relatório de pendências...")
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
# SELECIONAR FILIAIS
# ============================================================

def selecionar_filiais(driver, filiais):
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    wait = WebDriverWait(driver, TIMEOUT)

    for filial in filiais:
        log(f"Selecionando filial: {filial}...")
        select2 = wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, ".select2-container"))
        )
        driver.execute_script("arguments[0].click();", select2)
        time.sleep(1)

        campo = wait.until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, ".select2-input"))
        )
        campo.clear()
        campo.send_keys(filial)
        time.sleep(2)
        campo.send_keys(Keys.ENTER)
        time.sleep(1)

    log("Todas as filiais selecionadas.")

# ============================================================
# EXTRAIR
# ============================================================

def extrair(driver, pasta_download: Path):
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    import calendar

    wait = WebDriverWait(driver, TIMEOUT)

    # Data final = último dia do mês atual
    hoje = datetime.now()
    ultimo_dia = calendar.monthrange(hoje.year, hoje.month)[1]
    data_str = datetime(hoje.year, hoje.month, ultimo_dia).strftime("%d/%m/%Y")

    selecionar_filiais(driver, FILIAIS_TEXTO)

    log(f"Preenchendo data final: {data_str}...")
    campo_data = wait.until(
        EC.presence_of_element_located((By.ID, "dtVencimentoFinal"))
    )
    driver.execute_script("arguments[0].removeAttribute('readonly');", campo_data)
    driver.execute_script(f"arguments[0].value='{data_str}';", campo_data)
    time.sleep(2)

    antes = snapshot_xlsx(pasta_download)

    log("Disparando extração...")
    btn = wait.until(
        EC.element_to_be_clickable((By.ID, "btnExtrairRelatorio"))
    )
    driver.execute_script("arguments[0].click();", btn)

    return aguardar_download(pasta_download, antes)

# ============================================================
# RENOMEAR E MOVER
# ============================================================

def renomear_e_mover(caminho_original: str) -> tuple[str, str]:
    timestamp  = datetime.now().strftime("%d-%m-%y %H%M")
    novo_nome  = f"Todas - Relatorio de Pendencias - {timestamp}.xlsx"
    pasta_sem  = pasta_da_semana(datetime.now(), PASTA_REPO_DATA)
    pasta_sem.mkdir(parents=True, exist_ok=True)
    destino    = pasta_sem / novo_nome
    os.rename(caminho_original, str(destino))
    log(f"Arquivo salvo: {destino}")
    return str(destino), novo_nome

# ============================================================
# MAIN
# ============================================================

def main():
    if not USUARIO or not SENHA:
        log("ERRO: Variáveis MG_USUARIO e MG_SENHA não definidas!")
        sys.exit(1)

    log("=" * 50)
    log("MG Contécnica — Extrator Automático")
    log(f"Horário: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    log("=" * 50)

    # Pasta temporária de download (raiz do data/)
    pasta_download = PASTA_REPO_DATA
    pasta_download.mkdir(parents=True, exist_ok=True)

    driver = None
    try:
        log("Iniciando Chrome headless...")
        driver = iniciar_driver(pasta_download)

        fazer_login(driver)
        navegar_para_relatorio(driver)

        arquivo = extrair(driver, pasta_download)

        if arquivo:
            log("Download concluído!")
            destino, nome = renomear_e_mover(arquivo)
            log(f"Arquivo final: {nome}")
            log("Extração concluída com sucesso!")
        else:
            log("ERRO: Download não detectado dentro do tempo limite.")
            sys.exit(1)

    except Exception as e:
        log(f"ERRO: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

    finally:
        if driver:
            driver.quit()

if __name__ == "__main__":
    main()
