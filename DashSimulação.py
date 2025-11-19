import subprocess
import time
import os

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException  # <<< IMPORTANTE

URL = "https://microstrategy.itapevarec.com.br/MicroStrategy/asp/Main.aspx?evt=3001&src=Main.aspx.3001&Port=0&"
CAMINHO_CREDENCIAIS = r"\\fs01\ITAPEVA ATIVAS\DADOS\SA_Credencials.txt"

PASTA_DOWNLOAD_PADRAO = r"\\fs01\ITAPEVA ATIVAS\DADOS\Simulador\External Agencies Report - Guaranties - Autos"
PASTA_DOWNLOAD_AUTOS = PASTA_DOWNLOAD_PADRAO
PASTA_DOWNLOAD_SIMULACOES = r"\\fs01\ITAPEVA ATIVAS\DADOS\Simulador\External Agencies Report - Guaranties - Autos Simulations"

NOME_RELATORIO_AUTOS = "External Agencies Report - Guaranties - Autos"
NOME_RELATORIO_SIMULACOES = "External Agencies Report - Guaranties - Autos Simulations"


def carregar_credenciais_microstrategy(caminho):
    usuario = None
    senha = None
    with open(caminho, "r", encoding="utf-8") as f:
        for linha in f:
            linha = linha.strip()
            if not linha or linha.startswith("#"):
                continue
            if linha.lower().startswith("login:"):
                usuario = linha.split(":", 1)[1].strip()
            elif linha.lower().startswith("senha:"):
                senha = linha.split(":", 1)[1].strip()
    if not usuario or not senha:
        raise RuntimeError("Nao foi possivel ler login/senha do arquivo de credenciais.")
    return usuario, senha


def criar_driver():
    o = Options()
    o.add_argument("--start-maximized")
    o.add_argument("--disable-notifications")
    o.add_argument("--log-level=3")
    o.add_experimental_option("excludeSwitches", ["enable-logging"])

    os.makedirs(PASTA_DOWNLOAD_PADRAO, exist_ok=True)
    os.makedirs(PASTA_DOWNLOAD_SIMULACOES, exist_ok=True)

    prefs = {
        "download.default_directory": PASTA_DOWNLOAD_PADRAO,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    o.add_experimental_option("prefs", prefs)

    s = ChromeService(log_output=subprocess.DEVNULL)
    return webdriver.Chrome(service=s, options=o)


def login_microstrategy():
    usuario, senha = carregar_credenciais_microstrategy(CAMINHO_CREDENCIAIS)
    d = criar_driver()
    w = WebDriverWait(d, 30)

    d.get(URL)

    u = w.until(EC.visibility_of_element_located((By.ID, "Uid")))
    u.clear()
    u.send_keys(usuario)

    p = w.until(EC.visibility_of_element_located((By.ID, "Pwd")))
    p.clear()
    p.send_keys(senha)

    b = w.until(EC.element_to_be_clickable((By.ID, "3054")))
    b.click()

    return d, w


def abrir_projeto(d, w):
    seletor = "#projects_ProjectsStyle > table > tbody > tr > td:nth-child(1) > div > table > tbody > tr > td.mstrLargeIconViewItemText > a"
    link = w.until(EC.element_to_be_clickable((By.CSS_SELECTOR, seletor)))
    link.click()


def clicar_desktop(d, timeout=60):
    w = WebDriverWait(d, timeout)
    icone = w.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.mstr-dskt-icn")))
    d.execute_script("arguments[0].scrollIntoView(true);", icone)
    d.execute_script("arguments[0].click();", icone)


def abrir_pasta_external_agencies(d):
    link = (
        "https://microstrategy.itapevarec.com.br/MicroStrategy/asp/"
        "Main.aspx?Server=SRV-APPMS02&Project=Projeto+Corporativo&Port=0"
        "&evt=2001&src=Main.aspx.shared.fbb.fb.2001"
        "&folderID=B913B4B141968694EABA5A8059112102"
    )
    d.get(link)


def abrir_pasta_relatorios_padrao(d):
    link = (
        "https://microstrategy.itapevarec.com.br/MicroStrategy/asp/"
        "Main.aspx?Server=SRV-APPMS02&Project=Projeto+Corporativo&Port=0"
        "&evt=2001&src=Main.aspx.shared.fbb.fb.2001"
        "&folderID=1FA0A4F641D05A8AEB794290CFAEEC19"
    )
    d.get(link)


def abrir_pasta_cube_reports(d):
    link = (
        "https://microstrategy.itapevarec.com.br/MicroStrategy/asp/"
        "Main.aspx?Server=SRV-APPMS02&Project=Projeto+Corporativo&Port=0"
        "&evt=2001&src=Main.aspx.shared.fbb.fb.2001"
        "&folderID=F6FC064746B9CED48C7F0B876A9650E7"
    )
    d.get(link)


def abrir_relatorio_guaranties_autos(d):
    link = (
        "https://microstrategy.itapevarec.com.br/MicroStrategy/asp/"
        "Main.aspx?Server=SRV-APPMS02&Project=Projeto+Corporativo&Port=0"
        "&evt=4001&src=Main.aspx.4001&visMode=0&reportViewMode=1"
        "&reportID=0643B3DA4BAA63188229B69E48E5E363"
    )
    d.get(link)


def abrir_relatorio_simulacoes(d):
    link = (
        "https://microstrategy.itapevarec.com.br/MicroStrategy/asp/"
        "Main.aspx?Server=SRV-APPMS02&Project=Projeto+Corporativo&Port=0"
        "&evt=4001&src=Main.aspx.4001&visMode=0&reportViewMode=1"
        "&reportID=F69CEA1E478AB6C2AE75DD9544D34462"
    )
    d.get(link)


def clicar_exportar(d, w):
    handles_antes = d.window_handles
    botao = w.until(EC.element_to_be_clickable((By.ID, "tbExport")))
    d.execute_script("arguments[0].click();", botao)

    try:
        WebDriverWait(d, 15).until(lambda drv: len(drv.window_handles) > len(handles_antes))
        handles_depois = d.window_handles
        novo_handle = list(set(handles_depois) - set(handles_antes))
        if novo_handle:
            d.switch_to.window(novo_handle[0])
    except Exception:
        pass

    time.sleep(2)


def _tentar_click_botao_exportar_no_contexto(d, timeout=10):
    w = WebDriverWait(d, timeout)

    try:
        botao = w.until(EC.element_to_be_clickable((By.ID, "3131")))
        d.execute_script("arguments[0].scrollIntoView(true);", botao)
        d.execute_script("arguments[0].click();", botao)
        return True
    except Exception:
        pass

    try:
        xpath = "//input[@id='3131' and @value='Exportar']"
        botao = w.until(EC.element_to_be_clickable((By.XPATH, xpath)))
        d.execute_script("arguments[0].scrollIntoView(true);", botao)
        d.execute_script("arguments[0].click();", botao)
        return True
    except Exception:
        pass

    return False


def confirmar_exportar(d, timeout=60):
    if _tentar_click_botao_exportar_no_contexto(d, timeout=5):
        time.sleep(2)
        return

    iframes = d.find_elements(By.TAG_NAME, "iframe")
    for frame in iframes:
        try:
            d.switch_to.default_content()
            d.switch_to.frame(frame)
            if _tentar_click_botao_exportar_no_contexto(d, timeout=5):
                d.switch_to.default_content()
                time.sleep(2)
                return
        except Exception:
            continue
    d.switch_to.default_content()

    time.sleep(2)


def esperar_novo_download(pasta, arquivos_antes, timeout=120):
    inicio = time.time()
    while time.time() - inicio < timeout:
        atuais = set(os.listdir(pasta))
        novos = [f for f in atuais - arquivos_antes if not f.lower().endswith(".crdownload")]
        if novos:
            caminhos_novos = [os.path.join(pasta, f) for f in novos]
            caminho = max(caminhos_novos, key=os.path.getmtime)
            return caminho
        time.sleep(1)
    return None


def renomear_relatorio_com_data(caminho_original, nome_base, pasta_destino):
    if not caminho_original or not os.path.isfile(caminho_original):
        print("Nao foi possivel encontrar o arquivo baixado para renomear.")
        return

    os.makedirs(pasta_destino, exist_ok=True)
    _, ext = os.path.splitext(caminho_original)
    novo_nome = f"{nome_base}{ext}"
    novo_caminho = os.path.join(pasta_destino, novo_nome)

    if os.path.abspath(caminho_original) == os.path.abspath(novo_caminho):
        print(f"Relatorio salvo como: {novo_caminho}")
        return

    try:
        os.replace(caminho_original, novo_caminho)
    except Exception:
        try:
            if os.path.exists(novo_caminho):
                os.remove(novo_caminho)
            os.replace(caminho_original, novo_caminho)
        except Exception as e:
            print(f"Erro ao mover arquivo: {e}")
            return

    print(f"Relatorio salvo como: {novo_caminho}")


def clicar_mstr_logo(d, timeout=20):
    w = WebDriverWait(d, timeout)
    d.switch_to.default_content()
    logo = w.until(EC.element_to_be_clickable((By.ID, "mstrLogo")))
    d.execute_script("arguments[0].click();", logo)
    time.sleep(1)


# ============================================================
#  FUNÇÃO QUE DÁ DOUBLE-CLICK + ESPERA O LOAD (STALENESS)
# ============================================================

def clicar_por_id_em_qualquer_frame(d, element_id, timeout=30):
    """
    Procura o elemento por ID no documento principal e em todos
    os frames/iframes, dá um double-click REAL (ActionChains)
    e só retorna depois que o elemento ficar 'stale'
    (ou vencer o timeout de staleness).
    """
    fim = time.time() + timeout
    while time.time() < fim:
        # 1) tentar no documento principal
        try:
            d.switch_to.default_content()
            el = WebDriverWait(d, 2).until(
                EC.element_to_be_clickable((By.ID, element_id))
            )
            ActionChains(d).move_to_element(el).pause(0.1).double_click(el).perform()

            # Espera o "load" terminar: elemento ficar stale
            try:
                WebDriverWait(d, 10).until(EC.staleness_of(el))
            except TimeoutException:
                # se não ficou stale, segue a vida mesmo assim
                pass

            d.switch_to.default_content()
            return
        except Exception:
            pass

        # 2) tentar em cada frame/iframe de primeiro nível
        frames = []
        try:
            frames += d.find_elements(By.TAG_NAME, "frame")
        except Exception:
            pass
        try:
            frames += d.find_elements(By.TAG_NAME, "iframe")
        except Exception:
            pass

        for fr in frames:
            try:
                d.switch_to.default_content()
                d.switch_to.frame(fr)
                el = WebDriverWait(d, 2).until(
                    EC.element_to_be_clickable((By.ID, element_id))
                )
                ActionChains(d).move_to_element(el).pause(0.1).double_click(el).perform()

                # Espera o "load" terminar: elemento ficar stale
                try:
                    WebDriverWait(d, 10).until(EC.staleness_of(el))
                except TimeoutException:
                    pass

                d.switch_to.default_content()
                return
            except Exception:
                continue

        time.sleep(1)

    d.switch_to.default_content()
    # Se chegar aqui, não encontrou/clicou no tempo dado.


if __name__ == "__main__":
    driver, wait = login_microstrategy()
    abrir_projeto(driver, wait)
    clicar_desktop(driver)
    abrir_pasta_external_agencies(driver)
    abrir_pasta_relatorios_padrao(driver)
    abrir_pasta_cube_reports(driver)
    abrir_relatorio_guaranties_autos(driver)

    time.sleep(2)  # tempo pro relatório montar

    # Cada um desses agora:
    # - dá double-click
    # - espera o elemento ficar stale (load terminar)
    clicar_por_id_em_qualquer_frame(driver, "0643B3DA4BAA63188229B69E48E5E3630")
    clicar_por_id_em_qualquer_frame(driver, "0643B3DA4BAA63188229B69E48E5E3635")
    clicar_por_id_em_qualquer_frame(driver, "0643B3DA4BAA63188229B69E48E5E36318")
    clicar_por_id_em_qualquer_frame(driver, "0643B3DA4BAA63188229B69E48E5E36319")
    clicar_por_id_em_qualquer_frame(driver, "0643B3DA4BAA63188229B69E48E5E36320")
    clicar_por_id_em_qualquer_frame(driver, "0643B3DA4BAA63188229B69E48E5E36321")
    clicar_por_id_em_qualquer_frame(driver, "0643B3DA4BAA63188229B69E48E5E36322")
    clicar_por_id_em_qualquer_frame(driver, "0643B3DA4BAA63188229B69E48E5E36323")
    clicar_por_id_em_qualquer_frame(driver, "0643B3DA4BAA63188229B69E48E5E3638")

    arquivos_antes_autos = set(os.listdir(PASTA_DOWNLOAD_PADRAO))

    clicar_exportar(driver, wait)

    # ================== AQUI: selecionar o radio de Excel ==================
    try:
        # tenta no documento principal e em iframes
        contextos = [None]
        try:
            contextos += driver.find_elements(By.TAG_NAME, "iframe")
        except Exception:
            pass

        clicou_radio = False
        for ctx in contextos:
            try:
                driver.switch_to.default_content()
                if ctx is not None:
                    driver.switch_to.frame(ctx)

                rb = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.ID, "exportFormatGrids_excelPlaintextIServer"))
                )
                driver.execute_script("arguments[0].click();", rb)
                clicou_radio = True
                driver.switch_to.default_content()
                break
            except Exception:
                continue

        driver.switch_to.default_content()
    except Exception:
        driver.switch_to.default_content()
    # ======================================================================

    confirmar_exportar(driver)

    caminho_autos = esperar_novo_download(PASTA_DOWNLOAD_PADRAO, arquivos_antes_autos, timeout=180)
    renomear_relatorio_com_data(caminho_autos, NOME_RELATORIO_AUTOS, PASTA_DOWNLOAD_AUTOS)

    clicar_mstr_logo(driver)
    clicar_desktop(driver)
    abrir_pasta_external_agencies(driver)
    abrir_pasta_relatorios_padrao(driver)
    abrir_pasta_cube_reports(driver)
    abrir_relatorio_simulacoes(driver)

    # aquiii – clicar nos campos do relatório de simulações
    clicar_por_id_em_qualquer_frame(driver, "F69CEA1E478AB6C2AE75DD9544D344620")
    clicar_por_id_em_qualquer_frame(driver, "F69CEA1E478AB6C2AE75DD9544D344628")
    clicar_por_id_em_qualquer_frame(driver, "F69CEA1E478AB6C2AE75DD9544D3446216")
    clicar_por_id_em_qualquer_frame(driver, "F69CEA1E478AB6C2AE75DD9544D3446218")
    clicar_por_id_em_qualquer_frame(driver, "F69CEA1E478AB6C2AE75DD9544D3446219")
    clicar_por_id_em_qualquer_frame(driver, "F69CEA1E478AB6C2AE75DD9544D3446220")
    clicar_por_id_em_qualquer_frame(driver, "F69CEA1E478AB6C2AE75DD9544D3446221")
    clicar_por_id_em_qualquer_frame(driver, "F69CEA1E478AB6C2AE75DD9544D3446226")
    clicar_por_id_em_qualquer_frame(driver, "F69CEA1E478AB6C2AE75DD9544D3446253")
    clicar_por_id_em_qualquer_frame(driver, "F69CEA1E478AB6C2AE75DD9544D3446254")
    

    arquivos_antes_sim = set(os.listdir(PASTA_DOWNLOAD_PADRAO))

    clicar_exportar(driver, wait)

    # ================== AQUI TAMBÉM: selecionar o radio de Excel ===========
    try:
        contextos = [None]
        try:
            contextos += driver.find_elements(By.TAG_NAME, "iframe")
        except Exception:
            pass

        clicou_radio = False
        for ctx in contextos:
            try:
                driver.switch_to.default_content()
                if ctx is not None:
                    driver.switch_to.frame(ctx)

                rb = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.ID, "exportFormatGrids_excelPlaintextIServer"))
                )
                driver.execute_script("arguments[0].click();", rb)
                clicou_radio = True
                driver.switch_to.default_content()
                break
            except Exception:
                continue

        driver.switch_to.default_content()
    except Exception:
        driver.switch_to.default_content()
    # ======================================================================

    confirmar_exportar(driver)

    caminho_sim = esperar_novo_download(PASTA_DOWNLOAD_PADRAO, arquivos_antes_sim, timeout=180)
    renomear_relatorio_com_data(caminho_sim, NOME_RELATORIO_SIMULACOES, PASTA_DOWNLOAD_SIMULACOES)

    print("Script finalizado. Navegador ficara aberto. Pressione Ctrl+C para encerrar.")
    try:
        while True:
            time.sleep(1)
            break
    except KeyboardInterrupt:
        pass
