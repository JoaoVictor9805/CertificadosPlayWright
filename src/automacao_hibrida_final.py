import os
import sys
import warnings

# Define STA (Single-Threaded Apartment) antes de qualquer importação do pywinauto/comtypes
if not hasattr(sys, "coinit_flags"):
    sys.coinit_flags = 2
warnings.filterwarnings("ignore", message=".*Apply externally defined coinit_flags.*")
import re
import time
import shutil
import socket
import subprocess
import ctypes
import logging
import threading
from typing import Optional, Callable
from datetime import datetime
from pathlib import Path

import openpyxl
import pyautogui
from pywinauto import Desktop
from pywinauto.keyboard import send_keys
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

# ==============================================================================
# CONFIGURAÇÕES E CAMINHOS
# ==============================================================================
BASE_DIR = Path(r"Z:\Temp\SOCIETARIO\AutomacaoCertificados")
EXCEL_PATH = BASE_DIR / "CNPJs.xlsx"
SCREENSHOTS_DIR = BASE_DIR / "CapturasDeTela"

URL_LOGIN = "https://det.sit.trabalho.gov.br/login"
URL_SERVICOS = "https://det.sit.trabalho.gov.br/servicos"

# ModoTeste = True: considera notificações com data <= hoje (testes).
# Em produção, trocar para False (data >= hoje).
MODO_TESTE = True

# Configurações do Chrome Real e Porta CDP
CHROME_PATH = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
CHROME_USER_DATA = Path(os.environ["LOCALAPPDATA"]) / "Google" / "Chrome" / "User Data_RoboDET"
CDP_PORT = 9222

# Delays para acompanhamento visual do usuário (0.5s entre ações)
DELAY_ENTRE_ACOES_S = 0.5
DELAY_SLOW_MO_MS = 500

def obter_caminho_recurso(caminho_relativo: str) -> Path:
    """
    Retorna o caminho absoluto do recurso, compatível com execução direta (Python),
    PyInstaller em modo diretório (_internal / pasta do executável) ou modo temporário (_MEIPASS).
    """
    if getattr(sys, "frozen", False):
        base_meipass = getattr(sys, "_MEIPASS", None)
        if base_meipass:
            c = Path(base_meipass) / caminho_relativo
            if c.exists():
                return c

        exe_dir = Path(sys.executable).resolve().parent
        c_exe = exe_dir / caminho_relativo
        if c_exe.exists():
            return c_exe

        c_internal = exe_dir / "_internal" / caminho_relativo
        if c_internal.exists():
            return c_internal

    script_dir = Path(__file__).resolve().parent
    if (script_dir / caminho_relativo).exists():
        return script_dir / caminho_relativo
    if (script_dir.parent / caminho_relativo).exists():
        return script_dir.parent / caminho_relativo

    return script_dir.parent / caminho_relativo


# Imagens de referência para o login físico via PyAutoGUI (procura em assets/ e fallback)
_caminho_assets = obter_caminho_recurso("assets/capturasDeTela")
if _caminho_assets.exists():
    IMG_DIR = _caminho_assets
else:
    IMG_DIR = obter_caminho_recurso("capturasDeTela")

IMG_GOVBR = IMG_DIR / "EntrarComGov.png"
IMG_CERT = IMG_DIR / "SeuCertificadoDigital.png"
IMG_OK = IMG_DIR / "Ok.png"
IMG_DEBUG_ERRO = IMG_DIR / "debug_tela_erro.png"

# ==============================================================================
# LOGGER DA APLICAÇÃO (RoboDET)
# ==============================================================================
logger = logging.getLogger("RoboDET")


def log_etapa(item_id: str, etapa: str, msg: str):
    logger.info(f"[{item_id}] [{etapa}] {msg}")


# ==============================================================================
# FUNÇÕES DE SISTEMA OPERACIONAL (PORTAS, CHROME E PYAUTOGUI)
# ==============================================================================
def porta_em_uso(porta: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(("127.0.0.1", porta)) == 0


def liberar_porta_se_ocupada():
    """Fecha processos órfãos na porta CDP para garantir uma inicialização limpa."""
    if porta_em_uso(CDP_PORT):
        logger.info(f"Porta {CDP_PORT} em uso. Tentando fechar instância anterior...")
        try:
            with sync_playwright() as p:
                b = p.chromium.connect_over_cdp(f"http://127.0.0.1:{CDP_PORT}")
                ctx = b.contexts[0]
                page = ctx.pages[0] if ctx.pages else ctx.new_page()
                cdp = ctx.new_cdp_session(page)
                cdp.send("Browser.close")
                time.sleep(1)
        except Exception:
            pass


def iniciar_chrome_hibrido():
    """
    Inicia o executável oficial do Google Chrome com perfil persistente e porta CDP.
    O Playwright NÃO é conectado durante a inicialização, evitando qualquer flag ou
    sinalização de automação no gov.br.
    """
    CHROME_USER_DATA.mkdir(parents=True, exist_ok=True)
    cmd = [
        CHROME_PATH,
        f"--remote-debugging-port={CDP_PORT}",
        f"--user-data-dir={CHROME_USER_DATA}",
        "--start-maximized",
        "--no-first-run",
        "--no-default-browser-check",
        "--disable-blink-features=AutomationControlled",
        URL_LOGIN,
    ]
    subprocess.Popen(cmd)
    for _ in range(20):
        if porta_em_uso(CDP_PORT):
            break
        time.sleep(0.5)


def focar_janela_segura(win):
    """Garante foco na janela do Chrome para recepção de eventos físicos."""
    try:
        hwnd = win.handle
        if ctypes.windll.user32.IsIconic(hwnd):
            ctypes.windll.user32.ShowWindow(hwnd, 9)
        ctypes.windll.user32.SetForegroundWindow(hwnd)
        time.sleep(0.5)
    except Exception:
        pass


def obter_janela_chrome(timeout_segundos: int = 30):
    """Localiza a janela principal do Chrome aberta."""
    desktop = Desktop(backend="uia")
    fim = time.time() + timeout_segundos
    while time.time() < fim:
        try:
            for win in desktop.windows():
                txt = win.window_text()
                cls = win.class_name()
                if cls == "Chrome_WidgetWin_1" or any(k in txt for k in ["Chrome", "DET", "gov.br"]):
                    focar_janela_segura(win)
                    return win
        except Exception:
            pass
        time.sleep(1)
    raise TimeoutError("Janela do Google Chrome não encontrada dentro do tempo limite.")


def clicar_robusto(pos_x: int, pos_y: int):
    """Move o mouse e dispara um clique físico completo e deliberado (mouseDown + mouseUp)."""
    pyautogui.moveTo(pos_x, pos_y, duration=0.6)
    time.sleep(0.25)
    pyautogui.mouseDown()
    time.sleep(0.2)
    pyautogui.mouseUp()
    time.sleep(0.2)


def localizar_imagem(caminho_imagem: Path, confidence: float = 0.8):
    """Localiza o centro da imagem na tela com tolerância a ClearType (grayscale)."""
    if not caminho_imagem.exists():
        return None

    try:
        pos = pyautogui.locateCenterOnScreen(str(caminho_imagem), confidence=confidence, grayscale=True)
        if pos:
            return pos
    except Exception:
        pass

    try:
        pos = pyautogui.locateCenterOnScreen(str(caminho_imagem), confidence=max(0.65, confidence - 0.1), grayscale=True)
        if pos:
            return pos
    except Exception:
        pass

    return None


def clicar_botao_govbr_garantido(timeout_segundos: int = 25) -> bool:
    """
    Localiza e clica no botão 'Entrar com gov.br', garantindo que o clique
    realmente acionou a transição de tela.
    """
    fim = time.time() + timeout_segundos
    while time.time() < fim:
        pos = localizar_imagem(IMG_GOVBR, confidence=0.8)
        if pos:
            logger.info(f"[FASE 1 - SO] 'Entrar com gov.br' localizado em ({pos.x}, {pos.y}). Clicando...")
            clicar_robusto(pos.x, pos.y)
            time.sleep(2)

            # Verifica se o botão sumiu (indicando navegação iniciada)
            if not localizar_imagem(IMG_GOVBR, confidence=0.8):
                logger.info("[FASE 1 - SO] Botão 'Entrar com gov.br' clicado com sucesso!")
                return True
            else:
                logger.info("[FASE 1 - SO] Tentando segundo clique de ativação...")
                clicar_robusto(pos.x, pos.y)
                time.sleep(2)
                return True
        time.sleep(0.8)

    return False


def clicar_botao_certificado_garantido(timeout_segundos: int = 30) -> bool:
    """
    Aguarda e clica em 'Seu certificado digital' na tela do sso.acesso.gov.br.
    Se demorar, realiza uma leve rolagem na página para garantir visibilidade.
    """
    logger.info(f"[FASE 1 - SO] Procurando 'Seu certificado digital' na tela (timeout={timeout_segundos}s)...")
    fim = time.time() + timeout_segundos
    tentativa = 0

    while time.time() < fim:
        tentativa += 1
        pos = localizar_imagem(IMG_CERT, confidence=0.75)
        if pos:
            logger.info(f"[FASE 1 - SO] 'Seu certificado digital' ENCONTRADO em ({pos.x}, {pos.y})! Clicando...")
            clicar_robusto(pos.x, pos.y)
            # Delay estendido para dar tempo ao Windows/Chrome de carregar a lista de certificados
            time.sleep(2.5)
            return True

        if tentativa % 6 == 0:
            logger.info(f"[FASE 1 - SO] Tentando leve rolagem da página para baixo (tentativa {tentativa})...")
            try:
                largura, altura = pyautogui.size()
                pyautogui.moveTo(largura // 2, altura // 2)
                pyautogui.scroll(-200)
            except Exception:
                pass

        time.sleep(0.8)

    # Se falhar, salva captura para diagnóstico
    logger.warning(f"'Seu certificado digital' não foi localizado em {timeout_segundos}s.")
    try:
        pyautogui.screenshot(str(IMG_DEBUG_ERRO))
        logger.info(f"[DIAGNÓSTICO] Screenshot da tela salvo em: {IMG_DEBUG_ERRO.name}")
    except Exception:
        pass

    return False


def confirmar_popup_certificado(timeout_segundos: int = 25) -> bool:
    """
    Aguarda o popup nativo de seleção de certificado e clica em OK.
    Tenta por visão computacional (PyAutoGUI) e conta com fallbacks via UIA e ENTER.
    """
    logger.info("[FASE 1 - SO] Aguardando popup de seleção de certificado...")
    time.sleep(1.0)  # Delay para estabilização do popup na tela
    fim = time.time() + timeout_segundos
    tentativa = 0

    while time.time() < fim:
        tentativa += 1

        # 1. Tenta por imagem (PyAutoGUI)
        pos_ok = localizar_imagem(IMG_OK, confidence=0.7)
        if pos_ok:
            logger.info(f"[FASE 1 - SO] Botão 'OK' do certificado encontrado em ({pos_ok.x}, {pos_ok.y})!")
            time.sleep(0.5)  # Delay antes do clique
            clicar_robusto(pos_ok.x, pos_ok.y)
            time.sleep(2.0)
            return True

        # 2. Fallback: via UIA nativo do Windows
        try:
            desktop = Desktop(backend="uia")
            for win in desktop.windows():
                titulo = win.window_text().lower()
                palavras_chave = [
                    "gov.br", "acesse sua conta", "selecionar um certificado",
                    "certificado", "segurança do windows", "windows security", "segurança"
                ]
                if any(k in titulo for k in palavras_chave):
                    btns_ok = win.descendants(control_type="Button", title="OK")
                    if btns_ok:
                        logger.info("[FASE 1 - SO] Botão 'OK' do certificado acionado via UIA!")
                        try:
                            focar_janela_segura(win)
                        except Exception:
                            pass
                        time.sleep(0.5)
                        btns_ok[0].click_input()
                        time.sleep(2.0)
                        return True
        except Exception:
            pass

        # A cada 4 tentativas (~3s), tenta focar a janela do certificado e enviar ENTER
        if tentativa % 4 == 0:
            logger.info("[FASE 1 - SO] Tentando confirmação intermediária com ENTER no popup...")
            try:
                desktop = Desktop(backend="uia")
                for win in desktop.windows():
                    titulo = win.window_text().lower()
                    if any(k in titulo for k in ["certificado", "segurança", "gov.br"]):
                        focar_janela_segura(win)
                        break
            except Exception:
                pass
            time.sleep(0.5)
            send_keys("{ENTER}")
            time.sleep(1.5)

        time.sleep(0.8)

    # 3. Fallback final: Enviar ENTER
    logger.info("[FASE 1 - SO] Enviando tecla ENTER para confirmar certificado...")
    time.sleep(0.8)
    send_keys("{ENTER}")
    time.sleep(2.0)
    return True


def fechar_navegador(browser=None, context=None, page=None):
    """
    Fecha a janela do Chrome real via comando CDP 'Browser.close'
    e encerra a sessão do Playwright de forma limpa.
    """
    try:
        if context and page and not page.is_closed():
            cdp = context.new_cdp_session(page)
            cdp.send("Browser.close")
            time.sleep(1)
    except Exception:
        pass

    try:
        if browser:
            browser.close()
    except Exception:
        pass


def limpar_ou_criar_pasta(caminho: Path):
    if not caminho.exists():
        caminho.mkdir(parents=True, exist_ok=True)
    else:
        for item in caminho.iterdir():
            try:
                if item.is_file() or item.is_symlink():
                    item.unlink()
                elif item.is_dir():
                    shutil.rmtree(item)
            except Exception as e:
                logger.warning(f"Não foi possível limpar {item}: {e}")


def remover_pastas_vazias_do_item(pasta_item: Path):
    """Replica a limpeza de subpastas vazias e da pasta principal do item."""
    for sub in ["CaixaPostal", "Notificacao", "ArquivosBaixados"]:
        sub_path = pasta_item / sub
        if sub_path.exists() and not any(sub_path.iterdir()):
            try:
                sub_path.rmdir()
            except OSError:
                pass

    if pasta_item.exists() and not any(pasta_item.iterdir()):
        try:
            pasta_item.rmdir()
        except OSError:
            pass


def exibir_alerta_planilha_inexistente(caminho_criado: Path):
    mensagem = (
        "A Planilha de referência para consulta de CNPJs não existia.\n\n"
        f"A planilha foi criada no caminho:\n{caminho_criado}\n\n"
        "Por favor, preencha com os CNPJs desejados na coluna 'CNPJ', um abaixo do outro, e execute o robô novamente."
    )
    ctypes.windll.user32.MessageBoxW(0, mensagem, "Planilha de Referência não existe", 0x30)


# ==============================================================================
# FLUXO PRINCIPAL DA AUTOMAÇÃO HÍBRIDA
# ==============================================================================
def executar_robo_hibrido(
    base_dir: Optional[Path] = None,
    progress_callback: Optional[Callable[[int, int, str], None]] = None,
    stop_event: Optional[threading.Event] = None,
):
    if base_dir is None:
        base_dir = BASE_DIR
    else:
        base_dir = Path(base_dir)

    excel_path = base_dir / "CNPJs.xlsx"
    if not excel_path.exists():
        # Fallback de compatibilidade caso o arquivo anterior exista
        legado_path = base_dir / "CNPJs-e-CPFs.xlsx"
        if legado_path.exists():
            excel_path = legado_path

    screenshots_dir = base_dir / "CapturasDeTela"
    hoje_date = datetime.now().date()
    arquivos_pendentes_excluir = []

    # 1. Verificação da Planilha de Entrada
    if not excel_path.exists():
        logger.error(f"Planilha de CNPJs não encontrada em: {excel_path}")
        try:
            base_dir.mkdir(parents=True, exist_ok=True)
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Documentos"
            ws["A1"] = "CNPJ"
            wb.save(str(excel_path))
            wb.close()
            logger.info(f"Planilha modelo criada com sucesso em: {excel_path}")
        except Exception as e:
            logger.error(f"Falha ao criar planilha modelo: {e}")
        exibir_alerta_planilha_inexistente(excel_path)
        return

    try:
        wb = openpyxl.load_workbook(str(excel_path), data_only=True)
        sheet = wb.active
        raw_cells = [
            str(cell.value).strip()
            for row in sheet.iter_rows(values_only=False)
            for cell in row
            if cell.value is not None and str(cell.value).strip() != ""
        ]
        wb.close()
    except Exception as e:
        logger.error(f"Erro ao ler a planilha {excel_path}: {e}")
        return

    # Filtra cabeçalhos conhecidos e linhas vazias
    header_names = {"CNPJ_OU_CPF", "CNPJ_E_CPF", "CNPJS_E_CPFS", "CNPJ", "CPF", "DOCUMENTO", "DOCUMENTOS"}
    excel_data = [item for item in raw_cells if item.upper() not in header_names]

    if not excel_data:
        logger.warning(
            f"A planilha em '{excel_path}' não contém CNPJs ou CPFs para processar. Preencha e inicie novamente."
        )
        return

    total_itens = len(excel_data)
    logger.info(f"Planilha carregada com sucesso. Total de itens a processar: {total_itens}")
    if progress_callback:
        progress_callback(0, total_itens, f"Iniciando automação... (0/{total_itens})")

    limpar_ou_criar_pasta(screenshots_dir)

    playwright_mgr = None
    browser = None
    context = None
    page = None

    try:
        # ======================================================================
        # FASE 1: Login via Sistema Operacional (PyAutoGUI - Sem rastro de CDP)
        # ======================================================================
        logger.info("Iniciando Fase 1: Abertura do Chrome e Login Físico via Sistema...")
        liberar_porta_se_ocupada()
        iniciar_chrome_hibrido()

        chrome_win = obter_janela_chrome(timeout_segundos=30)
        focar_janela_segura(chrome_win)
        time.sleep(2)

        if stop_event and stop_event.is_set():
            logger.warning("Execução interrompida antes do login.")
            fechar_navegador(browser, context, page)
            return

        # Clica em 'Entrar com gov.br'
        if not clicar_botao_govbr_garantido(timeout_segundos=25):
            logger.error("Erro: Não foi possível localizar o botão 'Entrar com gov.br'.")
            fechar_navegador(browser, context, page)
            return

        if stop_event and stop_event.is_set():
            logger.warning("Execução interrompida após clique no gov.br.")
            fechar_navegador(browser, context, page)
            return

        # Clica em 'Seu certificado digital'
        if not clicar_botao_certificado_garantido(timeout_segundos=30):
            logger.error("Erro: Não foi possível localizar 'Seu certificado digital'.")
            fechar_navegador(browser, context, page)
            return

        if stop_event and stop_event.is_set():
            logger.warning("Execução interrompida antes da seleção do certificado.")
            fechar_navegador(browser, context, page)
            return

        # Confirma popup nativo de certificado
        confirmar_popup_certificado(timeout_segundos=25)

        logger.info("Fase 1 concluída: Autenticação enviada. Conectando Playwright...")

        # ======================================================================
        # FASE 2: Conexão e Controle do Playwright após autenticação
        # ======================================================================
        playwright_mgr = sync_playwright().start()

        # Conecta via CDP com retentativas para aguardar estabilização do Chrome
        for tentativa_cdp in range(1, 6):
            if stop_event and stop_event.is_set():
                break
            try:
                browser = playwright_mgr.chromium.connect_over_cdp(
                    f"http://127.0.0.1:{CDP_PORT}",
                    slow_mo=DELAY_SLOW_MO_MS
                )
                break
            except Exception:
                time.sleep(2)

        if not browser:
            logger.error("Erro fatal: Não foi possível conectar o Playwright via CDP após 5 tentativas.")
            if playwright_mgr:
                playwright_mgr.stop()
            return

        context = browser.contexts[0]
        page = context.pages[0] if context.pages else context.new_page()

        # Aguarda carregar a tela de Serviços (Caixa Postal)
        try:
            page.wait_for_url("**/servicos**", timeout=45000)
        except Exception:
            pass

        card_caixa_postal = page.locator(".cardListItem:has(span.title:has-text('CAIXA POSTAL'))")
        card_caixa_postal.wait_for(state="visible", timeout=30000)
        time.sleep(DELAY_ENTRE_ACOES_S)

        logger.info("Fase 2 estabelecida: Playwright conectado e tela de Serviços confirmada!")

    except Exception as e:
        logger.error(f"Erro fatal no robô (fase de preparação ou login): {e}")
        fechar_navegador(browser, context, page)
        if playwright_mgr:
            playwright_mgr.stop()
        raise

    # ==========================================================================
    # 3. Loop de Processamento de CPFs / CNPJs (Playwright rápido e preciso)
    # ==========================================================================
    perfil = 0
    falhas_seguidas = 0
    etapa = "Início"

    for idx_item, current_item in enumerate(excel_data, start=1):
        if stop_event and stop_event.is_set():
            logger.warning("Execução cancelada pelo usuário.")
            break

        if falhas_seguidas >= 3:
            logger.error("Execução interrompida: 3 falhas seguidas. Os itens restantes não foram processados.")
            break

        item_id = current_item.strip()
        if not item_id:
            continue

        if progress_callback:
            progress_callback(idx_item - 1, total_itens, f"Processando item {idx_item}/{total_itens}: {item_id}")

        try:
            etapa = "Lendo o item da planilha"
            log_etapa(item_id, etapa, "Iniciando o processamento do item")

            etapa = "Aguardando a tela de serviços"
            page.locator(".cardListItem:has(span.title:has-text('CAIXA POSTAL'))").wait_for(state="visible", timeout=60000)
            time.sleep(DELAY_ENTRE_ACOES_S)

            etapa = "Trocando de perfil"
            btn_trocar_perfil = page.locator("button:has-text('Trocar Perfil')").first
            btn_trocar_perfil.wait_for(state="visible", timeout=20000)
            time.sleep(DELAY_ENTRE_ACOES_S)
            btn_trocar_perfil.click()
            time.sleep(2)

            # Clica no dropdown de seleção de perfil e navega via teclado como no PAD
            dropdown_perfil = page.locator("ng-select, .br-select, span.ng-arrow-wrapper").first
            dropdown_perfil.click()
            time.sleep(DELAY_ENTRE_ACOES_S)

            if perfil == 0:
                perfil += 1
                page.keyboard.press("ArrowDown")
                time.sleep(DELAY_ENTRE_ACOES_S)
                page.keyboard.press("Enter")
                time.sleep(DELAY_ENTRE_ACOES_S)
                page.keyboard.press("Tab")
            else:
                page.keyboard.press("ArrowDown")
                time.sleep(DELAY_ENTRE_ACOES_S)
                page.keyboard.press("ArrowDown")
                time.sleep(DELAY_ENTRE_ACOES_S)
                page.keyboard.press("Enter")
                time.sleep(DELAY_ENTRE_ACOES_S)
                page.keyboard.press("Tab")

            time.sleep(DELAY_ENTRE_ACOES_S)
            page.keyboard.type(item_id, delay=100)
            time.sleep(DELAY_ENTRE_ACOES_S)

            etapa = "Selecionando o perfil e validando CPF/CNPJ"
            page.locator("button:has-text('Selecionar')").first.click()
            time.sleep(2)

            # Verifica se apareceu mensagem de CPF/CNPJ inválido no DOM
            msg_invalido = page.get_by_text(re.compile(r"(CNPJ|CPF)\s+inv.lido", re.IGNORECASE))
            if msg_invalido.count() > 0 and msg_invalido.first.is_visible():
                time.sleep(DELAY_ENTRE_ACOES_S)
                page.locator("button:has-text('Cancelar')").first.click()
                time.sleep(DELAY_ENTRE_ACOES_S)
                logger.warning(f"Documento inválido, item ignorado: {item_id}")
                if progress_callback:
                    progress_callback(idx_item, total_itens, f"Item {idx_item}/{total_itens} inválido")
                continue

            # Criação de pastas do item
            etapa = "Criando a pasta do item"
            pasta_item = screenshots_dir / item_id
            pasta_caixa_postal = pasta_item / "CaixaPostal"
            pasta_notificacao = pasta_item / "Notificacao"
            pasta_arquivos_baixados = pasta_item / "ArquivosBaixados"

            etapa = "Criando a subpasta CaixaPostal"
            pasta_caixa_postal.mkdir(parents=True, exist_ok=True)
            etapa = "Criando a subpasta Notificacao"
            pasta_notificacao.mkdir(parents=True, exist_ok=True)
            etapa = "Criando a subpasta ArquivosBaixados"
            pasta_arquivos_baixados.mkdir(parents=True, exist_ok=True)
            time.sleep(DELAY_ENTRE_ACOES_S)

            # --- CAIXA POSTAL ---
            etapa = "Abrindo a Caixa Postal"
            time.sleep(DELAY_ENTRE_ACOES_S)
            card_cp = page.locator(".cardListItem:has(span.title:has-text('CAIXA POSTAL'))")
            card_cp.wait_for(state="visible", timeout=30000)
            card_cp.click()
            time.sleep(2)
            formatted_dt = datetime.now().strftime("%d%m%Y_%H%M%S")

            # 1. Primeiro: clica em 'Exibir apenas não lidas'
            etapa = "Filtrando por mensagens não lidas"
            time.sleep(DELAY_ENTRE_ACOES_S)

            loc_nao_lidas_filtro = page.locator(
                "span:has-text('Exibir apenas não lidas'), "
                "a:has-text('Exibir apenas não lidas'), "
                "button:has-text('Exibir apenas não lidas')"
            ).first

            clicou_filtro = False
            try:
                loc_nao_lidas_filtro.wait_for(state="visible", timeout=8000)
                time.sleep(0.3)
                loc_nao_lidas_filtro.click()
                clicou_filtro = True
                log_etapa(item_id, etapa, "Filtro 'Exibir apenas não lidas' acionado com sucesso.")
                time.sleep(2)  # Aguarda o Angular carregar a lista filtrada
            except Exception:
                try:
                    btn_fallback = page.get_by_text("Exibir apenas não lidas", exact=False).first
                    if btn_fallback.is_visible():
                        btn_fallback.click()
                        clicou_filtro = True
                        log_etapa(item_id, etapa, "Filtro 'Exibir apenas não lidas' acionado via get_by_text.")
                        time.sleep(2)
                except Exception:
                    pass

            if clicou_filtro:
                try:
                    page.mouse.move(0, 0)
                except Exception:
                    pass
            else:
                btn_todas = page.locator("span:has-text('Exibir todas'), a:has-text('Exibir todas'), button:has-text('Exibir todas')").first
                if btn_todas.is_visible():
                    log_etapa(item_id, etapa, "Filtro de não lidas já está ativo ('Exibir todas' presente).")
                else:
                    log_etapa(item_id, etapa, "Aviso: opção 'Exibir apenas não lidas' não foi localizada.")

            # 2. Depois de filtrar: verifica se há mensagens não lidas na tela e tira o print
            etapa = "Tirando o print da Caixa Postal"
            time.sleep(0.5)

            # Seletores reais do DOM identificados nas capturas:
            # - Linha com classe específica: .linha.nao-lida
            # - Ícone do envelope azul ativo: i.fa-envelope.active
            # - Linhas da tabela de mensagens: .tabela.mensagens .linha
            # - Contador numérico na aba superior: Caixa de Entrada (1), Caixa de Entrada (2), etc.
            loc_nao_lidas = page.locator(
                ".linha.nao-lida, "
                ".linha:has(i.fa-envelope.active), "
                ".linha:has(i.fa-envelope), "
                "i.fa-envelope.active, "
                "div[class*='nao-lida']"
            )
            loc_contador_aba = page.get_by_text(re.compile(r"Caixa de Entrada\s*\([1-9]\d*\)", re.I))
            linhas_tabela = page.locator(".tabela.mensagens .linha, .tabela .linha")
            msg_vazio = (
                page.locator(".empty-row")
                .or_(page.get_by_text(re.compile(r"nenhum registro|nenhuma mensagem|não há mensagens", re.I)))
            )
            btn_exibir_todas = page.locator("span:has-text('Exibir todas'), a:has-text('Exibir todas'), button:has-text('Exibir todas')")

            tem_msg_nao_lida = False
            try:
                # Critério A: Linha com classe nao-lida ou envelope ativo visível
                if loc_nao_lidas.count() > 0 and loc_nao_lidas.first.is_visible():
                    tem_msg_nao_lida = True
                # Critério B: Contador na aba 'Caixa de Entrada (N)' onde N > 0
                elif loc_contador_aba.count() > 0 and loc_contador_aba.first.is_visible():
                    tem_msg_nao_lida = True
                else:
                    # Critério C: Filtro de não lidas ativo ('Exibir todas' presente) com registros na tabela
                    filtro_ativo = btn_exibir_todas.count() > 0 and btn_exibir_todas.first.is_visible()
                    tabela_tem_linhas = linhas_tabela.count() > 0 and linhas_tabela.first.is_visible()
                    tabela_vazia = msg_vazio.count() > 0 and msg_vazio.first.is_visible()

                    if (filtro_ativo or clicou_filtro) and tabela_tem_linhas and not tabela_vazia:
                        tem_msg_nao_lida = True
                    else:
                        try:
                            loc_nao_lidas.first.wait_for(state="visible", timeout=3000)
                            tem_msg_nao_lida = True
                        except PlaywrightTimeoutError:
                            tem_msg_nao_lida = False
            except Exception as e:
                logger.warning(f"Erro ao verificar mensagens não lidas: {e}")
                tem_msg_nao_lida = False

            if tem_msg_nao_lida:
                time.sleep(DELAY_ENTRE_ACOES_S)

                # Afasta o cursor do mouse para não ativar tooltips e remove balões visuais residuais
                try:
                    page.mouse.move(0, 0)
                    time.sleep(0.6)  # Aguarda o fade-out natural do tooltip
                    page.evaluate("""() => {
                        const tooltips = document.querySelectorAll('.tooltip, [role="tooltip"], br-tooltip, .br-tooltip, div[class*="tooltip"], div[class*="popover"]');
                        tooltips.forEach(t => t.remove());
                    }""")
                except Exception:
                    pass

                caminho_print_cp = pasta_caixa_postal / f"{item_id}_{formatted_dt}.jpg"
                page.screenshot(path=str(caminho_print_cp), type="jpeg")
                log_etapa(item_id, etapa, "Print da Caixa Postal realizado com sucesso (mensagem não lida encontrada após filtro).")
                time.sleep(DELAY_ENTRE_ACOES_S)
            else:
                log_etapa(item_id, etapa, "Nenhuma mensagem não lida encontrada na Caixa Postal.")

            # Volta para a tela de Serviços
            time.sleep(DELAY_ENTRE_ACOES_S)
            btn_voltar = page.locator("a:has-text('Serviços'), button:has-text('Voltar'), i.fa-arrow-left, i.fa-home").first
            if btn_voltar.is_visible():
                btn_voltar.click()
            else:
                page.goto(URL_SERVICOS, timeout=60000)
            time.sleep(DELAY_ENTRE_ACOES_S)

            # --- NOTIFICAÇÃO ---
            etapa = "Abrindo a tela de Notificação"
            card_notificacao = page.locator(".cardListItem:has(span.title:has-text('NOTIFICAÇÃO'))")
            card_notificacao.wait_for(state="visible", timeout=60000)
            time.sleep(DELAY_ENTRE_ACOES_S)
            card_notificacao.click()
            time.sleep(DELAY_ENTRE_ACOES_S)

            etapa = "Tirando o print da Notificação"
            linhas_notif = page.locator("ngx-datatable datatable-row-wrapper")
            try:
                linhas_notif.first.wait_for(state="visible", timeout=10000)
                tem_notificacoes = True
            except PlaywrightTimeoutError:
                tem_notificacoes = False

            etapa = "Lendo a lista de notificações"
            if tem_notificacoes:
                qtd_linhas = linhas_notif.count()
                log_etapa(
                    item_id,
                    etapa,
                    f"Lista de notificações lida: {qtd_linhas} linha(s) extraída(s). Data de hoje considerada: {hoje_date.strftime('%d/%m/%Y')}",
                )

                dados_linhas = []
                existe_notificacao_valida = False

                for idx in range(qtd_linhas):
                    row_loc = linhas_notif.nth(idx)

                    # Busca a data dinamicamente nas células da linha (suporta diferentes ordens de coluna)
                    texto_data = ""
                    todas_celulas = row_loc.locator("datatable-body-cell").all()
                    for cell in todas_celulas:
                        c_txt = cell.inner_text().strip()
                        m = re.search(r"\b(\d{2}/\d{2}/\d{2,4})\b", c_txt)
                        if m:
                            texto_data = m.group(1)
                            break

                    if not texto_data and len(todas_celulas) > 5:
                        try:
                            texto_data = todas_celulas[5].inner_text().strip()
                        except Exception:
                            pass

                    # Valida formato da data (aceita ano com 4 dígitos %d/%m/%Y ou com 2 dígitos %d/%m/%y como 17/12/20)
                    data_notif = None
                    valida = False
                    for fmt in ("%d/%m/%Y", "%d/%m/%y"):
                        try:
                            data_notif = datetime.strptime(texto_data, fmt).date()
                            break
                        except ValueError:
                            pass

                    if data_notif:
                        valida = (data_notif <= hoje_date) if MODO_TESTE else (data_notif >= hoje_date)

                    if valida:
                        existe_notificacao_valida = True
                    dados_linhas.append((idx, texto_data, valida))

                if not existe_notificacao_valida:
                    log_etapa(item_id, etapa, "Nenhuma notificação possui data válida e o print foi ignorado.")

                # Processa cada linha da lista de notificações
                for idx, texto_data, linha_valida in dados_linhas:
                    etapa = f"Lendo a data da linha {idx}"
                    if linha_valida:
                        row_loc = linhas_notif.nth(idx)

                        etapa = f"Rolando a página até a linha {idx}"
                        row_loc.scroll_into_view_if_needed()
                        time.sleep(DELAY_ENTRE_ACOES_S)

                        # Print recortado da linha da notificação
                        etapa = f"Tirando print recortado da linha {idx}"
                        caminho_print_linha = pasta_notificacao / f"{item_id}_linha_{idx}_{formatted_dt}.jpg"
                        row_loc.screenshot(path=str(caminho_print_linha), type="jpeg")
                        time.sleep(DELAY_ENTRE_ACOES_S)
                        log_etapa(
                            item_id,
                            etapa,
                            f"Print recortado da linha {idx} (data {texto_data}) realizado com sucesso.",
                        )

                        # Download do anexo caso disponível (suporta múltiplas linhas e ícones fa-print / fa-download)
                        etapa = f"Localizando o botão de download da linha {idx}"
                        candidatos_download = row_loc.locator(
                            "button:has(i.fa-download), a:has(i.fa-download), "
                            "button:has(i.fa-print), a:has(i.fa-print), "
                            "button[title*='Download'], button[title*='Baixar'], button[title*='Imprimir']"
                        )
                        btn_download = None
                        if candidatos_download.count() > 0:
                            btn_download = candidatos_download.first
                        else:
                            celula_acoes = row_loc.locator("datatable-body-cell").last
                            botoes_acoes = celula_acoes.locator("button, a")
                            if botoes_acoes.count() > 1:
                                btn_download = botoes_acoes.nth(1)
                            elif botoes_acoes.count() == 1:
                                btn_download = botoes_acoes.first

                        if btn_download and btn_download.count() > 0 and btn_download.is_visible():
                            log_etapa(item_id, etapa, f"Linha {idx} (data {texto_data}): clicando no botão de download")
                            try:
                                etapa = f"Aguardando o download da linha {idx}"
                                time.sleep(DELAY_ENTRE_ACOES_S)
                                with page.expect_download(timeout=60000) as download_info:
                                    btn_download.click()
                                    log_etapa(item_id, etapa, f"Linha {idx}: clique no botão de download realizado")

                                download = download_info.value
                                sug_name = download.suggested_filename
                                destino_final = pasta_arquivos_baixados / sug_name
                                # Se já existir arquivo com mesmo nome (múltiplas linhas baixando mesmo nome), diferencia pelo índice
                                if destino_final.exists():
                                    stem = destino_final.stem
                                    ext = destino_final.suffix
                                    destino_final = pasta_arquivos_baixados / f"{stem}_linha_{idx}{ext}"

                                download.save_as(str(destino_final))
                                time.sleep(DELAY_ENTRE_ACOES_S)
                                log_etapa(
                                    item_id,
                                    etapa,
                                    f"Linha {idx}: arquivo baixado e movido para ArquivosBaixados ({destino_final.name})",
                                )
                            except PlaywrightTimeoutError:
                                log_etapa(
                                    item_id,
                                    etapa,
                                    f"AVISO: download não detectado em 60s (linha {idx})",
                                )
                        else:
                            log_etapa(
                                item_id,
                                etapa,
                                f"Linha {idx} (data {texto_data}): notificação válida sem botão de download (apenas leitura).",
                            )
                    else:
                        log_etapa(
                            item_id,
                            etapa,
                            f"Linha {idx} ignorada: data {texto_data} fora do critério (ModoTeste={MODO_TESTE}, hoje={hoje_date.strftime('%d/%m/%Y')})",
                        )
            else:
                log_etapa(item_id, etapa, "Nenhuma notificação listada na tela de Notificação")

            # Retorna à tela de serviços
            time.sleep(DELAY_ENTRE_ACOES_S)
            page.goto(URL_SERVICOS, timeout=60000)
            time.sleep(DELAY_ENTRE_ACOES_S)

            # Limpeza de diretórios vazios e contabilização
            etapa = "Limpando subpastas vazias"
            remover_pastas_vazias_do_item(pasta_item)

            etapa = "Conferindo os arquivos e limpando a pasta vazia"
            arquivos_guardados = list(pasta_item.rglob("*")) if pasta_item.exists() else []
            qtd_arquivos = sum(1 for f in arquivos_guardados if f.is_file())

            logger.info(
                f"SUCESSO: O item '{item_id}' foi verificado. Arquivos guardados (prints + downloads): {qtd_arquivos}"
            )
            falhas_seguidas = 0

        except Exception as e:
            falhas_seguidas += 1
            logger.error(f"Erro no item '{current_item}' na etapa '{etapa}': {e}")
            try:
                page.goto(URL_SERVICOS, timeout=60000)
            except Exception:
                pass
        finally:
            if progress_callback:
                progress_callback(idx_item, total_itens, f"Item {idx_item}/{total_itens} finalizado")

    # ==========================================================================
    # 4. Encerramento, Logout e Fechamento do Navegador
    # ==========================================================================
    try:
        time.sleep(DELAY_ENTRE_ACOES_S)
        logger.info("Finalizando execução. Efetuando logout no sistema...")

        btn_sair = page.locator("br-button:has-text('Sair'), button:has-text('Sair'), a:has-text('Sair')").first
        if not btn_sair.is_visible():
            btn_sair = page.get_by_text(re.compile(r"^\s*Sair\s*$", re.IGNORECASE)).first

        if btn_sair.is_visible():
            time.sleep(DELAY_ENTRE_ACOES_S)
            btn_sair.click()
            logger.info("Clique no botão 'Sair' realizado.")
            time.sleep(DELAY_ENTRE_ACOES_S)

            btn_confirmar = page.locator("button:has-text('Sim'), button:has-text('Confirmar'), br-button:has-text('Sim')").first
            try:
                if btn_confirmar.is_visible(timeout=2000):
                    btn_confirmar.click()
                    time.sleep(DELAY_ENTRE_ACOES_S)
            except Exception:
                pass

            time.sleep(1.5)
        else:
            logger.info("Botão 'Sair' não localizado no cabeçalho.")
    except Exception as e:
        logger.warning(f"Aviso durante tentativa de logout: {e}")

    for arq in arquivos_pendentes_excluir:
        try:
            Path(arq).unlink(missing_ok=True)
        except Exception:
            pass

    logger.info("Fechando o navegador Chrome...")
    fechar_navegador(browser, context, page)
    if playwright_mgr:
        playwright_mgr.stop()
    logger.info("Automação finalizada com sucesso.")


if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(message)s",
        datefmt="%d/%m/%Y %H:%M:%S"
    )
    executar_robo_hibrido()

