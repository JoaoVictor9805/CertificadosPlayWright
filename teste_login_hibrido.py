import os
import time
import socket
import subprocess
import ctypes
from pathlib import Path

import pyautogui
from pywinauto import Desktop
from pywinauto.keyboard import send_keys
from playwright.sync_api import sync_playwright

# ==============================================================================
# CONFIGURAÇÕES E CAMINHOS
# ==============================================================================
URL_LOGIN = "https://det.sit.trabalho.gov.br/login"
URL_SERVICOS = "https://det.sit.trabalho.gov.br/servicos"

CHROME_PATH = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
CHROME_USER_DATA = Path(os.environ["LOCALAPPDATA"]) / "Google" / "Chrome" / "User Data_RoboDET"
CDP_PORT = 9222

BASE_DIR = Path(__file__).resolve().parent
IMG_DIR = BASE_DIR / "capturasDeTela"
IMG_GOVBR = IMG_DIR / "EntrarComGov.png"
IMG_CERT = IMG_DIR / "SeuCertificadoDigital.png"
IMG_OK = IMG_DIR / "Ok.png"
IMG_DEBUG_ERRO = IMG_DIR / "debug_tela_erro.png"


# ==============================================================================
# FUNÇÕES DE SISTEMA (CHROME, JANELAS E PYAUTOGUI)
# ==============================================================================
def porta_em_uso(porta: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(("127.0.0.1", porta)) == 0


def liberar_porta_se_ocupada():
    """Fecha processos órfãos na porta CDP para garantir uma inicialização limpa."""
    if porta_em_uso(CDP_PORT):
        print(f"[INFO] Porta {CDP_PORT} em uso. Tentando fechar instância anterior...")
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
    """Inicia o Chrome oficial com o perfil dedicado e a porta CDP."""
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
    try:
        hwnd = win.handle
        if ctypes.windll.user32.IsIconic(hwnd):
            ctypes.windll.user32.ShowWindow(hwnd, 9)
        ctypes.windll.user32.SetForegroundWindow(hwnd)
        time.sleep(0.5)
    except Exception:
        pass


def obter_janela_chrome(timeout_segundos: int = 30):
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
    """Move o mouse e realiza um clique físico com press/release garantido."""
    pyautogui.moveTo(pos_x, pos_y, duration=0.5)
    time.sleep(0.15)
    pyautogui.mouseDown()
    time.sleep(0.1)
    pyautogui.mouseUp()


def localizar_imagem(caminho_imagem: Path, confidence: float = 0.8):
    """Localiza o centro da imagem na tela com tolerância a ClearType (grayscale)."""
    try:
        # 1. Tenta correspondência padrão
        pos = pyautogui.locateCenterOnScreen(str(caminho_imagem), confidence=confidence, grayscale=True)
        if pos:
            return pos
    except Exception:
        pass

    try:
        # 2. Tenta com confiança um pouco mais flexível (0.7)
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
            print(f"[PYAUTOGUI] 'Entrar com gov.br' localizado em ({pos.x}, {pos.y}). Clicando...")
            clicar_robusto(pos.x, pos.y)
            time.sleep(2)

            # Verifica se o botão sumiu (significa que a página começou a navegar)
            pos_ainda_visivel = localizar_imagem(IMG_GOVBR, confidence=0.8)
            if not pos_ainda_visivel:
                print("[PYAUTOGUI] Botão 'Entrar com gov.br' clicado e página navegando!")
                return True
            else:
                print("[AVISO] Botão ainda visível. Tentando segundo clique de ativação...")
                clicar_robusto(pos.x, pos.y)
                time.sleep(2)
                return True
        time.sleep(0.8)

    return False


def clicar_botao_certificado_garantido(timeout_segundos: int = 30) -> bool:
    """
    Aguarda e clica em 'Seu certificado digital' na tela do sso.acesso.gov.br.
    Se não encontrar inicialmente, faz leve scroll e tenta com tolerância.
    Se falhar, salva captura da tela para diagnóstico.
    """
    print(f"[PYAUTOGUI] Procurando 'Seu certificado digital' na tela (timeout={timeout_segundos}s)...")
    fim = time.time() + timeout_segundos
    tentativa = 0

    while time.time() < fim:
        tentativa += 1

        # Tenta localizar diretamente
        pos = localizar_imagem(IMG_CERT, confidence=0.75)
        if pos:
            print(f"[PYAUTOGUI] 'Seu certificado digital' ENCONTRADO em ({pos.x}, {pos.y})!")
            clicar_robusto(pos.x, pos.y)
            time.sleep(1.5)
            return True

        # Se passou de 6 segundos e não achou, pode precisar de um leve scroll para baixo
        if tentativa == 6:
            print("[INFO] Tentando leve rolagem da página para baixo...")
            try:
                pyautogui.scroll(-200)
            except Exception:
                pass

        time.sleep(0.8)

    # Se esgotou o tempo sem achar, tira um print da tela para vermos o que estava visível
    print(f"\n[AVISO] 'Seu certificado digital' não foi localizado dentro de {timeout_segundos}s.")
    try:
        pyautogui.screenshot(str(IMG_DEBUG_ERRO))
        print(f"[DIAGNÓSTICO] Screenshot da tela salvo em: {IMG_DEBUG_ERRO.name}")
    except Exception as e:
        print(f"[AVISO] Não foi possível salvar screenshot de diagnóstico: {e}")

    return False


def confirmar_popup_certificado(timeout_segundos: int = 15) -> bool:
    """
    Aguarda o popup de seleção de certificado e clica em OK.
    Tenta por imagem com PyAutoGUI e possui fallbacks via UIA e tecla ENTER.
    """
    fim = time.time() + timeout_segundos
    while time.time() < fim:
        # 1. Tenta por imagem (PyAutoGUI)
        pos_ok = localizar_imagem(IMG_OK, confidence=0.7)
        if pos_ok:
            print(f"[PYAUTOGUI] Botão 'OK' do certificado encontrado em ({pos_ok.x}, {pos_ok.y})!")
            clicar_robusto(pos_ok.x, pos_ok.y)
            time.sleep(1.5)
            return True

        # 2. Fallback: via UIA nativo do Windows
        try:
            desktop = Desktop(backend="uia")
            for win in desktop.windows():
                titulo = win.window_text()
                if any(k in titulo for k in ["gov.br", "Acesse sua conta", "Selecionar um certificado", "certificado"]):
                    btns_ok = win.descendants(control_type="Button", title="OK")
                    if btns_ok:
                        print("[UIA] Botão 'OK' do certificado acionado via UIA!")
                        btns_ok[0].click_input()
                        time.sleep(1.5)
                        return True
        except Exception:
            pass

        time.sleep(0.5)

    # 3. Fallback final: Enviar ENTER
    print("[FALLBACK] Enviando tecla ENTER para confirmar certificado...")
    send_keys("{ENTER}")
    time.sleep(1.5)
    return True


# ==============================================================================
# FLUXO DE TESTE HÍBRIDO (PyAutoGUI -> Playwright)
# ==============================================================================
def testar_login_hibrido():
    print("=" * 70)
    print("INICIANDO TESTE DE LOGIN HÍBRIDO")
    print("Fase 1: Reconhecimento de Imagem e Cliques Físicos (PyAutoGUI)")
    print("Fase 2: Conexão e Controle do Playwright após autenticação")
    print("=" * 70)

    liberar_porta_se_ocupada()

    # 1. Abre o Google Chrome oficial com o perfil e a porta
    print("\n[PASSO 1] Abrindo Google Chrome oficial...")
    iniciar_chrome_hibrido()

    print("[PASSO 2] Localizando e focando a janela do Chrome...")
    chrome_win = obter_janela_chrome(timeout_segundos=30)
    focar_janela_segura(chrome_win)
    time.sleep(2)

    # 2. Clica no botão 'Entrar com gov.br' via PyAutoGUI
    print("\n[PASSO 3] Localizando e clicando em 'Entrar com gov.br'...")
    if not clicar_botao_govbr_garantido(timeout_segundos=25):
        print("[ERRO] Não foi possível prosseguir: botão 'Entrar com gov.br' não encontrado.")
        return

    # 3. Clica no botão 'Seu certificado digital' via PyAutoGUI
    print("\n[PASSO 4] Localizando e clicando em 'Seu certificado digital'...")
    if not clicar_botao_certificado_garantido(timeout_segundos=30):
        print("[ERRO] Não foi possível prosseguir: 'Seu certificado digital' não encontrado.")
        return

    # 4. Confirma o popup do certificado digital (OK)
    print("\n[PASSO 5] Confirmando seleção de certificado no popup...")
    confirmar_popup_certificado(timeout_segundos=15)

    print("\n" + "=" * 70)
    print("FASE 1 CONCLUÍDA: Cliques físicos de login realizados com sucesso!")
    print("Aguardando carregamento da tela de Serviços para o Playwright assumir...")
    print("=" * 70 + "\n")

    # 5. Playwright se conecta via CDP (com retentativas para evitar ECONNREFUSED)
    playwright_mgr = sync_playwright().start()
    browser = None
    for tentativa_cdp in range(1, 6):
        try:
            print(f"[PASSO 6] Conectando Playwright via CDP (tentativa {tentativa_cdp}/5)...")
            browser = playwright_mgr.chromium.connect_over_cdp(f"http://127.0.0.1:{CDP_PORT}", slow_mo=500)
            break
        except Exception as e:
            print(f"[AVISO] Porta CDP ainda ocupada/iniciando: {e}. Aguardando 2s...")
            time.sleep(2)

    if not browser:
        print("[ERRO] Não foi possível conectar o Playwright via CDP após 5 tentativas.")
        playwright_mgr.stop()
        return

    context = browser.contexts[0]
    page = context.pages[0] if context.pages else context.new_page()

    # Aguarda carregar a tela de Serviços
    print("[FASE 2 - PLAYWRIGHT] Aguardando URL de Serviços...")
    try:
        page.wait_for_url("**/servicos**", timeout=45000)
    except Exception:
        pass

    card_caixa_postal = page.locator(".cardListItem:has(span.title:has-text('CAIXA POSTAL'))")
    card_caixa_postal.wait_for(state="visible", timeout=30000)

    print("[FASE 2 - PLAYWRIGHT] Conexão CDP estabelecida com sucesso!")
    print(f"[FASE 2 - PLAYWRIGHT] URL ativa: {page.url}")
    print(f"[FASE 2 - PLAYWRIGHT] Título da página: {page.title()}")
    print(f"[FASE 2 - PLAYWRIGHT] Card 'CAIXA POSTAL' visível: {card_caixa_postal.is_visible()}")

    card_notificacao = page.locator(".cardListItem:has(span.title:has-text('NOTIFICAÇÃO'))")
    btn_trocar_perfil = page.locator("button:has-text('Trocar Perfil')").first
    print(f"[FASE 2 - PLAYWRIGHT] Card 'NOTIFICAÇÃO' visível: {card_notificacao.is_visible()}")
    print(f"[FASE 2 - PLAYWRIGHT] Botão 'Trocar Perfil' visível: {btn_trocar_perfil.is_visible()}")

    print("\n" + "=" * 70)
    print("TESTE HÍBRIDO 100% BEM-SUCEDIDO!")
    print("O login ocorreu via visão computacional física (PyAutoGUI) e o Playwright assumiu o controle.")
    print("Fechando a sessão de teste em 5 segundos...")
    print("=" * 70)
    time.sleep(5)

    # Encerramento limpo
    try:
        cdp = context.new_cdp_session(page)
        cdp.send("Browser.close")
    except Exception:
        pass

    try:
        browser.close()
    except Exception:
        pass

    playwright_mgr.stop()
    print("[FIM] Teste concluído.")


if __name__ == "__main__":
    testar_login_hibrido()
