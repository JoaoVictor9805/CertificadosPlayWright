import os
import re
import time
import shutil
import subprocess
import threading
import ctypes
from datetime import datetime
from pathlib import Path

import openpyxl
# pyrefly: ignore[missing-import]
from docx import Document
from pywinauto import Desktop, Application
from pywinauto.keyboard import send_keys
from PIL import ImageGrab

# ==============================================================================
# CONFIGURAÇÕES E CAMINHOS (Baseados no RoboSocietario-Principal-v10)
# ==============================================================================
BASE_DIR = Path(r"Z:\Temp\SOCIETARIO\AutomacaoCertificados")
LOG_PATH = BASE_DIR / "logs.docx"
EXCEL_PATH = BASE_DIR / "CNPJs-e-CPFs.xlsx"
SCREENSHOTS_DIR = BASE_DIR / "CapturasDeTela"
DOWNLOADS_DIR = Path.home() / "Downloads"

URL_LOGIN = "https://det.sit.trabalho.gov.br/login"
URL_SERVICOS = "https://det.sit.trabalho.gov.br/servicos"

# ModoTeste = True: considera notificações com data <= hoje (testes).
# Em produção, trocar para False (data >= hoje).
MODO_TESTE = True

# Configurações do Google Chrome Real (UIA nativo evita CAPTCHA do gov.br)
CHROME_PATH = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
CHROME_USER_DATA = Path(os.environ["LOCALAPPDATA"]) / "Google" / "Chrome" / "User Data_RoboDET"


# ==============================================================================
# GERENCIADOR DE LOGS EM WORD (Substitui RoboSocietario-Log-v10)
# ==============================================================================
class WordLogger:
    def __init__(self, doc_path: Path):
        self.doc_path = doc_path
        self.doc_path.parent.mkdir(parents=True, exist_ok=True)
        if self.doc_path.exists():
            try:
                self.doc_path.unlink()
            except OSError:
                pass
        self.doc = Document()

    def _now_str(self) -> str:
        return datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    def log_etapa(self, item_id: str, etapa: str, msg: str):
        texto = f"{self._now_str()} - [{item_id}] [{etapa}] {msg}"
        print(texto)
        try:
            self.doc.add_paragraph(texto)
            self.doc.save(str(self.doc_path))
        except Exception as e:
            print(f"[AVISO] Falha ao salvar logs.docx: {e}")

    def log_raw(self, msg: str):
        texto = f"{self._now_str()} - {msg}"
        print(texto)
        try:
            self.doc.add_paragraph(texto)
            self.doc.save(str(self.doc_path))
        except Exception as e:
            print(f"[AVISO] Falha ao salvar logs.docx: {e}")


# ==============================================================================
# FUNÇÕES AUXILIARES (Pastas, Janelas e Alertas)
# ==============================================================================
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
                print(f"[AVISO] Não foi possível limpar {item}: {e}")


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


def exibir_alerta_planilha_inexistente():
    mensagem = (
        "A Planilha de referência para consulta de CNPJs e CPFs não existia.\n\n"
        f"A planilha foi criada na pasta CERTIFICADODET no caminho {BASE_DIR}\n"
        "Por favor, a preencha com os valores desejados, um abaixo do outro e execute o robô novamente"
    )
    ctypes.windll.user32.MessageBoxW(0, mensagem, "Planilha de Referência não existe", 0x30)


def confirmar_popup_certificado_nativo(timeout_segundos: int = 15):
    """
    Monitora em background a janela nativa do Windows/Navegador de seleção de
    Certificado Digital e clica em 'OK' (UIAutomation nativo).
    """
    fim = time.time() + timeout_segundos
    while time.time() < fim:
        try:
            desktop = Desktop(backend="uia")
            for win in desktop.windows():
                titulo = win.window_text()
                if any(k in titulo for k in ["gov.br", "Acesse sua conta", "Selecionar um certificado", "certificado"]):
                    btns_ok = win.descendants(control_type="Button", title="OK")
                    if btns_ok:
                        btns_ok[0].click_input()
                        return True
        except Exception:
            pass
        time.sleep(0.5)
    return False


# ==============================================================================
# FUNÇÕES DE INTERAÇÃO COM CHROME VIA PYWINAUTO (UIA)
# ==============================================================================
def iniciar_chrome_pywinauto():
    """
    Inicia o Google Chrome com perfil dedicado e com acessibilidade forçada.
    Não utiliza flags de automação (evita detecção de CAPTCHA no gov.br).
    """
    CHROME_USER_DATA.mkdir(parents=True, exist_ok=True)
    cmd = [
        CHROME_PATH,
        f"--user-data-dir={CHROME_USER_DATA}",
        "--force-renderer-accessibility",
        "--start-maximized",
        "--no-first-run",
        "--no-default-browser-check",
        "--disable-blink-features=AutomationControlled",
        URL_LOGIN,
    ]
    subprocess.Popen(cmd)
    time.sleep(3)


def focar_janela_segura(win):
    """Foca a janela usando Win32 API para evitar COMError do Chrome UIA."""
    try:
        hwnd = win.handle
        if ctypes.windll.user32.IsIconic(hwnd):
            ctypes.windll.user32.ShowWindow(hwnd, 9)  # SW_RESTORE
        ctypes.windll.user32.SetForegroundWindow(hwnd)
        time.sleep(0.5)
    except Exception:
        pass


def obter_janela_chrome(timeout_segundos: int = 30):
    """Localiza e foca a janela principal do Google Chrome via UIA."""
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


def obter_documento_web(chrome_win, timeout_segundos: int = 30):
    """Obtém o elemento Document da página web no Chrome via UIA."""
    fim = time.time() + timeout_segundos
    while time.time() < fim:
        try:
            docs = chrome_win.descendants(control_type="Document")
            if docs:
                return docs[0]
        except Exception:
            pass
        time.sleep(1)
    return chrome_win


def navegar_para_url(chrome_win, url: str):
    """Navega para a URL desejada usando a barra de endereços do Chrome (Ctrl+L)."""
    focar_janela_segura(chrome_win)
    time.sleep(0.5)
    send_keys("^l")
    time.sleep(0.3)
    send_keys(url + "{ENTER}", with_spaces=True)
    time.sleep(3)


def encontrar_elemento(pai, timeout_segundos: int = 15, **kwargs):
    """Busca elemento por critérios UIA com tentativas e tratamento de exceção."""
    fim = time.time() + timeout_segundos
    while time.time() < fim:
        try:
            elems = pai.descendants(**kwargs)
            if elems:
                return elems[0]
        except Exception:
            pass
        time.sleep(0.5)
    return None


def clicar_elemento_seguro(elem):
    """Clica em um elemento UIA via click_input, invoke ou coordenadas do mouse."""
    try:
        elem.click_input()
        return True
    except Exception:
        try:
            elem.invoke()
            return True
        except Exception:
            try:
                rect = elem.rectangle()
                import pywinauto.mouse
                pywinauto.mouse.click(coords=(rect.mid_point().x, rect.mid_point().y))
                return True
            except Exception:
                return False


def tirar_print_recortado(elem, caminho_destino: Path):
    """Tira um print do retângulo exato do elemento via capture_as_image ou ImageGrab."""
    try:
        img = elem.capture_as_image()
        img.convert("RGB").save(str(caminho_destino), "JPEG")
        return True
    except Exception:
        try:
            rect = elem.rectangle()
            bbox = (rect.left, rect.top, rect.right, rect.bottom)
            img = ImageGrab.grab(bbox=bbox)
            img.convert("RGB").save(str(caminho_destino), "JPEG")
            return True
        except Exception:
            return False


def tirar_print_janela(chrome_win, caminho_destino: Path):
    """Tira um print completo da janela do navegador."""
    try:
        img = chrome_win.capture_as_image()
        img.convert("RGB").save(str(caminho_destino), "JPEG")
        return True
    except Exception:
        try:
            rect = chrome_win.rectangle()
            bbox = (rect.left, rect.top, rect.right, rect.bottom)
            img = ImageGrab.grab(bbox=bbox)
            img.convert("RGB").save(str(caminho_destino), "JPEG")
            return True
        except Exception:
            return False


def aguardar_e_mover_download(downloads_dir: Path, pasta_destino: Path, timeout_segundos: int = 60) -> Path | None:
    """
    Monitora a pasta de downloads padrão do Windows ou caixa de diálogo Salvar Como
    para mover o arquivo recém-baixado para a pasta ArquivosBaixados.
    """
    # 1. Verifica se abriu a janela nativa 'Salvar como'
    desktop = Desktop(backend="uia")
    for _ in range(5):
        try:
            for win in desktop.windows():
                if "Salvar como" in win.window_text():
                    campo_nome = win.descendants(control_type="Edit")
                    if campo_nome:
                        nome_arquivo = campo_nome[0].window_text()
                        destino = pasta_destino / nome_arquivo
                        campo_nome[0].set_text(str(destino))
                        time.sleep(0.5)
                        send_keys("{ENTER}")
                        time.sleep(2)
                        return destino
        except Exception:
            pass
        time.sleep(0.5)

    # 2. Monitora surgimento de arquivo .pdf recente na pasta Downloads
    fim = time.time() + timeout_segundos
    arquivos_iniciais = set(downloads_dir.glob("*"))

    while time.time() < fim:
        arquivos_atuais = set(downloads_dir.glob("*"))
        novos = arquivos_atuais - arquivos_iniciais
        for f in novos:
            # Ignora downloads em andamento (.crdownload)
            if f.suffix.lower() == ".pdf":
                time.sleep(1)
                destino_final = pasta_destino / f.name
                shutil.move(str(f), str(destino_final))
                return destino_final
        time.sleep(1)

    return None


# ==============================================================================
# FLUXO PRINCIPAL DA AUTOMAÇÃO COM PYWINAUTO
# ==============================================================================
def executar_robo_pywinauto():
    logger = WordLogger(LOG_PATH)
    hoje_date = datetime.now().date()

    # 1. Verificação da Planilha de Entrada
    if not EXCEL_PATH.exists():
        logger.log_raw("Planilha de CNPJs não encontrada.")
        wb = openpyxl.Workbook()
        wb.save(str(EXCEL_PATH))
        wb.close()
        exibir_alerta_planilha_inexistente()
        return

    try:
        wb = openpyxl.load_workbook(str(EXCEL_PATH), data_only=True)
        sheet = wb.active
        excel_data = [
            str(cell.value).strip()
            for row in sheet.iter_rows(values_only=False)
            for cell in row
            if cell.value is not None and str(cell.value).strip() != ""
        ]
        wb.close()

        limpar_ou_criar_pasta(SCREENSHOTS_DIR)

        # 2. Inicialização do Google Chrome Real (sem flags de automação)
        iniciar_chrome_pywinauto()
        chrome_win = obter_janela_chrome(timeout_segundos=30)
        doc = obter_documento_web(chrome_win, timeout_segundos=30)

        # Verifica botão "Entrar com gov.br"
        btn_govbr = encontrar_elemento(doc, timeout_segundos=15, control_type="Button", title_re=r".*Entrar com.*gov.*")
        if not btn_govbr:
            btn_govbr = encontrar_elemento(doc, timeout_segundos=10, title_re=r".*Entrar com.*gov.*")

        if not btn_govbr:
            # Tenta via navegação direta caso não tenha carregado
            navegar_para_url(chrome_win, URL_LOGIN)
            time.sleep(3)
            doc = obter_documento_web(chrome_win, timeout_segundos=20)
            btn_govbr = encontrar_elemento(doc, timeout_segundos=15, title_re=r".*Entrar com.*gov.*")

        if not btn_govbr:
            logger.log_raw("Não foi possível carregar a página de login do DET.")
            return

        clicar_elemento_seguro(btn_govbr)
        time.sleep(2)

        # Atualiza documento e localiza botão "Seu certificado digital"
        doc = obter_documento_web(chrome_win, timeout_segundos=20)
        btn_cert = encontrar_elemento(
            doc, timeout_segundos=30, control_type="Button", title_re=".*certificado digital.*"
        )

        # Inicia thread em segundo plano para confirmar o pop-up nativo do Windows
        thread_popup = threading.Thread(target=confirmar_popup_certificado_nativo, args=(15,), daemon=True)
        thread_popup.start()

        if btn_cert:
            clicar_elemento_seguro(btn_cert)
        else:
            send_keys("{ENTER}")

        # Aguarda carregar a tela de Serviços (identificada pelo card CAIXA POSTAL)
        time.sleep(5)
        doc = obter_documento_web(chrome_win, timeout_segundos=40)
        card_caixa_postal = encontrar_elemento(
            doc, timeout_segundos=40, title_re=".*CAIXA POSTAL.*"
        )
        if not card_caixa_postal:
            logger.log_raw("Não foi possível carregar a tela de serviços após autenticação do certificado.")
            return

    except Exception as e:
        logger.log_raw(f"Erro fatal no robô (fase de preparação ou login): {e}")
        raise

    # 3. Loop de Processamento de CPFs / CNPJs
    perfil = 0
    falhas_seguidas = 0
    etapa = "Início"

    for current_item in excel_data:
        if falhas_seguidas >= 3:
            logger.log_raw("Execução interrompida: 3 falhas seguidas. Os itens restantes não foram processados.")
            break

        item_id = current_item.strip()
        if not item_id:
            continue

        try:
            etapa = "Lendo o item da planilha"
            logger.log_etapa(item_id, etapa, "Iniciando o processamento do item")

            etapa = "Aguardando a tela de serviços"
            doc = obter_documento_web(chrome_win, timeout_segundos=30)
            card_caixa_postal = encontrar_elemento(doc, timeout_segundos=30, title_re=".*CAIXA POSTAL.*")

            etapa = "Trocando de perfil"
            btn_trocar_perfil = encontrar_elemento(doc, timeout_segundos=20, control_type="Button", title_re=".*Trocar Perfil.*")
            if btn_trocar_perfil:
                clicar_elemento_seguro(btn_trocar_perfil)
            time.sleep(3)

            # Clica no dropdown de perfil e navega com teclado (réplica do PAD)
            doc = obter_documento_web(chrome_win, timeout_segundos=10)
            dropdown = encontrar_elemento(doc, timeout_segundos=10, control_type="ComboBox")
            if dropdown:
                clicar_elemento_seguro(dropdown)
            else:
                send_keys("{TAB}")
            time.sleep(1)

            if perfil == 0:
                perfil += 1
                send_keys("{DOWN}")
                time.sleep(1.5)
                send_keys("{ENTER}")
                time.sleep(1.5)
                send_keys("{TAB}")
            else:
                send_keys("{DOWN}")
                time.sleep(1.5)
                send_keys("{DOWN}")
                time.sleep(1.5)
                send_keys("{ENTER}")
                time.sleep(1.5)
                send_keys("{TAB}")

            time.sleep(0.5)
            send_keys(item_id, with_spaces=True)
            time.sleep(0.5)

            etapa = "Selecionando o perfil e validando CPF/CNPJ"
            btn_selecionar = encontrar_elemento(doc, timeout_segundos=10, control_type="Button", title_re=".*Selecionar.*")
            if btn_selecionar:
                clicar_elemento_seguro(btn_selecionar)
            else:
                send_keys("{ENTER}")
            time.sleep(3)

            # Verifica mensagem de documento inválido
            doc = obter_documento_web(chrome_win, timeout_segundos=5)
            msg_invalido = encontrar_elemento(doc, timeout_segundos=3, title_re=r".*(CNPJ|CPF)\s*inv[aá]lido.*")
            if msg_invalido:
                btn_cancelar = encontrar_elemento(doc, timeout_segundos=5, control_type="Button", title_re=".*Cancelar.*")
                if btn_cancelar:
                    clicar_elemento_seguro(btn_cancelar)
                time.sleep(2)
                logger.log_raw(f"Documento inválido, item ignorado: {item_id}")
                continue

            # Criação de pastas do item
            etapa = "Criando a pasta do item"
            pasta_item = SCREENSHOTS_DIR / item_id
            pasta_caixa_postal = pasta_item / "CaixaPostal"
            pasta_notificacao = pasta_item / "Notificacao"
            pasta_arquivos_baixados = pasta_item / "ArquivosBaixados"

            pasta_caixa_postal.mkdir(parents=True, exist_ok=True)
            pasta_notificacao.mkdir(parents=True, exist_ok=True)
            pasta_arquivos_baixados.mkdir(parents=True, exist_ok=True)
            time.sleep(1)

            # --- CAIXA POSTAL ---
            etapa = "Abrindo a Caixa Postal"
            doc = obter_documento_web(chrome_win, timeout_segundos=20)
            card_cp = encontrar_elemento(doc, timeout_segundos=20, title_re=".*CAIXA POSTAL.*")
            if card_cp:
                clicar_elemento_seguro(card_cp)
            time.sleep(3)

            formatted_dt = datetime.now().strftime("%d%m%Y_%H%M%S")
            doc = obter_documento_web(chrome_win, timeout_segundos=20)

            # Marca 'Exibir apenas não lidas'
            chk_nao_lidas = encontrar_elemento(doc, timeout_segundos=10, title_re=".*Exibir apenas n[aã]o lidas.*")
            if chk_nao_lidas:
                clicar_elemento_seguro(chk_nao_lidas)
            time.sleep(2)

            etapa = "Tirando o print da Caixa Postal"
            # Verifica se há tabela / linhas com mensagens
            doc = obter_documento_web(chrome_win, timeout_segundos=10)
            tabela_cp = encontrar_elemento(doc, timeout_segundos=5, control_type="Table")
            linhas_cp = tabela_cp.descendants(control_type="DataItem") if tabela_cp else []

            if linhas_cp:
                caminho_print_cp = pasta_caixa_postal / f"{item_id}_{formatted_dt}.jpg"
                tirar_print_janela(chrome_win, caminho_print_cp)
                logger.log_etapa(item_id, etapa, "Print da Caixa Postal realizado: existe mensagem não lida.")

            # Retorna para a tela de Serviços
            navegar_para_url(chrome_win, URL_SERVICOS)
            time.sleep(2)

            # --- NOTIFICAÇÃO ---
            etapa = "Abrindo a tela de Notificação"
            doc = obter_documento_web(chrome_win, timeout_segundos=30)
            card_notif = encontrar_elemento(doc, timeout_segundos=30, title_re=".*NOTIFICA[CÇ][AÃ]O.*")
            if card_notif:
                clicar_elemento_seguro(card_notif)
            time.sleep(4)

            etapa = "Lendo a lista de notificações"
            doc = obter_documento_web(chrome_win, timeout_segundos=20)

            # Localiza todas as linhas de dados da tabela de notificações
            linhas_notif = []
            try:
                # Procura por linhas do datatable (DataItem, Custom ou Row)
                candidatos = doc.descendants(control_type="DataItem")
                if not candidatos:
                    candidatos = doc.descendants(control_type="Custom")
                # Filtra apenas linhas que contenham formato de data
                for c in candidatos:
                    txt = c.window_text()
                    if re.search(r"\d{2}/\d{2}/\d{4}", txt):
                        linhas_notif.append(c)
            except Exception:
                pass

            if linhas_notif:
                qtd_linhas = len(linhas_notif)
                logger.log_etapa(
                    item_id,
                    etapa,
                    f"Lista de notificações lida: {qtd_linhas} linha(s) extraída(s). Data de hoje considerada: {hoje_date.strftime('%d/%m/%Y')}",
                )

                dados_linhas = []
                existe_notificacao_valida = False

                for idx, row_elem in enumerate(linhas_notif):
                    texto_linha = row_elem.window_text()
                    match_data = re.search(r"\b(\d{2}/\d{2}/\d{4})\b", texto_linha)
                    texto_data = match_data.group(1) if match_data else ""
                    try:
                        data_notif = datetime.strptime(texto_data, "%d/%m/%Y").date()
                        valida = (data_notif <= hoje_date) if MODO_TESTE else (data_notif >= hoje_date)
                    except ValueError:
                        data_notif = None
                        valida = False

                    if valida:
                        existe_notificacao_valida = True
                    dados_linhas.append((idx, texto_data, valida, row_elem))

                if not existe_notificacao_valida:
                    logger.log_etapa(item_id, etapa, "Nenhuma notificação possui data válida e o print foi ignorado.")

                # Processa cada linha válida
                for idx, texto_data, linha_valida, row_elem in dados_linhas:
                    etapa = f"Lendo a data da linha {idx}"
                    if linha_valida:
                        # Rola/foca a linha antes de tirar o print
                        etapa = f"Rolando a página até a linha {idx}"
                        try:
                            row_elem.set_focus()
                        except Exception:
                            pass
                        time.sleep(0.5)

                        # Print recortado do retângulo da linha específica
                        etapa = f"Tirando print recortado da linha {idx}"
                        caminho_print_linha = pasta_notificacao / f"{item_id}_linha_{idx}_{formatted_dt}.jpg"
                        sucesso_print = tirar_print_recortado(row_elem, caminho_print_linha)
                        if sucesso_print:
                            logger.log_etapa(
                                item_id,
                                etapa,
                                f"Print recortado da linha {idx} (data {texto_data}) realizado com sucesso.",
                            )

                        # Verifica botão de download na linha
                        etapa = f"Localizando o botão de download da linha {idx}"
                        botoes_linha = row_elem.descendants(control_type="Button")
                        btn_download = None

                        # Procura botão com título de download ou pega o segundo botão da linha
                        for b in botoes_linha:
                            t = b.window_text().lower()
                            if any(k in t for k in ["baixar", "download", "arquivo", "salvar"]):
                                btn_download = b
                                break
                        if not btn_download and len(botoes_linha) >= 2:
                            btn_download = botoes_linha[1]

                        if btn_download:
                            logger.log_etapa(item_id, etapa, f"Linha {idx} (data {texto_data}): clicando no botão de download")
                            clicar_elemento_seguro(btn_download)

                            etapa = f"Aguardando o download da linha {idx}"
                            arquivo_baixado = aguardar_e_mover_download(
                                DOWNLOADS_DIR, pasta_arquivos_baixados, timeout_segundos=40
                            )
                            if arquivo_baixado:
                                logger.log_etapa(
                                    item_id,
                                    etapa,
                                    f"Linha {idx}: arquivo baixado e movido para ArquivosBaixados ({arquivo_baixado})",
                                )
                            else:
                                logger.log_etapa(
                                    item_id,
                                    etapa,
                                    f"AVISO: download não detectado em 40s (linha {idx})",
                                )
                        else:
                            logger.log_etapa(
                                item_id,
                                etapa,
                                f"Linha {idx} (data {texto_data}): notificação válida sem botão de download (apenas leitura).",
                            )
                    else:
                        logger.log_etapa(
                            item_id,
                            etapa,
                            f"Linha {idx} ignorada: data {texto_data} fora do critério (ModoTeste={MODO_TESTE}, hoje={hoje_date.strftime('%d/%m/%Y')})",
                        )
            else:
                logger.log_etapa(item_id, etapa, "Nenhuma notificação listada na tela de Notificação")

            # Retorna à tela de Serviços
            navegar_para_url(chrome_win, URL_SERVICOS)

            # Limpeza de pastas vazias e contabilização
            etapa = "Limpando subpastas vazias"
            remover_pastas_vazias_do_item(pasta_item)

            etapa = "Conferindo os arquivos e limpando a pasta vazia"
            arquivos_guardados = list(pasta_item.rglob("*")) if pasta_item.exists() else []
            qtd_arquivos = sum(1 for f in arquivos_guardados if f.is_file())

            logger.log_raw(
                f"SUCESSO: O item '{item_id}' foi verificado. Arquivos guardados (prints + downloads): {qtd_arquivos}"
            )
            falhas_seguidas = 0

        except Exception as e:
            falhas_seguidas += 1
            logger.log_raw(f"Erro no item '{current_item}' na etapa '{etapa}': {e}")
            try:
                navegar_para_url(chrome_win, URL_SERVICOS)
            except Exception:
                pass

    # 4. Encerramento e Logout
    try:
        doc = obter_documento_web(chrome_win, timeout_segundos=10)
        btn_sair = encontrar_elemento(doc, timeout_segundos=10, control_type="Button", title_re=".*Sair do sistema.*")
        if btn_sair:
            clicar_elemento_seguro(btn_sair)
    except Exception:
        pass


if __name__ == "__main__":
    executar_robo_pywinauto()
