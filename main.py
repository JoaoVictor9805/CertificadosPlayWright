import argparse
import logging
import os
import sys
import warnings
from datetime import datetime
from pathlib import Path

# Define STA (Single-Threaded Apartment) antes de qualquer biblioteca COM/pywinauto ser importada
if not hasattr(sys, "coinit_flags"):
    sys.coinit_flags = 2
warnings.filterwarnings("ignore", message=".*Apply externally defined coinit_flags.*")

# Adiciona o diretório raiz e 'src' ao sys.path para garantir resolução de módulos
ROOT_DIR = Path(__file__).resolve().parent
SRC_DIR = ROOT_DIR / "src"

if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from src.automacao_hibrida_final import executar_robo_hibrido


def configurar_logger(pasta_base: Path) -> Path:
    """Configura o logger 'RoboDET' para gravar no console e em arquivo .txt."""
    pasta_logs = pasta_base / "Logs"
    pasta_logs.mkdir(parents=True, exist_ok=True)

    dt_str = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    arquivo_log = pasta_logs / f"log_execucao_{dt_str}.txt"

    logger = logging.getLogger("RoboDET")
    logger.setLevel(logging.INFO)

    # Evita duplicação de handlers se reexecutado
    logger.handlers.clear()

    formatter = logging.Formatter(
        fmt="[%(asctime)s] [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    # Handler do Console (Terminal)
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    # Handler do Arquivo .txt
    file_handler = logging.FileHandler(str(arquivo_log), encoding="utf-8")
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    return arquivo_log


def obter_argumentos() -> tuple[Path, Optional[str]]:
    """
    Obtém o caminho da pasta base e argumentos opcionais de linha de comando.
    Se o caminho não for informado, encerra imediatamente com erro sem interação com o usuário.
    """
    parser = argparse.ArgumentParser(
        description="Automação DET - Consulta de Notificações e Certificados (Linha de Comando)"
    )
    parser.add_argument(
        "caminho_posicional",
        nargs="?",
        default=None,
        help="Caminho da pasta da automação (onde fica o CNPJs.xlsx e CapturasDeTela)",
    )
    parser.add_argument(
        "-c",
        "--caminho",
        dest="caminho_flag",
        default=None,
        help="Caminho alternativo da pasta da automação",
    )
    parser.add_argument(
        "-p",
        "--pendentes",
        dest="apenas_pendentes",
        action="store_true",
        help="Executa a automação processando a planilha 'CNPJs_Pendentes.xlsx'",
    )
    parser.add_argument(
        "-a",
        "--arquivo",
        dest="arquivo_planilha",
        default=None,
        help="Nome ou caminho de uma planilha Excel específica (ex: CNPJs_Pendentes.xlsx)",
    )

    args, _ = parser.parse_known_args()

    caminho_informado = args.caminho_flag or args.caminho_posicional

    if not caminho_informado or not str(caminho_informado).strip():
        print("=" * 70, file=sys.stderr)
        print("[ERRO] É obrigatório informar o caminho da pasta para executar a automação.", file=sys.stderr)
        print("Uso: python main.py \"C:\\caminho\\da\\pasta\"", file=sys.stderr)
        print("Opções adicionais:", file=sys.stderr)
        print("  -p, --pendentes   Executa apenas os CNPJs pendentes (CNPJs_Pendentes.xlsx)", file=sys.stderr)
        print("  -a, --arquivo     Informa o nome de uma planilha específica", file=sys.stderr)
        print("=" * 70, file=sys.stderr)
        sys.exit(1)

    planilha_especifica = None
    if args.apenas_pendentes:
        planilha_especifica = "CNPJs_Pendentes.xlsx"
    elif args.arquivo_planilha:
        planilha_especifica = args.arquivo_planilha

    return Path(caminho_informado).resolve(), planilha_especifica


def main():
    pasta_base, planilha_especifica = obter_argumentos()

    # Garante que a pasta base existe
    try:
        pasta_base.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        print(f"[ERRO CRÍTICO] Falha ao acessar ou criar a pasta base '{pasta_base}': {e}", file=sys.stderr)
        sys.exit(1)

    arquivo_log = configurar_logger(pasta_base)
    logger = logging.getLogger("RoboDET")

    logger.info("=" * 70)
    logger.info("INICIANDO AUTOMAÇÃO DET (MODO LINHA DE COMANDO)")
    logger.info(f"Pasta Base: {pasta_base}")
    if planilha_especifica:
        logger.info(f"Planilha Alvo: {planilha_especifica}")
    logger.info(f"Arquivo de Log: {arquivo_log}")
    logger.info("=" * 70)

    # Callback de progresso para exibição em log
    def callback_progresso(atual: int, total: int, status_texto: str):
        porcentagem = int((atual / total) * 100) if total > 0 else 0
        logger.info(f"[PROCESSO {atual}/{total} ({porcentagem}%)] {status_texto}")

    try:
        sucesso = executar_robo_hibrido(
            base_dir=pasta_base,
            progress_callback=callback_progresso,
            planilha_especifica=planilha_especifica,
        )
        if sucesso:
            logger.info("AUTOMAÇÃO FINALIZADA COM SUCESSO!")
        else:
            logger.warning("AUTOMAÇÃO CONCLUÍDA COM AVISOS OU FALHAS EM ITENS.")
    except Exception as e:
        logger.exception(f"FALHA INESPERADA NA EXECUÇÃO: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
