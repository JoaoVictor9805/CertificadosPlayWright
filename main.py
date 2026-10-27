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

from typing import Optional, Tuple

# Adiciona o diretório raiz e 'src' ao sys.path para garantir resolução de módulos
ROOT_DIR = Path(__file__).resolve().parent
SRC_DIR = ROOT_DIR / "src"

if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from src.automacao_hibrida_final import executar_robo_hibrido
from src.certificados import atualizar_aba_certificados_disponiveis


def configurar_logger(pasta_base: Path) -> Path:
    """Configura o logger 'RoboSocietario' para gravar no console e em arquivo .txt."""
    pasta_logs = pasta_base / "Logs"
    pasta_logs.mkdir(parents=True, exist_ok=True)

    dt_str = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    arquivo_log = pasta_logs / f"log_execucao_{dt_str}.txt"

    logger = logging.getLogger("RoboSocietario")
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


def obter_argumentos() -> Tuple[Path, str, Optional[str], bool]:
    """
    Obtém o caminho da pasta base, o tipo de operação ('automacao' ou 'certificados'),
    a planilha alvo e se a execução é interativa.
    No modo agendado/não-interativo (Task Scheduler), perguntas e pausas são desativadas.
    """
    parser = argparse.ArgumentParser(
        description="Robô Societário - Consulta DET / Gov.br e Gestão de Certificados"
    )
    parser.add_argument(
        "caminho_posicional",
        nargs="?",
        default=None,
        help="Caminho da pasta da automação (onde fica o CNPJs.xlsx e CapturasDeTela-Downloads)",
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
    parser.add_argument(
        "--atualizar-certificados",
        "--certificados",
        dest="apenas_certificados",
        action="store_true",
        help="Apenas coleta e atualiza os certificados instalados no Windows na planilha CNPJs.xlsx",
    )
    parser.add_argument(
        "--agendado",
        "--nao-interativo",
        dest="modo_agendado",
        action="store_true",
        help="Executa de forma não-interativa e sem pausas (para Agendador de Tarefas do Windows)",
    )

    args, _ = parser.parse_known_args()

    caminho_informado = args.caminho_flag or args.caminho_posicional

    # Se não foi informado caminho, assume automaticamente a pasta onde o .exe ou script está
    if not caminho_informado or not str(caminho_informado).strip():
        if getattr(sys, "frozen", False):
            pasta_base = Path(sys.executable).resolve().parent
        else:
            pasta_base = Path.cwd().resolve()
    else:
        pasta_base = Path(caminho_informado).resolve()

    # Detecta se a execução possui um terminal interativo com operador humano
    eh_interativo = bool(sys.stdin.isatty() and not args.modo_agendado)

    # Se a flag direta de certificados foi informada por linha de comando
    if args.apenas_certificados:
        return pasta_base, "certificados", None, eh_interativo

    # Se a flag de pendentes ou arquivo específico foi informada
    if args.apenas_pendentes:
        return pasta_base, "automacao", "CNPJs_Pendentes.xlsx", eh_interativo
    if args.arquivo_planilha:
        return pasta_base, "automacao", args.arquivo_planilha, eh_interativo

    # Operação padrão: Automação das consultas
    p_pendentes = pasta_base / "CNPJs_Pendentes.xlsx"
    p_principal = pasta_base / "CNPJs.xlsx"
    if not p_principal.exists():
        p_legado = pasta_base / "CNPJs-e-CPFs.xlsx"
        if p_legado.exists():
            p_principal = p_legado

    planilha_especifica = None
    # Apenas pergunta sobre pendentes se for uma execução interativa com usuário humano
    if eh_interativo and p_pendentes.exists() and p_principal.exists():
        print("\n" + "-" * 70)
        print(" [AVISO] Identificada planilha com itens pendentes da execução anterior:")
        print("   [1] Executar base completa (CNPJs.xlsx)")
        print("   [2] Reprocessar apenas pendentes anteriores (CNPJs_Pendentes.xlsx)")
        print("-" * 70)
        try:
            escolha_pend = input(" Digite a opção [Pressione ENTER para opção 1]: ").strip()
            if escolha_pend == "2":
                planilha_especifica = "CNPJs_Pendentes.xlsx"
            else:
                planilha_especifica = p_principal.name
        except (EOFError, KeyboardInterrupt):
            planilha_especifica = p_principal.name
        print()
    elif p_pendentes.exists() and not p_principal.exists():
        planilha_especifica = "CNPJs_Pendentes.xlsx"
    else:
        planilha_especifica = p_principal.name if p_principal.exists() else "CNPJs.xlsx"

    return pasta_base, "automacao", planilha_especifica, eh_interativo


def main():
    pasta_base, operacao, planilha_especifica, eh_interativo = obter_argumentos()

    # Ramo 1: Apenas Atualização de Certificados Digitais no Excel
    if operacao == "certificados":
        p_excel = pasta_base / "CNPJs.xlsx"
        print("\n" + "=" * 70)
        print(" ATUALIZAÇÃO DE CERTIFICADOS DIGITAIS (WINDOWS)")
        print(f" Pasta Base: {pasta_base}")
        print(f" Planilha:   {p_excel.name}")
        print("=" * 70)
        print(" Localizando certificados instalados em 'Certificados - Usuário Atual'...")
        codigo_saida = 0
        try:
            qtd = atualizar_aba_certificados_disponiveis(p_excel)
            print(f"\n [SUCESSO] {qtd} certificado(s) digital(is) encontrado(s) e gravados com sucesso!")
            print(f" Aba atualizada: 'Certificados_Disponiveis' na planilha:")
            print(f" -> {p_excel}")
            print("\n Instruções para preenchimento:")
            print(" 1. Abra a planilha 'CNPJs.xlsx'.")
            print(" 2. Consulte os certificados disponíveis na aba 'Certificados_Disponiveis'.")
            print(" 3. Na aba 'CNPJs_consultados', preencha a coluna 'Certificado' com a Razão Social ou CNPJ.")
            print(" 4. Salve e feche a planilha antes de rodar a automação do robô.")
        except PermissionError as pe:
            print(f"\n [AVISO DE ARQUIVO ABERTO] {pe}")
            codigo_saida = 1
        except Exception as e:
            print(f"\n [ERRO] Falha ao coletar/gravar certificados: {e}")
            codigo_saida = 1
        finally:
            print("\n" + "=" * 70)
            if eh_interativo:
                try:
                    input(" Pressione ENTER para fechar esta janela...")
                except (EOFError, KeyboardInterrupt):
                    pass
            if codigo_saida != 0:
                sys.exit(codigo_saida)
        return

    # Ramo 2: Execução da Automação (Robô)
    # Garante que a pasta base existe
    try:
        pasta_base.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        print(f"[ERRO CRÍTICO] Falha ao acessar ou criar a pasta base '{pasta_base}': {e}", file=sys.stderr)
        sys.exit(1)

    # Atualiza preventivamente a aba de certificados na planilha padrão se possível
    p_excel_padrao = pasta_base / "CNPJs.xlsx"
    if p_excel_padrao.exists():
        try:
            atualizar_aba_certificados_disponiveis(p_excel_padrao)
        except Exception:
            pass  # Se a planilha estiver aberta no momento, prossegue normalmente com a automação

    arquivo_log = configurar_logger(pasta_base)
    logger = logging.getLogger("RoboSocietario")

    logger.info("=" * 70)
    logger.info("INICIANDO AUTOMAÇÃO DET (ROBÔ SOCIETÁRIO)")
    logger.info(f"Pasta Base: {pasta_base}")
    if planilha_especifica:
        logger.info(f"Planilha Alvo: {planilha_especifica}")
    logger.info(f"Arquivo de Log: {arquivo_log}")
    logger.info("=" * 70)

    # Callback de progresso para exibição em log
    def callback_progresso(atual: int, total: int, status_texto: str):
        porcentagem = int((atual / total) * 100) if total > 0 else 0
        logger.info(f"[PROCESSO {atual}/{total} ({porcentagem}%)] {status_texto}")

    codigo_saida = 0
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
        codigo_saida = 1
    finally:
        print("\n" + "=" * 70)
        if eh_interativo:
            try:
                input(" Pressione ENTER para fechar esta janela...")
            except (EOFError, KeyboardInterrupt):
                pass
        else:
            logger.info("Execução finalizada (modo não-interativo / agendado). Encerrando processo.")
        if codigo_saida != 0:
            sys.exit(codigo_saida)


if __name__ == "__main__":
    main()
