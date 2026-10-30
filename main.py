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


def configurar_logger(pasta_base: Path, nome_prefixo: str = "log_execucao") -> Path:
    """Configura o logger 'RoboSocietario' para gravar no console e em arquivo .txt."""
    pasta_resultados = pasta_base / "Resultados"
    pasta_logs = pasta_resultados / "Logs"
    pasta_logs.mkdir(parents=True, exist_ok=True)

    dt_str = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    arquivo_log = pasta_logs / f"{nome_prefixo}_{dt_str}.txt"

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
    pasta_planilhas = pasta_base / "PlanilhasBase"
    pasta_planilhas.mkdir(parents=True, exist_ok=True)

    p_pendentes = pasta_planilhas / "CNPJs_Pendentes.xlsx"
    if not p_pendentes.exists() and (pasta_base / "CNPJs_Pendentes.xlsx").exists():
        p_pendentes = pasta_base / "CNPJs_Pendentes.xlsx"

    p_principal = pasta_planilhas / "CNPJs.xlsx"
    if not p_principal.exists() and (pasta_base / "CNPJs.xlsx").exists():
        p_principal = pasta_base / "CNPJs.xlsx"
    elif not p_principal.exists() and (pasta_planilhas / "CNPJs-e-CPFs.xlsx").exists():
        p_principal = pasta_planilhas / "CNPJs-e-CPFs.xlsx"
    elif not p_principal.exists() and (pasta_base / "CNPJs-e-CPFs.xlsx").exists():
        p_principal = pasta_base / "CNPJs-e-CPFs.xlsx"

    planilha_especifica = None
    # Apenas pergunta sobre pendentes se for uma execução interativa com usuário humano
    if eh_interativo and p_pendentes.exists() and p_principal.exists():
        print("\n" + "-" * 70)
        print(" [AVISO] Identificada planilha com itens pendentes da execução anterior:")
        print("   [1] Executar base completa (PlanilhasBase/CNPJs.xlsx)")
        print("   [2] Reprocessar apenas pendentes anteriores (PlanilhasBase/CNPJs_Pendentes.xlsx)")
        print("-" * 70)
        try:
            escolha_pend = input(" Digite a opção [Pressione ENTER para opção 1]: ").strip()
            if escolha_pend == "2":
                planilha_especifica = str(p_pendentes)
            else:
                planilha_especifica = str(p_principal)
        except (EOFError, KeyboardInterrupt):
            planilha_especifica = str(p_principal)
        print()
    elif p_pendentes.exists() and not p_principal.exists():
        planilha_especifica = str(p_pendentes)
    else:
        planilha_especifica = str(p_principal) if p_principal.exists() else "CNPJs.xlsx"

    return pasta_base, "automacao", planilha_especifica, eh_interativo


def main():
    # Oculta a pasta interna de bibliotecas (_internal) para manter a raiz limpa e elegante para o usuário
    try:
        if getattr(sys, "frozen", False):
            exe_dir = Path(sys.executable).resolve().parent
            internal_dir = exe_dir / "_internal"
            if internal_dir.exists():
                import subprocess
                subprocess.run(["attrib", "+h", str(internal_dir)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception:
        pass

    pasta_base, operacao, planilha_especifica, eh_interativo = obter_argumentos()

    # Ramo 1: Apenas Atualização de Certificados Digitais no Excel
    if operacao == "certificados":
        try:
            pasta_base.mkdir(parents=True, exist_ok=True)
            (pasta_base / "PlanilhasBase").mkdir(parents=True, exist_ok=True)
            (pasta_base / "Resultados").mkdir(parents=True, exist_ok=True)
        except Exception:
            pass

        arquivo_log = configurar_logger(pasta_base, nome_prefixo="log_certificados")
        logger = logging.getLogger("RoboSocietario")

        pasta_planilhas = pasta_base / "PlanilhasBase"
        pasta_planilhas.mkdir(parents=True, exist_ok=True)
        p_excel = pasta_planilhas / "CNPJs.xlsx"
        if not p_excel.exists() and (pasta_base / "CNPJs.xlsx").exists():
            p_excel = pasta_base / "CNPJs.xlsx"

        logger.info("=" * 70)
        logger.info("ATUALIZAÇÃO DE CERTIFICADOS DIGITAIS (WINDOWS)")
        logger.info(f"Pasta Base: {pasta_base}")
        logger.info(f"Planilha:   {p_excel}")
        logger.info(f"Arquivo de Log: {arquivo_log}")
        logger.info("=" * 70)
        logger.info("[CERTIFICADOS] Localizando certificados instalados em 'Certificados - Usuário Atual'...")
        codigo_saida = 0
        try:
            qtd = atualizar_aba_certificados_disponiveis(p_excel)
            logger.info(f"[CERTIFICADOS] [SUCESSO] {qtd} certificado(s) digital(is) encontrado(s) e gravados com sucesso!")
            logger.info(f"[CERTIFICADOS] Aba atualizada: 'Certificados_Disponiveis' na planilha: {p_excel}")
            logger.info("[CERTIFICADOS] Instruções:")
            logger.info("[CERTIFICADOS] 1. Abra a planilha em 'PlanilhasBase/CNPJs.xlsx'.")
            logger.info("[CERTIFICADOS] 2. Consulte os certificados disponíveis na aba 'Certificados_Disponiveis'.")
            logger.info("[CERTIFICADOS] 3. Na aba 'CNPJs_consultados', preencha a coluna 'Certificado' com a Razão Social ou CNPJ.")
            logger.info("[CERTIFICADOS] 4. Salve e feche a planilha antes de rodar a automação do robô.")
        except PermissionError as pe:
            logger.error(f"[CERTIFICADOS] [AVISO DE ARQUIVO ABERTO] {pe}")
            codigo_saida = 1
        except Exception as e:
            logger.error(f"[CERTIFICADOS] [ERRO] Falha ao coletar/gravar certificados: {e}")
            codigo_saida = 1
        finally:
            logger.info("=" * 70)
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
        (pasta_base / "PlanilhasBase").mkdir(parents=True, exist_ok=True)
        (pasta_base / "Resultados").mkdir(parents=True, exist_ok=True)
    except Exception as e:
        print(f"[ERRO CRÍTICO] Falha ao acessar ou criar as pastas base: {e}", file=sys.stderr)
        sys.exit(1)

    arquivo_log = configurar_logger(pasta_base, nome_prefixo="log_execucao")
    logger = logging.getLogger("RoboSocietario")

    # Atualiza preventivamente a aba de certificados na planilha padrão se possível
    pasta_planilhas = pasta_base / "PlanilhasBase"
    p_excel_padrao = pasta_planilhas / "CNPJs.xlsx"
    if not p_excel_padrao.exists() and (pasta_base / "CNPJs.xlsx").exists():
        p_excel_padrao = pasta_base / "CNPJs.xlsx"
    if p_excel_padrao.exists():
        try:
            atualizar_aba_certificados_disponiveis(p_excel_padrao)
        except Exception:
            pass  # Se a planilha estiver aberta no momento, prossegue normalmente com a automação

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
        logger.info("=" * 70)
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
