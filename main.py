import sys
import warnings

# Define STA (Single-Threaded Apartment) antes de qualquer biblioteca COM/pywinauto ser importada.
# Sem isso, pywinauto define MTA (coinit_flags = 0), travando o seletor de arquivos (filedialog) do Tkinter no Windows.
sys.coinit_flags = 2
warnings.filterwarnings("ignore", message=".*Apply externally defined coinit_flags.*")

from pathlib import Path

# Adiciona o diretório raiz e 'src' ao sys.path para garantir resolução de módulos
ROOT_DIR = Path(__file__).resolve().parent
SRC_DIR = ROOT_DIR / "src"

if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from src.app import main

if __name__ == "__main__":
    main()
