import os
from pathlib import Path

# Diretórios base
BASE_DIR = Path(__file__).resolve().parent.parent
TESSERACT_DIR = BASE_DIR / "tesseract"
POPPLER_DIR = BASE_DIR / "poppler" / "bin"

# Configurações do Tesseract
TESSERACT_PATHS = [
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    str(TESSERACT_DIR / "tesseract.exe")
]

# Caminhos possíveis para tessdata
TESSDATA_PATHS = [
    os.environ.get('TESSDATA_PREFIX', ''),  # variável de ambiente, se definida corretamente
    r"C:\Program Files\Tesseract-OCR\tessdata",
    r"C:\Program Files (x86)\Tesseract-OCR\tessdata",
    str(TESSERACT_DIR / "tessdata")
]

# Configurações do Poppler
POPPLER_PATHS = [
    r"C:\Program Files\Poppler\Library\bin",
    r"C:\Program Files (x86)\Poppler\Library\bin",
    str(POPPLER_DIR)
]

# Configurações de OCR
OCR_CONFIG = "--psm 12 --oem 2"
OCR_LANG = "por+osd+eng"

# Configurações de resolução para OCR
OCR_DPI_OPTIONS = {
    'inicial': 300,
    'max': 600,
    'passo': 100
}

# Nome padrão para o arquivo gerado
DEFAULT_XLSX_FILENAME = "POs.xlsx"

import re

ANSI_COLORS = {
    "preto": "\033[30m",
    "vermelho": "\033[31m",
    "verde": "\033[32m",
    "amarelo": "\033[33m",
    "azul": "\033[34m",
    "roxo": "\033[35m",
    "ciano": "\033[36m",
    "reset": "\033[0m"
}


def colorize_terminal(text: str) -> str:
    def replace_tag(match):
        tag = match.group(1)
        content = match.group(2)
        color = ANSI_COLORS.get(tag.lower(), "")
        return f"{color}{content}{ANSI_COLORS['reset']}"

    return re.sub(r"<(\w+)>(.*?)</\1>", replace_tag, text)