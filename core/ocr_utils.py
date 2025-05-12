import os
import pytesseract

from typing import Tuple, Optional, List, LiteralString

from PIL import ImageEnhance
from pdf2image import convert_from_path, pdf2image

from config import (
    OCR_CONFIG,
    OCR_LANG,
    OCR_DPI_OPTIONS,
    POPPLER_PATHS,
    TESSERACT_PATHS,
    TESSDATA_PATHS, colorize_terminal
)

def check_dependencies() -> Tuple[List[str], Optional[str]]:
    """Verifica e configura as dependências do sistema (Tesseract e Poppler).

    Returns:
        Tuple[List[str], Optional[str]]:
            - Lista de mensagens de erro (vazia se tudo estiver OK)
            - Caminho do Poppler configurado (None se não encontrado)
    """
    errors = []
    poppler_path = None

    # Configura Tesseract
    try:
        configurar_tesseract()
    except FileNotFoundError as e:
        errors.append(f"Erro no Tesseract: {e}")
    except RuntimeError as e:
        errors.append(f"Falha ao inicializar Tesseract: {e}")
    except Exception as e:
        errors.append(f"Erro inesperado no Tesseract: {e}")

    # Configura Poppler
    try:
        configurar_poppler()
        poppler_path = os.environ.get('POPPLER_PATH')
        if not poppler_path:
            errors.append("POPPLER_PATH foi configurado mas está vazio")
    except FileNotFoundError as e:
        errors.append(f"Erro no Poppler: {e}")
    except Exception as e:
        errors.append(f"Erro inesperado no Poppler: {e}")

    return errors, poppler_path


def configurar_tesseract():
    """Configura o caminho do Tesseract OCR e verifica se está funcionando.

    Raises:
        FileNotFoundError: Se o Tesseract ou arquivos *.traineddata não forem encontrados.
        RuntimeError: Se o Tesseract não puder ser inicializado.
    """
    # Configura caminho do Tesseract
    for path in TESSERACT_PATHS:
        if os.path.exists(path):
            pytesseract.pytesseract.tesseract_cmd = path
            break
    else:
        raise FileNotFoundError(
            f"Tesseract não encontrado. Paths testados: {TESSERACT_PATHS}"
        )

    # Configura caminho dos dados treinados
    for path in TESSDATA_PATHS:
        if path and os.path.isdir(path):
            # Verifica se há pelo menos um arquivo '.traineddata'
            traineddata_files = [
                f for f in os.listdir(path) if f.endswith(".traineddata")
            ]
            if traineddata_files:
                os.environ["TESSDATA_PREFIX"] = path
                break
    else:
        raise FileNotFoundError(
            f"Nenhum arquivo .traineddata encontrado. Paths testados: {TESSDATA_PATHS}"
        )

    # Valida instalação
    try:
        pytesseract.get_tesseract_version()
    except Exception as e:
        raise RuntimeError(f"Erro ao validar Tesseract: {e}")


def configurar_poppler():
    """Configura o caminho do Poppler e retorna o path configurado.

    Returns:
        str: Caminho do Poppler configurado.

    Raises:
        FileNotFoundError: Se o Poppler não for encontrado.
    """
    if "POPPLER_PATH" in os.environ and os.path.exists(os.environ["POPPLER_PATH"]):
        return os.environ["POPPLER_PATH"]

    # Testa os caminhos padrão
    for path in POPPLER_PATHS:
        if path and os.path.exists(path):
            os.environ["POPPLER_PATH"] = path
            return path

    raise FileNotFoundError(
        f"Poppler não encontrado. Paths testados: {POPPLER_PATHS}\n"
        "Baixe o Poppler em: https://github.com/oschwartz10612/poppler-windows/releases"
    )


def extrair_texto_ocr(filepath: str, dpi: int = OCR_DPI_OPTIONS['inicial']) -> None | LiteralString | str:
    """
    Extrai texto de um arquivo (PDF/imagem) usando OCR.

    Args:
        filepath: Caminho do arquivo a ser processado
        dpi: Resolução para processamento (valor entre 'inicial' e 'max' de OCR_DPI_OPTIONS)

    Returns:
        Texto extraído ou string vazia em caso de falha
    """
    try:
        # Validação inicial
        if not os.path.exists(filepath):
            raise FileNotFoundError(f"Arquivo não encontrado: {filepath}")

        # Garante que o DPI está dentro dos limites
        dpi_efetivo = min(max(dpi, OCR_DPI_OPTIONS['inicial']), OCR_DPI_OPTIONS['max'])

        # Configurações do Poppler
        poppler_path = os.environ.get('POPPLER_PATH')
        if not poppler_path:
            raise RuntimeError("Caminho do Poppler não configurado")

        # Processamento do PDF/Imagem
        print(colorize_terminal(
            f"[<roxo>DEBUG</roxo>] {f'Tentando: Extração por OCR com {dpi} DPI'}")
        )
        imagens = convert_from_path(
            filepath,
            poppler_path=poppler_path,
            dpi=dpi_efetivo,
            grayscale=True,
            fmt='png',
            thread_count=min(2, os.cpu_count() or 1),  # Limita a 2 threads, mas considera CPUs disponíveis
            strict=True  # Para tratar erros no PDF
        )

        texto_total = []

        for img in imagens:
            try:
                # Pré-processamento da imagem
                img = (
                    img.convert('L')
                    .point(lambda x: 0 if x < 180 else 255)
                )

                # Aplica contraste
                enhancer = ImageEnhance.Contrast(img)
                img = enhancer.enhance(1.5)

                # OCR
                texto = pytesseract.image_to_string(
                    img,
                    lang=OCR_LANG,
                    config=OCR_CONFIG
                )

                texto_total.append(texto.strip())

            finally:
                img.close()  # liberação de recursos

        print(colorize_terminal(f"[<roxo>DEBUG</roxo>] [<verde>SUCESSO</verde>] Texto extraído por OCR"))
        return "\n".join(filter(None, texto_total))

    except pdf2image.exceptions.PDFPageCountError:
        error_msg = "O arquivo PDF está corrompido ou vazio"
    except PermissionError:
        error_msg = "Sem permissão para acessar o arquivo"
    except Exception as e:
        error_msg = f"Erro no processamento OCR: {str(e)}"

    print(colorize_terminal(f"[<vermelho>FALHA</vermelho>] Texto extraído por OCR\n{error_msg}"))
    return ""


