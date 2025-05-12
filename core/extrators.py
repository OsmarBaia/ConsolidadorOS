import os
import re
from datetime import datetime
import pdfplumber
from config import colorize_terminal, OCR_DPI_OPTIONS
from core.ocr_utils import extrair_texto_ocr


def extrair_nf_do_nome_arquivo(nome_arquivo: str) -> str:
    match = re.search(r'NF[.\s]*(\d+)', nome_arquivo, re.IGNORECASE)
    return match.group(1) if match else ""


def extrair_nf_do_conteudo_arquivo(texto):
    patterns = [
        r'numero\s+da\s+nota[^\d]*(\d{4,})',
        r'\bnota\s*[.:]?\s*(\d{4,})',
    ]

    for pattern in patterns:
        match = re.search(pattern, texto, re.IGNORECASE)
        if match:
            num = match.group(1).strip()
            if num.isdigit() and len(num) >= 4:
                return num

    matches = re.findall(r'\b(?!0{4,})\d{4,}\b', texto)
    return matches[0] if matches else ""


def extrair_numero_nf(nome_arquivo: str, texto: str) -> str:
    nf_nome = extrair_nf_do_nome_arquivo(nome_arquivo)
    if not nf_nome:
        nf_nome = extrair_nf_do_conteudo_arquivo(texto)

    return nf_nome


def extrair_pos_do_nome_arquivo(nome_arquivo: str) -> list[str]:
    match = re.search(r'\bPO\s+([^\n\r]+)', nome_arquivo, re.IGNORECASE)
    if not match:
        return []

    trecho = match.group(1)
    return list(dict.fromkeys(re.findall(r'\b\d{6}\b', trecho)))


def extrair_pos_do_conteudo_arquivo(texto: str) -> list[str]:
    matches = re.findall(r'\bPO\s+(\d{6})\b', texto, re.IGNORECASE)
    return list(dict.fromkeys(matches))


def extrair_pos(nome_arquivo: str, texto: str) -> list[str]:
    pos_nome = extrair_pos_do_nome_arquivo(nome_arquivo)
    pos_conteudo = extrair_pos_do_conteudo_arquivo(texto)

    if sorted(pos_conteudo) != sorted(pos_nome):
        return [po for po in pos_conteudo if po.isdigit() and len(po) == 6]

    return pos_nome


def extrair_data_emissao(texto: str) -> str:
    patterns = [
        r'(?:data\s+e\s+hora\s+da\s+emiss[ao]\s*|emiss[ao]\s*:\s*)(\d{2}/\d{2}/\d{4})',
        r'(?:data\s*:\s*|emiss[ao]\s*em\s*)(\d{2}[\/\-\.]\d{2}[\/\-\.]\d{4})',
        r'\b(\d{2}[\/\-\.]\d{2}[\/\-\.]\d{4})\b'
    ]

    for pattern in patterns:
        match = re.search(pattern, texto, re.IGNORECASE)
        if match:
            data = match.group(1).replace('-', '/').replace('.', '/')
            try:
                datetime.strptime(data, '%d/%m/%Y')
                return data
            except ValueError:
                continue
    return ""


def extrair_linhas_pedido(texto: str):
    padrao_preciso = re.compile(
        r'(?:/)?\s*(?:PO\s+)?(?P<po>\d{6})\s+LINHA\s+(?P<linha>\d+)\s+VALOR\s+(?P<valor>[\d.,]+)\s*/\s*(?P<descricao>[^\n\r]*)',
        re.IGNORECASE
    )

    padrao_alternativo = re.compile(
        r'(?:/)?\s*(?:PO\s+)?(?P<po>\d{6})[\s\S]*?(?:linha|item)[\s:]*(?P<linha>\d+)[\s\S]*?'
        r'(?:valor|vlr?|total)[\s:]*[\$R]*(?P<valor>[\d.,]+)'
        r'(?:[\s/;-]*(?P<descricao>[^\n\r]*))?',
        re.IGNORECASE
    )

    matches = list(padrao_preciso.finditer(texto))
    return matches if matches else list(padrao_alternativo.finditer(texto))


def extrair_texto_pdf(filepath: str) -> str:
    print(colorize_terminal(
        f"[<roxo>DEBUG</roxo>] Tentando: Extração por Leitura Simples")
    )
    try:
        with pdfplumber.open(filepath) as pdf:
            return "\n".join(
                page.extract_text(x_tolerance=2, y_tolerance=2) or ""
                for page in pdf.pages
            ).strip()
    except Exception as e:
        print(colorize_terminal(f"[<vermelho>ERRO</vermelho>] Falha no processamento como PDF\n{e}"))
        return ""


def extrair_dados_recursivamente(
        filepath: str,
        dpi_atual: int = None,
        dados_parciais: dict = None,
        dpi_inicial: int = OCR_DPI_OPTIONS['inicial'],
        dpi_maximo: int = OCR_DPI_OPTIONS['max'],
        dpi_incremento: int = OCR_DPI_OPTIONS['passo']
):
    if dados_parciais is None:
        dados_parciais = {
            'numero_nf': '',
            'data_nf': '',
            'pos_nf': '',
            'linhas_nf': []
        }

    if not all(dados_parciais.values()):
        dpi_atual = dpi_inicial if dpi_atual is None else dpi_atual + dpi_incremento
        if dpi_atual <= dpi_maximo:
            dados_parciais = extrair_dados_texto(filepath, extrair_texto_ocr(filepath, dpi=dpi_atual))
            campos_obrigatorios = ['numero_nf', 'pos_nf', 'linhas_nf']

            if not all(dados_parciais[campo] for campo in campos_obrigatorios):
                extrair_dados_recursivamente(filepath, dpi_atual, dados_parciais)
            else:
                print(colorize_terminal(
                    f"[<roxo>DEBUG</roxo>] [<verde>SUCESSO</verde>] Dados Extraídos com SUCESSO"
                ))
                return dados_parciais

    return dados_parciais


def extrair_dados(filepath: str):
    print(colorize_terminal(f"[<roxo>DEBUG</roxo>] extrators.py <- extrair_dados"))
    print(colorize_terminal(f"[<roxo>DEBUG</roxo>] extrair_dados -> extrair_dados_texto"))
    dados = extrair_dados_texto(filepath, extrair_texto_pdf(filepath))
    if not all(dados.values()):
        campos_obrigatorios = ['numero_nf', 'pos_nf', 'linhas_nf']
        if not all(dados[campo] for campo in campos_obrigatorios):
            print(colorize_terminal(
                f"[<vermelho>FALHA</vermelho>] Falha na extração por leitura simples"
            ))
            dados = extrair_dados_recursivamente(filepath)


    return dados


def extrair_dados_texto(filepath, texto):
    print(colorize_terminal(f"[<roxo>DEBUG</roxo>] extrators.py <- extrair_dados_texto"))
    try:
        dados = {
            'numero_nf': extrair_numero_nf(filepath, texto) or '',
            'data_nf': extrair_data_emissao(texto) or '',
            'pos_nf': extrair_pos(filepath, texto) or [],
            'linhas_nf': extrair_linhas_pedido(texto) or [],
        }
        return dados
    except Exception as e:
        print(colorize_terminal(f"[<vermelho>ERRO</vermelho>] Falha ao extrair dados:\n{str(e)}"))
        return {}
