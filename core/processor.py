import os
import time
from typing import Callable, List, Dict
from config import colorize_terminal
from core.extrators import extrair_dados
from core.file_utils import mover_arquivo, salvar_em_xlsx


def process_xlsx(
        dados_por_po: Dict[str, List[Dict]],
        xlsx_output_dir: str,
        xlsx_filename: str,
        log_message: Callable
):
    log_message(f"<azul>Gerando Arquivo xlsx... </azul>: {xlsx_filename}")
    print(colorize_terminal(f"[<roxo>DEBUG</roxo>] processor.py <- process_xlsx"))

    # Debug: Check if dados_por_po has content
    if not dados_por_po:
        print(colorize_terminal("\n[<roxo>DEBUG</roxo>] process_xlsx <- dados_por_po\n"
                                "[<vermelho>ERRO</vermelho>] Nenhum dado encontrado em dados_por_po\n"))
        log_message("[<vermelho>ERRO</vermelho>] Nenhum dado válido para exportar")
        return

    # Debug: Show POs found
    print(colorize_terminal(f"\n[<roxo>DEBUG</roxo>] process_xlsx <- dados_por_po\n"
                            f"[<verde>OK</verde>] Dados encontrados em dados_por_po com SUCESSO\n"
                            f"\t-== Sumário ==-"
                            f"\t\tPOs encontrados: {len(dados_por_po)}"))
    for po, linhas in dados_por_po.items():
        print(f"\t\tPO {po}: {len(linhas)} linhas")

    print(colorize_terminal(f"\t-== Descritivo ==-"
                            f"\t\t{dados_por_po}"))

    try:
        if salvar_em_xlsx(xlsx_output_dir, dados_por_po, xlsx_filename):
            log_message(f"[<verde>SUCESSO</verde>] Arquivo {xlsx_filename} gerado...\n\n")
        else:
            log_message("[<vermelho>FALHA</vermelho>] Erro ao salvar arquivo XLSX")
    except Exception as e:
        log_message(f"[<vermelho>ERRO</vermelho>] Falha ao gerar XLSX: {str(e)}")


def process_directory(
        base_dir: str,
        termo_nome: str,
        termos_exclusao: List[str],
        xlsx_output_dir: str,
        xlsx_filename: str,
        update_progress: Callable,
        log_message: Callable,
        batch_size: int = 10
):
    log_message("<azul>Iniciando...</azul>")

    print(colorize_terminal(f"[<roxo>DEBUG</roxo>] processor.py <- process_directory"))

    arquivos_pdf = coletar_arquivos_pdf(
        base_dir=base_dir,
        termo_nome=termo_nome,
        termos_exclusao=termos_exclusao,
        log_message=log_message,
    )

    total = len(arquivos_pdf)

    print(colorize_terminal(
        f"[<roxo>{total}</roxo>] Arquivos PDFs encontrados"))

    dados_por_po = {}
    estatisticas = {
        'sucesso': 0,
        'falha': 0,
        'falhas': 0
    }

    for batch_start in range(0, total, batch_size):
        batch_files = arquivos_pdf[batch_start:batch_start + batch_size]

        for idx, filepath in enumerate(batch_files, 1):
            global_idx = batch_start + idx
            update_progress(global_idx, total)

            log_message(
                f"[<azul>{global_idx}</azul>/<roxo>{total}</roxo>] {os.path.basename(filepath)}")
            log_message(
                f"[<azul>{global_idx}</azul>/<roxo>{total}</roxo>] <amarelo>Processando...</amarelo>")

            print(colorize_terminal(
                f"[<azul>{global_idx}</azul>/<roxo>{total}</roxo>] {os.path.basename(filepath)}"))

            # Inicio do processamento
            try:
                resultado, status_nf = process_pdf(filepath)
                print(colorize_terminal(f"\nDados:\n{resultado}"))

                # Atualiza contadores estatísticos
                if status_nf in estatisticas:
                    estatisticas[status_nf] += 1

                if status_nf != "sucesso":
                    estatisticas["falhas"] += 1

                # Log resultado do processamento
                if isinstance(resultado, dict) and 'linhas' in resultado:
                    for linha in resultado['linhas']:
                        if linha.get('processado'):
                            po = linha["po"]
                            dados = {
                                'nota': resultado.get('numero_nf'),
                                'data_emissao': resultado.get('data_nf'),
                                'linha': linha['linha'],
                                'descricao': linha['descricao'],
                                'valor': linha['valor']
                            }
                            dados_por_po.setdefault(po, []).append(dados)

                status_msg = f'[<verde>{status_nf.upper()}</verde>]' if status_nf == 'sucesso' else f'[<vermelho>{status_nf.upper()}</vermelho>]'
                log_message(
                    f"[<azul>{global_idx}</azul>/<roxo>{total}</roxo>] {status_msg}")

                print(colorize_terminal(f"Status: '{status_msg}'"))

                # Mover arquivo para pastas
                pasta_destino = "Notas Processadas" if status_nf == "sucesso" else "Precisa Revisar"
                mover_arquivo(filepath, os.path.join(base_dir, pasta_destino))

                print(f"Movido para a pasta: '{pasta_destino}'\n")

            except Exception as e:
                estatisticas["falhas"] += 1
                nome_curto = os.path.basename(filepath)

                log_message(f"[<vermelho>ERRO</vermelho>] {nome_curto}: {str(e)}")
                print(colorize_terminal(f"[<vermelho>ERRO</vermelho>] {nome_curto}: {str(e)}"))

                pasta_dest = os.path.join(base_dir, "Precisa Revisar")
                mover_arquivo(filepath, pasta_dest)

                log_message(f"[<amarelo>Aviso</amarelo>] {nome_curto} movido para a pasta: '{pasta_dest}'")
                print(colorize_terminal(f"[<amarelo>Aviso</amarelo>] {nome_curto} movido para a pasta: '{pasta_dest}'\n"))

        if batch_start + batch_size < total:
            time.sleep(0.5)

    log_message("<azul>Processamento Finalizado...</azul>")
    print(f'Dados de Saida:\n{dados_por_po}')

    process_xlsx(dados_por_po, xlsx_output_dir,xlsx_filename, log_message)
    process_stats(estatisticas, log_message)

def process_stats(
        estatisticas: Dict,
        log_message: Callable
):
    print(colorize_terminal(f"[<roxo>DEBUG</roxo>] processor.py <- process_stats"))
    num_success = estatisticas.get('sucesso', 0)  # Changed 'sucessos' to 'sucesso'
    num_falhas = estatisticas.get('falha', 0)
    num_total = num_success + num_falhas

    pct_sucesso = (num_success / num_total) * 100 if num_total else 0
    pct_falha = (num_falhas / num_total) * 100 if num_total else 0

    log_message("<amarelo>=== RESUMO DO PROCESSAMENTO ===</amarelo>")
    log_message(f"Arquivos processados: <amarelo>{num_total}</amarelo>")
    log_message(
        f"Sucessos: [<verde>{pct_sucesso:.0f}%</verde>] [<verde>{num_success}</verde>/<amarelo>{num_total}</amarelo>]")
    log_message(
        f"Precisa de Revisão: [<vermelho>{pct_falha:.0f}%</vermelho>] [<vermelho>{num_falhas}</vermelho>/<amarelo>{num_total}</amarelo>]")

    print(colorize_terminal(f'[<roxo>DEBUG</roxo>]<amarelo>=== ESTATÍSTICAS ====</amarelo>'))
    for key, value in estatisticas.items():
        print(colorize_terminal(f'[<roxo>DEBUG</roxo>] {key}:{value}'))


def process_pdf(filepath):
    print(colorize_terminal(f"[<roxo>DEBUG</roxo>] processor.py <- process_pdf"))
    # Extração inicial dos dados
    dados_nf = extrair_dados(filepath) or {}
    numero_nf = dados_nf.get('numero_nf', '').strip()
    data_nf = dados_nf.get('data_nf', '').strip()
    pos_nf = dados_nf.get('pos_nf', [])
    linhas_nf = dados_nf.get('linhas_nf', [])

    # Validação inicial
    status_nf = 'sucesso'

    # Verifica se o número da NF tem pelo menos 4 dígitos
    if not (numero_nf.isdigit() and len(numero_nf) >= 4):
        status_nf = 'falha'

    linhas_processadas = []

    for match in linhas_nf:
        if not match:
            status_nf = 'falha'
            continue

        try:
            # Extração dos dados da linha
            po = match.group("po").strip() if match.group("po") else ""
            linha = match.group("linha").strip() if match.group("linha") else ""
            valor_raw = match.group("valor") if match.group("valor") else "0"
            descricao = match.group("descricao").strip() if match.group("descricao") else ""

            # Processamento do valor
            try:
                valor = processar_valor_monetario(valor_raw)
            except (ValueError, AttributeError):
                valor = 0.0
                print(colorize_terminal(f"[<vermelho>ERRO</vermelho>] Valor, {valor_raw}, inválido na linha: \n{match}"))
                status_nf = 'falha'

            # Criação do objeto de linha
            linha_processada = {
                'po': po,
                'linha': linha,
                'valor': valor,
                'descricao': descricao,
                'processado': bool(po and linha and valor > 0)
            }

            # Validação da linha
            if linha_processada['processado']:
                if po in pos_nf:  # Verifica se o PO está na lista de POs da NF
                    linhas_processadas.append(linha_processada)
                else:
                    print(colorize_terminal(f"[<amarelo>AVISO</amarelo>] PO {po} não encontrado na lista: {pos_nf}"))

                    if not all(linha_processada.values()):
                        campos_obrigatorios = ['po', 'linha', 'valor']
                        if all(linha_processada[campo] for campo in campos_obrigatorios):
                            linhas_processadas.append(linha_processada)
                    else:
                        linhas_processadas.append(linha_processada)

                    status_nf = 'falha'
            else:
                print(f"[<vermelho>ERRO</vermelho>] Linha incompleta: {linha_processada}")
                status_nf = 'falha'

        except Exception as e:
            print(f"[<vermelho>ERRO</vermelho>] Erro ao processar linha: {e}")
            status_nf = 'falha'
            continue

    # Verificação final se há linhas válidas
    if not linhas_processadas:
        status_nf = 'falha'

    # Retorno organizado
    resultado = {
        'numero_nf': numero_nf,
        'data_nf': data_nf,
        'linhas': linhas_processadas,
        'status': status_nf
    }

    # print(colorize_terminal(
    #     f"[<azul>INFO</azul>] === Resumo do processamento ===\n"
    #     f"Arquivo: {os.path.basename(filepath)}\n"
    #     f"Status: {status_nf}\n"
    #     f"NF: {numero_nf}\n"
    #     f"Data: {data_nf}\n"
    #     f"Linhas válidas:\n"))
    # for idx, l in enumerate(linhas_processadas):
    #     print(colorize_terminal(f"[<cinza>{idx}</cinza>] {l}"))

    return resultado, status_nf

def processar_valor_monetario(valor_str: str) -> float:
    try:
        return float(valor_str.replace('.', '').replace(',', '.'))
    except (ValueError, AttributeError):
        return 0.0


def coletar_arquivos_pdf(
        base_dir: str,
        termo_nome: str,
        termos_exclusao: List[str],
        log_message: Callable[[str], None] = print  # Parâmetro opcional para logging
) -> List[str]:
    """
    Coleta arquivos PDF válidos para processamento de forma recursiva.

    Args:
        base_dir: Diretório raiz para busca
        termo_nome: Termo que deve estar no nome do arquivo
        termos_exclusao: Lista de termos que invalidam o arquivo
        log_message: Função para registrar mensagens (opcional)

    Returns:
        Lista de caminhos completos dos arquivos válidos
    """
    arquivos_pdf = []
    total_encontrados = 0

    log_message("<azul>Buscando arquivos...</azul>")
    for root, _, files in os.walk(base_dir):
        for file in files:
            if file.lower().endswith('.pdf'):
                caminho_completo = os.path.join(root, file)

                if eh_arquivo_pra_processamento(caminho_completo, termo_nome, termos_exclusao):
                    arquivos_pdf.append(caminho_completo)
                    total_encontrados += 1

    if log_message:
        log_message(f"\n<roxo>{total_encontrados}</roxo> <azul>Arquivos encontrados</azul>")

    return arquivos_pdf


def eh_arquivo_pra_processamento(
        filepath: str,
        termo_nome: str,
        termos_exclusao: List[str]
) -> bool:
    """
    Verifica se um arquivo deve ser processado com base no nome.

    Args:
        filepath: Caminho completo do arquivo
        termo_nome: Termo obrigatório que deve estar no nome do arquivo
        termos_exclusao: Lista de termos que invalidam o arquivo se presentes

    Returns:
        True se o arquivo deve ser processado, False caso contrário
    """
    # Extrai o nome do arquivo em minúsculas
    nome_arquivo = os.path.basename(filepath).lower()

    # Remove a extensão do arquivo
    nome_sem_extensao = os.path.splitext(nome_arquivo)[0]

    # Converte termos de exclusão para minúsculas uma única vez
    termos_exclusao_lower = [termo.lower() for termo in termos_exclusao]

    # Verifica critérios de exclusão
    contem_termo_exclusao = any(
        termo in nome_sem_extensao
        for termo in termos_exclusao_lower
    )

    # Verifica se contém o termo obrigatório
    contem_termo_obrigatorio = termo_nome.lower() in nome_arquivo

    # Arquivo é válido se:
    # 1. NÃO contém termos de exclusão E
    # 2. CONTÉM o termo obrigatório
    return not contem_termo_exclusao and contem_termo_obrigatorio



