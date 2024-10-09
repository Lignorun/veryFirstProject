import os
import re
from typing import List
import PyPDF2
import openpyxl

# Constantes
PDF_DIR = r"C:/Users/dlins/Python/CV_Automação/corporate_empregare/CV_PDF"
TXT_DIR = r"C:/Users/dlins/Python/CV_Automação/corporate_empregare/CV_TXT"
EXCEL_FILE = 'zCurriculos.xlsx'
EXCEL_SHEET = 'CV_Bruto'

# Títulos a serem reconhecidos
TITLES = [
    'OBJETIVO',
    'SÍNTESE',
    'FORMAÇÃO',
    'EXPERIÊNCIA PROFISSIONAL',
    'CURSOS EXTRACURRICULARES',
    'IDIOMAS',
    'INFORMÁTICA',
    'INFORMAÇÕES ADICIONAIS'
]

def ler_pdf(arquivo_pdf: str) -> str:
    """
    Lê um arquivo PDF e extrai seu conteúdo como string, removendo textos indesejados.
    """
    texto_completo = []
    with open(arquivo_pdf, 'rb') as pdf_file:
        leitor_pdf = PyPDF2.PdfReader(pdf_file)
        for pagina in leitor_pdf.pages:
            texto = pagina.extract_text() or ''
            texto = re.sub(
                r'Currículo por EMPREGARE\.com - Software de Recrutamento e Seleção|Criado por EMPREGARE\.com - Software de Recrutamento e Seleção|•',
                '',
                texto
            ).strip()
            texto_completo.append(texto)
    return ' '.join(texto_completo)

def criar_e_preencher_txt(titulo: str, conteudo: str) -> None:
    """
    Cria um arquivo TXT com o conteúdo fornecido.
    """
    with open(titulo, 'w', encoding='utf-8') as arquivo:
        arquivo.write(conteudo)

def limpar_txt(caminho_arq: str) -> None:
    """
    Remove linhas vazias do início do arquivo TXT.
    """
    with open(caminho_arq, "r+", encoding='utf-8') as f:
        linhas = f.readlines()
        f.seek(0)
        linhas_filtradas = [linha for linha in linhas if linha.strip()]
        f.truncate(0)
        f.writelines(linhas_filtradas)

def separar_nome(caminho_arq: str) -> None:
    """
    Separa o nome do sexo no arquivo TXT, inserindo uma quebra de linha antes do sexo.
    """
    with open(caminho_arq, "r+", encoding='utf-8') as f:
        conteudo = f.read()
        padrao = r"(Masculino|Feminino|Não Informado)"
        match = re.search(padrao, conteudo)
        if match:
            pos = match.start()
            f.seek(0)
            f.write(conteudo[:pos] + '\n' + conteudo[pos:])
            f.truncate()

def adicionar_espaco(caminho_arq: str) -> None:
    """
    Adiciona uma linha em branco antes de cada título, se não houver.
    """
    with open(caminho_arq, "r", encoding='utf-8') as f:
        linhas = f.readlines()

    with open(caminho_arq, "w", encoding='utf-8') as f:
        anterior = ''
        for linha in linhas:
            linha_trim = linha.strip()
            if linha_trim in TITLES and anterior.strip() != '':
                f.write('\n')
            f.write(linha)
            anterior = linha

def eh_titulo(linha: str) -> bool:
    """
    Verifica se a linha é um dos títulos definidos.
    """
    return linha.strip() in TITLES

def inicializa_dados_cv() -> List[str]:
    """
    Inicializa uma lista com espaços reservados para os dados do CV.
    """
    return [''] * 29

def processar_txt(caminho_arq: str) -> List[str]:
    """
    Processa o arquivo TXT e extrai os dados estruturados do CV.
    """
    dados_cv = inicializa_dados_cv()

    with open(caminho_arq, "r", encoding='utf-8') as f:
        linhas = [linha.strip() for linha in f]

    # Mapeamento dos índices
    campos = {
        'Nome': 0,
        'Sexo': 1,
        'EstadoCivil': 2,
        'Nascimento': 3,
        'Cep': 4,
        'Telefone': 5,
        'Email': 6,
        'PretencaoSalarial': 7,
        'Objetivo': 8,
        'Sintese': 9,
        'Formacao': [10, 11, 12],
        'ExperienciaProfissional': [13, 14, 15, 16, 17],
        'CursoExtraCurricular': 18,
        'Idiomas': [19, 20, 21, 22, 23, 24, 25, 26],
        'Informatica': 27,
        'InfoAdicional': 28
    }

    # Extração dos dados básicos
    dados_cv[campos['Nome']] = linhas[0]

    sexo_estado_nasc = linhas[1].split(',')
    if len(sexo_estado_nasc) >= 3:
        dados_cv[campos['Sexo']] = sexo_estado_nasc[0].strip()
        dados_cv[campos['EstadoCivil']] = sexo_estado_nasc[1].strip()
        dados_cv[campos['Nascimento']] = sexo_estado_nasc[2].strip()[:10]

    # Extração do CEP
    endereco = linhas[2]
    cep_match = re.search(r'CEP[:\s]*([\d.-]+)', endereco)
    if cep_match:
        cep = re.sub(r'[.-]', '', cep_match.group(1))
        dados_cv[campos['Cep']] = cep

    # Telefone
    dados_cv[campos['Telefone']] = linhas[3] if linhas[3].endswith('R') == False else linhas[4]

    # Email
    email_match = re.search(r'Email\s*:\s*(\S+)', linhas[5])
    if email_match:
        dados_cv[campos['Email']] = email_match.group(1).strip()

    # Pretensão Salarial
    pretencao = linhas[6]
    if pretencao != '':
        pretencao_match = re.search(r'\$[\d,]+', pretencao)
        if pretencao_match:
            dados_cv[campos['PretencaoSalarial']] = pretencao_match.group()

    # Processamento dos títulos e conteúdo
    indice = 7
    while indice < len(linhas):
        linha = linhas[indice]
        if eh_titulo(linha):
            titulo = linha
            indice += 1
            conteudo = []
            while indice < len(linhas) and not eh_titulo(linha := linhas[indice]):
                if linha != '':
                    conteudo.append(linha)
                indice += 1

            if titulo == 'OBJETIVO':
                dados_cv[campos['Objetivo']] = ' '.join(conteudo)
            elif titulo == 'SÍNTESE':
                dados_cv[campos['Sintese']] = ' '.join(conteudo)
            elif titulo == 'FORMAÇÃO':
                for i, formacao in enumerate(conteudo[:3]):
                    dados_cv[campos['Formacao'][i]] = formacao
            elif titulo == 'EXPERIÊNCIA PROFISSIONAL':
                for i, experiencia in enumerate(conteudo[:5]):
                    dados_cv[campos['ExperienciaProfissional'][i]] = experiencia
            elif titulo == 'CURSOS EXTRACURRICULARES':
                dados_cv[campos['CursoExtraCurricular']] = ' '.join(conteudo)
            elif titulo == 'IDIOMAS':
                for idioma in conteudo[:4]:
                    partes = idioma.split('-')
                    if len(partes) == 2:
                        dados_cv[campos['Idiomas'][0]] = partes[0].strip()
                        dados_cv[campos['Idiomas'][1]] = partes[1].strip()
            elif titulo == 'INFORMÁTICA':
                dados_cv[campos['Informatica']] = ' '.join(conteudo)
            elif titulo == 'INFORMAÇÕES ADICIONAIS':
                dados_cv[campos['InfoAdicional']] = ' '.join(conteudo)
        else:
            indice += 1

    return dados_cv

def atualizar_excel(dados: List[str], arquivo_excel: str, sheet_name: str) -> None:
    """
    Adiciona os dados do CV à planilha Excel especificada.
    """
    wb = openpyxl.load_workbook(arquivo_excel)
    ws = wb[sheet_name]
    ws.append(dados)
    wb.save(arquivo_excel)
    wb.close()

def main():
    # Lista todos os arquivos PDF no diretório especificado
    arquivos_pdf = [arquivo for arquivo in os.listdir(PDF_DIR) if arquivo.lower().endswith('.pdf')]

    for arquivo_pdf in arquivos_pdf:
        caminho_pdf = os.path.join(PDF_DIR, arquivo_pdf)
        cv_texto = ler_pdf(caminho_pdf)

        nome_txt = os.path.splitext(arquivo_pdf)[0] + '.txt'
        caminho_txt = os.path.join(TXT_DIR, nome_txt)

        # Processo de conversão de PDF para TXT
        criar_e_preencher_txt(caminho_txt, cv_texto)
        limpar_txt(caminho_txt)
        separar_nome(caminho_txt)
        adicionar_espaco(caminho_txt)

        # Processa o arquivo TXT para extrair os dados do CV
        dados_cv = processar_txt(caminho_txt)

        # Atualiza a planilha Excel com os dados extraídos
        atualizar_excel(dados_cv, EXCEL_FILE, EXCEL_SHEET)

        print(f"Processado: {arquivo_pdf}")

if __name__ == "__main__":
    main()
