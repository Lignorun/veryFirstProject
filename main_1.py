import os
import PyPDF2
import openpyxl

__doc


# Função que lê o PDF e transforma em string, atribuindo a uma variável que será manipulada
def LerPDF (aqrquivoPDF):
    pdf_file = open(aqrquivoPDF, 'rb')
    read_pdf = PyPDF2.PdfFileReader(pdf_file)
    number_of_pages = read_pdf.getNumPages()
    page = read_pdf.getPage(0)
    page_content = page.extractText().replace('Currículo por EMPREGARE.com - Software de Recrutamento e Seleção', '').replace('Criado por EMPREGARE.com - Software de Recrutamento e Seleção', '').strip().replace('•','')
    parsed = ''.join(page_content) 
    parsed2=''
    if number_of_pages == 2:
        page = read_pdf.getPage(1)
        page_content = page.extractText().replace('Currículo por EMPREGARE.com - Software de Recrutamento e Seleção', '').replace('Criado por EMPREGARE.com - Software de Recrutamento e Seleção', '').strip().replace('•','')
        parsed2 = ''.join(page_content)
    CV_String = parsed + parsed2
    return CV_String


#Só cria os arquivos de txt
def Cria_e_preenche_TXT (Titulo, conteudo):
    arquivo = open(Titulo, 'w') 
    arquivo.write(conteudo)
    arquivo.close()
    return arquivo

#Retira os espaços vazios do início do arquivo
def LimpaTXT(arq):
    with open(arq,"r+") as f:
        new_f = f.readlines()
        f.seek(0)
        for line in new_f:
            if line !='\n':
                f.write(line) 
        f.close()
    return

#corrige o erro que junta o nome com o sexo
def SeparaNome(arq):
        with open(arq,"r+") as f:
            f.seek(0)
            linha1 = f.read()
            if linha1.find("Masculino") != -1:
                f.seek(linha1.find("Masculino"))
                f.write('\n' + linha1[linha1.find("Masculino"):])
                f.close()
                return
            elif linha1.find("Feminino") != -1:
                f.seek(linha1.find("Feminino"))
                f.write('\n' + linha1[linha1.find("Feminino"):])
                f.close()
                return
            elif linha1.find("Não Informado") != -1:
                f.seek(linha1.find("Não Informado"))
                f.write('\n' + linha1[linha1.find("Não Informado"):])
                f.close()
                return
            else:
                f.close()
                return

#Procura pelos títulos. Caso não encontre, preenche com um espaço vazio
#Caso encontre e não tenha um espaço antes dele, adiciona o espaço 



def AdicionaEspaco(arq):           #ferifica cada linha para ver se é um título e adiciona um espaço, caso não tenha
    with open(arq,"r+") as f:
        new_f = f.readlines()
        f.seek(0)
        titulo = 0
        espaco = 0
        for line in new_f:
            if line == ' \n':
                f.write(line)
                espaco = 1
            elif line == "OBJETIVO\n":
                titulo = 1
            elif line == 'SÍNTESE\n':
                titulo = 1
            elif line == 'FORMAÇÃO\n':
                titulo = 1
            elif line == 'EXPERIÊNCIA PROFISSIONAL\n':
                titulo = 1
            elif line == 'CURSOS EXTRACURRICULARES\n':
                titulo = 1
            elif line == 'IDIOMAS\n':
                titulo = 1
            elif line == 'INFORMÁTICA\n':
                titulo = 1
            elif line == 'INFORMAÇÕES ADICIONAIS\n':
                titulo = 1    
            else:
                f.write(line)
                titulo = 0
                espaco = 0
            if titulo == 1 and espaco == 0:
                f.write(' \n' + line)
            elif titulo == 1 and espaco == 1:
                f.write(line)
        f.close()
    return
'''

def AdicionaTituloFaltante(arq):           #Varre o texto e procura algum título que não tenha e adiciona ele com uma linha em branco
    with open(arq,"r+") as f:
        new_f = f.readlines()
        f.seek(0)
        titulo1 = 0
        titulo2 = 0
        for line in new_f:
            if line == ' \n':
                f.write(line)
                espaco = 1
            elif line == "OBJETIVO\n":
                titulo = 1
            elif line == 'SÍNTESE\n':
                titulo = 1
            elif line == 'FORMAÇÃO\n':
                titulo = 1
            elif line == 'EXPERIÊNCIA PROFISSIONAL\n':
                titulo = 1
            elif line == 'CURSOS EXTRACURRICULARES\n':
                titulo = 1
            elif line == 'IDIOMAS\n':
                titulo = 1
            elif line == 'INFORMÁTICA\n':
                titulo = 1
            elif line == 'INFORMAÇÕES ADICIONAIS\n':
                titulo = 1    
            else:
                f.write(line)
                titulo = 0
                espaco = 0
            if titulo == 1 and espaco == 0:
                f.write(' \n' + line)
            elif titulo == 1 and espaco == 1:
                f.write(line)
        f.close()
    return

'''

# OBJETIVO
# SÍNTESE
# FORMAÇÃO
# EXPERIÊNCIA PROFISSIONAL
# CURSOS EXTRACURRICULARES
# IDIOMAS
# INFORMÁTICA
# INFORMAÇÕES ADICIONAIS

def ehTitulo(linhaTitulo):    #Índice de títulos para adicionar

    if linhaTitulo == 'OBJETIVO\n':
        return True
    elif linhaTitulo == 'SÍNTESE\n':
        return True
    elif linhaTitulo == 'FORMAÇÃO\n':
        return True
    elif linhaTitulo == 'EXPERIÊNCIA PROFISSIONAL\n':
        return True
    elif linhaTitulo == 'CURSOS EXTRACURRICULARES\n':
        return True
    elif linhaTitulo == 'IDIOMAS\n':
        return True
    elif linhaTitulo == 'INFORMÁTICA\n':
        return True
    elif linhaTitulo == 'INFORMAÇÕES ADICIONAIS\n':
        return True
    else:
        return False

    
def IniciaDadosCV():
    ListaCV = []
    # Nome = 0
    # Sexo = 1
    # EstadoCivil = 2
    # Nascimento = 3
    # Cep = 4
    # Telefone = 5
    # Email = 6
    # PretencaoSalarial = 7
    # Objetivo = 8
    # Sintese = 9
    # Formação1 = 10
    # Formação2 = 11
    # Formação3 = 12
    # Experiencia1 = 13
    # Experiencia2 = 14
    # Experiencia3 = 15
    # Experiencia4 = 16
    # Experiencia5 = 17 
    # CursoExtraCurricular = 18
    # Idioma1 = 19
    # NivelIdioma1 = 20
    # Idioma2 = 21
    # NivelIdioma2 = 22
    # Idioma3 = 23
    # NivelIdioma3 = 24
    # Idioma4 = 25
    # NivelIdioma4 = 26
    # Informatica = 27
    # InfoAdicional = 28

    for i in range(29):
        ListaCV.append('') 
    return ListaCV




pathPDF = r"C:/Users/dlins/Python/CV_Automação/corporate_empregare/CV_PDF" 
pathTXT = r"C:/Users/dlins/Python/CV_Automação/corporate_empregare/CV_TXT"

#Inicia o programa convertendo todos os .docx para txt em outra pasta, conferindo se o arquivo já foi criado

arquivosDOC = os.listdir()  #Cria uma lista com os nomes de todos os .pdf para processar

for arquivo in arquivosDOC:
    if arquivo.endswith(".pdf"):
        CV = LerPDF(arquivo)
        nomeTXT = arquivo.replace(".pdf", ".txt")
        Cria_e_preenche_TXT (nomeTXT, CV)
        LimpaTXT(nomeTXT)
        SeparaNome(nomeTXT)
        AdicionaEspaco(nomeTXT)
       

        #Inicializa todas as variáveis que cada CV terá do candidato em string, que será feito uma lista com cada conteúdo e, 
        # em seguida adicionada ao final do arquivo de excel

        #Dados
        DadosCV = IniciaDadosCV()
        #Indices
        id = 0
        INome = id
        id += 1
        ISexo = id
        id += 1
        IEstadoCivil = id
        id += 1
        INascimento = id
        id += 1
        ICep = id
        id += 1
        ITelefone = id
        id += 1
        IEmail = id
        id += 1
        IPretencaoSalarial = id
        id += 1
        IObjetivo = id
        id += 1
        ISintese = id
        id += 1
        IFormação1 = id
        id += 1
        IFormação2 = id
        id += 1
        IFormação3 = id
        id += 1
        IExperiencia1 = id
        id += 1
        IExperiencia2 = id
        id += 1
        IExperiencia3 = id
        id += 1
        IExperiencia4 = id
        id += 1
        IExperiencia5 = id
        id += 1 
        ICursoExtraCurricular = id
        id += 1
        IIdioma1 = id
        id += 1
        INivelIdioma1 = id
        id += 1
        IIdioma2 = id
        id += 1
        INivelIdioma2 = id
        id += 1
        IIdioma3 = id
        id += 1
        INivelIdioma3 = id
        id += 1
        IIdioma4 = id
        id += 1
        INivelIdioma4 = id
        id += 1
        IInformatica = id
        id += 1
        IInfoAdicional = id
        
        f = open(nomeTXT,"r+")              #Começa a varrer o arquivo para adicionar no excel
        New_f = f.readlines()
        f.seek(0)
        DadosCV[INome] = f.readline()                   #1a linha e adiciona: Nome
        linha = f.readline().split(',')                 #2a linha
        DadosCV[ISexo] = linha[0].strip()               #Adiciona: Sexo            
        DadosCV[IEstadoCivil] = linha[1].strip()        #Adiciona: Estado civil
        DadosCV[INascimento] = linha[2].strip()[:10]    #Aduciona: Data de nascimento
        endereco = f.readline()                         #3a linha, endereço
        if endereco.endswith('R\n') == False:
            endereco = endereco[:len(endereco)-1] + ' '
            endereco = endereco + f.readline()


        if endereco.find('CEP:') != -1:
            cep = endereco.replace('.', '')
            cep = cep.replace('-', '')
            indiceCEP = cep.find('CEP:')
            DadosCV[ICep] = cep[indiceCEP + 5 :indiceCEP + 13]

        linha = f.readline()                            #4a Linha
        if linha.endswith('R\n'):
            linha = f.readline()
        DadosCV[ITelefone] = linha                      #Adiciona: Telefone
        linha = f.readline().split(':')
        DadosCV[IEmail] = linha[1].strip()              #Adiciona: Email
        linha = f.readline() 
        if linha != ' \n':                              #Caso tenha, adiciona a pretenção salarial
            linha = linha.split(':')
            linha = linha[1].strip()
            DadosCV[IPretencaoSalarial] = linha[linha.index('$')-1:]
            linha = f.readline()
        #Finalizado a leitura da primeira etapa do documento
        #Aparti daqui, Verifica o título para fazer a leitura diferenciada
        linha = f.readline()
        if linha == 'OBJETIVO\n':
            linha = f.readline()
            while linha != ' \n':
                DadosCV[IObjetivo] =  DadosCV[IObjetivo] + linha
                linha = f.readline()
            linha = f.readline()
      

        if linha == 'SÍNTESE\n':
            linha = f.readline()
            while linha != ' \n':
                DadosCV[ISintese] =  DadosCV[ISintese] + linha
                linha = f.readline()
            linha = f.readline()
        

        if linha == 'FORMAÇÃO\n':
            linha = f.readline()
            while linha != ' \n':
                DadosCV[IFormação1] =  DadosCV[IFormação1] + linha
                linha = f.readline()
                IFormação1 += 1
            linha = f.readline()
        

        if linha == 'EXPERIÊNCIA PROFISSIONAL\n':
            linha = f.readline()
            while linha != ' \n' and linha != None and linha != '':
                DadosCV[IExperiencia1] =  DadosCV[IExperiencia1] + linha
                linha = f.readline()
            IExperiencia1 += 1
            while (ehTitulo(linha) == False and linha != None and linha != ''):
                DadosCV[IExperiencia1] =  DadosCV[IExperiencia1] + linha
                linha = f.readline()
                if linha == ' \n':
                    IExperiencia1 += 1
               

        if linha == 'CURSOS EXTRACURRICULARES\n':
            linha = f.readline()
            while linha != ' \n' and linha != None and linha != '':
                DadosCV[ICursoExtraCurricular] =  DadosCV[ICursoExtraCurricular] + linha
                linha = f.readline()
            if linha != None:
                linha = f.readline() 
            

        if linha == 'IDIOMAS\n':
            linha = f.readline()
            while linha != ' \n' and linha != None and linha != '':          
                if linha.find('-') != -1:
                    idiomas = linha.split('-')
                    DadosCV[IIdioma1] = idiomas[0].strip()                        
                    DadosCV[INivelIdioma1] = idiomas[1].strip()        
                    IIdioma1 += 2
                    INivelIdioma1 += 2
                linha = f.readline()
            if linha != None:
                linha = f.readline() 


        if linha == 'INFORMÁTICA\n':
            linha = f.readline()
            while linha != ' \n' and linha != None and linha != '':
                DadosCV[IInformatica] =  DadosCV[IInformatica] + linha
                linha = f.readline()
            if linha != None:
                linha = f.readline() 


        if linha == 'INFORMAÇÕES ADICIONAIS\n':
            linha = f.readline()
            while linha != ' \n' and linha != None and linha != '':
                DadosCV[IInfoAdicional] =  DadosCV[IInfoAdicional] + linha
                linha = f.readline()


        #print(DadosCV)



        f.close()



        wb = openpyxl.load_workbook('zCurriculos.xlsx')  
        ws = wb['CV_Bruto']
        ws = wb.active
        ws.append(DadosCV)
        wb.save('zCurriculos.xlsx')

        wb.close()

        #print(DadosCV)
        #print('\n')
    
