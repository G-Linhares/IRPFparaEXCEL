#! python3.6
import sys
import time
import os
import openpyxl
from openpyxl.worksheet.dimensions import ColumnDimension

##Decidi setar todas as variáveis pra 0 ou vazio antes que os problemas começassem a surgir.
n_cpf = 0
nome_cliente = ""
data_nascimento = 0
numero_recibo = 0
cliente = [n_cpf,nome_cliente,data_nascimento,numero_recibo] 
cont = 0
linha_excel = 2
numero_linha = 0
coluna_excel = 'A'
letra_coluna_excel = ""  
coordenadas_excel = ""
##
os.system("cls")
os.system("mode con cols=100 lines=40")
os.system("color 1f")
os.system("Title IRPF PARA EXCEL")
print ('''
AVISOS:\n\n *** CERTIFIQUE-SE DE TER ALGUM PROGRAMA QUE ABRA\n ARQUIVOS .XLSX INSTALADO NA MÁQUINA OU O PROGRAMA DARÁ ERRO.\n
 *** PREFIRA EXECUTAR O PROGRAMA COM PRIVILÉGIOS DE ADMINISTRADOR, \n POIS PODEM HAVER PROBLEMAS COM RELAÇÃO A PERMISSÃO DE ACESSO \nA PASTAS DO WINDOWS.\n  
 *** ESTE PROGRAMA ASSUME QUE VOCÊ TENHA O PROGRAMA DA RECEITA \n INSTALADO NO DIRETÓRIO PADRÃO, CASO ESTE NÃO SEJA O CASO, ENVIE-ME UM E-MAIL:\n Lenharesg@outlook.com \n 
 *** ESTE PROGRAMA FOI DISTRIBUIDO DE FORMA GRATUÍTA.\n
''')
print("Digite qual o ano do IR que deseja consultar... \n(Necessário possuir o programa do ano a pesquisar no computador...)\n")
try:
    
    while True:
        ano_consulta = input()
        ano = ano_consulta
        ano_consulta = "".join(['IRPF',ano_consulta])
        caminho = ('C:\\Arquivos de Programas RFB\\')
        if ano_consulta.isdecimal():
            if int(ano_consulta) <= 2007:
                print("Ano de consulta inserido é inválido!\n Ano de verificação inserido é menor ou igual a 2007!\n")
                continue
            continue
        if ano_consulta not in os.listdir(caminho):
            print("Ano de consulta inserido inválido!\nVerifique se o programa do ano requisitado existe na máquina e tente novamente!")
            continue  
        break
    ##Cria arquivo do Excel e define as colunas e seus respectivos nomes.
    wb = openpyxl.Workbook()
    wb.save('C:\\Arquivos de Programas RFB\RelatórioIRPF' + str(ano) + '.xlsx')
    sheet = wb["Sheet"]
    sheet['A1'] = "C.P.F"
    sheet['B1'] = "Nome"
    sheet['C1'] = "Nascimento"
    sheet['D1'] = "Recibo do ano anterior"
    sheet['E1'] = "Fonte de rendimentos"
    sheet['F1'] = "Segunda fonte de rendimentos"
    ##
    ano_consulta = "".join([ano_consulta,'\\transmitidas'])
    caminho = 'C:\\Arquivos de Programas RFB\\'
    caminho_transmitidas = ''.join([caminho,ano_consulta,'\\'])
    print(ano_consulta)
    sheet.title = ano
    sheet = wb[ano]
    inicio_programa = time.time()

    #Obtem lista de todos os arquivos da Receita na pasta.
    pasta_IRPF = os.listdir(caminho_transmitidas)
    ##
    for file in pasta_IRPF:
    #Seleciona apenas os arquivos com informações (pois na pasta existem arquivos sem informações úteis).
        if file.endswith(".DEC"):
            nome_retif = ''.join([file[:29],'RETIF.DEC'])
    ##
            ##Verifica se o arquivo do imposto atual tem retificadora na mesma pasta (para sempre obter os dados mais recentes)
            if (nome_retif in pasta_IRPF and file == nome_retif) or (nome_retif not in pasta_IRPF):
            ##
             ##Após todas as verificações, insere "máscara" no CPF atual e começa a extrair os dados.
                cpf = file[:11]
                cpflista = list(cpf)
                cpflista.insert(3, '.')
                cpflista.insert(7, '.')
                cpflista.insert(11, '-')
                cpf = ''.join(cpflista)
                caminho_arquivo = "".join([caminho_transmitidas, file])
                abrir_arquivo = open(caminho_arquivo)
                conteudo = abrir_arquivo.read()
                nome_cliente = conteudo[39:80]
                nome_cliente = nome_cliente.strip()        
                print('Processando o CPF : ',cpf)
                if (conteudo[213:215] == '10') or ("RETIF.DEC" in file):
                    localizacao_recibo = conteudo.find("10SS")
                    if localizacao_recibo != -1:
                        localizacao_recibo += 6
                        numero_recibo = conteudo[localizacao_recibo:localizacao_recibo+12]
                    else:    
                        numero_recibo = "000000000000"
                else:              
                    numero_recibo = conteudo[203:215]
                with open(caminho_arquivo) as fp:  
                     for line in fp:
                        numero_linha += 1
                        if numero_linha == 5:
                            empresa = line[27:80]
                            empresa.strip(" ")
                            if empresa.isdecimal() or line[27:40].isdecimal():
                                empresa = "Fonte de rendimentos não cadastrada ou inexistente."
                        if numero_linha == 6:
                            empresa2 = line[27:80]
                            empresa2.strip(" ")
                            try:
                                teste_logico = int(line[:27])
                                try:
                                    empresa2 = int(empresa2)
                                    empresa2 = " "
                                except ValueError:
                                    try:
                                        int(empresa2[:4])
                                        empresa2 = ""
                                    except ValueError:
                                        pass
                            except ValueError:
                                empresa2 = " "
                data_nascimento = conteudo[112:120]
                #adiciona as barras à data de nascimento:
                n_cpf = cpf
                data_lista = list(data_nascimento)
                data_lista.insert(2, '/')
                data_lista.insert(5, '/')
                data_nascimento = ''.join(data_lista)
                coordenadas_excel = "".join([coluna_excel,str(linha_excel)])
                #---------------
                sheet[coordenadas_excel] = n_cpf
                coluna_excel = coluna_excel = chr(ord(coluna_excel)+1)
                coordenadas_excel = "".join([coluna_excel,str(linha_excel)])
                #---------------
                sheet[coordenadas_excel] = nome_cliente            
                coluna_excel = chr(ord(coluna_excel)+1)
                coordenadas_excel = "".join([coluna_excel,str(linha_excel)])
                #---------------
                sheet[coordenadas_excel] = data_nascimento            
                coluna_excel = chr(ord(coluna_excel)+1)
                coordenadas_excel = "".join([coluna_excel,str(linha_excel)])
                #---------------
                sheet[coordenadas_excel] = numero_recibo            
                coluna_excel = chr(ord(coluna_excel)+1)
                coordenadas_excel = "".join([coluna_excel,str(linha_excel)])
                #---------------
                sheet[coordenadas_excel] = empresa            
                coluna_excel = coluna_excel = chr(ord(coluna_excel)+1)
                coordenadas_excel = "".join([coluna_excel,str(linha_excel)])
                #---------------
                sheet[coordenadas_excel] = empresa2            
                coluna_excel = coluna_excel = chr(ord(coluna_excel)+1)
                coordenadas_excel = "".join([coluna_excel,str(linha_excel)])
                #---------------
             ##
                ##Define onde vai serão armazenadas as informações obtidas e salva o arquivo
                linha_excel += 1
                coluna_excel = 'A'
                coordenadas_excel = "".join([coluna_excel,str(linha_excel)])
                wb.save('C:\\Arquivos de Programas RFB\RelatórioIRPF' + str(ano) + '.xlsx')
                ##
                abrir_arquivo.close()
                numero_linha = 0
    fim_programa = time.time()
    tempo = fim_programa - inicio_programa
    os.system("cls")
    print("Programa criado por Lenhares\nE-mail: Lenharesg@outlook.com\n Favor não distribuir sem os devidos créditos.")
    time.sleep(1.5)
    print("Transferência de dados concluida em", round(tempo,2) ,"segundos!\n O arquivo com a saída de dados foi salvo no\n\"C:\\Arquivos de Programas RFB\" com o nome /'RelatórioIRPF" + str(ano) + ".xlsx'. \nAperte Enter para continuar...")
    input()                          
except:
    print("Ocorreu um erro! Verifique se a pasta do imposto de renda está em branco ou se os arquivos estão corrompidos!")
