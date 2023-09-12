#########################################################
#################### LIBS UTILIZADAS ####################
#########################################################

import datetime
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from cabecalho import criar_cabecalho
from formulas import escrever_formulas
from estilizar import aplicar_estilos
from envio_email import anexar_e_enviar
from conexao_banco import consultar

#########################################################
####### CAPTURA DO MÊS E ANO ATUAIS PARA CONSULTA #######
#########################################################

data_atual = datetime.datetime.now()
data_formatada = data_atual.strftime("%d/%m/%Y")
mes_atual = data_atual.month
ano_atual = data_atual.year

if(mes_atual == 1): #CONDICIONAL PARA A VIRADA DE ANO
    mes_atual = 12
    ano_atual = ano_atual - 1

#########################################################
######## CRIAÇÃO DA PLANILHA E ABERTURA DA MESMA ########
#########################################################

workbook = Workbook()
sheet = workbook.active

#########################################################
######## LAÇO DE IMPRESSÃO DO DATASET E ATRIBUTOS #######
#########################################################

intervalo_cabecalho = 60  #LINHAS DE DATASET A SEREM REPRODUZIDAS A CADA CABEÇALHO

linha_inicial = 0   #LINHA INICIAL DA PLANILHA
linha_atual = linha_inicial #LINHA ATUAL A RECEBE COMO PARAMETRO
contador_linhas_dados = 0  #CONTADOR DE LINHAS QUE FORAM ESCRITAS COM DADOS

resultado = consultar(mes_atual) #MES ATUAL PASSADO COMO PARAMETRO PARA A CONSULTA DO DATAFRAME QUE SERÁ ARMAZENADO NO RESULTADO
total_linhas = len(resultado) #LEN UTILIZADO PARA SABER O TOTAL DE LINHAS DESSE DATAFRAME


def imprimir_cabecalho(linha): #FUNÇÃO QUE CHAMA O CABEÇALHO
    criar_cabecalho(sheet, linha + 1, data_formatada)


imprimir_cabecalho(linha_atual) #CHAMADA INICIAL DO CABEÇALHO
linha_atual += 10 #LINHA ATUAL RECEBE O SEU VALOR MAIS O TAMANHO DO CABEÇALHO

while linha_atual < total_linhas:

    linhas_a_escrever = min(intervalo_cabecalho - contador_linhas_dados, total_linhas - linha_atual)

    for _ in range(linhas_a_escrever):
        linha_dados = resultado.iloc[linha_atual - linha_inicial]
        sheet.append(linha_dados.to_list())
        linha_atual += 1
        contador_linhas_dados += 1
        escrever_formulas(sheet, linha_atual)
        aplicar_estilos(sheet, linha_atual - 1)

    if contador_linhas_dados == intervalo_cabecalho: #VERIFICANDO SE É NECESSÁRIO ESCREVER O CABEÇALHO MAIS UMA VEZ
        imprimir_cabecalho(linha_atual)
        contador_linhas_dados = 0 
        linha_atual += 10

#########################################################
############# AJUSTE DE LARGURA DAS COLUNAS #############
#########################################################

sheet.column_dimensions['A'].auto_size = True
sheet.column_dimensions['J'].width = 35

#########################################################
############### SALVAR E ENVIAR PLANILHA ################
#########################################################

workbook.save("relatorios/LAGOSTA - RELATÓRIO SECRETARIA DE AQUICULTURA E PESCA - MES - {} - ANO - {}.xlsx".format(mes_atual, ano_atual))
anexar_e_enviar(mes_atual, ano_atual)