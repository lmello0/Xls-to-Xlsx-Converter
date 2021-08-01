import xlrd
import xlsxwriter as xlsx
import os
import glob
from time import strftime

vDIRETORIO = 'C:\\Users\\'

#VAI PARA DIRETORIO
os.chdir(vDIRETORIO)

for dirs in os.walk(vDIRETORIO, topdown=False):
    for vXLS in glob.glob('*.xls'):
        #TRATAMENTO DO NOME
        vNOME_SAIDA = str(vXLS)
        vNOME_SAIDA = vNOME_SAIDA.replace(".xls", "") + strftime("_%Y%m%d_%H%M") + ".xlsx"

        #NOME ARQUIVO SAIDA
        vSAIDA_XLSX   = xlsx.Workbook(vNOME_SAIDA)
        #CRIA PLANILA P REGISTRO
        vWORKSHEET    = vSAIDA_XLSX.add_worksheet('LOGS')
        #ABRE XLS
        vABRE_ARQUIVO = xlrd.open_workbook(vXLS) ##
        #USA XLS POR INDICE
        vARQ          = vABRE_ARQUIVO.sheet_by_index(0)

        for x in range(vARQ.nrows):
            vCONT = 0
            #COLETA DE CONTEUDO
            vCONTEUDO_ENTRADA = vARQ.row_values(x, start_colx=0, end_colx=None)

            #ESCRITA NO ARQUIVO
            vCONTEUDO_SAIDA = vWORKSHEET.write_row(x, vCONT, vCONTEUDO_ENTRADA)
            vCONT += 1

        vSAIDA_XLSX.close()