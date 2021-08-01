# .xls to .xlsx converter
# by Lucas Mello dos Santos - Brazil
# github.com/lmello0/

import xlrd
import xlsxwriter as xlsx
import os
from time import strftime
import time

def fPROGRAMA(vDIRETORIO, vCONT_ARQ):
    for root, dirs, files in os.walk(vDIRETORIO):
        for name in files:
            if name[-4:] == '.xls':
                vCONT_ARQ = vCONT_ARQ + 1
                print("[ FILE: {} ]".format(name))

                vXLS = str(root) + "\\" + str(name)
                
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

    return vCONT_ARQ


vDIRETORIO = str(input('Path: '))

vINICIO_TEMPO = time.time()
vDIRETORIO = vDIRETORIO.replace("\\", "\\\\")
vCONT_ARQ = 0


#VAI PARA DIRETORIO
os.chdir(vDIRETORIO)

print("\n")
vCONT_ARQ = fPROGRAMA(vDIRETORIO, vCONT_ARQ)
vFIM_TEMPO = time.time()

vEXEC_TIME = (vFIM_TEMPO - vINICIO_TEMPO) * 1000

print("\nConversion completed\n{} files converted".format(vCONT_ARQ))
print("Execution time: {:.2f}ms\n".format(vEXEC_TIME))