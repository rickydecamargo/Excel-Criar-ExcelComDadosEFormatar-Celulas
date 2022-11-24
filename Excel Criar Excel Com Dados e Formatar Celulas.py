#IMPORTANDO IMAGEM DO COMPUTADOR PARA O EXCEL

import xlsxwriter as opcoesDoXlsxWriter
import os

#1 - indicando onde será criado o arquivo, seu nome e sua extensão. Importante a questão das barras duplas (testar).
nomeCaminhoArquivo = 'C:\\Users\\Windows\\Desktop\\Python Projetos\\xlsxwriter\ArquivoImagem.xlsx'
workbook = opcoesDoXlsxWriter.Workbook(nomeCaminhoArquivo)
sheetPadrao = workbook.add_worksheet("Dados") #Para renomear o nome da Sheet1 para Dados.

sheetPadrao.write("A1", "Ricky")

sheetPadrao.insert_image('B2', 'C:\\Users\\Windows\\Desktop\\Python Projetos\\xlsxwriter\Foto.jpg' )

#3 - Para fechar e salvar as informações
workbook.close()

#4 - Abrir o arquivo para verificar o resultado
os.startfile(nomeCaminhoArquivo)
