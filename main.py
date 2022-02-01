# Importação das bibliotecas

import os
from win32com.client import Dispatch
from numpy import trapz, mean, std, max, min, sqrt, zeros, power
import xlsxwriter as xlsx


# Levantamento dos arquivos e respectivos nomes
files = []

for file in os.listddeactivateir():

    # Confere a extensão do arquivo
    if file.endswith(".LTX") or file.endswith(".LTD") or file.endswith(".TEM"):
        files.append(file)


# Coloca em ordem crescente
files = sorted(files)

# Criação da planilha onde serão salvos os dados
table = xlsx.Workbook("Resultados.xlsx")
sheet = table.add_worksheet(name="Estátisticas")


# Abrindo o server para manipular arquivos Lynx
oFileTS = Dispatch("LynxFile.FileTS")

# Abre um arquivo para retirar alguns parâmetros
r = oFileTS.OpenFile(files[0])

# Número de arquivos
n_files = len(files)

# Número de canais dos arquivos
n_channels = oFileTS.nChannels

# Frequência de aquisição dos sinais
freq_aq = oFileTS.SampleFreq

# Número total de amostras dos arquivos
n_samples = oFileTS.nSamples

# Número de amostras necessárias para 30 segundos de análise
samples = int(30 * freq_aq)

# Iteração entre os arquivos
for i in range(0, n_files):

    # Posiciona o marcador de linha
    row = 10 * i
    col = 0

    # Cria a coluna de nomes
    sheet.write(row, col, files[i])
    sheet.write(row + 2, col, "Amostras")
    sheet.write(row + 3, col, "Máximo")
    sheet.write(row + 4, col, "Mínimo")
    sheet.write(row + 5, col, "Máx - Mín")
    sheet.write(row + 6, col, "Média")
    sheet.write(row + 7, col, "Desvio Padrão")
    sheet.write(row + 8, col, "Área")
    sheet.write(row + 9, col, "RMS")

    # Abre o Arquivo
    r = oFileTS.OpenFile(files[i])

    for j in range(0, n_channels):

        # Posiciona o marcador da coluna
        col = j

        # Cria uma array para armazenar o sinal do canal
        Buf = zeros(samples)

        # Leitura das medições do canal
        r, Buf, NOut = oFileTS.ReadBuffer(j, (n_samples - samples), samples, Buf)

        # Processamento dos dados do canal
        channel_name = oFileTS.SnName(j)
        sheet.write(row, col + 1, channel_name)

        channel_unit = oFileTS.SnUnit(j)
        sheet.write(row + 1, col + 1, channel_unit)

        channel_samples = samples
        sheet.write(row + 2, col + 1, channel_samples)

        channel_max = max(Buf)
        sheet.write(row + 3, col + 1, channel_max)

        channel_min = min(Buf)
        sheet.write(row + 4, col + 1, channel_min)

        channel_max_min = channel_max - channel_min
        sheet.write(row + 5, col + 1, channel_max_min)

        channel_mean = mean(Buf)
        sheet.write(row + 6, col + 1, channel_mean)

        channel_mean = std(Buf)
        sheet.write(row + 7, col + 1, channel_mean)

        channel_area = trapz(Buf, dx=(1 / freq_aq))
        sheet.write(row + 8, col + 1, channel_area)

        channel_rms = sqrt(mean(power(Buf, 2)))
        sheet.write(row + 9, col + 1, channel_rms)

table.close()
