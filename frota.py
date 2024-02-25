import openpyxl
import matplotlib.pyplot as plt
import numpy as np

workbook = openpyxl.load_workbook('dados (4).xlsx')
worksheet = workbook.active

valores_est = []
valores_loc = []
valores_lar = []
valores_tot = []
datas = []

for row_number, row in enumerate(worksheet.iter_rows(values_only=True), start=1):
    if row_number > 133:
        break
    data, valor_est, valor_loc, valor_tot, valor_lar = row[0:5]
    valores_est.append(valor_est)
    valores_loc.append(valor_loc)
    valores_lar.append(valor_lar)
    valores_tot.append(valor_tot)
    datas.append(data)

valores_est = valores_est[::-1]
valores_loc = valores_loc[::-1]
valores_lar = valores_lar[::-1]
valores_tot = valores_tot[::-1]
datas = datas[::-1]

coeficientes = np.polyfit(range(len(valores_tot)), valores_tot, 1)
linha_tendencia = np.polyval(coeficientes, range(len(valores_tot)))

plt.plot(datas, valores_est, color="green", linestyle="solid", markersize=6)
plt.plot(datas, valores_loc, color="blue", linestyle="solid", markersize=6)
plt.plot(datas, valores_lar, color="cyan", linestyle="solid", markersize=6)
plt.plot(datas, valores_tot, color="salmon", linestyle="solid", markersize=6)
plt.plot(datas, linha_tendencia, color="purple", linestyle="dashed", label="Linha de Tendência")

plt.xlabel('Data de pesquisa'
plt.ylabel('Frota')
plt.title('Frota de ônibus da SPTrans')

plt.xticks(datas[::12], rotation=45, fontsize=6)
plt.tight_layout()

plt.savefig('result.png')
plt.close('all')
