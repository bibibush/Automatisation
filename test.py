import openpyxl
import math
import os

dir_path = '//srvlabreche/Dossier semaine commun'
stock_path = 'Stock/Stock Produits finis.xlsm'

stock_wb = openpyxl.load_workbook(os.path.join(dir_path, stock_path), data_only= True, keep_vba=True)

hm = stock_wb.worksheets[4]
halal = stock_wb.worksheets[1]



specialite_data = [halal['AK84'].value,]
specialite_s = [8]
specialite_TOTAL = [halal['R85'].value,]
specialite_melee = [(halal['AK82'].value * halal['AA82'].value)]
specialite_m = [halal['AA82'].value,]

data = []
for i in range(len(specialite_data)):
    if specialite_data[i] >= specialite_s[i]:
        x = 0
        data.append(x)
    else:
        y = specialite_s[i] * specialite_melee[i] + specialite_melee[i] - specialite_TOTAL[i]
        x = y / specialite_m[i]
        x = round(x, 1)
        z = math.ceil(x)
        if (z * 10) - (x * 10) >= 5:
            data.append(z - 0.5)
        elif z == 0:
            z = 0.5
            data.append(z)
        else:
            data.append(z)

print(data)