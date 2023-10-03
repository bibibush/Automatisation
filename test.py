import openpyxl
import math
import os

dir_path = '//srvlabreche/Dossier semaine commun'
planning_path = 'Production/Planification de la production/Semaine 40 - Planning Fabrication.xlsm'
save_path = 'Production/Planification de la production/Test Semaine 40 - Planning Fabrication.xlsm'
stock_path = 'Stock/Stock Produits finis.xlsm'

stock_wb = openpyxl.load_workbook(os.path.join(dir_path, stock_path), data_only= True, keep_vba=True)
planning_wb = openpyxl.load_workbook(os.path.join(dir_path, planning_path), keep_vba= True)

planning_ws = planning_wb.worksheets[0]

specialite = stock_wb.worksheets[5]

specialite_data = [specialite['AJ36'].value,]
specialite_s = [5.5]
specialite_TOTAL = [specialite['R39'].value,]
specialite_melee = [(specialite['AJ38'].value * specialite['AA38'].value)]
specialite_m = [specialite['AA38'].value,]

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

print(data[0])