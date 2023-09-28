import openpyxl
import math

planning_path = 'planning fab/Semaine 40 - Planning Fabrication.xlsm'
# save_path = 'planning fab/Semaine 33 2 - Planning Fabrication.xlsm'

stock_wb = openpyxl.load_workbook('Stock Produits finis.xlsm', data_only= True, keep_vba=True)
planning_wb = openpyxl.load_workbook(planning_path, keep_vba= True)

planning_ws = planning_wb.worksheets[0]

# Halal
halal = stock_wb.worksheets[1]

halal_data = [halal['AK28'].value, halal['AK51'].value, halal['AK128'].value,
               halal['AK135'].value, halal['AK84'].value, halal['AK116'].value] 

halal_s = [3, 3, 3, 3, 6, 6]
halal_total = [
    halal['R28'].value,
    halal['R51'].value,
    halal['R129'].value,
    halal['R136'].value,
    halal['R85'].value,
    halal['R117'].value,
]
halal_melee = [
    (halal['AK25'].value * halal['AA25'].value),
    (halal['AK48'].value * halal['AA48'].value), 
    (halal['AK126'].value * halal['AA126'].value), 
    (halal['AK133'].value * halal['AA133'].value), 
    (halal['AK82'].value * halal['AA82'].value), 
    (halal['AK114'].value * halal['AA114'].value), 
    ]
halal_m = [
    halal['AA25'].value,
    halal['AA48'].value,
    halal['AA126'].value,
    halal['AA133'].value,
    halal['AA82'].value,
    halal['AA114'].value,
]

data = []
for i in range(len(halal_data)):
    if halal_data[i] >= halal_s[i]:
        x = 0
        data.append(x)
    else:
        y = halal_s[i] * halal_melee[i] + halal_melee[i] - halal_total[i]
        x = y / halal_m[i]
        x = round(x, 1)
        z = math.ceil(x)
        if (z * 10) - (x * 10) >= 5:
            data.append(z - 0.5)
        else:
            data.append(z)

for i in range(7, 12 + 1):
    planning_ws[f'E{i}'].value = data[i - 7]


# H.G, S.A.G
HG_SAG = stock_wb.worksheets[2]
HG_data = [
    HG_SAG['AJ118'].value,
    HG_SAG['AJ82'].value,
    HG_SAG['AJ10'].value,
    HG_SAG['AJ37'].value,
    HG_SAG['AJ60'].value,
    HG_SAG['AJ18'].value,
]
HG_S = [5, 3, 3.5, 5.9, 6, 4.5]
HG_TOTAL = [
    HG_SAG['R119'].value,
    HG_SAG['R83'].value,
    HG_SAG['R11'].value,
    HG_SAG['R38'].value,
    HG_SAG['R61'].value,
    HG_SAG['R19'].value,
]
HG_melee = [
    (HG_SAG['AJ116'].value * HG_SAG['AA118'].value),
    (HG_SAG['AJ80'].value * HG_SAG['AA82'].value),
    (HG_SAG['AJ8'].value * HG_SAG['AA8'].value),
    (HG_SAG['AJ35'].value * HG_SAG['AA35'].value),
    (HG_SAG['AJ58'].value * HG_SAG['AA60'].value),
    (HG_SAG['AJ16'].value * HG_SAG['AA16'].value),
]
HG_m = [
    HG_SAG['AA118'].value,
    HG_SAG['AA82'].value,
    HG_SAG['AA8'].value,
    HG_SAG['AA35'].value,
    HG_SAG['AA60'].value,
    HG_SAG['AA16'].value,
]

data = []
for i in range(len(HG_data)):
    if HG_data[i] >= HG_S[i]:
        x = 0
        data.append(x)
    else:
        y = HG_S[i] * HG_melee[i] + HG_melee[i] - HG_TOTAL[i]
        x = y / HG_m[i]
        x = round(x, 1)
        z = math.ceil(x)
        if (z * 10) - (x * 10) >= 5:
            data.append(z - 0.5)
        else:
            data.append(z)

planning_ws['E16'] = data[0]       
planning_ws['E32'] = data[1]
planning_ws['E33'] = data[2]
planning_ws['E34'] = data[3]
planning_ws['E35'] = data[4]
planning_ws['E37'] = data[5]

# HM, BN
BN = stock_wb.worksheets[3]
HM = stock_wb.worksheets[4]

HM_data = [
    HM['AJ46'].value,
    HM['AJ70'].value,
    HM['AJ27'].value,
    HM['AJ3'].value,
    HM['AJ11'].value,
    HM['AJ19'].value,
    BN['AJ68'].value,
]
HM_s = [5.24, 6, 5, 3.5, 5, 7, 7.3]
HM_TOTAL = [
    HM['R51'].value,
    HM['R75'].value,
    HM['R33'].value,
    HM['R8'].value,
    HM['R16'].value,
    HM['R24'].value,
    BN['R69'].value,
]
HM_melee = [
    (HM['AJ48'].value * HM['AA48'].value),
    (HM['AJ72'].value * HM['AA72'].value),
    (HM['AJ29'].value * HM['AA29'].value),
    (HM['AJ5'].value * HM['AA5'].value),
    (HM['AJ13'].value * HM['AA13'].value),
    (HM['AJ21'].value * HM['AA21'].value),
    (BN['AJ66'].value * BN['AA66'].value),
]
HM_m = [
    HM['AA48'].value,
    HM['AA72'].value,
    HM['AA29'].value,
    HM['AA5'].value,
    HM['AA13'].value,
    HM['AA21'].value,
    BN['AA66'].value,
]

data = []
for i in range(len(HM_data)):
    if HM_data[i] >= HM_s[i]:
        x = 0
        data.append(x)
    else:
        y = HM_s[i] * HM_melee[i] + HM_melee[i] - HM_TOTAL[i]
        x = y / HM_m[i]
        x = round(x, 1)
        z = math.ceil(x)
        if (z * 10) - (x * 10) >= 5:
            data.append(z - 0.5)
        else:
            data.append(z)

planning_ws['E44'] = data[0]
planning_ws['E45'] = data[1]
planning_ws['E56'] = data[2]
planning_ws['E57'] = data[3]
planning_ws['E58'] = data[4]
planning_ws['E59'] = data[5]
planning_ws['E61'] = data[6]


# SPECIALITE
specialite = stock_wb.worksheets[5]

specialite_data = [
    specialite['AJ13'].value,
    specialite['AJ36'].value,
    specialite['AJ26'].value,
    specialite['AJ51'].value,
    specialite['AJ64'].value,
    specialite['AJ80'].value,
    specialite['AJ105'].value,
    specialite['AJ141'].value,
    specialite['AJ115'].value,
    specialite['AJ132'].value,
    specialite['AJ149'].value,
    specialite['AJ90'].value,
    specialite['AJ124'].value,
    specialite['AJ158'].value,
    specialite['AJ174'].value,
    specialite['AJ263'].value,
    specialite['AJ195'].value,
    specialite['AJ230'].value,
    specialite['AJ207'].value,
]
specialite_s = [5.5, 5.5, 5.5, 5.5, 5.5, 5.5, 5.5, 5.5, 5.5, 5.5, 5.5, 5.5, 5.5, 5.5, 5.5, 5.5, 5.5, 5.5, 5.5]
specialite_TOTAL = [
    specialite['R16'].value,
    specialite['R39'].value,
    specialite['R29'].value,
    specialite['R54'].value,
    specialite['R68'].value,
    specialite['R83'].value,
    specialite['R108'].value,
    specialite['R144'].value,
    specialite['R118'].value,
    specialite['R135'].value,
    specialite['R152'].value,
    specialite['R93'].value,
    specialite['R127'].value,
    specialite['R161'].value,
    specialite['R177'].value,
    specialite['R266'].value,
    specialite['R198'].value,
    specialite['R233'].value,
    specialite['R210'].value,
]
specialite_melee = [
    (specialite['AJ15'].value * specialite['AA15'].value),
    (specialite['AJ38'].value * specialite['AA38'].value),
    (specialite['AJ28'].value * specialite['AA28'].value),
    (specialite['AJ53'].value * specialite['AA53'].value),
    (specialite['AJ66'].value * specialite['AA66'].value),
    (specialite['AJ82'].value * specialite['AA82'].value),
    (specialite['AJ107'].value * specialite['AA107'].value),
    (specialite['AJ143'].value * specialite['AA143'].value),
    (specialite['AJ117'].value * specialite['AA117'].value),
    (specialite['AJ134'].value * specialite['AA134'].value),
    (specialite['AJ151'].value * specialite['AA151'].value),
    (specialite['AJ92'].value * specialite['AA92'].value),
    (specialite['AJ126'].value * specialite['AA126'].value),
    (specialite['AJ160'].value * specialite['AA160'].value),
    (specialite['AJ176'].value * specialite['AA176'].value),
    (specialite['AJ265'].value * specialite['AA265'].value),
    (specialite['AJ197'].value * specialite['AA197'].value),
    (specialite['AJ232'].value * specialite['AA232'].value),
    (specialite['AJ209'].value * specialite['AA209'].value),
]
specialite_m = [
    specialite['AA15'].value,
    specialite['AA38'].value,
    specialite['AA28'].value,
    specialite['AA53'].value,
    specialite['AA66'].value,
    specialite['AA82'].value,
    specialite['AA107'].value,
    specialite['AA143'].value,
    specialite['AA117'].value,
    specialite['AA134'].value,
    specialite['AA151'].value,
    specialite['AA92'].value,
    specialite['AA126'].value,
    specialite['AA160'].value,
    specialite['AA176'].value,
    specialite['AA265'].value,
    specialite['AA197'].value,
    specialite['AA232'].value,
    specialite['AA209'].value,
]
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
        else:
            data.append(z)

for i in range(66, 81 + 1):
    planning_ws[f'E{i}'] = data[i - 66]

for i in range(87, 89 + 1):
    planning_ws[f'E{i}'] = data[i - 71]

# FILET MIGNON
data = []
if HG_SAG['AJ122'].value >= 5:
    data.append(100)
else:
    data.append(0)

planning_ws['E3'] = data[0]

planning_wb.save(planning_path)
    