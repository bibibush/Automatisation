import xlwings as xw
import math
import os
from tkinter import *
from tkinter.messagebox import *
import pyautogui


root = Tk()
root.title('Automatisation de planning de fabrication')
label = Label(root, text='Mettez le numero de la semaine de planning')
label_vide = Label(root, text='', width=5, height=10)
entry = Entry(root, width=15)
btn = Button(root, width=10 ,text='Lancer')
label.grid(row=0, column=0, columnspan=3)
label_vide.grid(row=1, column=0)
label_vide.grid(row=1, column=1)
label_vide.grid(row=1, column=2)
entry.grid(row=2, column=0)
label_vide.grid(row=2,column=1)
btn.grid(row=2, column=2)

entry.insert(0, int())

def automatiser():

    dir_path = '//srvlabreche/Dossier semaine commun'
    planning_path = f'Production/Planification de la production/Semaine {entry.get()} - Planning Fabrication.xlsm'
    save_path = f'Production/Planification de la production/Semaine {entry.get()} - Planning Fabrication.xlsm'
    stock_path = 'Stock/Stock Produits finis.xlsm'

    with xw.App(visible=False) as app:
        stock_wb = xw.Book(os.path.join(dir_path, stock_path))
        planning_wb = xw.Book(os.path.join(dir_path, planning_path))

        planning_ws = planning_wb.sheets[0]

        # Halal
        halal = stock_wb.sheets[1]

        halal_data = [halal['AK28'].value, halal['AK51'].value, halal['AK84'].value, halal['AK116'].value] 

        halal_s = [4, 4, 8, 8]
        halal_total = [
            halal['R28'].value,
            halal['R51'].value,
            halal['R85'].value,
            halal['R117'].value,
        ]
        halal_melee = [
            (halal['AK25'].value * halal['AA25'].value),
            (halal['AK48'].value * halal['AA48'].value), 
            (halal['AK82'].value * halal['AA82'].value), 
            (halal['AK114'].value * halal['AA114'].value), 
            ]
        halal_m = [
            halal['AA25'].value,
            halal['AA48'].value,
            halal['AA82'].value,
            halal['AA114'].value,
        ]
        halal_planning_s_en_cours = [
            planning_ws['F7'].value,
            planning_ws['F8'].value,
            planning_ws['F11'].value,
            planning_ws['F12'].value,
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
                elif z == 0:
                    z = 0.5
                    data.append(z)
                else:
                    data.append(z)
        
        order = [7, 8, 11, 12]
        for i in range(len(halal_data)):
            if halal_planning_s_en_cours[i] >= data[i]:
                planning_ws[f'E{order[i]}'].value = 0
            else:
                planning_ws[f'E{order[i]}'].value = data[i] - halal_planning_s_en_cours[i]            

        # H.G, S.A.G
        HG_SAG = stock_wb.sheets[2]
        HG_data = [
            HG_SAG['AJ118'].value,
            HG_SAG['AJ82'].value,
            HG_SAG['AJ10'].value,
            HG_SAG['AJ37'].value,
            HG_SAG['AJ60'].value,
            HG_SAG['AJ18'].value,
        ]
        HG_S = [5, 3, 4, 5.9, 7, 4.5]
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
        HG_planning_s_en_cours = [
            planning_ws['F16'].value,
            planning_ws['F32'].value,
            planning_ws['F33'].value,
            planning_ws['F34'].value,
            planning_ws['F35'].value,
            planning_ws['F37'].value,
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
                elif z == 0:
                    z = 0.5
                    data.append(z)
                else:
                    data.append(z)

        order = [16, 32, 33, 34, 35, 37]

        for i in range(len(HG_planning_s_en_cours)):
            if HG_planning_s_en_cours[i] >= data[i]:
                planning_ws[f'E{order[i]}'].value = 0
            else:
                planning_ws[f'E{order[i]}'].value = data[i] - HG_planning_s_en_cours[i]

        # HM, BN
        BN = stock_wb.sheets[3]
        HM = stock_wb.sheets[4]

        HM_data = [
            HM['AJ46'].value,
            HM['AJ70'].value,
            HM['AJ27'].value,
            HM['AJ3'].value,
            HM['AJ11'].value,
            HM['AJ19'].value,
            BN['AJ68'].value,
        ]
        HM_s = [5.9, 7, 5, 3.5, 5, 7, 7.3]
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
        HM_planning_s_en_cours = [
            planning_ws['F44'].value,
            planning_ws['F45'].value,
            planning_ws['F56'].value,
            planning_ws['F57'].value,
            planning_ws['F58'].value,
            planning_ws['F59'].value,
            planning_ws['F61'].value,
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
                elif z == 0:
                    z = 0.5
                    data.append(z)
                else:
                    data.append(z)
        order = [44, 45, 56, 57, 58, 59, 61]

        for i in range(len(HM_planning_s_en_cours)):
            if HM_planning_s_en_cours[i] >= data[i]:
                planning_ws[f'E{order[i]}'].value = 0
            else:
                planning_ws[f'E{order[i]}'].value = data[i] - HM_planning_s_en_cours[i]

        # SPECIALITE
        specialite = stock_wb.sheets[5]

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
            specialite['AJ176'].value,
            specialite['AJ265'].value,
            specialite['AJ197'].value,
            specialite['AJ232'].value,
            specialite['AJ209'].value,
        ]
        specialite_s = [6, 6, 6, 6, 6, 6, 6, 6, 6, 6, 6, 6, 6, 6, 6, 6, 6, 6, 6]
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
            specialite['R179'].value,
            specialite['R268'].value,
            specialite['R200'].value,
            specialite['R235'].value,
            specialite['R212'].value,
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
            (specialite['AJ178'].value * specialite['AA178'].value),
            (specialite['AJ267'].value * specialite['AA267'].value),
            (specialite['AJ199'].value * specialite['AA199'].value),
            (specialite['AJ234'].value * specialite['AA234'].value),
            (specialite['AJ211'].value * specialite['AA211'].value),
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
            specialite['AA178'].value,
            specialite['AA267'].value,
            specialite['AA199'].value,
            specialite['AA234'].value,
            specialite['AA211'].value,
        ]
        specialite_planning_s_en_cours = []
        for i in range(66, 81 + 1):
            specialite_planning_s_en_cours.append(planning_ws[f'F{i}'].value)
        for i in range(87, 89 + 1):
            specialite_planning_s_en_cours.append(planning_ws[f'F{i}'].value)

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

        for i in range(66, 81 + 1):
            if specialite_planning_s_en_cours[i - 66] >= data[i - 66]:
                planning_ws[f'E{i}'].value = 0
            else:
                planning_ws[f'E{i}'].value = data[i - 66] - specialite_planning_s_en_cours[i - 66]
            
        for i in range(87, 89 + 1):
            if specialite_planning_s_en_cours[i - 71] >= data[i - 71]:
                planning_ws[f'E{i}'].value = 0
            else:
                planning_ws[f'E{i}'].value = data[i - 71] - specialite_planning_s_en_cours[i - 71]
            
        # FILET MIGNON
        data = []
        if HG_SAG['AJ122'].value <= 5:
            data.append(100)
        else:
            data.append(0)

        planning_ws['E3'].value = data[0]

        planning_wb.save(os.path.join(dir_path, save_path))
        stock_wb.close()
        planning_wb.close()
    print('Parfait !')
    pyautogui.alert('Parfait !')


def check_automatiser(event):
    if askyesno(title='Confirmation', message='Vous voulez lancer le programme ?'):
        try:
            automatiser()
        except Exception as e:
            print(e)
            pyautogui.alert('Appelez Jin hyeong !')

btn.bind('<Button-1>', check_automatiser)
root.mainloop()