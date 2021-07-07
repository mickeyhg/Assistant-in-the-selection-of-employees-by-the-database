# -*- coding: utf8 -*-
import pandas as pd
import time
import xlrd
import openpyxl
import tkinter
from tkinter import *
from tkinter import messagebox

input_type = 'gui'
instruction_text = 'Для использования прогрыммы вам нужна база данных кандидатов, в виде таблицы с названием "1.xlsx", в которой столбцы означают характеристики кандидатов, а строки - отдельных кандидатов. Характеристики в таблице должны совпадать с характеристиками из пункта меню "Формат ввода". Эту таблицу нужно поместить в папку с программой. \n \nКритерии для отбора кандидатов вы можете ввести через интерфейс,или создать таблицу 2.xlsx, в которой столбцы означают характеристики, а строки - требуемые параметры. Вы можете указывать только те характеристики, которые вас интересуют. Значения характеристик должны соответствовать значениям из пункта меню "Формат ввода". Для смены типа ввода(ввод через интерфейс или через таблицу) однократно нажмите на кнопку "Смена типа ввода". По умолчанию стоит ввод через интерфейс. \n \nДля запуска программы нажмите кнопку "Старт". После запуска в папке с программой появится файл "Кандидаты.txt".'
format_text = 'Пол - муж или жен \nВозраст - ввести число \nВысшее образование - есть или нет \nСреднее образование - есть или нет \nОпыт работы - ввести число \nНаличие судимостей - есть или нет \nЗаболевания - есть или нет \nНаличие детей - да или нет \nВладение английским языком - да или нет \nСпециальность - ввести название специальности \nСостоит в браке - да или нет \nВодительское удостоверение - есть или нет \n\nТакже в таблицу с базой данных(1.xlsx) нужно добавить столбец "Фамилия И.О."'
def main_script(criteria):
    # Основной алгоритм
    time_start = time.time()
    # Подготовка данных
    database = pd.read_excel('1.xlsx')
    if input_type == 'table':
        criteria = pd.read_excel('2.xlsx')
        conf_data = database[criteria.columns]
    conf_data = database[criteria.keys()]
    # conf_data[0]
    candidates = []
    # Анализ кандидатов
    for case in range(0, len(conf_data)):
        check = 0
        mismatched_criteria = []
        matched_criteria = []
        if input_type == 'gui':
            # При вводе с интерфейса
            if 'Возраст' in criteria:
                if str(conf_data['Возраст'][case]) > criteria['Возраст']:
                    check += 1
                    mismatched_criteria.append('Возраст')
                else:
                    matched_criteria.append('Возраст')
            if 'Высшее образование' in criteria:
                if conf_data['Высшее образование'][case] == criteria['Высшее образование']:
                    matched_criteria.append('Высшее образование')
                else:
                    check += 1
                    mismatched_criteria.append('Высшее образование')
            if 'Среднее образование' in criteria:
                if conf_data['Среднее образование'][case] == criteria['Среднее образование']:
                    matched_criteria.append('Среднее образование')
                else:
                    check += 1
                    mismatched_criteria.append('Среднее образование')
            if 'Опыт работы' in criteria:
                if conf_data['Опыт работы'][case] < criteria['Опыт работы']:
                    check += 1
                    mismatched_criteria.append('Опыт работы')
                else:
                    matched_criteria.append('Опыт работы')
                    
            if 'Наличие судимостей' in criteria:
                if conf_data['Наличие судимостей'][case] == criteria['Наличие судимостей']:
                    matched_criteria.append('Наличие судимостей')
                else:
                    check += 1
                    mismatched_criteria.append('Наличие судимостей')
                    
            if 'Психические заболевания' in criteria:
                if conf_data['Психические заболевания'][case] == criteria['Психические заболевания']:
                    matched_criteria.append('Психические заболевания')
                else:
                    check += 1
                    mismatched_criteria.append('Психические заболевания')
            if 'Водительское удостоверение' in criteria:
                if conf_data['Водительское удостоверение'][case] == criteria['Водительское удостоверение']:
                    matched_criteria.append('Водительское удостоверение')
                else:
                    check += 1
                    mismatched_criteria.append('Водительское удостоверение')
            if 'Наличие детей' in criteria:
                if conf_data['Наличие детей'][case] == criteria['Наличие детей']:
                    matched_criteria.append('Наличие детей')
                else:
                    check += 1
                    mismatched_criteria.append('Наличие детей')
            if 'Владение английским' in criteria:
                if conf_data['Владение английски'][case] == criteria['Владение английски']:
                    matched_criteria.append('Владение английски')
                else:
                    check += 1
                    mismatched_criteria.append('Владение английски')
            if 'Пол' in criteria:
                if conf_data['Пол'][case] == criteria['Пол']:
                    matched_criteria.append('Пол')
                else:
                    check += 1
                    mismatched_criteria.append('Пол')
            if 'Состоит в браке' in criteria:
                if conf_data['Состоит в браке'][case] == criteria['Состоит в браке']:
                    matched_criteria.append('Состоит в браке')
                else:
                    check += 1
                    mismatched_criteria.append('Состоит в браке')
            if 'Специальность' in criteria:
                if conf_data['Специальность'][case] == criteria['Специальность']:
                    matched_criteria.append('Специальность')
                else:
                    check += 1
                    mismatched_criteria.append('Специальность')
        elif input_type == 'table':
            # При вводе с таблицы
            if 'Возраст' in criteria:
                if conf_data['Возраст'][case] > criteria['Возраст'][0]:
                    check += 1
                    mismatched_criteria.append('Возраст')
                else:
                    matched_criteria.append('Возраст')
            if 'Высшее образование' in criteria:
                if conf_data['Высшее образование'][case] == criteria['Высшее образование'][0]:
                    matched_criteria.append('Высшее образование')
                else:
                    check += 1
                    mismatched_criteria.append('Высшее образование')
            if 'Среднее образование' in criteria:
                if conf_data['Среднее образование'][case] == criteria['Среднее образование'][0]:
                    matched_criteria.append('Среднее образование')
                else:
                    check += 1
                    mismatched_criteria.append('Среднее образование')
            if 'Опыт работы' in criteria:
                if conf_data['Опыт работы'][case] < criteria['Опыт работы'][0]:
                    check += 1
                    mismatched_criteria.append('Опыт работы')
                else:
                    matched_criteria.append('Опыт работы')                    
            if 'Наличие судимостей' in criteria:
                if conf_data['Наличие судимостей'][case] == criteria['Наличие судимостей'][0]:
                    matched_criteria.append('Наличие судимостей')
                else:
                    check += 1
                    mismatched_criteria.append('Наличие судимостей')                    
            if 'Психические заболевания' in criteria:
                if conf_data['Психические заболевания'][case] == criteria['Психические заболевания'][0]:
                    matched_criteria.append('Психические заболевания')
                else:
                    check += 1
                    mismatched_criteria.append('Психические заболевания')
            if 'Водительское удостоверение' in criteria:
                if conf_data['Водительское удостоверение'][case] == criteria['Водительское удостоверение'][0]:
                    matched_criteria.append('Водительское удостоверение')
                else:
                    check += 1
                    mismatched_criteria.append('Водительское удостоверение')
            if 'Наличие детей' in criteria:
                if conf_data['Наличие детей'][case] == criteria['Наличие детей'][0]:
                    matched_criteria.append('Наличие детей')
                else:
                    check += 1
                    mismatched_criteria.append('Наличие детей')
            if 'Владение английским' in criteria:
                if conf_data['Владение английски'][case] == criteria['Владение английски'][0]:
                    matched_criteria.append('Владение английски')
                else:
                    check += 1
                    mismatched_criteria.append('Владение английски')
            if 'Пол' in criteria:
                if conf_data['Пол'][case] == criteria['Пол'][0]:
                    matched_criteria.append('Пол')
                else:
                    check += 1
                    mismatched_criteria.append('Пол')
            if 'Состоит в браке' in criteria:
                if conf_data['Состоит в браке'][case] == criteria['Состоит в браке'][0]:
                    matched_criteria.append('Состоит в браке')
                else:
                    check += 1
                    mismatched_criteria.append('Состоит в браке')
            if 'Специальность' in criteria:
                if conf_data['Специальность'][case] == criteria['Специальность'][0]:
                    matched_criteria.append('Специальность')
                else:
                    check += 1
                    mismatched_criteria.append('Специальность')  
        result = (check*100)/len(criteria.keys())
        candidates.append([case, result, matched_criteria, mismatched_criteria])
    candidates = sorted(candidates, key=lambda student: student[1])
    # Создание выходного файла
    f=open('Кандидаты.txt','w', encoding = 'utf-8')
    for candidate in candidates:     
        f.write(str([database['Фамилия И. О.'][candidate[0]]])[2:-2])
        f.write('\n')
        string = 'Процент соответствующих параметров: ' + str(100-candidate[1]) + '%;'
        f.write(string)
        f.write('\n')
        f.write('Подходящие параметры: ')
        for i in range(0, len(candidate[2])):
            string = candidate[2][i] + ': ' + str(database[candidate[2][i]][candidate[0]])
            if i < len(candidate[2]):
                string = string + '; '
            f.write(string)
        f.write('\n')
        
        if len(candidate[3]) > 0:
            f.write('Неподходящие параметры: ')
            for i in range(0, len(candidate[3])):
                string = candidate[3][i] + ': ' + str(database[candidate[3][i]][candidate[0]])
                if i < len(candidate[3]):
                    string = string + '; '
                f.write(string)
        f.write('\n')    
        f.write('\n')
    f.close()
    messagebox.showinfo(title='Уведомление', message='Файл успешно создан')
    time_end = time.time()
    print(time_end - time_start)

def input_changing():
    # Изменение метода ввода
    global input_type
    if input_type == 'gui':
        input_type = 'table'
        messagebox.showinfo(title='Уведомление', message='Выбран метод ввода через таблицу')
    elif input_type == 'table':
        input_type = 'gui'
        messagebox.showinfo(title='Уведомление', message='Выбран метод ввода через интерфейс')
arr1 = ['есть', 'нет']
arr2 = ['да', 'нет']
arr3 = ['муж', 'жен']
all_is_fine = False
def launching():
    # Чтение критериев с интерфейса
    global all_is_fine
    criteria = {}
    if age.get() != '':
        if age.get().isdigit() == True: 
            criteria['Возраст'] = age.get()
            all_is_fine = True
        else:
            if all_is_fine != False:
                messagebox.showinfo(title='Уведомление', message='Поле "Возраст" поддерживает тоько ввод чисел. Для справки войдите в "Формат ввода"')
            all_is_fine = False

    if high_ed.get() != '':
        if high_ed.get() in arr1 and all_is_fine == True: 
            criteria['Высшее образование'] = high_ed.get()
            all_is_fine = True
        else:
            if all_is_fine != False:
                messagebox.showinfo(title='Уведомление', message='Поле "Высшее образование" поддерживает варианты ввода: есть , нет. Для справки войдите в "Формат ввода"')
            all_is_fine = False

    if second_ed.get() != '':
        if second_ed.get() in arr1 and all_is_fine == True: 
            criteria['Среднее образование'] = second_ed.get()
            all_is_fine = True
        else:
            if all_is_fine != False:
                messagebox.showinfo(title='Уведомление', message='Поле "Среднее образование" поддерживает только варианты ввода: есть , нет. Для справки войдите в "Формат ввода"')
            all_is_fine = False
        
    if expirience.get() != '':
        if high_ed.get().isdigit() == True and all_is_fine == True: 
            criteria['Опыт работы'] = expirience.get()
            all_is_fine = True
        else:
            messagebox.showinfo(title='Уведомление', message='Поле "Опыт работы" поддерживает тоько ввод чисел. Для справки войдите в "Формат ввода"')
            all_is_fine = False

    if sex.get() != '':
        if sex.get() in arr3 and all_is_fine == True: 
            criteria['Пол'] = sex.get()
            all_is_fine = True
        else:
            if all_is_fine != False:
                messagebox.showinfo(title='Уведомление', message='Поле "Пол" поддерживает варианты ввода: муж , жен. Для справки войдите в "Формат ввода"')
            all_is_fine = False

    if conviction.get() != '':
        if conviction.get() in arr1 and all_is_fine == True: 
            criteria['Наличие судимостей'] = conviction.get()
            all_is_fine = True
        else:
            if all_is_fine != False:
                messagebox.showinfo(title='Уведомление', message='Поле "Наличие судимостей" поддерживает варианты ввода: есть , нет. Для справки войдите в "Формат ввода"')
            all_is_fine = False

    if illness.get() != '':
        if illness.get() in arr1 and all_is_fine == True: 
            criteria['Психические заболевания'] = illness.get()
            all_is_fine = True
        else:
            if all_is_fine != False:
                messagebox.showinfo(title='Уведомление', message='Поле "Психические заболевания" поддерживает варианты ввода: есть , нет. Для справки войдите в "Формат ввода"')
            all_is_fine = False

    if religiosity.get() != '':
        if religiosity.get() in arr2 and all_is_fine == True: 
            criteria['Наличие детей'] = religiosity.get()
            all_is_fine = True
        else:
            if all_is_fine != False:
                messagebox.showinfo(title='Уведомление', message='Поле "Наличие детей" поддерживает варианты ввода: да , нет. Для справки войдите в "Формат ввода"')
            all_is_fine = False

    if spec.get() != '':
        criteria['Специальность'] = spec.get()
        all_is_fine = True     

    if english.get() != '':
        if english.get() in arr2 and all_is_fine == True: 
            criteria['Владение английским языком'] = english.get()
            all_is_fine = True
        else:
            if all_is_fine != False:
                messagebox.showinfo(title='Уведомление', message='Поле "Владение английским языком" поддерживает варианты ввода: да , нет. Для справки войдите в "Формат ввода"')
            all_is_fine = False
  
    if mar_status.get() != '':
        if mar_status.get() in arr2 and all_is_fine == True: 
            criteria['Семейное положение'] = mar_status.get()
            all_is_fine = True
        else:
            if all_is_fine != False:
                messagebox.showinfo(title='Уведомление', message='Поле "Семейное положение" поддерживает варианты ввода: да , нет. Для справки войдите в "Формат ввода"')
            all_is_fine = False

    if license.get() != '':
        if license.get() in arr1 and all_is_fine == True: 
            criteria['Водительское удостоверение'] = license.get()
            all_is_fine = True
        else:
            if all_is_fine != False:
                messagebox.showinfo(title='Уведомление', message='Поле Водительское удостоверение"" поддерживает варианты ввода: есть , нет. Для справки войдите в "Формат ввода"')
            all_is_fine = False
    print(criteria)
    if all_is_fine == True or input_type == 'table':
        main_script(criteria)
        criteria = []
    all_is_fine = False
    
def instruction():
    # Вывод инструкции
    messagebox.showinfo(title='Инструкция', message=instruction_text)

def format_info():
    # Вывод формата
    messagebox.showinfo(title='Формат ввода', message=format_text)

# Интерфейс
program_name = 'Автоподбор кандидатов'
main_color= '#f7f7f7'
field_color= '#d6efff'
button_color = '#e1eaf0'
root = Tk()
root['bg'] = main_color
# root.iconbitmap('icon.ico')
root.title(program_name)
root.geometry('1000x400')
root.resizable(width = False, height = False)

canvas = Canvas(root, height= 1720, width=1400, bg = main_color)
canvas.pack()

frame1 = Frame(root, bg=main_color)
frame1.place(relwidth = 0.25, relheight = 0.15)
title = Label(frame1, text = 'Возраст', bg = main_color, font = 15)
# title.place(rely = 0.2, relx = 0.25)
title.pack()
age = Entry(frame1, bg = field_color)
age.pack()

frame2 = Frame(root, bg=main_color)
frame2.place(relwidth = 0.25, relheight = 0.15,relx = 0.25)
title = Label(frame2, text = 'Высшее образование', bg = main_color, font = 15)
title.pack()
high_ed = Entry(frame2, bg = field_color)
high_ed.pack()

frame3 = Frame(root, bg=main_color)
frame3.place(relwidth = 0.25, relheight = 0.15,relx = 0.50)
title = Label(frame3, text = 'Среднее образование', bg = main_color, font = 15)
title.pack()
second_ed = Entry(frame3, bg = field_color)
second_ed.pack()

frame4 = Frame(root, bg=main_color)
frame4.place(relwidth = 0.25, relheight = 0.15,relx = 0.75)
title = Label(frame4, text = 'Стаж работы(в годах)', bg = main_color, font = 15)
title.pack()
expirience = Entry(frame4, bg = field_color)
expirience.pack()

frame5 = Frame(root, bg=main_color)
frame5.place(relwidth = 0.25, relheight = 0.15,rely = 0.15)
title = Label(frame5, text = 'Пол', bg = main_color, font = 15)
# title.place(rely = 0.2, relx = 0.25)
title.pack()
sex = Entry(frame5, bg = field_color)
sex.pack()

frame6 = Frame(root, bg=main_color)
frame6.place(relwidth = 0.25, relheight = 0.15,rely = 0.15, relx = 0.25)
title = Label(frame6, text = 'Наличие судимостей', bg = main_color, font = 15)
# title.place(rely = 0.2, relx = 0.25)
title.pack()
conviction = Entry(frame6, bg = field_color)
conviction.pack()

frame7 = Frame(root, bg=main_color)
frame7.place(relwidth = 0.25, relheight = 0.15,rely = 0.15, relx = 0.5)
# customFont = tkFont.Font(family="Helvetica", size=12)
title = Label(frame7, text = 'Психические расстройства', bg = main_color, font=15)
# title.place(rely = 0.2, relx = 0.25)

title.pack()
illness = Entry(frame7, bg = field_color)
illness.pack()

frame8 = Frame(root, bg=main_color)
frame8.place(relwidth = 0.25, relheight = 0.15,rely = 0.15, relx = 0.75)
title = Label(frame8, text = 'Наличие детей', bg = main_color, font = 15)

title.pack()
religiosity = Entry(frame8, bg = field_color)
religiosity.pack()

frame9 = Frame(root, bg=main_color)
frame9.place(relwidth = 0.25, relheight = 0.15,rely = 0.30)
title = Label(frame9, text = 'Специальность', bg = main_color, font = 15)
title.pack()
spec = Entry(frame9, bg = field_color)
spec.pack()

frame10 = Frame(root, bg=main_color)
frame10.place(relwidth = 0.25, relheight = 0.15,rely = 0.30, relx = 0.25)
title = Label(frame10, text = 'Владение английским', bg = main_color, font = 15)
title.pack()
english = Entry(frame10, bg = field_color)
english.pack()

frame11 = Frame(root, bg=main_color)
frame11.place(relwidth = 0.25, relheight = 0.15,rely = 0.30, relx = 0.5)
title = Label(frame11, text = 'Состоит в браке', bg = main_color, font = 15)
title.pack()
mar_status = Entry(frame11, bg = field_color)
mar_status.pack()

frame12 = Frame(root, bg=main_color)
frame12.place(relwidth = 0.25, relheight = 0.15,rely = 0.30, relx = 0.75)
title = Label(frame12, text = 'Водительское удостоверение', bg = main_color, font = 15)
title.pack()
license = Entry(frame12, bg = field_color)
license.pack()

frame13 = Frame(root, bg=main_color)
frame13.place(relwidth = 0.3, relheight = 0.15, rely = 0.75, relx = 0.6)
btn = Button(frame13, text='Старт', bg = button_color, command = launching)
btn.place(relwidth = 1, relheight = 0.75)

frame13 = Frame(root, bg=main_color)
frame13.place(relwidth = 0.3, relheight = 0.15, rely = 0.75, relx = 0.1)
btn2 = Button(frame13, text='Смена типа ввода', bg = button_color, command = input_changing)
btn2.place(relwidth = 1, relheight = 0.75)

frame14 = Frame(root, bg=main_color)
frame14.place(relwidth = 0.3, relheight = 0.15, rely = 0.55, relx = 0.1)
btn3 = Button(frame14, text='Формат ввода', bg = button_color, command = format_info)
btn3.place(relwidth = 1, relheight = 0.75)

frame15 = Frame(root, bg=main_color)
frame15.place(relwidth = 0.3, relheight = 0.15, rely = 0.55, relx = 0.6)
btn4 = Button(frame15, text='Инструкция', bg = button_color, command = instruction)
btn4.place(relwidth = 1, relheight = 0.75)

root.mainloop()


# © Л. В. Мануйлов, 2021