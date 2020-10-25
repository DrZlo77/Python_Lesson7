
#1) Вручную создать текстовый файл с данными (например, марка авто, модель авто, расход топлива, стоимость)
# файл с данными hw_data_ship
#2) Создать doc шаблон, где будут использованы данные параметры.
# файл-шаблон hw_data_ship hw_ship_report.docx
'''
3) Автоматически сгенерировать отчет о машине в формате doc (как в видео 7.2).
'''
#==================================================================================

import csv
import json
import datetime
from docxtpl import DocxTemplate
from docxtpl import InlineImage
from docx.shared import Cm
from docxtpl import DocxTemplate, InlineImage
from time import time


# Получаем данные из файла hw_data_ship чтобы заполнить им шаблон hw_ship_report.docx
def get_data_file():
    start_time = time()

    list_data = get_list_data()

    context = get_context(list_data)

    template_name = 'hw_template.docx'

    template = DocxTemplate(template_name)

    acc1 = InlineImage(template, 'acc1.jpg', Cm(15))

    context['acc1'] = acc1  # adds the InlineImage object to the context

    acc2 = InlineImage(template, 'acc2.jpg', Cm(15))

    context['acc2'] = acc2  # adds the InlineImage object to the context

    template.render(context)

    template.save('hw_report.docx')

    end_time = time()
    spend_time = end_time - start_time
    print(f'Генерация docx, время- {spend_time}')

def get_context (list_data):
    context_dict = {}
    #context_dict.update({'2':1})
    counter_list = 1
    arg_1_name = "type"
    arg_2_name = "team"
    arg_3_name = "long"
    for elem_dict in list_data:
        #print(type(elem_dict))
        counter_tuple = 1
        for tuple_el in elem_dict.items():
            if counter_tuple == 1:
                context_dict.update({str(arg_1_name)+str(counter_list):tuple_el[1]})
            elif counter_tuple == 2:
                context_dict.update({str(arg_2_name) + str(counter_list): tuple_el[1]})
            else:
                context_dict.update({str(arg_3_name) + str(counter_list): tuple_el[1]})
            if counter_tuple == 3:
                counter_tuple = 1

            counter_tuple+=1

        counter_list+=1
    return context_dict

def get_list_data ():
    list_data = []

    with open('hw_data_ship.txt') as file:
        reader = csv.DictReader(file)
        for row in reader:
            list_data.append(dict(row))
    return  list_data


#==================================================================================

'''
4) Создать csv файл с данными о машине.
'''
def get_csv_file ():
    start_time = time()
    list_data = get_list_data()
    fieldnames = ['type','team','long']
    with open('csv_data.csv','w') as file:
        writer = csv.DictWriter(file,delimiter = ';',fieldnames=fieldnames)
        writer.writeheader()
        for index in range(len(list_data)):
            writer.writerow(list_data[index])
    end_time = time()
    spend_time = end_time - start_time
    print(f'Генерация csv, время- {spend_time}')
#==================================================================================

'''
5) Создать json файл с данными о машине. 
'''
def get_json_file():
    start_time = time()
    list_data = get_list_data()
    context = get_context(list_data)
    with open('json_data.json', 'w') as file:
        json.dump(context, file)
    end_time = time()
    spend_time = end_time-start_time
    print(f'Генерация json, время- {spend_time}')


get_data_file()
get_csv_file()
get_json_file()