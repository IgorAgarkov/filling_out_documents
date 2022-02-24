# -*- coding: utf8 -*-

from docxtpl import DocxTemplate
from docx2pdf import convert
from datetime import date
import re

num1, num2 = 117, 118                     # первый и последний номера диапазона без года "-21"
template = 'example KOR-9_template.docx'  # шаблон документа docx


mark_search = re.search('\w{2,3}-\d{1,}[^_ ]?\d*', template)  # '\w{2,3}' - любая буква (цифра), '-' - дейфис, '\d{1,}' - любая цифра от 1 и более, '[^_ ]?' любой знак кроме _ и пробела от 0 до 1, '\d*' - любая цифра от 0 и более (жадный)
mark = mark_search.group()                # маркировка, или название продукции
# mark = 'ЗВК-200'                        
today_date = date.today().strftime("%d.%m.%Y")    # дата
# today_date = '26.03.2021'                       
surn = 'Агарков И.Н.'                     # ФИО

for i in range(num1, num2 + 1):  
    num = str(i) + '-' + today_date[-2:]
    doc = DocxTemplate(template)       
    contex = {'marking' : mark, 'number' : num, 'surname' : surn, 'date' : today_date}
    doc.render(contex)
    file_name = 'Паспорт-' + mark + '_' + num + '_' + today_date
    doc.save(file_name + '.docx')
    convert(file_name + '.docx')
    