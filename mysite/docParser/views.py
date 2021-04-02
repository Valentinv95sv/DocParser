import os
from django.shortcuts import HttpResponse
from django.shortcuts import render
from . import Docx_insert_data_class

inserted_fields = {
    'name/number_of_tender' : 'Supply tender AOA',
    'day.month_of_signed' : '22.04',
    'name_of_tender' : 'new tender of 2020',
    'time_of_term' : '11:30',
    'day_of_term': '23',
    'month_of_term':'апреля'
}

insert_table = [
    {
    'position1': 'Boss',
    'full_name1':'Valentin V'
    },
    {
    'position1': 'intern',
    'full_name1':'Andrey'
    }
]
insert_table2 = [
    {
        'TRU_name1':'only',
        'Small_describe1':'bad',
        'netto1':'3000'
    },
    {
        'TRU_name1':'society',
        'Small_describe1':'Good',
        'netto1':'4500'
    }
]

insert_table3 = [
    {
        'suppliers_name':'Valentin',
        'suppliers_location':'Almaty',
        'time_of_bid':'12.12.1998',
        'lots_in_bid':'android, ios'
    },
    {
        'suppliers_name':'Michail',
        'suppliers_location':'Piter',
        'time_of_bid':'12.12.1992',
        'lots_in_bid':'android, windows'
    }
]

insert_table4 = [
    {
        'describe_of_product':'good phone',
        'Final_price':'199.999'
    },
    {
        'describe_of_product':'bad phone',
        'Final_price':'150.999'
    },
    {
        'describe_of_product':'blue phone',
        'Final_price':'19.999'
    },
    {
        'describe_of_product':'red phone',
        'Final_price':'200.999'
    },
    {
        'describe_of_product':'dreen phone',
        'Final_price':'9.999'
    }
]
insert_table5 = [
    {
        'p/p_number':'12',
        'supplier':'Andrey',
        'rreason_for_rejection':'not enough docs'
    },
{
        'p/p_number':'2',
        'supplier':'Sven',
        'rreason_for_rejection':'Bad name'
    }
]

insert_table6 = [
    {
        'TRU_name2':'test',
        'describ2e':'only letters',
        'netto2':'676'
    },
    {
        'TRU_name2':'wow!',
        'describ2e':'many letters',
        'netto2':'16'
    }
]

insert_table7 = [
    {
        'TRU_name3':'qwerty',
        'describe3':'lots of pain',
        'netto3':'123'
    },
    {
        'TRU_name3':'ios',
        'describe3':'lots of pain',
        'netto3':'23'
    }
]

insert_table8 = [
    {
        'position2':'dj',
        'full_name2':'david guetta'
    },
    {
        'position2':'dj',
        'full_name2':'david guetta'
    },
    {
        'position2':'dj',
        'full_name2':'david guetta'
    },
    {
        'position2':'dj',
        'full_name2':'david guetta'
    },
    {
        'position2':'dj',
        'full_name2':'david guetta'
    },
    {
        'position2':'dj',
        'full_name2':'david guetta'
    },
]


template = 'docParser/Shablon/ПИ ТЕНДЕР (ШАБЛОН)#.docx'
path = 'blanks/'

a = Docx_insert_data_class.docx_insert_data_class()

income = []
income.append(insert_table)
income.append(insert_table2)
income.append(insert_table3)
income.append(insert_table4)
income.append(insert_table5)
income.append(insert_table6)
income.append(insert_table7)
income.append(insert_table8)

pict_path = {'david guetta': r'C:/MyWorkFlow/123.png', 'we': r'C:/MyWorkFlow/234.png'}

def index(request):
    return render(request, 'test.html')

def download(request):
    file_path = a.merge_file_fields_tables(template, inserted_fields, income, path + a.get_file_name())

    a.add_picture(file_path, pict_path)

    if os.path.exists(file_path):
        with open(file_path, 'rb') as fh:
            response = HttpResponse(fh.read(), content_type="application/vnd.ms-excel")
            response['Content-Disposition'] = 'inline; filename=' + os.path.basename(file_path)
            return response

