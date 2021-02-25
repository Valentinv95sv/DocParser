import os
from django.shortcuts import HttpResponse
import datetime
from mailmerge import MailMerge
from django.shortcuts import render
from . import Docx_insert_data_class

inserted_data = {
    'toWhom' : '123',
    'fromWho' : '234',
    'Tender_Number' : 'nony',
    'max_work' : 'qwerrty'
}

a = Docx_insert_data_class.docx_insert_data_class('docParser/Doc_form.docx', inserted_data)

def index(request):
    return render(request, 'test.html')

def get_new_file():
    path = 'blanks/'
    filename = 'Blank_' + str(datetime.datetime.now().strftime("%H%M%S")) + '.docx'
    dict = inserted_data
    template = 'docParser/Doc_form.docx'
    document = MailMerge(template)
    document.merge_pages([dict])
    document.write(path + filename)
    return path + filename



def download(request):
    file_path = get_new_file()
    if os.path.exists(file_path):
        with open(file_path, 'rb') as fh:
            response = HttpResponse(fh.read(), content_type="application/vnd.ms-excel")
            response['Content-Disposition'] = 'inline; filename=' + os.path.basename(file_path)
            return response