import os
from django.shortcuts import HttpResponse
from django.shortcuts import render
from . import Docx_insert_data_class

inserted_data = {
    'toWhom' : '123',
    'fromWho' : '234',
    'Tender_Number' : 'nony',
    'max_work' : 'qwerrty'
}
template = 'docParser/Doc_form.docx'
path = 'blanks/'
a = Docx_insert_data_class.docx_insert_data_class(template, inserted_data)

def index(request):
    return render(request, 'test.html')

def download(request):
    file_path = a.get_new_file(path + a.get_file_name())
    if os.path.exists(file_path):
        with open(file_path, 'rb') as fh:
            response = HttpResponse(fh.read(), content_type="application/vnd.ms-excel")
            response['Content-Disposition'] = 'inline; filename=' + os.path.basename(file_path)
            return response