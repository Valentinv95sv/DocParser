from mailmerge import MailMerge
import datetime

class docx_insert_data_class:

    def __init__(self, doc_path, inserted_data):
        self.doc_path = doc_path
        self.inserted_data = inserted_data

    def get_new_file(self, filename):
        dict = self.inserted_data
        template = self.doc_path
        document = MailMerge(template)
        document.merge_pages([dict])
        document.write(filename)
        return filename

    def get_fields(self):
        template = self.doc_path
        document = MailMerge(template)
        return document.get_merge_fields()

    def get_file_name(self):
        filename = 'Blank_' + str(datetime.datetime.now().strftime("%H%M%S")) + '.docx'
        path = '/blanks'
        return path + filename




