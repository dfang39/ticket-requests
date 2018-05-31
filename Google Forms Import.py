import datetime
import os
import openpyxl

downloadFolder = 'C:\\Users\\DanFang\\Downloads'
os.chdir(downloadFolder)

dated_files = [(os.path.getmtime(fn), os.path.basename(fn))
               for fn in os.listdir(downloadFolder) if fn.lower().endswith('.xlsx')]
dated_files.sort()
dated_files.reverse()
excel = dated_files[0][1]

workbook = openpyxl.load_workbook(excel, data_only = True, read_only = True)
sheet = workbook['Form Responses 1']
print(sheet)

class Request(object):
    def __init__(self, agency_name, mc_name, mc_email, game1):
        self.event = event
        self.agency = agency
        self.preferred_date = preferred_date
        self.tickets_requested = tickets_requested
        self.chaperone_name = chaperone_name
        self.notes = notes
        self.status = status
        self.created_on = created_on