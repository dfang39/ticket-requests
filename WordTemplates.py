"""for generating ticket request forms"""
import os
import xlrd
import datetime
import mailmerge

today = datetime.date.today().strftime("%m/%d/%Y")

downloadFolder = 'C:\\Users\\DanFang\\Downloads'
os.chdir(downloadFolder)

dated_files = [(os.path.getmtime(fn), os.path.basename(fn))
               for fn in os.listdir(downloadFolder) if fn.lower().endswith('.xlsx')]
dated_files.sort()
dated_files.reverse()
excel = dated_files[0][1]

open_excel = xlrd.open_workbook(excel)
sheet = open_excel.sheet_by_index(0)
keys = [sheet.cell(0, col_index).value for col_index in range(sheet.ncols)]

ticket_requests = []
for row_index in range(1, sheet.nrows):
    d = {keys[col_index]: sheet.cell(row_index, col_index).value for col_index in range(sheet.ncols)}
    ticket_requests.append(d)

#print(ticket_requests)

class Request(object):
    def __init__(self, event = "", agency = "", preferred_date_total = "", preferred_date = "", preferred_time = "", alternate_date = "", alternate_time = "", tickets_requested = "", number_children = "", number_adults = "", chaperone_name = "", chaperone_cell = "", notes = "", kids_youngest = "", kids_oldest = "", address = "", city = "", state = "", postal = "", email = "", phone = "", fax = "", primary_contact = "", transportation_type = ""):
        self.event = event
        self.agency = agency
        self.preferred_date_total = preferred_date_total
        self.preferred_date = preferred_date
        self.preferred_time = preferred_time
        self.alternate_date = alternate_date
        self.alternate_time = alternate_time
        self.tickets_requested = tickets_requested
        self.number_children = number_children
        self.number_adults = number_adults
        self.chaperone_name = chaperone_name
        self.chaperone_cell = chaperone_cell
        self.notes = notes
        self.kids_youngest = kids_youngest
        self.kids_oldest = kids_oldest
        self.address = address
        self.city = city
        self.state = state
        self.postal = postal
        self.email = email
        self.phone = phone
        self.fax = fax
        self.primary_contact = primary_contact
        self.transportation_type = transportation_type


#print(ticket_requests[0])

#request_objects = []
#for request in ticket_requests:
    #request_objects.append(Request(request["Event"]))

#print(len(request_objects))

request_objects = []

for request in ticket_requests:
    pref_date_string = xlrd.xldate.xldate_as_datetime(request['Preferred Date'], open_excel.datemode)
    request["Pref Date"] = datetime.datetime.date(pref_date_string)
    request["Pref Time"] = datetime.datetime.time(pref_date_string)
    request["Alt Date"] = ""
    request["Alt Time"] = ""
    if request['Alternate Date'] != '':
        alt_date_string = xlrd.xldate.xldate_as_datetime(request['Alternate Date'], open_excel.datemode)
        request["Alt Date"] = datetime.datetime.date(alt_date_string)
        request["Alt Time"] = datetime.datetime.time(alt_date_string)

    primary_contact = " ".join(request["Primary Contact (Agency) (Account)"].split(", ")[::-1])
    request["Primary Contact (Agency) (Account)"] = primary_contact
    #print(request["Primary Contact (Agency) (Account)"])

    request_objects.append(Request(request["Event"], request["Agency"], pref_date_string, request["Pref Date"], request["Pref Time"], request["Alt Date"], request["Alt Time"], request["Tickets Requested"], request["Number of Children"], request["Total Adults"], request["Chaperone Name"], request["Chaperone Cell Phone"], request["Notes"], request["Kids Youngest Age"], request["Kids Oldest Age"], request["Address Line 1 (Agency) (Account)"], request["City (Agency) (Account)"], request["State or Province (Agency) (Account)"], request["Postal Code (Agency) (Account)"], request["Email (Agency) (Account)"], request["Main Phone (Agency) (Account)"], request["Fax (Agency) (Account)"], request["Primary Contact (Agency) (Account)"], request["Transportation Type"]))

out_folder = "C:\\Users\\DanFang\\Desktop\\Requests"
os.chdir(out_folder)

#Region Carnegie Science Center

csc_requests = []
for request in request_objects:
    if "Carnegie Science Center" in request.event:
        csc_requests.append(request)

print('CSC:', len(csc_requests))

csc_requests.sort(key=lambda r: r.preferred_date)

if len(csc_requests) > 0:
    cscdoc = mailmerge.MailMerge('csctemplate.docx')
    #print(cscdoc.get_merge_fields())

    csc_merge = []
    for request in csc_requests:
        departure_string = request.preferred_date_total + datetime.timedelta(hours=4)
        departure_time = datetime.datetime.time(departure_string)

        if request.alternate_date != "":
            csc_merge.append({'PreferredTime': request.preferred_time.strftime("%#I:%M %p"), 'DepartureTime': departure_time.strftime("%#I:%M %p"), 'DateToday': today, 'PreferredDate': request.preferred_date.strftime("%#m/%#d/%Y"), 'Notes': request.notes, 'MainContact': request.primary_contact, 'Agency': request.agency, 'city': request.city, 'Chaperone': request.chaperone_name, 'Zip': request.postal, 'Alternate': request.alternate_date.strftime("%#m/%#d/%Y") + " " + request.alternate_time.strftime("%#I:%M %p"), 'NumberAdults': str(int(request.number_adults)), 'Number': str(int(request.number_children)), 'State': request.state, 'ChaperoneCell': request.chaperone_cell, 'Address': request.address, 'Total': str(int(request.tickets_requested)), 'TransportationType': request.transportation_type, 'AgencyPhone': request.phone})
        else:
            csc_merge.append({'PreferredTime': request.preferred_time.strftime("%#I:%M %p"),
                              'DepartureTime': departure_time.strftime("%#I:%M %p"), 'DateToday': today,
                              'PreferredDate': request.preferred_date.strftime("%#m/%#d/%Y"), 'Notes': request.notes,
                              'MainContact': request.primary_contact, 'Agency': request.agency, 'city': request.city,
                              'Chaperone': request.chaperone_name, 'Zip': request.postal,
                              'Alternate': "",
                              'NumberAdults': str(int(request.number_adults)),
                              'Number': str(int(request.number_children)), 'State': request.state,
                              'ChaperoneCell': request.chaperone_cell, 'Address': request.address,
                              'Total': str(int(request.tickets_requested)),
                              'TransportationType': request.transportation_type, 'AgencyPhone': request.phone})

    cscdoc.merge_pages(csc_merge)

    csc_request_title = 'TFK CSC Ticket Requests ' + str(today).replace("/", "-") + '.docx'

    cscdoc.write(csc_request_title)

#Region Children's Museum

cmpgh_requests = []
for request in request_objects:
    if "Children's Museum" in request.event:
        cmpgh_requests.append(request)

print('Childrens Museum:', len(cmpgh_requests))

cmpgh_requests.sort(key=lambda r: r.preferred_date)

if len(cmpgh_requests) > 0:
    cmpghdoc = mailmerge.MailMerge('childrensmuseumtemplate.docx')
    #print(cmpghdoc.get_merge_fields())

    cmpgh_merge = []
    for request in cmpgh_requests:
        departure_string = request.preferred_date_total + datetime.timedelta(hours=4)
        departure_time = datetime.datetime.time(departure_string)
        cmpgh_merge.append({'PreferredTime': request.preferred_time.strftime("%#I:%M %p"), 'DepartureTime': departure_time.strftime("%#I:%M %p"), 'PreferredDate': request.preferred_date.strftime("%#m/%#d/%Y"), 'Notes': request.notes, 'Agency': request.agency, 'City': request.city, 'Chaperone': request.chaperone_name, 'Zip': request.postal, 'NumberAdults': str(int(request.number_adults)), 'NumberChildren': str(int(request.number_children)), 'State': request.state, 'ChaperonePhone': request.chaperone_cell, 'Address': request.address, 'Transportation': request.transportation_type, 'KidsYoungest': str(int(request.kids_youngest)), 'KidsOldest': str(int(request.kids_oldest)) })

    cmpghdoc.merge_pages(cmpgh_merge)

    cmpgh_request_title = 'TFK Children\'s Museum Ticket Requests ' + str(today).replace("/", "-") + '.docx'

    cmpghdoc.write(cmpgh_request_title)

#Region History Center

heinz_requests = []
for request in request_objects:
    if "Heinz History Center" in request.event:
        heinz_requests.append(request)

print('Heinz History:', len(heinz_requests))

heinz_requests.sort(key=lambda r: r.preferred_date)

if len(heinz_requests) > 0:
    heinzdoc = mailmerge.MailMerge('heinztemplate.docx')
    # print(heinzdoc.get_merge_fields())

    heinz_merge = []
    for request in heinz_requests:
        departure_string = request.preferred_date_total + datetime.timedelta(hours=2)
        departure_time = datetime.datetime.time(departure_string)
        if request.alternate_date != "":
            heinz_merge.append({'PreferredTime': request.preferred_time.strftime("%#I:%M %p"), 'DepartureTime': departure_time.strftime("%#I:%M %p"), 'PreferredDate': request.preferred_date.strftime("%#m/%#d/%Y"), 'AlternateDate': request.alternate_date.strftime("%#m/%#d/%Y") + " " + request.alternate_time.strftime("%#I:%M %p"), 'Notes': request.notes, 'Agency': request.agency, 'Chaperone': request.chaperone_name, 'NumberAdults': str(int(request.number_adults)), 'NumberChildren': str(int(request.number_children)), 'ChaperonePhone': request.chaperone_cell, 'KidsYoungest': str(int(request.kids_youngest)), 'KidsOldest': str(int(request.kids_oldest)), 'Fax': request.fax})
        else:
            heinz_merge.append({'PreferredTime': request.preferred_time.strftime("%#I:%M %p"),
                                'DepartureTime': departure_time.strftime("%#I:%M %p"),
                                'PreferredDate': request.preferred_date.strftime("%#m/%#d/%Y"),
                                'AlternateDate': "",
                                'Notes': request.notes, 'Agency': request.agency, 'Chaperone': request.chaperone_name,
                                'NumberAdults': str(int(request.number_adults)),
                                'NumberChildren': str(int(request.number_children)),
                                'ChaperonePhone': request.chaperone_cell,
                                'KidsYoungest': str(int(request.kids_youngest)),
                                'KidsOldest': str(int(request.kids_oldest)), 'Fax': request.fax})

    heinzdoc.merge_pages(heinz_merge)

    heinz_request_title = 'TFK Heinz History Ticket Requests ' + str(today).replace("/", "-") + '.docx'

    heinzdoc.write(heinz_request_title)

# Region Mattress Factory

mattress_requests = []
for request in request_objects:
    if "Mattress Factory" in request.event:
        mattress_requests.append(request)

print('Mattress Factory:', len(mattress_requests))

mattress_requests.sort(key=lambda r: r.preferred_date)

if len(mattress_requests) > 0:
    mattressdoc = mailmerge.MailMerge('mattresstemplate.docx')
    # print(mattressdoc.get_merge_fields())

    mattress_merge = []
    for request in mattress_requests:
        mattress_merge.append({'PreferredDateTime': request.preferred_date.strftime("%#m/%#d/%Y") + " " + request.preferred_time.strftime("%#I:%M %p"),'NumberAdults': str(int(request.number_adults)), 'Agency': request.agency, 'Chaperone': request.chaperone_name, 'NumberChildren': str(int(request.number_children)), 'ChaperonePhone': request.chaperone_cell, "Main": request.primary_contact, "Email": request.email, "TotalTickets": str(int(request.tickets_requested))})

    mattressdoc.merge_pages(mattress_merge)

    mattress_request_title = 'TFK Mattress Factory Ticket Requests ' + str(today).replace("/", "-") + '.docx'

    mattressdoc.write(mattress_request_title)