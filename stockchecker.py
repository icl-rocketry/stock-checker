import pygsheets
from mouser.api import MouserPartSearchRequest
from datetime import datetime
from openpyxl import load_workbook

API_KEYS_FILE = 'C:\stockchecker\mouser_api_keys.yaml'

#Google Sheets
#gc = pygsheets.authorize(service_file='auth.json')
#sh = gc.open('Component Stock Checking')
#wks = sh[0]
#read = wks.get_as_df()

#Excel
stocksheet = load_workbook(filename="stock.xlsx")
sheet = stocksheet.active
datalen = sheet.max_row-1

stock = []
for i in range(0,datalen):
    args = []
    partno = str(sheet['A'+str(2+i)].value) #excel
    #print(read.iloc[i][0]) #google sheets
    print(partno)
    request = MouserPartSearchRequest('partnumber', API_KEYS_FILE, *args)
    if request.url:
            # Run request
            search = request.part_search(partno)
            if search:
                # Print result
                respraw = request.get_clean_response()
                stock.append([respraw["Availability"]])
                sheet["F"+str(2+i)] = respraw["Availability"] #excel

now = datetime.now()
current_time = now.strftime("%H:%M:%S on %d/%m/%Y")

#Google Sheets
#wks.update_values((2,6),stock) 
#wks.update_values('L2', [[current_time]])

#Excel
sheet["L2"] = current_time
stocksheet.save(filename="stock.xlsx")