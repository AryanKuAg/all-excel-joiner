from re import template
import xlrd
from openpyxl import Workbook
import re

#
# # For row 0 and column 0
# print(type(sheet.cell_value(3, 5)))


workbook = Workbook()
ws = workbook.active
numberPattern = re.compile('[0-9]')

for file in range(1001,2001):
    tempList = []
   #  Give the location of the file
    loc = ("Andhra General Scan (Box 4)/Andhra General Scan (Box 2)_Page_"+ str(file).zfill(4) +".xlsx")

    # To open Workbook
    try:
        wb = xlrd.open_workbook(loc)
    except:
        print('no file')
    sheet = wb.sheet_by_index(0)

    ws.append(["Andhra General Scan (Box 2)_Page_"+ str(file).zfill(4)])
    for row in range(0, 23):
        tempList.append([])
        
        reversedAddressTwo = ''
        reversedAddressFirst = ''
        addressTwo = ''
        addressOne = ''
        for col in range(0, 18):
            addressTwoBool = True
            ##########################################
            try:
                tempData = str(sheet.cell_value(row, col))
                ###################ADDRESS######################
                if col == 7:
                   
                    #print('col 7 ', tempData)
                    for i in tempData[::-1]:
                                            
                        if numberPattern.search(i) and addressTwoBool:
                           
                            addressTwoBool = False
                            
                        if addressTwoBool:
                            reversedAddressTwo += i
                           
                        else:
                            reversedAddressFirst += i
                          
            
                    for j in reversedAddressFirst[::-1]:
                        addressOne += j

                    for j in reversedAddressTwo[::-1]:
                        addressTwo += j
                
                    tempData = addressOne
                    

                if col == 8:
                    #print('col 8 ', tempData)
                    tempData = addressTwo
                    
                ########################################
                if col == 3 or col == 16 or col == 4:
                    if len(tempData) < 8:
                        tempData = '0' + tempData

                if col == 6:
                    name = ''
                    for yoyo in tempData.split():
                        name += yoyo + "  "

                    tempData = name.strip()

                if col == 7:
                    tempData = tempData.capitalize()

                
                        
            except:
                tempData = ''
            ##########################################

            tempList[row].append(str(tempData).replace(".0", '').strip())


    for i in tempList:
        ws.append(i)

    
  
    ws.append([])




workbook.save("Andhra General Scan (Box 4).xlsx")
