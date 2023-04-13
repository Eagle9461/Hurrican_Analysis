from bs4 import BeautifulSoup
import csv
import datetime
import urllib.request
import pandas as pd
import openpyxl
import os
import sys

strInputState = input('Input your state\n')

html_doc = ''
url = 'https://www.aoml.noaa.gov/hrd/hurdat/uststorms.html'
try:
    request = urllib.request.Request(url)
    response = urllib.request.urlopen(request)
    html_doc = response.read()
except Exception as ee:
    print(ee)
    print("Request Error.")
    input('Press Enter to exit.')
    exit()

def handleCellStr(_strCell, _bDelBar = True):
    strCell = _strCell.replace('\t', '')
    strCell = strCell.replace('\n', '')
    if _bDelBar == True:
        strCell = strCell.replace('-', '')
    strCell = strCell.replace(' ', '')
    return strCell

title = ['Number', 'Name of Storm' , 'Latitude', 'Longitude', 'Date', 'Estimated Landfall', 'Estimated Max Winds']

export_date = datetime.datetime.now()
export_xlsfname = export_date.strftime('%m') + '.' + \
               export_date.strftime('%d') + '.' + \
               export_date.strftime('%Y') + '_0_' + strInputState + '.xlsx'
if os.path.exists(export_xlsfname):
    try:
        os.rename(export_xlsfname, export_xlsfname + '_')
        os.rename(export_xlsfname + '_', export_xlsfname)
    except OSError as e:
        print("Excel file opened. After close, and then Try.")
        exit()

results = []

# html_file = open("response.txt", "r")
# html_doc = html_file.read()
# bs_soup = BeautifulSoup(html_doc, "html.parser")
bs_soup = BeautifulSoup(html_doc, "lxml")
tag_lines = None
try:
    tag_lines = bs_soup.find('td', {'id': 'tdcontent'}).find('div').find('center').find('table').findAll('tr')
    
    # tag_lines = bs_soup.find('td', {'id': 'tdcontent'}).find_all('tr')
except Exception as ee:
    print(ee)
    print("No lines.")
    input('Press Enter to exit.')
    exit()

if tag_lines == None or len(tag_lines) == 0:
    print("No Content")
    input('Press Enter to exit.')
    exit()

tag_lines = tag_lines[2:]
idx = 0
for tr_line in tag_lines:
    idx += 1
    line_cells = tr_line.find_all('td')
    if len(line_cells) != 7 and len(line_cells) != 8:
        continue
    #print(line_cells)
    strStorm = line_cells[0].text
    strDate = line_cells[1].text
    strTime = line_cells[2].text
    strLatitude = line_cells[3].text
    strLongitude = line_cells[4].text
    strMaxWinds = line_cells[5].text
    strStates = line_cells[6].text
    strStormNames = ""
    if len(line_cells) > 7:
        strStormNames = line_cells[7].text

    strStrom = handleCellStr(strStorm)
    strDate = handleCellStr(strDate, False)
    strTime = handleCellStr(strTime)
    strLatitude = handleCellStr(strLatitude)
    strLongitude = handleCellStr(strLongitude)
    strMaxWinds = handleCellStr(strMaxWinds)
    strStates = handleCellStr(strStates)
    strStormNames = handleCellStr(strStormNames)

        
    while True:
        try:
            temp = strDate[len(strDate) - 1:]
            nTemp = int(temp)
            break
        except Exception as e:
            strDate = strDate[:len(strDate) - 1]
        
   # strNoFront = format(int(strDate.split('-')[0]), '02d')
    converted_num = int(strStorm)
    if converted_num > 10:
        strNumber = strDate[len(strDate) - 2:] + '-' + strStorm
    strNumber = strDate[len(strDate) - 2:] + '-' + '0' + strStorm
 #   strModifedDate = strDate.split('-')[1]
        
    bState = False
    strStateLast = ''
    strSeverity = ''
    if strStates.find(',') > -1:
        nList = [0]
        strStateItems = strStates.split(',')
        for strStateOne in strStateItems:
            if strStateOne.find(strInputState) > -1:
                strSeverity = strStateOne
                bState = True
                nList.append(int(strStateOne[len(strStateOne) - 1:]))
        strStateLast = str(max(nList))
    else:
        if strStates.find(strInputState) > -1:
            bState = True
        strStateLast = strStates[len(strStates) - 1:]
        
    if bState == False:
        continue
        
    strType = ''
    if strStateLast == '1':
        strType = 'TS'
    elif strStateLast == '2':
        strType = 'H'
    elif strStateLast == '3':
        strType = 'MH'
    elif strStateLast == '4':
        strType = 'MH'
    elif strStateLast == '5':
        strType = 'MH'
#'Number', 'Name of Storm' , 'Latitude', 'Longitude', 'Estimated Landfall', 'Estimated Max Winds'
    line = [strNumber, strStormNames, strLatitude, strLongitude, strDate, strStates, strMaxWinds]
    results.append(line)

df = pd.DataFrame(results, columns=title)
df.to_excel(export_xlsfname, sheet_name='hurricane', index=False)
input('Press Enter to exit.')