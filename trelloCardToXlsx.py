"""
Requirements:
- python 3.6+
- xlsxwriter

Usage:
python3 trelloCardToXlsx.py <inputfile>
"""

import configparser
import json
import math
import re
import sys
import unicodedata
import xlsxwriter

from dateutil.parser import parse
from dateutil import tz
from datetime import datetime
from os.path import exists as file_exists
from os import remove


def convertUTCtoLocalDatetime(strDate):
    """
    Convert UTC date to local datetime
    :param strDate: String date
    :return: datetime
    """
    from_zone = tz.gettz(config['Dates']['tz_from_zone'])
    to_zone = tz.gettz(config['Dates']['tz_to_zone'])
    utc = parse(strDate)
    utc = utc.replace(tzinfo=from_zone)
    local = utc.astimezone(to_zone)
    #return local.replace(tzinfo=None)
    return local

def countNbLine(cellText, lineLimit):
    """
    Count number of lines for in a text
    :param cellText: String
    :param lineLimit: Integer max char in a line
    :return: Integer number of lines
    """
    nbLine = math.ceil(len(cellText) / lineLimit)
    nbLineBreak = 0
    if cellText.count('\n') > 0:
        nbLineBreak = cellText.count('\n') + 1
    
    if nbLine > nbLineBreak:
        return nbLine
    
    return nbLineBreak


# Get fileInput from command-line argument
if len(sys.argv) < 2:
    print("usage:", sys.argv[0], "<inputfile>\n")
    sys.exit()
inputFilename = sys.argv[1]

# Check if inputFilename exists
if not file_exists(inputFilename):
    print("error:", inputFilename, "does not exist\n")
    sys.exit()

# Read configuration file (config.ini)
if not file_exists('config.ini'):
    print("error : config.ini missing\n")
    sys.exit()
config = configparser.ConfigParser()
config.read('config.ini', encoding='utf-8')






# Read json card file
# using the with keyword to make sure that the file is properly closed.
with open(inputFilename, 'r', encoding='utf-8') as trelloCardFile:
    #data=trelloCardFile.read()
    data = json.load(trelloCardFile)

# Set (clean) outputFileName
outputFileName = unicodedata.normalize('NFKD', data['name']).encode('ascii', 'ignore').decode('utf8')
outputFileName = re.sub(r'[^\w\s-]', '', outputFileName)
outputFileName = re.sub(r'[-\s]+', '-', outputFileName)
outputFileName = outputFileName.strip('-_')
outputFileName = outputFileName[:250] + ".xlsx"

# Remove outputFileName if exists
if file_exists(outputFileName):
    try:
        remove(outputFileName)
    except:
        print("error : outputFileName", outputFileName, "already exists.\nYou should rename or delete outputFileName.")
        sys.exit()
        
# Create Workbook and add sheet
workbook = xlsxwriter.Workbook(outputFileName)
worksheet = workbook.add_worksheet(config['Labels']['sheet_name'])

# Page setup
cellHeight = 15
worksheet.set_portrait()
worksheet.set_paper(9) # A4
worksheet.set_margins(0.6, 0.6, 0.75, 0.75) # L R T B

# Set column properties 
worksheet.set_column('A:A', 5)
worksheet.set_column('B:B', 12)
worksheet.set_column('C:F', 17)

# Set row properties
worksheet.set_default_row(cellHeight)
worksheet.set_row(0, 60)    # Title

# Set cell format
cell_title_n1 = workbook.add_format({
    'font_color': 'white', 'font_name': 'Calibri', 'font_size': 16, 'bold': True,
    'bg_color' : '#3969AD',
    'text_wrap': True,
    'indent' : 1
})
cell_title_n1.set_align('left')
cell_title_n1.set_align('top')

cell_title_n2 = workbook.add_format({
    'font_color': 'black', 'font_name': 'Calibri', 'font_size': 14, 'bold': True,
    'indent' : 1
})
cell_title_n2.set_align('left')
cell_title_n2.set_align('top')

cell_title_n3 = workbook.add_format({
    'font_color': 'black', 'font_name': 'Calibri', 'font_size': 12, 'bold': True, 'italic': True,
    'indent' : 2
})
cell_title_n3.set_align('left')
cell_title_n3.set_align('top')

cell_description = workbook.add_format({
    'text_wrap': True,
    'indent' : 1
})
cell_description.set_align('left')
cell_description.set_align('top')


cell_activity_date = workbook.add_format()
cell_activity_date.set_align('right')
cell_activity_date.set_align('top')
cell_activity_date.set_text_wrap()

cell_activity_comment = workbook.add_format()
cell_activity_comment.set_align('left')
cell_activity_comment.set_align('top')
cell_activity_comment.set_text_wrap()

cell_complete_status = workbook.add_format({
    'bold': True, 'italic': True,
    'font_color': '#B6D7A8',
})
cell_complete_status.set_align('right')
cell_complete_status.set_align('top')

cell_incomplete_status = workbook.add_format({
    'bold': True, 'italic': True,
    'font_color': 'red',
})
cell_incomplete_status.set_align('right')
cell_incomplete_status.set_align('top')

cell_complete_checkItem = workbook.add_format()
cell_complete_checkItem.set_align('left')
cell_complete_checkItem.set_align('top')
cell_complete_checkItem.set_text_wrap()
cell_complete_checkItem.set_font_strikeout()

cell_incomplete_checkItem = workbook.add_format()
cell_incomplete_checkItem.set_align('left')
cell_incomplete_checkItem.set_align('top')
cell_incomplete_checkItem.set_text_wrap()


cell_percent_format = workbook.add_format({'num_format': '0%'})

cell_blue = workbook.add_format({
    'font_color': 'white',
    'align': 'left',
    'align': 'top',
    'text_wrap': True,
    'bg_color' : '#2D5389',
    'indent' : 1
})

cell_grey = workbook.add_format({
    'font_color': 'black',
    'align': 'left',
    'align': 'top',
    'text_wrap': True,
    'bg_color' : '#EFEFEF',
    'indent' : 1
})

cell_green = workbook.add_format({
    'font_color': 'black',
    'align': 'center',
    'text_wrap': True,
    'bg_color' : '#B6D7A8'
})

cell_orange = workbook.add_format({
    'font_color': 'black',
    'align': 'center',
    'text_wrap': True,
    'bg_color' : '#FABF8F'
})


# Write Sheet

# Title
# --------------------
worksheet.merge_range('A1:F1', data['name'], cell_title_n1)
# separators
worksheet.merge_range('A2:F2', "" , cell_blue)
worksheet.merge_range('A3:F3', "" , cell_grey)

# List
# --------------------
strList = ""
if data['idList'] in config['TrelloLists']:
    strList = config['Labels']['in_list'] + " " + config['TrelloLists'][data['idList']]
worksheet.merge_range('A4:F4', strList, cell_grey)

# Labels
# --------------------
strLabels = config['Labels']['labels'] + " : "
if 'labels' in data:
    if len(data['labels']) > 0:
        arrLabels = []
        for index in range(len(data['labels'])):
            if data['labels'][index]['name'] != "":
                arrLabels.append(data['labels'][index]['name'])
        strLabels += ", ".join(arrLabels)
worksheet.merge_range('A5:F5', strLabels, cell_grey )
worksheet.merge_range('A6:F6', "" , cell_grey)

# Dates
# --------------------
# start date
strStartDate = config['Labels']['start_date'] + " : "
if 'start' in data:
    if data['start'] is not None:
        startDate = convertUTCtoLocalDatetime(data['start'])
        strStartDate += startDate.strftime(config['Dates']['str_date_format'])
worksheet.merge_range('A7:C7', strStartDate, cell_grey)

# due date
strDueDate = config['Labels']['due_date'] + " : "
if 'due' in data:
    if data['due'] is not None:
        dueDate = convertUTCtoLocalDatetime(data['due'])
        strDueDate += dueDate.strftime(config['Dates']['str_datetime_format']) 
worksheet.merge_range('D7:E7', strDueDate, cell_grey)

# due complete
strDueComplete = ""
cell_format = cell_grey
if 'dueComplete' in data:
    if data['dueComplete']:
        strDueComplete = config['Labels']['due_date_complete']
        cell_format = cell_green
    else:
        # we have to compare today date with due date
        if 'due' in data:
            if data['due'] is not None:
                dueDate = convertUTCtoLocalDatetime(data['due'])
                nowDate = convertUTCtoLocalDatetime(str(datetime.utcnow()))
                if dueDate < nowDate:
                    strDueComplete = config['Labels']['due_date_overdue']
                    cell_format = cell_orange
worksheet.write('F7', strDueComplete, cell_format)

# last activity date
strdLastActivityDate = config['Labels']['last_activity_date'] + " : "
if 'dateLastActivity' in data:
    if data['dateLastActivity'] is not None:
        lastActivityDate = convertUTCtoLocalDatetime(data['dateLastActivity'])
        strdLastActivityDate += lastActivityDate.strftime(config['Dates']['str_datetime_format']) 
worksheet.merge_range('A8:C8', strdLastActivityDate, cell_grey)
worksheet.merge_range('D8:F8', "", cell_grey)

# Separator
worksheet.merge_range('A9:F9', "" , cell_grey)


# Description
# --------------------
worksheet.write('A11', config['Labels']['description'], cell_title_n2)
worksheet.merge_range('A12:F12', data['desc'], cell_description)
rowHeight = (countNbLine(data['desc'], 80) + 1) * cellHeight
worksheet.set_row(11, rowHeight)


# Checklists
# --------------------
cellRow = 13
if 'checklists' in data:
    if len(data['checklists']) > 0:
        arrChecklists = []
        # get checklists
        for i in range(len(data['checklists'])):
            if data['checklists'][i]['name'] != "":
                arrCheckItems = []
                # get checkItems
                if data['checklists'][i]['checkItems']:
                    if len(data['checklists'][i]['checkItems']) > 0:
                        for j in range(len(data['checklists'][i]['checkItems'])):
                            arrCheckItems.append([
                                data['checklists'][i]['checkItems'][j]['name'],
                                data['checklists'][i]['checkItems'][j]['pos'],
                                data['checklists'][i]['checkItems'][j]['state']
                            ])
                         
                         # order checkItems by position (pos)
                        if len(arrCheckItems) > 0:
                             arrCheckItems =  sorted(arrCheckItems, key=lambda x: x[1])
                            
                # push checklist
                arrChecklists.append([
                    data['checklists'][i]['name'],
                    data['checklists'][i]['pos'],
                    arrCheckItems
                ])
        
        # order checklists by position
        if len(arrChecklists) > 0:
            arrChecklists = sorted(arrChecklists, key=lambda x: x[1])
        
        # write Checklists
        # Title
        worksheet.write(cellRow, 0, config['Labels']['checklists'], cell_title_n2)
        for i in range(len(arrChecklists)):
            # checklist Title
            cellRow += 1
            worksheet.write(cellRow, 0, str( arrChecklists[i][0] ), cell_title_n3)
            # checklist checkItems
            if len(arrChecklists[i][2]) > 0:
                checklistRow = cellRow
                checklistComplete = 0
                for j in range(len(arrChecklists[i][2])):
                    cellRow += 1
                    strStatus = "X"
                    cell_status = cell_incomplete_status
                    cell_checkItem = cell_incomplete_checkItem
                    if str( arrChecklists[i][2][j][2] ) == "complete":
                        checklistComplete += 1
                        strStatus = "V"
                        cell_status = cell_complete_status
                        cell_checkItem = cell_complete_checkItem
                    worksheet.write(cellRow, 0, strStatus, cell_status )
                    worksheet.merge_range(cellRow, 1, cellRow, 5, str( arrChecklists[i][2][j][0] ), cell_checkItem )
                    rowHeight = countNbLine(str( arrChecklists[i][2][j][0] ), 70) * cellHeight
                    worksheet.set_row(cellRow, rowHeight)
                pcentComplet = checklistComplete / len(arrChecklists[i][2])
                worksheet.write(checklistRow, 5, pcentComplet, cell_percent_format)
            cellRow += 1


# Activity
# --------------------
cellRow += 1
if 'actions' in data:
    if len(data['actions']) > 0:
        # Title
        worksheet.write(cellRow, 0, config['Labels']['activity'], cell_title_n2)
        
        # Items
        for i in range(len(data['actions'])):
            if data['actions'][i]["type"] == "commentCard":
                cellRow += 1
                # Date + fullName|initials
                activityDate = convertUTCtoLocalDatetime(data['actions'][i]['date'])
                strActivityDate = activityDate.strftime(config['Dates']['str_datetime_format']) + "\n" + str(data['actions'][i]['memberCreator'][config['Labels']['user_fullName']])
                worksheet.merge_range(cellRow, 0, cellRow, 1, strActivityDate, cell_activity_date)
                
                # Comment
                worksheet.merge_range(cellRow, 2, cellRow, 5, str(data['actions'][i]["data"]["text"]), cell_activity_comment)
                rowHeight = (countNbLine(str(data['actions'][i]["data"]["text"]), 60) + 1) * cellHeight
                if rowHeight < 60:
                    rowHeight = 60
                worksheet.set_row(cellRow, rowHeight)


# Close Workbook
# --------------------
workbook.close()
print("Done: file", outputFileName, "created\n")
