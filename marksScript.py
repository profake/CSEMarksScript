import os
import re
from datetime import datetime
from openpyxl import Workbook

# assign directory, getcwd returns current working dir
directory = os.getcwd()
 
# User Variables
idPrefix = '222-115'
idStart = 161
idEnd = 200 

# My variables
dateMatchRegex = '\d+-\d+-\d+'
fileDates = []
workbook = Workbook()
sheet = workbook.active
sheet['A1'] = "Date"
sheet['A2'] = "Class No"

# iterate over files in
# that directory
for filename in os.listdir(directory):
    file = os.path.join(directory, filename)
    # checking if it is a file
    if os.path.isfile(file):
        fileName = file.split('.')[0]
        fileExtension = file.split('.')[1]
        if (fileExtension == 'txt' and re.search(dateMatchRegex, fileName)):
          # checking if file is a txt and follows the dd-mm format or d-mm or dd-m or d-m format
          fileDates.append(re.search(dateMatchRegex, fileName).group(0))

fileDates.sort(key=lambda date: datetime.strptime(date, "%d-%m-%y"))
# now the text files are sorted by date

# Column A1 for IDs
# Column B2, C2, D2 ... Z2 for Dates. First date is B2, second C2...

index = 3

for id in range(idStart, idEnd+1):
  curCell = 'A{}'.format(index)
  sheet[curCell] = idPrefix+'-'+str(id)
  index+=1
  # Adds student IDs serially in the sheet 

letterIndex = 'B'
for date in range(0, len(fileDates)):
  curCell = '{}1'.format(letterIndex)
  sheet[curCell] = fileDates[date]
  letterIndex = chr(ord(letterIndex) + 1) # Increments letterIndex
  # adds dates serially

# Loop through all files 
# Read the IDs and match with studentsIds
# If match, set the cell as 'P'
column = 'B'
fileCount = 1
for date in fileDates:
  fileName = os.path.join(directory, date + '.txt')
  with open (fileName, 'r') as f:
    for line in f:
      studentId = line.strip()
      index = 3
      sheet['{}2'.format(column)] = fileCount # number of class
      # Loop through sheet values
      curCell = 'A{}'.format(index)
      while(sheet[curCell].value):
        curCell = 'A{}'.format(index)
        if (sheet[curCell].value) == studentId:
          placeCell = '{}{}'.format(column, index)
          sheet[placeCell] = 'P'
        index+=1
  column = chr(ord(column) + 1)
  fileCount+=1

curDateTime = datetime.now().strftime("%d_%m_%Y %H_%M_%S")
workbook.save(filename=curDateTime + ".xlsx")