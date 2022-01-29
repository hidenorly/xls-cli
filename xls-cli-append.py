#   Copyright 2022 hidenorly
#
#   Licensed under the Apache License, Version 2.0 (the "License");
#   you may not use this file except in compliance with the License.
#   You may obtain a copy of the License at
#
#       http://www.apache.org/licenses/LICENSE-2.0
#
#   Unless required by applicable law or agreed to in writing, software
#   distributed under the License is distributed on an "AS IS" BASIS,
#   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
#   See the License for the specific language governing permissions and
#   limitations under the License.

import sys
import os
import argparse
import unicodedata
import openpyxl as xl

def ljust_jp(value, length, pad = " "):
  count_length = 0
  for char in value.encode().decode('utf8'):
    if ord(char) <= 255:
      count_length += 1
    else:
      count_length += 2
  return value + pad * (length-count_length)

def isRobustNumeric(data):
  if data is not None:
    try:
      float(data)
    except ValueError:
      return False
    return True
  else:
    return False

def getCsvJsonCommon(row):
  result = ""

  for aData in row:
    if result != "":
      result = result + ", "
    if aData is None:
      aData = ""
    if isRobustNumeric(aData):
      result = result + str(aData)
    else:
      result = result + "\"" + str(aData) + "\""

  return result

def getCsv(row):
  result = getCsvJsonCommon(row) + ","
  return result

def getJson(row):
  result = "[ " + getCsvJsonCommon(row) + " ],"
  return result

def dumpSheet(aSheet, args):
  for aRow in aSheet.values:
    if args.json:
      print( "  " + getJson(aRow) )
    elif args.csv:
      print( getCsv(aRow) )

def openBook(fileName):
  result = None

  if os.path.exists( fileName ):
    result = xl.load_workbook( fileName )
  else:
    result = xl.Workbook()

  return result

def openSheet(book, sheetName):
  result = None

  if book:
    for aSheet in book:
      if sheetName == "*" or aSheet.title == sheetName:
        result = aSheet

    if result == None:
      result = book.create_sheet()
      if not sheetName == "*":
        result.title = sheetName

  return result

def getValueArrayFromCells(cells):
  result = []
  for aCol in cells:
    result.append( aCol.value )
  return result

def getLastPosition( sheet ):
  if sheet.max_column==1 and sheet.max_row == 1:
    return 1, 1
  else:
    return 1, sheet.max_row+1

def setCells(targetSheet, startPosX, startPosY, rows):
  print("hello from " + str(startPosX) + ":" + str(startPosY) )
  if isinstance(targetSheet, xl.worksheet.worksheet.Worksheet):
    y = startPosY
    for aRow in rows:
      x = startPosX
      for aCell in aRow:
        targetSheet.cell(row=y, column=x).value = aCell.value
        x = x + 1
      y = y + 1

def dumpRows(rows):
  for aRow in rows:
    data = getValueArrayFromCells( aRow )
    print( getCsv( data ) )

def getAllSheets(book):
  result = []
  for aSheet in book:
    result.append( aSheet )
  return result

if __name__=="__main__":
  parser = argparse.ArgumentParser(description='Parse command line options.')
  parser.add_argument('args', nargs='*', help='Specify input.xlsx output.xlsx')
  parser.add_argument('-i', '--inputsheet', action='store', default="*", help='Specfy sheet name e.g. Sheet1')
  parser.add_argument('-o', '--outputsheet', action='store', default="*", help='Specfy sheet name e.g. Sheet1')
  parser.add_argument('-m', '--merge', action='store_true', default=False, help='Specify if you want to merge all of sheets of input book')

  args = parser.parse_args()

  if len(args.args)==2:
    targetSheets = []
    targetSheets.append( args.inputsheet )
    targetSheets.append( args.outputsheet )

    books=[]
    i = 0
    for aFile in args.args:
      aBook = openBook( aFile )
      books.append( aBook )
      i = i + 1

    inputSheets = []
    if args.merge:
      inputSheets = getAllSheets( books[0] )
      if len(inputSheets) == 0:
       inputSheets.append( books[0].create_sheet() )
    else:
      inputSheets.append( openSheet( books[0], targetSheets[0] ) )

    outputSheet = openSheet( books[1], targetSheets[1] )

    if len(inputSheets)>0 and outputSheet:
      for anInputSheet in inputSheets:
        sourceRows = anInputSheet.rows
        startPosX, startPosY = getLastPosition( outputSheet )
        setCells( outputSheet, startPosX, startPosY, sourceRows )

    books[1].save( args.args[1] )
