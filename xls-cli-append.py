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


if __name__=="__main__":
  parser = argparse.ArgumentParser(description='Parse command line options.')
  parser.add_argument('args', nargs='*', help='Specify input.xlsx output.xlsx')
  parser.add_argument('-i', '--inputsheet', action='store', default="*", help='Specfy sheet name e.g. Sheet1, all sheet if not specified')
  parser.add_argument('-o', '--outputsheet', action='store', default="*", help='Specfy sheet name e.g. Sheet1, 1st found sheet if not specified')

  args = parser.parse_args()

  if len(args.args)==2:
    targetSheets = []
    targetSheets.append( args.inputsheet )
    targetSheets.append( args.outputsheet )

    books=[]
    sheets=[]
    i = 0
    for aFile in args.args:
      aBook = openBook( aFile)
      books.append( aBook )
      sheets.append( openSheet( aBook, targetSheets[i] ) )
      i = i + 1

    i = 0
    for aSheet in sheets:
      if aSheet:
        print( "sheet[" + str(i) + "]:" + aSheet.title )
      i = i + 1

    if len(books)==2:
      books[1].save( args.args[1] )
