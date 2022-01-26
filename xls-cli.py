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


if __name__=="__main__":
  parser = argparse.ArgumentParser(description='Parse command line options.')
  parser.add_argument('args', nargs='*', help='Specify xlsx files e.g. book1.xlsx book2.xls')
  parser.add_argument('-j', '--json', action='store_true', default=False, help='Output as json')
  parser.add_argument('-c', '--csv', action='store_true', default=False, help='Output as csv')
  parser.add_argument('-m', '--merge', action='store_true', default=False, help='Output table as merged')
  parser.add_argument('-s', '--sheet', action='store', default="*", help='Specfy sheet name e.g. Sheet1')

  args = parser.parse_args()

  if not args.json and not args.csv:
    args.csv=True

  if args.json and args.merge:
    print("[")
  for aFile in args.args:
    if os.path.exists( aFile ):
      workBook = xl.load_workbook( aFile, read_only=True, keep_vba=False, data_only=True, keep_links=False )
      for aSheet in workBook:
        if args.sheet == "*" or aSheet.title == args.sheet:
          if args.json and not args.merge:
            print("[")
          dumpSheet(aSheet, args)
          if args.json and not args.merge:
            print("]")
    if not args.merge:
      print("")
  if args.json and args.merge:
    print("]")
