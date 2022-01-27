# How to use

```
% python3 xls-cli.py --help                         
usage: xls-cli.py [-h] [-j] [-c] [-m] [-s SHEET] [args ...]

Parse command line options.

positional arguments:
  args                  Specify xlsx files e.g. book1.xlsx book2.xls

optional arguments:
  -h, --help            show this help message and exit
  -j, --json            Output as json
  -c, --csv             Output as csv
  -m, --merge           Output table as merged
  -s SHEET, --sheet SHEET
                        Specfy sheet name e.g. Sheet1
```

## output xlsx data as csv (default)

```
% python3 xls-cli.py sample.xlsx
```

### output all of sheets as merged as csv

```
% python3 xls-cli.py sample.xlsx --merge
```

### output specified sheet as csv

```
% python3 xls-cli.py sample.xlsx --sheet="Sheet1" --csv
```

only output the sheet of "Sheet1" as csv


## output xlsx data as json

```
% python3 xls-cli.py sample.xlsx --json
```


# Setup

```
$ pip install openpyxl
```