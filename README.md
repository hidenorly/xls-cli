# How to use : xls-cli.py

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


# xls-cli-append.py

```
usage: xls-cli-append.py [-h] [-i INPUTSHEET] [-o OUTPUTSHEET] [-m] [-r RANGE] [-s] [args ...]

Parse command line options.

positional arguments:
  args                  Specify output.xlsx input.xlsx

optional arguments:
  -h, --help            show this help message and exit
  -i INPUTSHEET, --inputsheet INPUTSHEET
                        Specfy sheet name e.g. Sheet1
  -o OUTPUTSHEET, --outputsheet OUTPUTSHEET
                        Specfy sheet name e.g. Sheet1
  -m, --merge           Specify if you want to merge all of sheets of input book
  -r RANGE, --range RANGE
                        Specify range e.g. A1:C3 if you want to specify input range
  -s, --swap            Specify if you want to swap row and column
```

## append data to specified xls sheet from specified xls sheet

```
% python3 xls-cli-append.py output.xlsx input.xlsx
```

### all of input.xls's sheets

```
% python3 xls-cli-append.py output.xlsx input.xlsx --merge
```

### with range

```
% python3 xls-cli-append.py output.xlsx input.xlsx --range="A1:C2"
```

### column, row swapped

```
% python3 xls-cli-append.py output.xlsx input.xlsx --swap
```

### range specified all of sheets and col&row swapped

```
% python3 xls-cli-append.py output.xlsx input.xlsx --range="A1:C2" --swap --merge
```


## append data to specified xls sheet from specified csv file

```
% python3 xls-cli-append.py sample.csv sample2.xlsx
```

### with range

```
% python3 xls-cli-append.py sample.csv sample2.xlsx --range="A1:C2"
```

### column, row swapped

```
% python3 xls-cli-append.py sample.csv sample2.xlsx --swap
```

### column, row swapped with ranged

```
% python3 xls-cli-append.py sample.csv sample2.xlsx --swap --range="A1:C2"
```



# Setup

```
$ pip install openpyxl
```