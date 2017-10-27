## Auto Ledger

A python program to parse through a general accounting journal excel sheet and create a ledger from the data

Use:
```
$ python3 autoledger.py input_file sheet_number output_file_name
```

for example:
```
$ python3 autoledger.py journal.xlsx 1 output.xlsx
```

## Required Libraries
* [Openpyxl](https://bitbucket.org/openpyxl/openpyxl/src) by Eric Gazoni, Charlie Clark

## Todo:
* make the style look nice so no manual work has to be done
* get formulas working for running balance 
* make sure account object know what type of account they are
* ???
* profit

