# Package to read/write Excel for python
**Context**: If you search on google  `Working with Excel Files in Python`, you could see `openpyxl`, `xlsxwriter`, `xlrd`, `xlwt`
and your purpose is read/write data Excel on python

Maybe, you must take a lot of time for choosing which package. For this repo, I hope I could help you to choose what package to read/write excel.


## SUMMARY

### 1. xlsxwriter
- I feel that xlsxwriter's simple, very easy to write, good at formatting.
- I can write and work nearly normal excel (MS). 
- Odoo framework's used to make for reporting module for many version until now.

### 2. xlrd
- This package for reading data and formatting information from Excel files (xlsx, xls).

### 3. openpyxl 
- It write and make data to excel file similar to `xlsxwriter`.
- It write to existing excel file, `xlsxwriter` can't do it.
- It read excel file similar to `xlrd`.

**Finally**:
- `openpyxl` includes read/write features of 2 package (`xlsxwriter` + `xlrd`)  and I think using one package to work help us easier to handel.
- I'll choose `openpyxl` package to write/read excel file.


## xlsxwriter
- This package for writing data, formatting information in particular, charts in the Excel 2010 format 
- It supports Python 2.7, 3.4+
- Document: https://xlsxwriter.readthedocs.io.
- Significant features:
    - Full formatting (Conditional formatting, Rich multi-format strings)
    - Autofilters.
    - Merged cells.
    - Data validation and drop down lists.
    - Worksheet PNG/JPEG/BMP/WMF/EMF images.
    - Easy writing different types of data (https://xlsxwriter.readthedocs.io/tutorial03.html)
    - Memory optimization mode.
    - `write_formula`

- **Code example**:
```python
import xlsxwriter

# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook('demo.xlsx')
worksheet = workbook.add_worksheet(worksheet_name) # Defaults to Sheet1

# Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold': True})
# Add a number format for cells with money.
money = workbook.add_format({'num_format': '$#,##0'})
# Some data we want to write to the worksheet.
expenses = (
    ['Rent', 1000],
    ['Gas',   100],
    ['Food',  300],
    ['Gym',    50],
)
# Start from the first cell. Rows and columns are zero indexed.
row = 0
col = 0

# Iterate over the data and write it out row by row.
for item, cost in (expenses):
    worksheet.write(row, col,     item)
    worksheet.write(row, col + 1, cost)
    row += 1

# Write a total using a formula.
worksheet.write(row, 0, 'Total')
worksheet.write(row, 1, '=SUM(B1:B4)')

workbook.close()

```
![result](https://xlsxwriter.readthedocs.io/_images/tutorial01.png)

## xlrd
- This package for reading data and formatting information from Excel files (xlsx, xls)
- It supports Python 2.7, 3.4+
- Document: https://xlrd.readthedocs.io/en/latest/
- Significant features:
    - Handling of Unicode
    - Dates in Excel spreadsheets
    - Named references, constants, formulas, and macros
    - Formatting information in Excel Spreadsheets
    - Loading worksheets on demand
    - XML vulnerabilities and Excel files
    - API Reference
- Code example:
```python
import xlrd
book = xlrd.open_workbook("your_path_file.xls")
print("The number of worksheets is {0}".format(book.nsheets))
print("Worksheet name(s): {0}".format(book.sheet_names()))
sh = book.sheet_by_index(0)
print("{0} {1} {2}".format(sh.name, sh.nrows, sh.ncols))
print("Cell D30 is {0}".format(sh.cell_value(rowx=29, colx=3)))
for rx in range(sh.nrows):
    print(sh.row(rx))
```
- Reading excel file and returning list of dictionary example:

```python
import xlrd
COLUMN_MAPPING = {
    'Column Name A1': 'mapped_key_a1',
    'Column Name A2': 'mapped_key_a2',
}

wb = xlrd.open_workbook("your_path_file.xlsx")
ws = wb.sheet_by_index(0)

raw_headers     = [cell.value for cell in ws.row(0)]
v_fields        = [ COLUMN_MAPPING.get(h.strip()) for h in raw_headers ]
col_count       = len(raw_headers)

returning_data  = []
for ri in range(1, ws.nrows):  # ri aka row_index - we skip the 0th row which is the header rows
    values = [ ws.cell_value(rowx=ri, colx=ci) for ci in range(col_count) ]  # ci aka column_index
    rvr    = dict(zip(v_fields, values))
    returning_data.append(rvr)
return returning_data
```


## openpyxl
- This package for reading and writing Excel files 
- It supports Python 2 and 3
- Document: https://openpyxl.readthedocs.io/en/stable/index.html
- Code example for writing a excel file:
```python
from openpyxl import Workbook

wb=Workbook()
filepath="/home/ubuntu/demo.xlsx"
sheet=wb.active
data=[
    ('Id','Name','Marks'),
    (1,ABC,50),
    (2,CDE,100)
]

# append all rows
for row in data:
    sheet.append(row)

# save file
wb.save(filepath)
```

- Code example for reading a excel file:
```python
from openpyxl import load_workbook
filepath="/home/ubuntu/demo.xlsx"
wb=load_workbook(filepath)
sheet=wb.active

max_row=sheet.max_row
max_column=sheet.max_column

# iterate over all cells
for i in range(1, max_row+1):
     for j in range(1, max_column+1):
          # get particular cell value    
          cell_obj=sheet.cell(row=i, column=j)
          # print cell value     
          print(cell_obj.value, end=' | ')
```

- Code example for formatting:
```python
from openpyxl.writer.dump_worksheet import WriteOnlyCell

filepath="/home/ubuntu/demo.xlsx"
wb = Workbook(optimized_write=True)
ws = wb.create_sheet()

user_style = Style(font=Font(name='Courrier', size=36))

cell       = WriteOnlyCell(ws, value='hello')
cell.style = user_style

ws.append([cell, 3.14, None])
wb.save(filepath)

# assert user_style in wb.shared_styles
wb2 = load_workbook(filepath)
ws2 = wb2[ws.title]
assert ws2['A1'].style == user_style 
```


## INSTALL ENVIRONMENT AND RUN PYTEST FOR THESE ABOVE CODE

### 1. Install environment
- Install pyenv for mac: 
```
# Install Homebrew if it isn't already available
/usr/bin/ruby -e "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/master/install)" 
brew install pyenv 

# Add pyenv initializer to shell startup script
echo 'eval "$(pyenv init -)"' >> ~/.bash_profile 
source ~/.bash_profile # Reload your profile
pyenv versions
```

- Install pyenv for ubuntu: 
```
sudo apt update -y
sudo apt install -y make build-essential libssl-dev zlib1g-dev libbz2-dev \
  libreadline-dev libsqlite3-dev wget curl llvm libncurses5-dev libncursesw5-dev \
  xz-utils tk-dev libffi-dev

git clone https://github.com/pyenv/pyenv.git ~/.pyenv

echo 'export PYENV_ROOT="$HOME/.pyenv"'    >> ~/.bashrc
echo 'export PATH="$PYENV_ROOT/bin:$PATH"' >> ~/.bashrc
echo -e 'if command -v pyenv 1>/dev/null 2>&1; then\n  eval "$(pyenv init -)"\nfi' >> ~/.bashrc

# restart shell & aftermath check
exec "$SHELL"
pyenv --version
```

- Install python 3.6.7
```
pyenv install 3.6.7
pyenv global 3.6.7
python --version # should be 3.6.7
```

- Install pipenv
```
# install pipenv
curl https://raw.githubusercontent.com/kennethreitz/pipenv/master/get-pipenv.py | python
echo "
# pipenv setup"                                   >> ~/.bashrc
echo 'export PATH="~/.local:$PATH"'               >> ~/.bashrc
echo 'export PIPENV_VENV_IN_PROJECT=1'            >> ~/.bashrc # project's venv location will be in the project folder as .venv ref. https://pipenv.readthedocs.io/en/latest/advanced/#pipenv.environments.PIPENV_VENV_IN_PROJECT
echo 'export PIPENV_DEFAULT_PYTHON_VERSION=3.6.7' >> ~/.bashrc # project's venv location will be in the project folder as .venv ref. https://pipenv.readthedocs.io/en/latest/advanced/#pipenv.environments.PIPENV_VENV_IN_PROJECT
echo 'eval "$(pipenv --completion)"'              >> ~/.bashrc # shell auto-completion ref. https://pipenv.readthedocs.io/en/latest/advanced/#shell-completion

#restart shell & aftermath check
exec "$SHELL"
pipenv --version
```

- `cd` to `you_project_path` and run `pipenv sync` to install all dependencies of project.

### 2. Run pytest
- `cd` to `you_project_path` and run `pipenv run pytest`