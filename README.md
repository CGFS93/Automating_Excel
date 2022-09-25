# Automating_Excel

## Overview of Automating Excel:
Automating Excel workbook using Python to correct values and create a bar chart. Using openpyxl, a for-loop, and structuring a fuction, python enabled the ability to automate tedious work on excel. 


### Purpose:
The purpose of this project is to show case pythons abilities to automate corrections to workbooks on excel. And, automation of visualization via python onto excel as a bar chart.


## Resources
**Data Source:** Transaction.xlsx

**Software:** Python3, openpyxl.

# Results:

### Correcting Values

<p align="center">
  <img src="https://user-images.githubusercontent.com/98966503/192129093-d96ed84e-2626-4719-9ff9-2066d1f73ff2.png">
</p>

-The code below is the for-loop responsable for the prices correction on the wb transactions.xlsx.<br>
-
for row in range(2, sheet.max_row + 1):
        cell = sheet.cell(row, 3)
        corrected_price = cell.value * 0.9
        correctted_price_cell = sheet.cell(row, 4)
        correctted_price_cell.value = corrected_price.<br>
-The code begins by calling the range of the cell and rows within the ws. The reange starts from cell 2 since the cell a1 is the titles.
3 values are created, cell is the row the code will input to corrected price, corrected_price_cell is a new row where the corrected values will be inserted. The final value is equating the revised value to the new row.<br>
<br><br>

### Adding BarChart

<p align="center">
  <img src="https://user-images.githubusercontent.com/98966503/192129094-c8a6c3de-663c-4f1b-a7f5-0add1921eb82.png">
</p>

-# Creating Value reference
    values = Reference(sheet,
                       min_row=2,
                       max_row=sheet.max_row,
                       min_col=4,
                       max_col=4)

    # Creating BarChart on Workbook
    chart = BarChart()
    chart.add_data(values)
    sheet.add_chart(chart, "e2") .<br>
<br><br>

### Refactoring Code for Automation
</p>
 Dependencies
import openpyxl as xl
from openpyxl.chart import BarChart, Reference
<br>
 Creating a Funtion, Loading Workbook
<br><br>
def process_workbook(filename):
    wb = xl.load_workbook(filename)
<br>
     Creating Values
    sheet = wb["Sheet1"]
<br>
     For loop correcting Prices
    for row in range(2, sheet.max_row + 1):
        cell = sheet.cell(row, 3)
        corrected_price = cell.value * 0.9
        correctted_price_cell = sheet.cell(row, 4)
        correctted_price_cell.value = corrected_price
<br>
     Creating Value reference
    values = Reference(sheet,
                       min_row=2,
                       max_row=sheet.max_row,
                       min_col=4,
                       max_col=4)
<br>
     Creating BarChart on Workbook
    chart = BarChart()
    chart.add_data(values)
    sheet.add_chart(chart, "e2")
<br>
     Saving New Workbook
    wb.save(filename)
<br><br>



# Summary:

