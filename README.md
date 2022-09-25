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

      # For loop correcting Prices
        for row in range(2, sheet.max_row + 1):
          cell = sheet.cell(row, 3)
          corrected_price = cell.value * 0.9
          correctted_price_cell = sheet.cell(row, 4)
          correctted_price_cell.value = corrected_price


### Adding BarChart

<p align="center">
  <img src="https://user-images.githubusercontent.com/98966503/192129094-c8a6c3de-663c-4f1b-a7f5-0add1921eb82.png">


    # Creating BarChart on Workbook
     chart = BarChart()
     chart.add_data(values)
     sheet.add_chart(chart, "e2") .


### Refactoring Code for Automation
    
    # Dependencies
    import openpyxl as xl
    from openpyxl.chart import BarChart, Reference

    # Creating a Funtion, Loading Workbook


    def process_workbook(filename):
    wb = xl.load_workbook(filename)

    # Creating Values
    sheet = wb["Sheet1"]

    # For loop correcting Prices
    for row in range(2, sheet.max_row + 1):
        cell = sheet.cell(row, 3)
        corrected_price = cell.value * 0.9
        correctted_price_cell = sheet.cell(row, 4)
        correctted_price_cell.value = corrected_price

    # Creating Value reference
    values = Reference(sheet,
                       min_row=2,
                       max_row=sheet.max_row,
                       min_col=4,
                       max_col=4)

    # Creating BarChart on Workbook
    chart = BarChart()
    chart.add_data(values)
    sheet.add_chart(chart, "e2")

    # Saving New Workbook
    wb.save(filename)
<br><br>



# Summary:
The Example used in this project holds very little data and could have easily been fixed individually by hand. But, this example is ment to show case pythons versatility. The code for this project can still be used for a larger data set with very little refactoring.  
