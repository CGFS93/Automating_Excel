# Automating_Excel

## Overview of Automating Excel:
Automating Excel workbook with Python to correct values and create  bar charts. Using openpyxl, Python for-loop, and structuring a fuction, python enables automation of tedious excel work. 


### Purpose:
Showcase pythons abilities to automate maintanace corrections to excel workbooks. Allowing visualization via python onto excel as a bar chart.


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
  
  The Refactorization of the code allows new wb inputs for automation.
    
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
<br> 



# Summary:
Python is an interpreted, object-oriented, high-level programming language with dynamic semantics. Its high-level built in data structures, combined with dynamic typing and dynamic binding, make it very attractive for Rapid Application Development, as well as for use as a scripting or glue language to connect existing components together.
openpyxl is a Python library to read/write Excel 2010 xlsx/xlsm/xltx/xltm files. 
  for loops are used when you have a block of code which you want to repeat a fixed number of times. The for-loop is always used in combination with an iterable object, like a list or a range. The Python for statement iterates over the members of a sequence in order, executing the block each time. Contrast the for statement with the ''while'' loop, used when a condition needs to be checked each iteration or to repeat a block of code forever. In bar charts values are plotted as either horizontal bars or vertical columns. 
  Implementing these tools in conjunction will grant the ability to automate repetitive tasks. 
