# Notes
# ws.row(1).height_mismatch = True - use if row height doesn't match font height
# ALWAYS RUN ON PO_Script.py ... NOT PO_Styles !!

import xlwt
import xlrd
import datetime, xlrd
import math

# Reading/copying values

workbook = xlrd.open_workbook('input.xlsx')
sheet = workbook.sheet_by_index(0)

read_MPO_Number = sheet.cell_value(2,1)
read_MPO_Number = int(read_MPO_Number) # Takes out the decimal place of PO#
read_WB_or_SF = sheet.cell_value(2,2)
read_Date = sheet.cell_value(2,3) # This is just a number

# Converting the read_Date number to the full date

read_Date_as_datetime = datetime.datetime(*xlrd.xldate_as_tuple(read_Date, workbook.datemode))

# Converting the read_Date full date to a string
read_Date_as_datetime_string = str(read_Date_as_datetime)

# Deleting bits of string
Date_Month = read_Date_as_datetime_string.replace("00:00:00","")
Date_Month = Date_Month.replace("2018-0","") # Months 1-9

#Author Note, make a string replacement here if you want to get rid of "0#" as the day for the output

Date_Month = Date_Month.replace("2018-","") # Months 10-12

# Here is print(Date_Month, "Date_Month")
Date_Month = Date_Month.replace("-","/") # slash


# Reading Data into Arrays

# The full column information, until it hits nothing
read_PO_data = [sheet.cell_value(row, 1) for row in range(sheet.nrows)]
read_product_data = [sheet.cell_value(row, 2) for row in range(sheet.nrows)]
read_qty_data = [sheet.cell_value(row, 3) for row in range(sheet.nrows)]
read_price_data = [sheet.cell_value(row, 4) for row in range(sheet.nrows)]
read_gauge_data = [sheet.cell_value(row, 7) for row in range(sheet.nrows)]
read_material_data = [sheet.cell_value(row, 8) for row in range(sheet.nrows)]
read_size_data = [sheet.cell_value(row, 9) for row in range(sheet.nrows)]


# Deleting unecessary infromation in array for all data
read_PO_data = read_PO_data[5:]
read_product_data = read_product_data[5:]
read_qty_data = read_qty_data[5:]
read_price_data = read_price_data[5:]
read_gauge_data = read_gauge_data[5:]
read_material_data = read_material_data[5:]
read_size_data = read_size_data[5:]

# Getting rid of blanks or "None" in lists

read_material_data = list(filter(None, read_material_data))
read_size_data = list(filter(None, read_size_data))
read_gauge_data = list(filter(None, read_gauge_data))


# Number of Products, which yields how long the product range will iterate
data_range = len(read_PO_data)

# Number of Sheet Variations, which yields how long the material range will iterate
data_range_material = len(read_material_data)




# Putting specific product lines in a list (ex: GP100, UAD200, etc.)

read_product_line = []


for x in range(data_range):
    mylist_new_var = read_product_data[x].split(" ",1)[0]
    x = x + 1
    testing = read_product_line.append(mylist_new_var)

print("\n", read_product_line, "read_product_line")


# Putting specific product heights in a list (ex: 24, 12, 32, etc.)

read_specific_height = []

for x in range(data_range):
    mylist_new_var = float(read_product_data[x].split("x",1)[-1])
    x = x + 1
    testing = read_specific_height.append(mylist_new_var)

print(read_specific_height, "read_specific_height")


## Hinge Caluclation

Hinge_Calculation = []


for x in range(data_range):
    if read_product_line[x] == 'GP100':
        if read_specific_height[x] <= 8: # x1 4" Hinge
            hinges = (int(read_qty_data[x] * 1))
            testing = Hinge_Calculation.append("x" + str(hinges) + ' - 4" Hinges')
        elif read_specific_height[x] <= 20: # x2 4" Hinge
            hinges = (int(read_qty_data[x] * 2))
            testing = Hinge_Calculation.append("x" + str(hinges) + ' - 4" Hinges')
        elif read_specific_height[x] <= 24: # x3 4" Hinge
            hinges = (int(read_qty_data[x] * 3))
            testing = Hinge_Calculation.append("x" + str(hinges) + ' - 4" Hinges')
        else: # Piano Hinge
            piano_hinge_height = read_specific_height[x] - .625
#            print(piano_hinge_height, "piano_hinge_height")
            piano_hinge_per_96hinge = int(96 / piano_hinge_height)
#            print(piano_hinge_per_96hinge, 'piano_hinge_per_96hinge')
            piano_hinge_needed = math.ceil((int(read_qty_data[x])) / piano_hinge_per_96hinge)
#            print(piano_hinge_needed, "piano_hinge_needed") 
            testing = Hinge_Calculation.append(int(piano_hinge_needed))

print(Hinge_Calculation, "Hinge_Calculation")














# PO Price Calculation - QTY and Price

# Multiplying both arrays to create one array   
data_times_price_array = [read_qty_data*read_price_data for read_qty_data,read_price_data in zip(read_qty_data,read_price_data)]


# Summing the combined array, the total price
read_total_price = sum(data_times_price_array)

# Taking out the decimal for the above value
# The Total Price (Exact)
read_total_price = int(read_total_price)

# The Total Price (Use for PO)
read_total_price_ceil = int(math.ceil(read_total_price / 100))*100



# Title of PO, creating the variables into strings

po_header = "MPO#" + str(read_MPO_Number) + " - " + str(read_WB_or_SF) + " - " + str(Date_Month) + "- $" + str(read_total_price_ceil)

# This is just an idea: read_data = [sheet.cell_value(5, col) for col in range(sheet.ncols)]
























# WRITING THE PO

wb = xlwt.Workbook()
ws = wb.add_sheet('PO To Floor')

xlwt.add_palette_colour("custom_gray", 0x21)
wb.set_colour_RGB(0x21, 231, 230, 230)

import PO_Styles

ws.col(0).width = 935 #A 
ws.col(1).width = 3100 #B 
ws.col(2).width = 7150 #C 
ws.col(3).width = 3250 #D 
ws.col(4).width = 3220 #E 
ws.col(5).width = 2700 #F 
ws.col(6).width = 2720 #G
ws.col(7).width = 3020 #H


# set height of cells
# **** Thinking about using nrows instead of a number....
for x in range(0,100):
    ws.row(x).height_mismatch = True
    ws.row(x).height = 287



# Creating the Header of the PO

# ws.write(y,x)
ws.write(1, 2, po_header, PO_Styles.style_header)
# ws.merge(start-y,final-y,start-x,final-x)
ws.merge(1, 2, 2, 6, PO_Styles.style_header)


# Product Headers (Don't Change)

ws.write(4, 1, 'P.O.#', PO_Styles.style_gray_fill_left)
ws.write(4, 2, 'Product', PO_Styles.style_gray_fill_left)
ws.write(4, 3, 'QTY Needed', PO_Styles.style_gray_fill_left)
ws.write(4, 4, 'QTY Made', PO_Styles.style_gray_fill_left)
ws.write(4, 5, 'Supervisor Sign Off	', PO_Styles.style_gray_fill_left)
ws.merge(4, 4, 5, 6, PO_Styles.style_gray_fill_left)

# Product Data (Does Change)


row_x = 5 # row_x is the row variable 

# x will be repeated based on data_range (number of products listed)
# for numbers, an int will be used to get rid of deicmal
for x in range(data_range):  
    a = int(read_PO_data[x])
    b = read_product_data[x]
    c = read_qty_data[x]
    ws.write(row_x, 1, a, PO_Styles.style_normal_small_center)
    ws.write(row_x, 2, b, PO_Styles.style_normal_left)
    ws.write(row_x, 3, c, PO_Styles.style_normal_center)
    ws.write(row_x, 4, "", PO_Styles.style_normal_center)
    ws.write(row_x, 5, "", PO_Styles.style_normal_center)
    ws.merge(row_x, row_x, 5, 6, PO_Styles.style_normal_center)
    row_x = row_x + 1
else: # if loop ends, add another +2 to row_x
    row_x = row_x + 2

# Material Quantities Main Header
ws.write(row_x, 1, "MATERIAL QUANTITIES", PO_Styles.style_header_normal)
ws.merge(row_x, row_x, 1, 7, PO_Styles.style_header_normal)

row_x = row_x + 2

# Material Quantities Headers

ws.write(row_x, 1, 'Material', PO_Styles.style_gray_fill_left)
ws.write(row_x, 2, 'Size', PO_Styles.style_gray_fill_left)
ws.write(row_x, 3, 'Gauge', PO_Styles.style_gray_fill_left)
ws.write(row_x, 4, 'QTY Expected', PO_Styles.style_gray_fill_left)
ws.merge(row_x, row_x, 4, 5, PO_Styles.style_gray_fill_left)
ws.write(row_x, 6, 'QTY Used', PO_Styles.style_gray_fill_left)
ws.merge(row_x, row_x, 6, 7, PO_Styles.style_gray_fill_left)

row_x = row_x + 1

# Material Quantities Data

for x in range(data_range_material):
    a = read_material_data[x]
    b = read_size_data[x]
    c = read_gauge_data[x]
    ws.write(row_x, 1, a, PO_Styles.style_normal_left)
    ws.write(row_x, 2, b, PO_Styles.style_normal_left)
    ws.write(row_x, 3, c, PO_Styles.style_normal_left)

    ws.write(row_x, 4, '', PO_Styles.style_normal_left)
    ws.merge(row_x, row_x, 4, 5, PO_Styles.style_normal_left)
    ws.write(row_x, 6, '', PO_Styles.style_normal_left)
    ws.merge(row_x, row_x, 6, 7, PO_Styles.style_normal_left)

    row_x = row_x + 1
else: # if loop ends, add another +1 to row_x
    row_x = row_x + 1

# Hinge Main Header
ws.write(row_x, 1, "HINGE QTY (Raw 8ft Hinge/Rod)", PO_Styles.style_header_normal)
ws.merge(row_x, row_x, 1, 7, PO_Styles.style_header_normal)

row_x = row_x + 2

ws.write(row_x, 1, 'Product', PO_Styles.style_gray_fill_left)
ws.merge(row_x, row_x, 1, 3, PO_Styles.style_gray_fill_left)
ws.write(row_x, 4, 'QTY Expected', PO_Styles.style_gray_fill_left)
ws.merge(row_x, row_x, 4, 5, PO_Styles.style_gray_fill_left)
ws.write(row_x, 6, 'QTY Used', PO_Styles.style_gray_fill_left)
ws.merge(row_x, row_x, 6, 7, PO_Styles.style_gray_fill_left)

row_x = row_x + 1

for x in range(data_range):  
    a = read_product_data[x]
    b = Hinge_Calculation[x]
    ws.write(row_x, 1, a, PO_Styles.style_normal_left)
    ws.merge(row_x, row_x, 1, 3, PO_Styles.style_normal_left)
    ws.write(row_x, 4, b, PO_Styles.style_normal_center)
    ws.merge(row_x, row_x, 4, 5, PO_Styles.style_normal_center)
    ws.write(row_x, 6, '', PO_Styles.style_normal_center)
    ws.merge(row_x, row_x, 6, 7, PO_Styles.style_normal_center)
    row_x = row_x + 1
else: # if loop ends, add another +2 to row_x
    row_x = row_x + 2
    ws.write(row_x, 1, "Copy/paste from PO Master List", PO_Styles.style_normal_left_highlight)
    ws.merge(row_x, row_x, 1, 7, PO_Styles.style_normal_left_highlight)




# Start reading from "PO Master Parts List" inserted in input

'''
workbook = xlrd.open_workbook('PO Master Parts List.xlsx')
sheet = workbook.sheet_by_index(0)

GP_beginning = 567
GP_end = 574

read_GP100_data = [sheet.cell_value(row, 2) for row in range(GP_beginning,GP_end)]

print("\n", read_GP100_data, "read_GP100_data")
'''

# Cutting

# Original Excel

# Work

# PO Information (to copy to Quality Control Sheet)

wb.save('PO_Script.xls')