'''  
This python file extracts HansaWorld Excel sheets into organized information specific to AG Laser Technology workflow
'''

'''
How I would improve this project if I were to do it again:
- MORE CLASSES
# Create classes for better organization, readability and encapsulation of variables
# Try to hide variables to prevent any unwanted access to variables in global

- Have better variable names
# Keep them short & sweet, managing long variables = frustrating

- Break this file into smaller files or modules
# Maybe have my "writing the PO" part of the code in its own seperate file!

- Create more loops to condense code
# There are a few areas where I can use a loop to condense similar lines of code
# Generally - A lot of copy/paste = there's probably a better way to write it

What I do like
- The comments are helpful
- The algorithms I came up to solve this with were really cool

'''

import xlwt
import xlrd
import datetime, xlrd
import math
import arial10
import time

''' Reading Excel File'''

hardware_spec_list = ["WB-9","WB-20", "WB-13C"] # List of WB Hardware

''' Notes for the user to check '''

print("\n\tNote: All Workbooks must have a .xlsx extension")
print("\tNote: the current hardware/special list is: ", hardware_spec_list)
print("\t\tIf extra hardware or special (ie: WB-100 or 'Painted Orange') is included,")
print("\t\tit will be erased when combined\n")

''' In case the user forgets to read the hardware/special list '''

while True:
    aware = input("Are you aware of the hardware/special list above? Enter 'yes' to continue: ")
    if aware == 'yes':
        break
    else:
        print("\tPlease enter 'yes' to confirm awareness")

workbook_to_extract = input('What workbook would you like to extract? ')
print("")

start_time = time.time()

workbook = xlrd.open_workbook(workbook_to_extract + '.xlsx')
sheet = workbook.sheet_by_index(0)

''' 
myDataDict stores the combined data.  Whenever there is a repeated product, myDataDict will not add another key-value pair, but rather, will add the additional QTY, price and PO#
'''

myDataDict = {}

'''Below are lists to hold singular data'''

mySingularProduct = [] # Product Code/ID
mySingularDescription = []
mySingularQTY = []
mySingularPrice = []
mySingularPO = []
mySingularRow = []

active = False
n = 1
i = 0

PO_Number = int(sheet.cell_value(1,0))

def exception_error():
    ''' 
    If PO_Number cannot be converted into an integer, it basically means that the excel sheet is complete.
    '''
    try:
        if sheet.cell_value(i-1,0) == 'OUT':
            PO_Number = int(sheet.cell_value(i,0))
        # logic to load the page. If it is successful, it will not go to except.
        return True
    except ValueError:
        print("COMPLETE")
        # will come to this clause when page will throw error.
        return False

while True:
    i = i + 1

    ''' 
    Can we read the next PO_Number? 
    If we can't we will break out of the while loop
    '''
    if exception_error() == False:
        break

    ''' 
    If the Column B doesn't have a value, chances are it's just a price, not a product 
    '''
    if sheet.cell_value(i,3) == 'NT':
        continue

    ''' If one cell in Col A is blank, continue - go back to while'''
    if sheet.cell_value(i,0) == '':
        continue

    ''' Activate dictionary and while when USD is read '''
    ''' Deactivate dictionary and while when FREIGHT or OUT is read '''
    if sheet.cell_value(i-1,0) == 'USD':
        #print("Here is USD!")
        active = True
    elif sheet.cell_value(i,0) == 'FREIGHT':
        #print("Here is FREIGHT!")
        active = False
    elif sheet.cell_value(i,0) == 'OUT':
        #print("Here is OUT!")
        active = False        
    
    ''' 
    The PO Number will be gathered from the current cell IF the current cell and cell after are both floats 
    '''
    try:
        float(sheet.cell_value(i+1,0))
        PO_Number = int(sheet.cell_value(i,0))
    except ValueError:
        # print("Not a float")
        pass   
    except IndexError: # Getting index error, break at index error
        break     

    ''' Go into while if active is activated'''

    while active:
        
        ''' Lines below are useful for debugging:
        print(sheet.cell_value(i,0))    # This prints out the product
        print("\t Value:", i+1)         # This prints out the row location
        print("\t PO:", PO_Number)      # This prints out the PO number
        '''

        product = sheet.cell_value(i,0)
        product_description = sheet.cell_value(i,2)
        product_quantity = int(sheet.cell_value(i,1))
        product_PO_num = str(PO_Number)

        mySingularRow.append(i+1) # singular value to denote row location

        ''' 
        Sometimes there won't be a product price in col [3], it goes to col [0] of next line. Below is the case of the price being under the product.
        '''
        if sheet.cell_value(i,3) == '':
            product_price = sheet.cell_value(i,1) * sheet.cell_value(i+1,0)
        else:
            product_price = sheet.cell_value(i,1) * sheet.cell_value(i,3)

        ''' 
        If the product we are dealing with is a SPECIAL.  We will add consecutive numbers so that each SPECIAL is different.  This also includes the string of SPEC as well  
        '''
        if 'SPEC' in product:
            product = product + " #" + str(n)
            n = n + 1

        mySingularPrice.append(product_price) # singular list append for price    
        mySingularQTY.append(product_quantity) # singular list append for QTY
        mySingularProduct.append(product) # singular list append for product ID
        mySingularDescription.append(product_description) # singular list append for product description

        mySingularPO.append(product_PO_num) # singular list append for PO Number

        '''
        while loop for combined data:
        If the next line is a piece of hardware (ex: WB-9, WB-20 ...) add the 'hardware name' to the product ID and product description, as well as sum the price

        while loop for singular data:
        Take the data about the hardware and append to list, just like any ordinary product.
        '''

        while sheet.cell_value(i+1,0) in hardware_spec_list:
            product = product + " - " + str(sheet.cell_value(i+1,0))
            product_description = product_description + " - " + str(sheet.cell_value(i+1,0))
            product_price = product_price + sheet.cell_value(i+1,5)

            mySingularRow.append(i+2) # relation to the previous (i+1)
            mySingularPO.append(product_PO_num)
            mySingularPrice.append(sheet.cell_value(i+1,5))  
            mySingularQTY.append(sheet.cell_value(i+1,1)) 
            mySingularProduct.append(sheet.cell_value(i+1,0))
            mySingularDescription.append(sheet.cell_value(i+1,2))

            i = i + 1

        '''
        Construct a dictionary of all products.
        Key = Product ID
        Values = (product description, product quantity, product price, PO number)
            All values are lists

        Dictionary will update keys as it reads, so it must start with a setdefault() to let us "update" the first time around.  If there is already information in the dictionary, execution will bypass setdefault().
        '''

        myDataDict.setdefault(product, ("", 0, 0,""))

        myDataDict[product] = product_description, myDataDict[product][1] + product_quantity, myDataDict[product][2] + product_price, str(myDataDict[product][3]) + " " + product_PO_num
        break  

''' 
# If you would like to print everything from myDataDict:

for key,value in myDataDict.items():
    print(key)
    print(value)
'''

''' Writing the PO below '''

wb = xlwt.Workbook()
ws = wb.add_sheet('PO To Floor')

'''Adding A custom gray color for my HW_styles'''

xlwt.add_palette_colour("custom_gray", 0x21)
wb.set_colour_RGB(0x21, 231, 230, 230)

import HW_Styles # Note to self: imports SHOULD be at the top of .py file...

ws.col(0).width = 3100 #A 
ws.col(1).width = 5100 #B 

ws.col(2).width_mismatch = True
ws.col(2).width = 10100 #C 

ws.col(3).width = 3100 #D 
ws.col(4).width = 3100 #E 
ws.col(5).width = 3100 #F 
ws.col(6).width = 3100 #G
ws.col(7).width = 3100 #H

ws.write(1, 1, 'Singular Data - "' + workbook_to_extract + '"' , HW_Styles.style_header)
ws.merge(1, 2, 1, 6, HW_Styles.style_header)

ws.row(1).height_mismatch = True
ws.row(1).height = 350
ws.row(2).height_mismatch = True
ws.row(2).height = 350

ws.write(4, 1, 'Product Code', HW_Styles.style_gray_fill_left)
ws.write(4, 2, 'Description', HW_Styles.style_gray_fill_left)
ws.write(4, 3, 'QTY', HW_Styles.style_gray_fill_left)
ws.write(4, 4, 'Price', HW_Styles.style_gray_fill_left)
ws.write(4, 5, 'PO#', HW_Styles.style_gray_fill_left)
ws.write(4, 6, 'Row#', HW_Styles.style_gray_fill_left)

row_y = 5

for x in range(len(mySingularProduct)):
    ws.write(row_y, 1, mySingularProduct[x], HW_Styles.style_normal_left)
    ws.write(row_y, 2, mySingularDescription[x], HW_Styles.style_normal_left)
    ws.write(row_y, 3, mySingularQTY[x], HW_Styles.style_normal_left_no_dec)
    ws.write(row_y, 4, mySingularPrice[x], HW_Styles.style_normal_left)
    ws.write(row_y, 5, mySingularPO[x], HW_Styles.style_normal_left_no_dec)
    ws.write(row_y, 6,mySingularRow[x], HW_Styles.style_normal_left_no_dec)

    row_y = row_y + 1

row_y = row_y + 2

ws.write(row_y, 1, 'Combined Data - "' + workbook_to_extract + '"', HW_Styles.style_header)
ws.merge(row_y, row_y + 1, 1, 6, HW_Styles.style_header)

ws.row(row_y).height_mismatch = True
ws.row(row_y).height = 350
ws.row(row_y + 1).height_mismatch = True
ws.row(row_y + 1).height = 350

row_y = row_y + 3

ws.write(row_y, 1, 'Product Code', HW_Styles.style_gray_fill_left)
ws.write(row_y, 2, 'Description', HW_Styles.style_gray_fill_left)
ws.write(row_y, 3, 'QTY', HW_Styles.style_gray_fill_left)
ws.write(row_y, 4, 'Price', HW_Styles.style_gray_fill_left)
ws.write(row_y, 5, 'PO#s that include this Product', HW_Styles.style_gray_fill_left)
ws.merge(row_y, row_y, 5, 7, HW_Styles.style_gray_fill_left)

row_y = row_y + 1

'''
Below are empty lists to solve the max characters in indexes 1,2 and 7
'''

index_product_code = []
index_description = []
index_PO_num = []

'''
The below for loop extracts all keys and values (lists) from myDataDict into Combined Results
'''

for key,value in myDataDict.items():
    print(key)
    ws.write(row_y, 1, key, HW_Styles.style_normal_left)
    ws.write(row_y, 2, value[0], HW_Styles.style_normal_left)
    ws.write(row_y, 3, value[1], HW_Styles.style_normal_left_no_dec)
    ws.write(row_y, 4, value[2], HW_Styles.style_normal_left)
    ws.write(row_y, 5, value[3], HW_Styles.style_normal_left_no_wrap)
    ws.merge(row_y, row_y, 5, 7, HW_Styles.style_normal_left_no_wrap)
    print("\t",value)

    index_product_code.append(key)
    index_description.append(value[0])
    index_PO_num.append(value[3])

    row_y = row_y + 1

''' Author's Note '''

row_y = row_y + 2
ws.write(row_y, 2, "AUTHOR'S NOTE", HW_Styles.style_bold_cent_no_wrap)
row_y = row_y + 1
ws.write(row_y, 2, "The Current Hardware/Special List is:", HW_Styles.style_plain_left_no_wrap)
row_y = row_y + 1

for hardware in hardware_spec_list:
    ws.write(row_y, 2, "• " + str(hardware), HW_Styles.style_plain_left_no_wrap)
    row_y = row_y + 1

author_note_string1 = "- If extra hardware or special (ie: WB-100 or 'Painted Orange') is included, it WILL be erased when combined"
author_note_string2 = "- Extra hardware or special must be updated into list for it to NOT be erased when combined"
row_y = row_y + 1
ws.write(row_y, 1, author_note_string1, HW_Styles.style_plain_left_no_wrap)
row_y = row_y + 1
ws.write(row_y, 1, author_note_string2, HW_Styles.style_plain_left_no_wrap)

'''
The below function finds the max number of characters for indexes: 1,2 and 7
We can use the answer to solve how long the width of the cells need to be using the arial10 module
'''

def get_longest_name(a_list):
    return max((name for name in a_list), key=len, default='')

longest_index_1 = get_longest_name(index_product_code)
longest_index_2 = get_longest_name(index_description)
longest_index_7 = get_longest_name(index_PO_num)

'''
Using the arial10 module, we will find the correct fit of inexes 1,2 and 5
'''

ws.col(1).width = int(arial10.fitwidth(longest_index_1))  
ws.col(2).width = int(arial10.fitwidth(longest_index_2)) 
a =  int(arial10.fitwidth(longest_index_7))

'''
for index 7 below, I am subtracting the 2 merged cell lengths and adding a little bit extra room that is RELATED to the initial value, because arial10 is smaller than the font I'm using
'''

ws.col(7).width = (a - 6200) + int(a*0.1)

wb.save(workbook_to_extract + ' - Extract.xls')

print("\nExtract Completed.")
print("My program took", time.time() - start_time, "to run")