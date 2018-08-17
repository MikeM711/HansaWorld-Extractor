label = "helloasdfasdfasdf asdfasdf asdfadfasdf dsfasdf asdfasdfasdfasdfasdf werqwescasdfa ewradfs"

ws.write(1, 2, label,HW_Styles.style_normal_center_HW)

ws.col(2).width = int(arial10.fitwidth(label)) 



'''''''

read_MPO_Number = sheet.cell_value(2,1)
read_MPO_Number = int(read_MPO_Number) # Takes out the decimal place of PO#
read_WB_or_SF = sheet.cell_value(2,2)
read_Date = sheet.cell_value(2,3) # This is just a number

