import xlrd
import xlsxwriter
# import pandas as pd
# import numpy as np

# saving location file
location = r'C:\Users\User\Desktop\project-store\Group2_Yesodot\workOnExcel\membership.xlsx'
# variable that present the file we will work with
members_file = xlrd.open_workbook(location)
# the specific sheet we need from the file:
sheet = members_file.sheet_by_index(0)

row_list = []
members_list = []


# copy the file to list:
for i in range(0, sheet.nrows):
    row_list = sheet.row_values(i)
    members_list.append(row_list)

print (members_list)

# add new costumer to members list:
first_name = input('First name:')
last_name = input('Last name:')
id = input('ID: ')
address = input('Address:')
birthday = input('Date of birth: ')
phone = input('Phone number: ')
members_list.append([first_name, last_name, id, address, birthday, phone])

# update excel file by new members list:
workbook = xlsxwriter.Workbook('membership.xlsx')
worksheet = workbook.add_worksheet('membership')

i = 0
for i in range(len(members_list)):
    for j in range (6):
        worksheet.write(i, j, members_list[i][j])


workbook.close()
print('The customer was successfully added to the customer club')