import xlrd
'''
find a custumer in the members club
'''
def find_custumer(name, last):
    file_loc = r'C:\Users\micha\Desktop\קוד מיכל\Group2_Yesodot\Hackathon\membership.xlsx'
    workbook = xlrd.open_workbook(file_loc)
    worksheet = workbook.sheet_by_index(0)

    worksheet.cell_value(0, 0)

    for i in range(worksheet.nrows):
        if worksheet.cell_value(i, 0) == name and worksheet.cell_value(i, 1) == last:
            return('Exist!')
    return('Doesnt Exist!')
name, last=input('enter the first name: '), input('enter the last name: ')
print(find_custumer(name, last))