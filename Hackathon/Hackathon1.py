#ITSEMIL
#this function sends the message to the manager

def alerts(m):
    if (m!='\0'):
        print("Dear manager,you have new alert:")
        print(m)
    else:
       print(" ")

# this func recieves the meassage from the shift manager
def MessageForManager():
    print("Enter here you message:")
    mes=input()
    alerts(mes)



MessageForManager()


import xlrd
<<<<<<< HEAD
import string


def Error_page():
    exit(0)


def Log_In():
    file_loc = r'C:\Users\User\Desktop\project\Group2_Yesodot\Hack\passwarde.xlsx'

    pas_file = xlrd.open_workbook(file_loc)
    sheet = pas_file.sheet_by_index(0)
    sheet.cell_value(0, 0)
    # print(sheet.nrows)
    # print(sheet.ncols)
    flag = 0

    def check_name (flag):
        name = input('enter user name-english letters only ')
        for i in range(1, sheet.ncols+1):
            check = sheet.cell_value(i, 0)
            if check == name:
                flag=1
                Password = int(input('Enter a 6-digit password-'))
                index= i
                for j in range(2):
                    if Password == (sheet.cell_value(index, 1)):
                        worker_access = sheet.cell_value(index, 2)
                        Open_Menu(worker_access)
                        break
                    else:
                        Password = int(input('wrong password, try again'))

                if j == 2:
                    print("soory, too many tries")
                    Error_page()
                break
            else:
                continue

        return flag

    ans=check_name(0)
    if ans==0:
        for i in range (2):
            print("Name does not exist on the system, Try again")
            if check_name(0)!= 0:
                break




def Open_Menu(access):
    access_manage = 'manager'
    access_Responsible = 'shift r'
    access_worker ='worker'

    if access == access_manage:
        manager_menu()
    if access == access_Responsible:
        Responsible_menu()
    if access == access_worker:
        worker_menu()

def manager_menu():

    print('manager menu:')
    print('Select the desired action ')
    print('1- sell item')
    print('2- Issue reports')
    print('3- Cancelling a transaction\ Refund')
    print('4- Replenishment')
    print('5- Remove item inventory')
    print('6- Changes in work arrangements')
    print('7- add customer to customer club')
    print('8- remove customer from customer club')


def Responsible_menu():
    print('responsible menu:')
    print('Select the desired action ')
    print('1- sell item')
    print('2- Issue reports')
    print('3- Submit messages to the administrator')
    print('4- Submission of constraints')
    print('5- add customer to customer club')
    print('6- remove customer from customer club')

def worker_menu():
    print('worker menu:')
    print('Select the desired action ')
    print('1- sell item')
    print('2- Issue reports')
    print('3- Closing the POS')
    print('4- Submission of constraints')
    print('5- add customer to customer club')
    print('6- find customer in customer club')


Log_In()














# =======
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
>>>>>>> f8668815022f0658d33d5822fab5c24e45f1a60f
