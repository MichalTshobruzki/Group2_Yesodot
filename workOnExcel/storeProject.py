import xlrd
import xlsxwriter
import time
from time import gmtime, strftime



def arrival_to_work(access):
    name = input('enter your first name: ')
    last = input('enter your last name: ')
    date_now = time.localtime()
    presence_list = []
    row_list = []

    presence_loc = r'C:\Users\User\Desktop\project-store\Group2_Yesodot\workOnExcel\presence1.xlsx'
    presence_loc = r'C:\Users\User\Desktop\project-store\Group2_Yesodot\workOnExcel\presence2.xlsx'

    presence_file = xlrd.open_workbook(presence_loc)
    sheet = presence_file.sheet_by_index(0)

    for i in range(0, sheet.nrows):
        row_list = sheet.row_values(i)
        if i > 0:
            row_list[0] = int(row_list[0])
        presence_list.append(row_list)

    month = ('{0}'.format(date_now[1]))
    day = ('{0}'.format(date_now[2]))
    weekDay = ('{0}'.format(date_now[6] + 2))
    presence_list.append(
        [sheet.nrows, name, last, strftime("%a, %d %b %Y %H:%M:%S", time.localtime()), month, day, weekDay])
    presence_workbook = xlsxwriter.Workbook('presence2.xlsx')
    worksheet = presence_workbook.add_worksheet('presence')

    for i in range(len(presence_list)):
        for j in range(len(presence_list[i])):
            worksheet.write(i, j, presence_list[i][j])
    presence_workbook.close()
    Open_Menu(access)

    # name = input('enter your first name: ')
    # last = input('enter your last name: ')
    # presence_list = []
    # row_list = []
    # presence_loc = r'C:\Users\micha\Desktop\project\Group2_Yesodot\workOnExcel\presence2.xlsx'
    # presence_file = xlrd.open_workbook(presence_loc)
    # sheet = presence_file.sheet_by_index(0)
    #
    # for i in range(0, sheet.nrows):
    #     row_list = sheet.row_values(i)
    #     if i > 0:
    #         row_list[0] = int(row_list[0])
    #     presence_list.append(row_list)
    #
    # presence_list.append([sheet.nrows, name, last, strftime("%a, %d %b %Y %H:%M:%S", time.localtime())])
    # presence_workbook = xlsxwriter.Workbook('presence1.xlsx')
    # worksheet = presence_workbook.add_worksheet('presence')
    #
    # for i in range(len(presence_list)):
    #     for j in range(len(presence_list[i])):
    #         worksheet.write(i, j, presence_list[i][j])
    # presence_workbook.close()
    # Open_Menu(access)



def departure(access):
    name = input('enter your first name: ')
    last = input('enter your last name: ')

    presence_list = []
    row_list = []
    presence_loc = r'C:\Users\User\Desktop\project-store\Group2_Yesodot\workOnExcel\presence1.xlsx'
    presence_loc = r'C:\Users\User\Desktop\project-store\Group2_Yesodot\workOnExcel\presence2.xlsx'

    presence_file = xlrd.open_workbook(presence_loc)
    sheet = presence_file.sheet_by_index(0)

    for i in range(0, sheet.nrows):
        row_list = sheet.row_values(i)
        if i > 0:
            row_list[0] = int(row_list[0])
        presence_list.append(row_list)

    for i in range(1, len(presence_list)):
        if presence_list[i][1] == name and presence_list[i][2] == last:
            presence_list[i][7] = strftime("%a, %d %b %Y %H:%M:%S", time.localtime())
            worker = i
    worker_arrival = presence_list[worker][3]
    worker_departure = presence_list[worker][7]

    #calculates the second of each time
    arrival_time = (int(worker_arrival[17]) * 10 + int(worker_arrival[18])) * 3600 + \
                   (int(worker_arrival[20]) * 10 + int(worker_arrival[21])) * 60 + \
                   (int(worker_arrival[23]) * 10 + int(worker_arrival[24]))
    departure_time = (int(worker_departure[17]) * 10 + int(worker_departure[18])) * 3600 + \
                     (int(worker_departure[20]) * 10 + int(worker_departure[21])) * 60 + \
                     (int(worker_departure[23]) * 10 + int(worker_departure[24]))
    delta = departure_time - arrival_time
    presence_list[worker][8] = delta

    presence_workbook = xlsxwriter.Workbook('presence2.xlsx')
    worksheet = presence_workbook.add_worksheet('presence')

    for i in range(len(presence_list)):
        for j in range(len(presence_list[i])):
            worksheet.write(i, j, presence_list[i][j])
    presence_workbook.close()
    Open_Menu(access)

#     name = input('enter your first name: ')
#     last = input('enter your last name: ')
#     presence_list = []
#     row_list = []
#     presence_loc = r'C:\Users\micha\Desktop\project\Group2_Yesodot\workOnExcel\presence2.xlsx'
#     presence_file = xlrd.open_workbook(presence_loc)
#     sheet = presence_file.sheet_by_index(0)
#
#     for i in range(0, sheet.nrows):
#         row_list = sheet.row_values(i)
#         if i > 0:
#             row_list[0] = int(row_list[0])
#         presence_list.append(row_list)
#     for i in range(1, len(presence_list)):
#         if presence_list[i][1] == name and presence_list[i][2] == last:
#             presence_list[i][4] = strftime("%a, %d %b %Y %H:%M:%S", time.localtime())
#             worker = i
#     worker_arrival = presence_list[worker][3]
#     worker_departure = presence_list[worker][4]
#
# ######calculates the second of each time
#     arrival_time = (int(worker_arrival[17])*10 + int(worker_arrival[18]))*3600 +\
#                    (int(worker_arrival[20])*10 + int(worker_arrival[21]))*60 +\
#                    (int(worker_arrival[23])*10 + int(worker_arrival[24]))
#     departure_time = (int(worker_departure[17]) * 10 + int(worker_departure[18])) * 3600 +\
#                      (int(worker_departure[20]) * 10 + int(worker_departure[21])) * 60 +\
#                      (int(worker_departure[23]) * 10 + int(worker_departure[24]))
#     delta = departure_time - arrival_time
#     presence_list[worker][5] = delta
#
#     presence_workbook = xlsxwriter.Workbook('presence1.xlsx')
#     worksheet = presence_workbook.add_worksheet('presence')
#
#     for i in range(len(presence_list)):
#         for j in range(len(presence_list[i])):
#             worksheet.write(i, j, presence_list[i][j])
#     presence_workbook.close()
 #   Open_Menu(access)


'''this func recieves the meassage from the shift manager'''
def MessageForManager(access):
    messages_list = []
    row_list = []

    message_loc = r'C:\Users\User\Desktop\project-store\Group2_Yesodot\workOnExcel\messages.xlsx'

    message_loc = r'C:\Users\User\Desktop\project-store\Group2_Yesodot\workOnExcel\messages.xlsx'

    message_file = xlrd.open_workbook(message_loc)
    sheet = message_file.sheet_by_index(0)
    for i in range(0, sheet.nrows):
        row_list = sheet.row_values(i)
        if i > 0:
            row_list[0] = int(row_list[0])
        messages_list.append(row_list)
    print("Enter here you message: ")
    messageFromWorker = input()
    messages_list.append([sheet.nrows, messageFromWorker])

    message_workbook = xlsxwriter.Workbook('messages.xlsx')
    worksheet = message_workbook.add_worksheet('messages')

    for i in range(len(messages_list)):
        worksheet.write(i, 0, messages_list[i][0])
        worksheet.write(i, 1, messages_list[i][1])
    message_workbook.close()

    Open_Menu(access)



'''find a custumer in the members club'''
def find_custumer(access):
    name, last = input('enter the first name: '), input('enter the last name: ')

    file_loc = r'C:\Users\User\Desktop\project-store\Group2_Yesodot\workOnExcel\membership.xlsx'
    file_loc = r'C:\Users\User\Desktop\project-store\Group2_Yesodot\workOnExcel\membership.xlsx'

    workbook = xlrd.open_workbook(file_loc)
    worksheet = workbook.sheet_by_index(0)
    worksheet.cell_value(0, 0)
    for i in range(worksheet.nrows):
        if worksheet.cell_value(i, 0) == name and worksheet.cell_value(i, 1) == last:
            return True
    return False



'''add worker Constraints'''
def add_worker_Constraints(access):
    constraints_list = []
    row_list = []

    constraints_loc = r'C:\Users\User\Desktop\project-store\Group2_Yesodot\workOnExcel\Constraints1.xlsx'
    constraints_loc = r'C:\Users\User\Desktop\project-store\Group2_Yesodot\workOnExcel\Constraints1.xlsx'

    constraints_file = xlrd.open_workbook(constraints_loc)
    amount_sheets = constraints_file.nsheets
    for i in range(amount_sheets):
        sheet = constraints_file.sheet_by_index(i)
        sheet_list = [sheet.name]
        for j in range(sheet.nrows):
            row_list = sheet.row_values(j)
            sheet_list.append(row_list)
        constraints_list.append(sheet_list)

    workbook = xlsxwriter.Workbook('Constraints1.xlsx')

    for i in range(len(constraints_list)):  #runs on 2 sheets - michal and shir
        worksheet = workbook.add_worksheet(constraints_list[i][0])   #constraints_list[i][0]- sheet name
        for j in range(1, len(constraints_list[i])):  # number of rows in sheet
            for k in range(len(constraints_list[i][j])):
                worksheet.write(j-1, k, constraints_list[i][j][k])

    name = input('Enter your name: ')

    worksheet = workbook.add_worksheet(name)

    worksheet.write(1, 0, 'Morning')
    worksheet.write(2, 0, 'Evening')
    worksheet.write(0, 1, 'Sunday')
    worksheet.write(0, 2, 'Monday')
    worksheet.write(0, 3, 'Tuesday')
    worksheet.write(0, 4, 'Wednesday')
    worksheet.write(0, 5, 'Thursday')
    worksheet.write(0, 6, 'Friday')
    worksheet.write(0, 7, 'Saturday')

    constraint1_day, constraint1_shift = input('enter your first constraint-> day: '), \
                                         input('enter your first constraint-> shift: ')
    constraint2_day, constraint2_shift = input('enter your second constraint-> day: '), \
                                         input('enter your second constraint-> shift: ')
    if constraint1_day == 'Sunday':
        constraint1_day = 1
    if constraint1_day == 'Monday':
        constraint1_day = 2
    if constraint1_day == 'Tuesday':
        constraint1_day = 3
    if constraint1_day == 'Wednesday':
        constraint1_day = 4
    if constraint1_day == 'Thursday':
        constraint1_day = 5
    if constraint1_day == 'Friday':
        constraint1_day = 6
    if constraint1_day == 'Saturday':
        constraint1_day = 7
    if constraint1_shift == 'Morning':
        constraint1_shift = 1
    if constraint1_shift == 'Evening':
        constraint1_shift = 2

    if constraint2_day == 'Sunday':
        constraint2_day = 1
    if constraint2_day == 'Monday':
        constraint2_day = 2
    if constraint2_day == 'Tuesday':
        constraint2_day = 3
    if constraint2_day == 'Wednesday':
        constraint2_day = 4
    if constraint2_day == 'Thursday':
        constraint2_day = 5
    if constraint2_day == 'Friday':
        constraint2_day = 6
    if constraint2_day == 'Saturday':
        constraint2_day = 7
    if constraint2_shift == 'Morning':
        constraint2_shift = 1
    if constraint2_shift == 'Evening':
        constraint2_shift = 2

    worksheet.write(constraint1_shift, constraint1_day, 'NO')
    worksheet.write(constraint2_shift, constraint2_day, 'NO')
    workbook.close()
    Open_Menu(access)

'''
order new stock
'''
def add_new_inventory(access):
    inventory_list = []
    row_list = []
    users_list = []

    inventory_loc = r'C:\Users\User\Desktop\project-store\Group2_Yesodot\workOnExcel\inventory.xlsx'

    #inventory_loc = r'C:\Users\micha\Desktop\project\Group2_Yesodot\workOnExcel\inventory.xlsx'

    inventory_file = xlrd.open_workbook(inventory_loc)
    sheet = inventory_file.sheet_by_index(0)

    for i in range(sheet.nrows):
        row_list = sheet.row_values(i)
        inventory_list.append(row_list)

    for i in range(1, len(inventory_list)):
        num = inventory_list[i][0]
        inventory_list[i][0] = int(num)
        num = inventory_list[i][3]
        inventory_list[i][3] = int(num)

    print(len(inventory_list))
    users_list.append(len(inventory_list))
    input_string = input('enter the name of the product: ')
    users_list.append(input_string)
    input_string = input('enter the size of the product: ')
    users_list.append(input_string)
    input_string = input('enter the amount of the product: ')
    users_list.append(input_string)
    input_string = input('enter the color of the product: ')
    users_list.append(input_string)
    inventory_list.append(users_list)

    inventory_workbook = xlsxwriter.Workbook('inventory.xlsx')
    worksheet = inventory_workbook.add_worksheet('inventory1')
    for i in range(len(inventory_list)):
        print(inventory_list[i])
        for j in range(len(inventory_list[i])):
            worksheet.write(i, j, inventory_list[i][j])

    inventory_workbook.close()
    Open_Menu(access)


def Add_custumer (access):
    # saving location file
    location = r'C:\Users\User\Desktop\project-store\Group2_Yesodot\workOnExcel\membership.xlsx'
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
        for j in range(6):
            worksheet.write(i, j, members_list[i][j])

    workbook.close()
    print('The customer was successfully added to the customer club')
    Open_Menu(access)


def Delete_customer (access):
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

    # get costumer id:
    ID = input('please enter costumer id: ')


    # find the index of the id in membership list:
    index = None
    for i in range(len(members_list)):
        for j in range(6):
            if members_list[i][j] == ID:
                index = i

    if index == None:
        print("id doesn't exists in membership club")


    # update excel file by new members list without the removed costumer:
    workbook = xlsxwriter.Workbook('membership.xlsx')
    worksheet = workbook.add_worksheet('membership')

    for i in range(len(members_list)):
        if i != index:
            for j in range(6):
                worksheet.write(i, j, members_list[i][j])

    workbook.close()
    print('The customer was successfully removed from customer club')
    Open_Menu(access)


def GetPrice(product_code):
    inventory_list = []
    inventory_loc = r'C:\Users\User\Desktop\project-store\Group2_Yesodot\workOnExcel\inventory.xlsx'
    inventory_file = xlrd.open_workbook(inventory_loc)
    sheet = inventory_file.sheet_by_index(0)
    price_index = 0
    for i in range(sheet.nrows):
            row_list = sheet.row_values(i)
            inventory_list.append(row_list)
    for i in range(1, len(inventory_list)):
        num = inventory_list[i][0]
        inventory_list[i][0] = int(num)
        num = inventory_list[i][3]
        inventory_list[i][3] = int(num)

    for i in range(len(inventory_list)):
        for j in range(len(inventory_list[i])):
                if product_code == inventory_list[i][j]:
                    price_index=inventory_list[i][j+5]
                    return price_index

def check_validation_of_product_code(code):
    inventory_list = []
    inventory_loc = r'C:\Users\User\Desktop\project-store\Group2_Yesodot\workOnExcel\inventory.xlsx'
    inventory_file = xlrd.open_workbook(inventory_loc)
    sheet = inventory_file.sheet_by_index(0)

    # copy the file to list:
    for i in range(0, sheet.nrows):
        row_list = sheet.row_values(i)
        inventory_list.append(row_list)

    #find the product in list:
    for i in range(len(inventory_list)):
        if inventory_list[i][0] == code:
            return True

    return False

def update_stock_with_sale(code_product, amount):
    inventory_list = []
    inventory_loc = r'C:\Users\User\Desktop\project-store\Group2_Yesodot\workOnExcel\inventory.xlsx'
    inventory_file = xlrd.open_workbook(inventory_loc)
    sheet = inventory_file.sheet_by_index(0)
    amount_index = 0
    for i in range(sheet.nrows):
            row_list = sheet.row_values(i)
            inventory_list.append(row_list)
    for i in range(1, len(inventory_list)):
        num = inventory_list[i][0]
        inventory_list[i][0] = int(num)
        num = inventory_list[i][3]
        inventory_list[i][3] = int(num)

    for i in range(len(inventory_list)):
        for j in range(len(inventory_list[i])):
            if code_product == inventory_list[i][j]:
                amount_index = inventory_list[i][j+3] - amount
                inventory_list[i][j+3] = amount_index

    inventory_workbook = xlsxwriter.Workbook('inventory.xlsx')
    worksheet = inventory_workbook.add_worksheet('inventory1')
    for i in range(len(inventory_list)):
        print(inventory_list[i])
        for j in range(len(inventory_list[i])):
            worksheet.write(i, j, inventory_list[i][j])
    print(inventory_list[i])
    inventory_workbook.close()

    # ============ function for create recipt and save her at recipects data=======================

def make_recipect(date, price):
        # saving location file
        location = r'C:\Users\User\Desktop\project-store\Group2_Yesodot\workOnExcel\recipects.xlsx'
        # variable that present the file we will work with
        recipects_file = xlrd.open_workbook(location)
        # the specific sheet we need from the file:
        sheet = recipects_file.sheet_by_index(0)

        row_list = []
        recipects_list = []
        item = []

        # copy the file to list:
        for i in range(0, sheet.nrows):
            row_list = sheet.row_values(i)
            recipects_list.append(row_list)

        # add new recipect to list:
        recipect_number = sheet.nrows
        item = [recipect_number, date, price]
        recipects_list.append(item)

        # update excel file by copy the update list:
        workbook = xlsxwriter.Workbook('recipects.xlsx')
        worksheet = workbook.add_worksheet('recipects')

        for i in range(len(recipects_list)):
            for j in range(3):
                worksheet.write(i, j, recipects_list[i][j])

#        workbook.close()
        print('Sale completed successfully')
        print('\n\n\n')

# def get_sales_report():
#         print('*****Sales Report For Manager:*****')
#         sales_loc = r'C:\Users\User\Desktop\project-store\Group2_Yesodot\workOnExcel.xlsx'
#         sales_file = xlrd.open_workbook(sales_loc)
#         sheet = sales_file.sheet_by_index(0)
#         sales_list = []
#         temp_list = []
#         for i in range(1, sheet.nrows):
#             row_list = sheet.row_values(i)
#             temp_list.append(row_list)
#
#         for i in range(0, len(temp_list)):
#             temp_list[i][1] = int(temp_list[i][1])
#             temp_list[i][3] = int(temp_list[i][3])
#             # row_list[4] = int(row_list[4])
#         sales_list = temp_list
#
#         print('**Date**   **Code Product**   **Name**    **Amount**   **Price**')
#         for j in range(0, len(sales_list)):
#             print('{0}       {1}           {2}            {3}         {4}'.format(sales_list[j][0], sales_list[j][1],
#                                                                                   sales_list[j][2], sales_list[j][3],
#                                                                                  sales_list[j][4]))

def GetName(product_code):
    inventory_list = []
    inventory_loc = r'C:\Users\User\Desktop\project-store\Group2_Yesodot\workOnExcel\inventory.xlsx'
    inventory_file = xlrd.open_workbook(inventory_loc)
    sheet = inventory_file.sheet_by_index(0)
    name_index = 0
    for i in range(sheet.nrows):
            row_list = sheet.row_values(i)
            inventory_list.append(row_list)
    for i in range(1, len(inventory_list)):
        num = inventory_list[i][0]
        inventory_list[i][0] = int(num)
        num = inventory_list[i][3]
        inventory_list[i][3] = int(num)

    for i in range(len(inventory_list)):
        for j in range(len(inventory_list[i])):
                if product_code == inventory_list[i][j]:
                    name_index = inventory_list[i][j+1]
                    return name_index



def sell_items(access):
    assumption = 0.15
    #================================= Account Execution ===========================

    total_price = 0
    item_list = []
    item = []

    flag = 1
    while flag == 1:
        #date:
        date = strftime("%d %b %Y", time.localtime())
        product_code = int(input('Enter product code:'))

        # check if product code exists:
        while (check_validation_of_product_code(product_code) == False):
            product_code = int(input("product code isn't exists,try again"))

        # product name:
        product_name = GetName(product_code)
        # Quantity:
        Quantity = int(input('Enter quantity:'))
        # price of product:
        price = GetPrice(product_code)

        item = [product_code,product_name, Quantity, price]

        item_list.append(item)

        # if there's more items:
        flag = int(input('for add more items press 1, else enter 0'))

    # ========================== Remove item from the list ===================================
    flag = 1
    while flag == 1:
        #show items list:
        print('************** products list is:**************\n')
        for i in range(len(item_list)):
            print('{0}) {1}'.format(i+1, item_list[i]))
        flag = int(input('To delete items from list press 1, for continue press 0'))
        if flag == 1:
            index = int(input('Enter index of item you want to remove'))
            cnt = 0
            del(item_list[index-1])
            print('item removed')

    # =========================== calculate total price ===================================
    total_price = 0
    for i in range(len(item_list)):
        item_amount= item_list[i][2]
        item_price= item_list[i][3]
        total_price += (item_price * item_amount)

    #============================= update sell of items====================================
    # for i in range(len(item_list)):
    #     update_sales(item_list[i])


    # ============================ print recipect =========================================
    print('\n\n****costumer recipect****')
    print('Date:{0}'.format(date))
    for i in range (len(item_list)):
        for j in range(1,4):
            print(item_list[i][j], end=" ")


    print('total price is:{0}₪'.format(total_price))
    print('Assumption is: {0}%' . format(assumption))
    total_price= total_price-(total_price*assumption)
    total_price= round(total_price,2)
    print('update price: {0}₪'.format(total_price))
    tax= total_price*0.18
    tax=round(tax,2)
    print('Tax=18%: {0}₪'.format(tax))
    total_price= total_price + tax
    print('===========================')
    total_price= total_price= round(total_price,2)
    print('Sum is:{0}₪'. format(total_price))

    # ===================================== make recipect =======================================
    make_recipect(date, total_price)

    Open_Menu(access)


def Open_Menu(access):
    access_manage = 'manager'
    access_Responsible = 'shift r'
    access_worker = 'worker'

    if access == access_manage:
        manager_menu(access)
    if access == access_Responsible:
        Responsible_menu(access)
    if access == access_worker:
        worker_menu(access)






def manager_menu(access):

    file_loc = r'C:\Users\User\Desktop\project-store\Group2_Yesodot\workOnExcel\messages.xlsx'

    file_loc = r'C:\Users\User\Desktop\project-store\Group2_Yesodot\workOnExcel\messages.xlsx'

    workbook = xlrd.open_workbook(file_loc)
    worksheet = workbook.sheet_by_index(0)
    print('-----------------------------------------------')
    print("*****Dear manager,you have new alert*****")
    print('*****{0}*****'.format(worksheet.cell_value(worksheet.nrows-1, 1)))
    print('manager menu:')
    print('Select the desired action ')
    print('1- sell item')
    print('2- Issue reports')
    print('3- Cancelling a transaction\ Refund')
    print('4- Order new stock')
    print('5- Remove item inventory')
    print('6- Changes in work arrangements')
    print('7- add customer to customer club')
    print('8- remove customer from customer club')
    print('-----------------------------------------------')

    choice = input('your choice: ')
    if choice == '1':
        sell_items(access)
    if choice == '4':
        add_new_inventory(access)
    if choice == '7':
        Add_custumer(access)
    if choice == '8':
        Delete_customer(access)


def Responsible_menu(access):
    print('-----------------------------------------------')
    print('responsible menu:')
    print('Select the desired action ')
    print('1- sell item')
    print('2- Issue reports')
    print('3- Submit messages to the administrator')
    print('4- Submission of constraints')
    print('5- add customer to customer club')
    print('6- remove customer from customer club')
    print('-----------------------------------------------')

    choice = input('your choice: ')
    if choice == '1':
        sell_items(access)
    if choice == '4':
        print('1- submission of constrains')
        print('2- Viewing constraints')
        print('-----------------------------------------------')
        choice = input()
        if choice == '1':
            add_worker_Constraints(access)
    if choice == '3':
        MessageForManager(access)
    if choice == '5':
        Add_custumer(access)
    if choice == '6':
        Delete_customer(access)


def worker_menu(access):
    print('-----------------------------------------------')
    print('worker menu:')
    print('Select the desired action ')
    print('1- sell item')
    print('2- Issue reports')
    print('3- Closing the POS')
    print('4- Submission of constraints')
    print('5- add customer to customer club')
    print('6- find customer in customer club')
    print('7- Entry to work')
    print('8- Departing from work')
    print('-----------------------------------------------')
    choice = input()
    if choice == '1':
        sell_items(access)
    if choice == '4':
        print('1- submission of constrains')
        print('2- Viewing constraints')
        print('-----------------------------------------------')
        choice = input()
        if choice == '1':
            add_worker_Constraints(access)
    if choice == '5':
        Add_custumer(access)
    if choice == '6':
        if find_custumer(access):
            print('Costumer exists in membership club')
        else:
            print("Costumer doesn't exists in membership club")
            ans = input('would you like to add the costumer to membership club?')
            if ans == 'yes':
                Add_custumer(access)
            else:
                Open_Menu(access)
    if choice == '7':
        arrival_to_work(access)
    if choice == '8':
        departure(access)


def Error_page():
    exit(0)

def Log_In():

    file_loc = r'C:\Users\User\Desktop\project-store\Group2_Yesodot\workOnExcel\passwarde.xlsx'
    file_loc = r'C:\Users\User\Desktop\project-store\Group2_Yesodot\workOnExcel\passwarde.xlsx'


    pas_file = xlrd.open_workbook(file_loc)
    sheet = pas_file.sheet_by_index(0)
    sheet.cell_value(0, 0)

    flag = 0

    def check_name (flag):
        name = input('enter user name-english letters only: ')
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
                    print("sorry, too many tries")
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

Log_In()





