import xlrd
import xlsxwriter
import time
# import DateTime
from time import gmtime, strftime
from datetime import date
import random
from tabulate import tabulate


def get_total_price_of_recipect(rec_num):
    # ============================== get lists of recipects =====================================
    # saving location file
    location = r'C:\Users\micha\PycharmProjects\yesodotFinish\recipects.xlsx'
    # variable that present the file we will work with
    recipects_file = xlrd.open_workbook(location)
    # the specific sheet we need from the file:
    sheet = recipects_file.sheet_by_index(0)

    row_list = []
    recipects_list = []

    # copy the file to list:
    for i in range(0, sheet.nrows):
        row_list = sheet.row_values(i)
        recipects_list.append(row_list)

    total_price = 0
    # search for recipect number:
    for i in range(len(recipects_list)):
        if recipects_list[i][0] == rec_num:
            total_price += recipects_list[i][2]

    return total_price


def update_stock_with_cancellation(items_list):
    # saving location file
    location = r'C:\Users\micha\PycharmProjects\yesodotFinish\inventory.xlsx'
    # variable that present the file we will work with
    inventory_file = xlrd.open_workbook(location)
    # the specific sheet we need from the file:
    sheet = inventory_file.sheet_by_index(0)

    # copy the file to list
    inventory_list = []
    # list- code, name, amount, price
    for i in range(sheet.nrows):
        row_list = sheet.row_values(i)
        inventory_list.append(row_list)

    # update the inventory by update the list:
    for i in range(len(items_list)):
        product_code = items_list[i][0]
        amount = int(items_list[i][3])
        for j in range(len(inventory_list)):
            if inventory_list[j][0] == product_code:
                inventory_list[j][3] = inventory_list[j][3] + amount

    # update the file by copy the new list:
    inventory_workbook = xlsxwriter.Workbook('inventory.xlsx')
    worksheet = inventory_workbook.add_worksheet('inventory1')
    for i in range(len(inventory_list)):
        # print(inventory_list[i])
        for j in range(len(inventory_list[i])):
            worksheet.write(i, j, inventory_list[i][j])
    # print(inventory_list[i])
    inventory_workbook.close()
####################################################################################


def update_cancelled_report(data_list):
    # saving location file
    location = r'C:\Users\micha\PycharmProjects\yesodotFinish\sales.xlsx'
    # variable that present the file we will work with
    cancelled_sales_file = xlrd.open_workbook(location)
    # the specific sheet we need from the file:
    sheet1 = cancelled_sales_file.sheet_by_index(0)
    sheet2 = cancelled_sales_file.sheet_by_index(1)

    sales_list = []
    row = []
    # copy the file- sheet 1 to list:
    for i in range(sheet1.nrows):
        row = sheet1.row_values(i)
        sales_list.append(row)

    # copy the file to list:
    cancelled_sales_list = []
    row_list = []
    for i in range(0, sheet2.nrows):
        row_list = sheet2.row_values(i)
        cancelled_sales_list.append(row_list)

    cancelled_sales_list.append(data_list)

    # copy the new list to the file:
    sales_workbook = xlsxwriter.Workbook('sales.xlsx')
    workseet_sales = sales_workbook.add_worksheet('sales')
    worksheet_cancel = sales_workbook.add_worksheet('cancelled sales')

    for i in range(len(sales_list)):
        for j in range(len(sales_list[i])):
            workseet_sales.write(i, j, sales_list[i][j])

    for i in range(len(cancelled_sales_list)):
        for j in range(len(cancelled_sales_list[i])):
            worksheet_cancel.write(i, j, cancelled_sales_list[i][j])

    sales_workbook.close()
####################################################################################


def cancel_sell(access):
    # first, recognition process of recipect:
    recipect_num = int(input('Please enter recepict number: '))
    while check_recipect_number_validation(recipect_num) == False:
        recipect_num = int(input('recepict number wrong, enter again: '))

    today = str(date.today())
    recipect_date = get_recipect_date(recipect_num)

    if recipect_date == today:

        # ==================get from sales report the items in the recipct========================

        # saving location file
        location = r'C:\Users\micha\PycharmProjects\yesodotFinish\sales.xlsx'
        # variable that present the file we will work with
        sales_file = xlrd.open_workbook(location)
        # the specific sheet we need from the file:
        sheet = sales_file.sheet_by_index(0)

        row_list1 = []
        sales_list = []
        canceled_list = []
        recipect_item_list = []
        item = []

        # copy the file to list- copy sheet 1:
        for i in range(0, sheet.nrows):
            row_list1 = sheet.row_values(i)
            sales_list.append(row_list1)

        # find the relevant items by recipect number, at the end of the process i get list of items belong to the recipect
        for i in range(len(sales_list)):
            if sales_list[i][7] == recipect_num:
                code = sales_list[i][3]
                name = sales_list[i][4]
                amount = sales_list[i][5]
                price = sales_list[i][6]
                item = [code, name, amount, price]
                recipect_item_list.append(item)

        # =================== get total price we need to refund =======================
        total_sum = get_total_price_of_recipect(recipect_num)

        'The cancellation fee according to the law is 5% or 100 NIS, whichever is lower'
        check_5_Percent = 0.05 * total_sum
        Cancellation_fee = 0
        if check_5_Percent < 100:
            Cancellation_fee = check_5_Percent
        else:
            Cancellation_fee = 100

        Refund = total_sum - Cancellation_fee
        Refund = round(Refund)
        print('recipect items list is:\n')
        for i in range(len(recipect_item_list)):
            print(recipect_item_list[i])

        print('-------------------')
        print('Dear employee, the customer must receive a refund of {0} NIS'.format(Refund))

        # ====================== update inventory after items are returned ===============

        update_stock_with_cancellation(recipect_item_list)
        print('cancellation process finished successfully')
        date_now = time.localtime()
        current_year, current_month, current_day = date_now[0], date_now[1], date_now[2]
        data_list = [recipect_num, current_year, current_month, current_day, Refund]
        update_cancelled_report(data_list)

    else:
        print('You can not cancel a sell.\nCancellation of sell is valid only for one day.')

    Open_Menu(access)
####################################################################################


def clear_constraints(access):
    constraints_loc = r'C:\Users\micha\PycharmProjects\yesodotFinish\Constraints1.xlsx'
    constraints_file = xlrd.open_workbook(constraints_loc)
    sheet = constraints_file.sheet_by_index(0)
    sheet_list = []
    for i in range(sheet.nrows):
        row_list = sheet.row_values(i)
        sheet_list.append(row_list)
    workbook_constraints = xlsxwriter.Workbook('Constraints1.xlsx')
    worksheet = workbook_constraints.add_worksheet('shifts')
    for i in range(len(sheet_list)):
        for j in range(len(sheet_list[i])):  # number of rows in sheet
            worksheet.write(i, j, sheet_list[i][j])

    workbook_constraints.close()
    Open_Menu(access)
####################################################################################


def add_2_workers_to_shifts(worker1, worker2):
    constraints_loc = r'C:\Users\micha\PycharmProjects\yesodotFinish\Constraints1.xlsx'
    constraints_file = xlrd.open_workbook(constraints_loc)
    screwed_loc = r'C:\Users\micha\PycharmProjects\yesodotFinish\Screwed.xlsx'
    screwed_file = xlrd.open_workbook(screwed_loc)
    amount_sheets_constraints = constraints_file.nsheets

    constraints_list = []
    row_list = []
    screwed_list = []

    # add the sheets of constraints to list##########
    for i in range(amount_sheets_constraints):
        sheet = constraints_file.sheet_by_index(i)
        sheet_list = [sheet.name]
        for j in range(sheet.nrows):
            row_list = sheet.row_values(j)
            sheet_list.append(row_list)
        constraints_list.append(sheet_list)
    count = 0

    # change the no one can to worker
    for i in range(len(constraints_list[0])):
        for j in range(len(constraints_list[0][i])):
            if constraints_list[0][i][j] == 'no one can' and count == 1:
                constraints_list[0][i][j] = worker2
            elif constraints_list[0][i][j] == 'no one can':
                constraints_list[0][i][j] = worker1
                count += 1

    # copy the list to excel
    workbook_constraints = xlsxwriter.Workbook('Constraints1.xlsx')
    for i in range(len(constraints_list)):
        worksheet = workbook_constraints.add_worksheet(constraints_list[i][0])  # constraints_list[i][0]- sheet name
        for j in range(1, len(constraints_list[i])):  # number of rows in sheet
            for k in range(len(constraints_list[i][j])):
                worksheet.write(j - 1, k, constraints_list[i][j][k])
    workbook_constraints.close()

    # add the sheets of screwed to list##########
    for i in range(2):
        sheet = screwed_file.sheet_by_index(i)
        sheet_list = [sheet.name]
        for j in range(sheet.nrows):
            row_list = sheet.row_values(j)
            sheet_list.append(row_list)
        screwed_list.append(sheet_list)

    # add the one/two workers to screwed list
    if count == 1:
        screwed_list[0].append([screwed_file.sheet_by_index(0).nrows, worker1])
    elif count == 2:
        screwed_list[0].append([screwed_file.sheet_by_index(0).nrows, worker1])
        screwed_list[0].append([screwed_file.sheet_by_index(0).nrows + 1, worker2])

    # copy the list to excel
    workbook = xlsxwriter.Workbook('Screwed.xlsx')
    for i in range(len(screwed_list)):
        worksheet = workbook.add_worksheet(screwed_list[i][0])  # constraints_list[i][0]- sheet name
        for j in range(1, len(screwed_list[i])):  # number of rows in sheet
            for k in range(len(screwed_list[i][j])):
                worksheet.write(j - 1, k, screwed_list[i][j][k])
    workbook.close()
####################################################################################


def max_val(var):
    maximum = 0
    for i in range(len(var)):
        if var[i] > maximum:
            maximum = var[i]
    return maximum
####################################################################################


def find_2_workers_when_no_one_can():
    screwed_loc = r'C:\Users\micha\PycharmProjects\yesodotFinish\Screwed.xlsx'
    screwed_file = xlrd.open_workbook(screwed_loc)
    number_of_shifts_sheet = screwed_file.sheet_by_index(1)
    screwed_sheet = screwed_file.sheet_by_index(0)

    screwed_worker1 = screwed_sheet.cell_value(screwed_sheet.nrows - 1, 1)
    screwed_worker2 = screwed_sheet.cell_value(screwed_sheet.nrows - 2, 1)

    list_of_number_of_shifts = []
    for i in range(1, number_of_shifts_sheet.nrows):
        list_of_number_of_shifts.append(number_of_shifts_sheet.cell_value(i, 1))

    maximum_shifts = max_val(list_of_number_of_shifts)
    index_of_max = list_of_number_of_shifts.index(maximum_shifts)

    first = second = maximum_shifts
    first_worker = second_worker = number_of_shifts_sheet.cell_value(index_of_max + 1, 0)
    for i in range(len(list_of_number_of_shifts)):
        # If current element is smaller than first then update both first and second
        if list_of_number_of_shifts[i] < first and list_of_number_of_shifts[i] != screwed_worker1 and \
                list_of_number_of_shifts[i] != screwed_worker2:
            second = first
            second_worker = first_worker
            first = list_of_number_of_shifts[i]
            first_worker = number_of_shifts_sheet.cell_value(i + 1, 0)

        # If list_of_number_of_shifts[i] is in between first and second then update second
        elif list_of_number_of_shifts[i] < second and list_of_number_of_shifts[i] != first and list_of_number_of_shifts[
            i] != screwed_worker1 and list_of_number_of_shifts[i] != screwed_worker2:
            second = list_of_number_of_shifts[i]
            second_worker = number_of_shifts_sheet.cell_value(i + 1, 0)

    add_2_workers_to_shifts(first_worker, second_worker)
####################################################################################


def count_shift_for_worker():
    constraints_loc = r'C:\Users\micha\PycharmProjects\yesodotFinish\Constraints1.xlsx'
    constraints_file = xlrd.open_workbook(constraints_loc)
    # screwed_loc = r'C:\Users\micha\Desktop\project_new\Group2_Yesodot\final project\Screwed.xlsx'
    screwed_loc = r'C:\Users\micha\PycharmProjects\yesodotFinish\Screwed.xlsx'
    screwed_file = xlrd.open_workbook(screwed_loc)

    shifts_sheet = constraints_file.sheet_by_index(0)
    shiftsForWorker_sheet = screwed_file.sheet_by_index(1)
    workers_dict = {}

    for i in range(1, shiftsForWorker_sheet.nrows):
        worker = shiftsForWorker_sheet.cell_value(i, 0)
        workers_dict[worker] = 0
        for j in range(shifts_sheet.nrows):
            for k in range(shifts_sheet.ncols):
                if worker == shifts_sheet.cell_value(j, k):
                    workers_dict[worker] = workers_dict[worker] + 1
    return workers_dict
####################################################################################


def write_number_of_shifts_to_sheet():
    shifts_dict = count_shift_for_worker()
    row_list = []
    screwed_list = []
    sheet_list = []
    screwed_loc = r'C:\Users\micha\PycharmProjects\yesodotFinish\Screwed.xlsx'
    screwed_file = xlrd.open_workbook(screwed_loc)
    amount_sheets = screwed_file.nsheets

    # copy the excel to list
    for i in range(amount_sheets):
        sheet = screwed_file.sheet_by_index(i)
        sheet_list = [sheet.name]
        for j in range(sheet.nrows):
            row_list = sheet.row_values(j)
            sheet_list.append(row_list)
        screwed_list.append(sheet_list)

    # add the numbers of shifts to each worker
    for i in range(2, len(screwed_list[1])):
        screwed_list[1][i][1] = shifts_dict[screwed_list[1][i][0]]

    workbook = xlsxwriter.Workbook('Screwed.xlsx')
    for i in range(len(screwed_list)):
        worksheet = workbook.add_worksheet(screwed_list[i][0])
        for j in range(1, len(screwed_list[i])):  # number of rows in sheet
            for k in range(len(screwed_list[i][j])):
                worksheet.write(j - 1, k, screwed_list[i][j][k])
    workbook.close()
####################################################################################


# make list of constraints of shift manager
def build_list_of_constraints_of_shift_manager(name):
    shiftManager_constraints = []
    constraints_loc = r'C:\Users\micha\PycharmProjects\yesodotFinish\Constraints1.xlsx'
    constraints_file = xlrd.open_workbook(constraints_loc)
    # find the sheet of the shift manager- michal
    for i in range(constraints_file.nsheets):
        sheet = constraints_file.sheet_by_index(i)
        if sheet.name == name:
            shiftManager_constraints.append(name)
            break
    # add the constraints to a list
    for j in range(sheet.nrows):
        for k in range(sheet.ncols):
            if sheet.cell_value(j, k) == 'NO':
                shiftManager_constraints.append([j, k])  # j- the shift, k- the day
    return shiftManager_constraints
####################################################################################


def make_shifts_for_shift_manager(list_of_constraints):
    list = []
    michal_number_of_shifts = 0
    emilia_number_of_shifts = 0
    for i in range(1, 3):
        row_list = []
        for j in range(1, 8):
            if (list_of_constraints[0][1][0] == i and list_of_constraints[0][1][1] == j) or (
                    list_of_constraints[0][2][0] == i and list_of_constraints[0][2][
                1] == j):  # if michal cant work, put emilia
                row_list.append(list_of_constraints[1][0])
                emilia_number_of_shifts += 1
            elif (list_of_constraints[1][1][0] == i and list_of_constraints[1][1][1] == j) or (
                    list_of_constraints[1][2][0] == i and list_of_constraints[1][2][1] == j):
                row_list.append(list_of_constraints[0][0])
                michal_number_of_shifts += 1
            elif emilia_number_of_shifts < michal_number_of_shifts:
                row_list.append(list_of_constraints[1][0])
                emilia_number_of_shifts += 1
            else:
                row_list.append(list_of_constraints[0][0])
                michal_number_of_shifts += 1
        list.append(row_list)
    return list
####################################################################################


# the function gets list of people who can work in the shift and return list by random of two workers
def make_shift_by_random(day):
    shift = []
    if len(day) == 0:
        shift.append('no one can')
        shift.append('no one can')
        return shift
    elif len(day) == 1:
        shift.append(day[0])
        shift.append('no one can')
        return shift
    while len(shift) < 2:
        rand = random.randint(0, len(day) - 1)
        if len(shift) == 0:
            shift.append(day[rand])
        else:
            for i in range(len(shift)):
                if day[rand] != shift[i]:
                    shift.append(day[rand])
                    break
    return shift
####################################################################################


# get from build_shifts the col and row of the cell that represent the shift
# and append all the workers that can work in this shift
def build_one_shift(row, col):
    shift = []
    constraints_loc = r'C:\Users\micha\PycharmProjects\yesodotFinish\Constraints1.xlsx'
    constraints_file = xlrd.open_workbook(constraints_loc)
    amount_sheets = constraints_file.nsheets - 2
    for i in range(1, amount_sheets):
        sheet = constraints_file.sheet_by_index(i)
        if sheet.cell_value(row, col) != 'NO':
            shift.append(sheet.name)
    return shift
####################################################################################


def build_shifts(access):
    constraints_list = []
    row_list = []
    constraints_loc = r'C:\Users\micha\PycharmProjects\yesodotFinish\Constraints1.xlsx'
    constraints_file = xlrd.open_workbook(constraints_loc)
    amount_sheets = constraints_file.nsheets
    # add the sheets of constraints to list##########
    if amount_sheets < 9:
        print('You still can not build shifts, not all employees have submitted constraints')
        Open_Menu(access)
    for i in range(1, amount_sheets):
        sheet = constraints_file.sheet_by_index(i)
        sheet_list = [sheet.name]
        for j in range(sheet.nrows):
            row_list = sheet.row_values(j)
            sheet_list.append(row_list)
        constraints_list.append(sheet_list)

    # every workers that can work in every shift:
    # send to function the cell of the shift that he will check with each worker
    # and add him to the list if he can work
    sunday_morning = build_one_shift(1, 1)
    sunday_evening = build_one_shift(2, 1)
    monday_morning = build_one_shift(1, 2)
    monday_evening = build_one_shift(2, 2)
    tuesday_morning = build_one_shift(1, 3)
    tuesday_evening = build_one_shift(2, 3)
    wednesday_morning = build_one_shift(1, 4)
    wednesday_evening = build_one_shift(2, 4)
    thursday_morning = build_one_shift(1, 5)
    thursday_evening = build_one_shift(2, 5)
    friday_morning = build_one_shift(1, 6)
    saturday_evening = build_one_shift(2, 7)

    # list of workers in the shift
    sunday_morning_shift = make_shift_by_random(sunday_morning)
    sunday_evening_shift = make_shift_by_random(sunday_evening)
    monday_morning_shift = make_shift_by_random(monday_morning)
    monday_evening_shift = make_shift_by_random(monday_evening)
    tuesday_morning_shift = make_shift_by_random(tuesday_morning)
    tuesday_evening_shift = make_shift_by_random(tuesday_evening)
    wednesday_morning_shift = make_shift_by_random(wednesday_morning)
    wednesday_evening_shift = make_shift_by_random(wednesday_evening)
    thursday_morning_shift = make_shift_by_random(thursday_morning)
    thursday_evening_shift = make_shift_by_random(thursday_evening)
    friday_morning_shift = make_shift_by_random(friday_morning)
    saturday_evening_shift = make_shift_by_random(saturday_evening)

    workbook = xlsxwriter.Workbook('Constraints1.xlsx')
    worksheet = workbook.add_worksheet('shifts')

    # input the table of shifts
    worksheet.write(1, 0, 'Morning')
    worksheet.write(2, 0, 'Morning')
    worksheet.write(3, 0, 'shift manager')
    worksheet.write(4, 0, 'Evening')
    worksheet.write(5, 0, 'Evening')
    worksheet.write(6, 0, 'shift manager')
    worksheet.write(0, 1, 'Sunday')
    worksheet.write(0, 2, 'Monday')
    worksheet.write(0, 3, 'Tuesday')
    worksheet.write(0, 4, 'Wednesday')
    worksheet.write(0, 5, 'Thursday')
    worksheet.write(0, 6, 'Friday')
    worksheet.write(0, 7, 'Saturday')

    # add the workers to each shift in the sheet
    for i in range(2):
        worksheet.write(i + 1, 1, sunday_morning_shift[i])
    for i in range(2):
        worksheet.write(i + 4, 1, sunday_evening_shift[i])
    for i in range(2):
        worksheet.write(i + 1, 2, monday_morning_shift[i])
    for i in range(2):
        worksheet.write(i + 4, 2, monday_evening_shift[i])
    for i in range(2):
        worksheet.write(i + 1, 3, tuesday_morning_shift[i])
    for i in range(2):
        worksheet.write(i + 4, 3, tuesday_evening_shift[i])
    for i in range(2):
        worksheet.write(i + 1, 4, wednesday_morning_shift[i])
    for i in range(2):
        worksheet.write(i + 4, 4, wednesday_evening_shift[i])
    for i in range(2):
        worksheet.write(i + 1, 5, thursday_morning_shift[i])
    for i in range(2):
        worksheet.write(i + 4, 5, thursday_evening_shift[i])
    for i in range(2):
        worksheet.write(i + 1, 6, friday_morning_shift[i])
    for i in range(2):
        worksheet.write(i + 4, 7, saturday_evening_shift[i])

    # add the shifts of the shifts managers:
    shift_managers_constraints = []
    shift_managers_constraints.append(build_list_of_constraints_of_shift_manager('michal'))
    shift_managers_constraints.append(build_list_of_constraints_of_shift_manager('emilia'))

    list_of_shifts_for_sManager = (make_shifts_for_shift_manager(shift_managers_constraints))
    for i in range(1, 7):
        worksheet.write(3, i, list_of_shifts_for_sManager[0][i - 1])
    for i in range(1, 8):
        if i == 6:
            continue
        elif i == 7:
            worksheet.write(6, i, list_of_shifts_for_sManager[1][i - 2])
        else:
            worksheet.write(6, i, list_of_shifts_for_sManager[1][i - 1])

    # copy the sheets of constraints
    for i in range(len(constraints_list)):
        worksheet = workbook.add_worksheet(constraints_list[i][0])  # constraints_list[i][0]- sheet name
        for j in range(1, len(constraints_list[i])):  # number of rows in sheet
            for k in range(len(constraints_list[i][j])):
                worksheet.write(j - 1, k, constraints_list[i][j][k])
    workbook.close()
    find_2_workers_when_no_one_can()
    write_number_of_shifts_to_sheet()
    Open_Menu(access)
####################################################################################


# manager can put cell and the name he want to change for working
def make_changes_in_shifts(access):
    constraints_loc = r'C:\Users\micha\PycharmProjects\yesodotFinish\Constraints1.xlsx'
    constraints_file = xlrd.open_workbook(constraints_loc)
    sheet = constraints_file.sheet_by_index(0)
    row_list = [' ', '0', '1', '2', '3', '4', '5', '6', '7']
    sheet_list = []
    sheet_list.append(row_list)
    amount_sheets = constraints_file.nsheets
    constraints_list = []

    # add the shifts to list
    for i in range(sheet.nrows):
        row_list = [i]
        row_list.extend(sheet.row_values(i))
        sheet_list.append(row_list)
    print(tabulate(sheet_list, tablefmt="fancy_grid"))

    # add the sheets of constraints to list##########
    for i in range(amount_sheets):
        sheet = constraints_file.sheet_by_index(i)
        sheet_list = [sheet.name]
        for j in range(sheet.nrows):
            row_list = sheet.row_values(j)
            sheet_list.append(row_list)
        constraints_list.append(sheet_list)

    print('enter row and col of the cell you want to change, for end entet- done')

    passworde_loc = r'C:\Users\micha\PycharmProjects\yesodotFinish\passwarde.xlsx'
    passworde_file = xlrd.open_workbook(passworde_loc)
    sheet_passworde = passworde_file.sheet_by_index(0)

    while True:
        flag = 0
        row = input('row- ')
        if row == 'done' or row == 'Done':
            break
        row = int(row)
        if row > 6 or row < 1:
            print('try again')
            continue
        col = int(input('col- '))
        if col > 7 or col < 1:
            print('try again')
            continue
        worker = str(input('worker- '))
        for i in range(sheet_passworde.nrows):
            if worker == sheet_passworde.cell_value(i, 0):
                flag = 1
                break
        if flag == 1:
            constraints_list[0][row + 1][col] = worker
        else:
            print('the worker does not exist')

    workbook = xlsxwriter.Workbook('Constraints1.xlsx')
    for i in range(len(constraints_list)):
        worksheet = workbook.add_worksheet(constraints_list[i][0])  # constraints_list[i][0]- sheet name
        for j in range(1, len(constraints_list[i])):  # number of rows in sheet
            for k in range(len(constraints_list[i][j])):
                worksheet.write(j - 1, k, constraints_list[i][j][k])
    workbook.close()
    Open_Menu(access)
####################################################################################


# shows to the screen table of shifts
def shifts_report(access):
    constraints_loc = r'C:\Users\micha\PycharmProjects\yesodotFinish\Constraints1.xlsx'
    constraints_file = xlrd.open_workbook(constraints_loc)
    sheet = constraints_file.sheet_by_index(0)
    row_list = []
    sheet_list = []

    ##add the shifts to list
    for i in range(sheet.nrows):
        row_list = sheet.row_values(i)
        sheet_list.append(row_list)

    print(tabulate(sheet_list, tablefmt="fancy_grid"))
    Open_Menu(access)
####################################################################################


# returns the total amount of the sales in the current day
def Daily_Money_amount(year, month, day):
    sales_loc = r'C:\Users\micha\PycharmProjects\yesodotFinish\recipects.xlsx'
    sales_file = xlrd.open_workbook(sales_loc)
    sheet = sales_file.sheet_by_index(0)
    temp_list = []
    total_money_amount = 0
    # this part is for coping the existing data thats in the file already
    for i in range(1, sheet.nrows):
        row_list = sheet.row_values(i)
        temp_list.append(row_list)

    # this part summs all the money from all the sells  that were made the current day.
    for i in range(1, len(temp_list)):
        if date == temp_list[i][1]:
            total_money_amount = total_money_amount + temp_list[i][2]
    return total_money_amount


####################################################################################


#  this function writes the daily money amount with the current date into EOD excel file and returns the daily amount of money
def EOD_report(access):
    EOD_loc = r'C:\Users\micha\PycharmProjects\yesodotFinish\EOD.xlsx'
    EOD_file = xlrd.open_workbook(EOD_loc)
    sheet = EOD_file.sheet_by_index(0)
    date1 = str(date.today())
    date_now = strftime("%d %b %Y", time.localtime())
    EOD_list = []
    t_list = []
    # this part is for coping the existing data thats in the file already
    for j in range(sheet.nrows):
        row_list = sheet.row_values(j)
        EOD_list.append(row_list)
    # this part makes a new list with the following data: the current date and the total amont of money(using the above function) and appending it with the rest of the information.
    total = Daily_Money_amount(date1)
    t_list.append(date_now)
    t_list.append(total)
    EOD_list.append(t_list)
    print('The total money for this day is:', end=' ')
    print(total)
    # this part writes all the lists to the excel file.
    EOD_workbook = xlsxwriter.Workbook('EOD.xlsx')
    worksheet = EOD_workbook.add_worksheet('EOD01')
    for i in range(len(EOD_list)):
        for j in range(len(EOD_list[i])):
            worksheet.write(i, j, EOD_list[i][j])

    EOD_workbook.close()
    Open_Menu(access)
####################################################################################


# this function checks if the money that was counted in the EOD report matches the money that was counted by the worker
def Closing_The_Register(access):
    flag = 0
    while flag == 0:
        date1 = str(date.today())
        money_from_register = input('Enter the amount you counted:')
        if float(money_from_register) == Daily_Money_amount(date1):
            print('All Valid.')
        else:
            print('Please inform the shift manager that the money you entered does not match the EOD report.'
                  'Would you to close the register anyway?[yes/no]')
            # only if the worker enters yes the function will close and the menu will be shown again.
            answer = input()
            if answer == 'yes':
                flag == 1
                Open_Menu(access)
            elif answer == 'no':
                flag == 0
            else:  # if the worker wont enter the right answer than the whole process will start again.
                print('not valid input.')
                flag == 0
    Open_Menu(access)
####################################################################################


# print table of all the inventory
def get_inventory_report(access):
    inventory_loc = r'C:\Users\micha\PycharmProjects\yesodotFinish\Inventory.xlsx'
    inventory_file = xlrd.open_workbook(inventory_loc)
    sheet = inventory_file.sheet_by_index(0)
    row_list = []
    inventory_list = []
    for i in range(0, sheet.nrows):
        row_list = sheet.row_values(i)
        inventory_list.append(row_list)

    # print table report
    print(tabulate(inventory_list, tablefmt="fancy_grid"))
    Open_Menu(access)
####################################################################################


# Prints all names and hours of all employees
def get_manager_presence_report(access):
    month = int(input('enter the number of month of the report you want: '))
    while True:
        if month > 0 and month < 13:
            break
        month = int(input('wrong choice, try again: '))
    print('*****  Presence Report For Manager  *****')
    presence_list = []
    presence_list.append(['worker', 'arrival', 'departure', 'total'])
    presence_loc = r'C:\Users\micha\PycharmProjects\yesodotFinish\presence1.xlsx'
    presence_file = xlrd.open_workbook(presence_loc)
    sheet = presence_file.sheet_by_index(0)
    for i in range(0, sheet.nrows):
        if sheet.cell_value(i, 7) == '':
            continue
        if sheet.cell_value(i, 4) == str(month):
            total_sec = int(sheet.cell_value(i, 8))
            sec = total_sec % 60
            total_sec = total_sec // 60
            mint = total_sec % 60
            hour = total_sec // 60
            row_list = [sheet.cell_value(i, 1), sheet.cell_value(i, 3), sheet.cell_value(i, 7),
                        ('%02d:%02d:%02d' % (hour, mint, sec))]
            presence_list.append(row_list)
    # print table report
    print(tabulate(presence_list, tablefmt="fancy_grid"))
    Open_Menu(access)
####################################################################################


# print presence report for the worker who asked for
def get_monthly_presence_report(access):
    name = input('enter your name: ')
    now = time.localtime()
    month = now[1]
    print('          *****  Presence Report  *****')
    presence_list = []
    presence_list.append(['arrival time', 'departure time', 'total work time'])

    presence_loc = r'C:\Users\micha\PycharmProjects\yesodotFinish\presence1.xlsx'
    presence_file = xlrd.open_workbook(presence_loc)
    sheet = presence_file.sheet_by_index(0)

    for i in range(0, sheet.nrows):
        if sheet.cell_value(i, 7) == '':
            continue
        if sheet.cell_value(i, 1) == name and sheet.cell_value(i, 4) == str(month):
            total_sec = sheet.cell_value(i, 8)
            sec = total_sec % 60
            total_sec = total_sec // 60
            mint = total_sec % 60
            hour = total_sec // 60
            row_list = [sheet.cell_value(i, 3), sheet.cell_value(i, 7), ('%02d:%02d:%02d' % (hour, mint, sec))]
            presence_list.append(row_list)
    print(tabulate(presence_list, tablefmt="fancy_grid"))

    Open_Menu(access)


####################################################################################


# returning some product to the Storage
def return_inventory(access):
    flag = 0
    while flag == 0:
        inventory_list = []
        updated_stock_list = []
        inventory_loc = r'C:\Users\micha\PycharmProjects\yesodotFinish\inventory.xlsx'
        inventory_file = xlrd.open_workbook(inventory_loc)
        sheet = inventory_file.sheet_by_index(0)
        k, l = 0, 0
        # coping the existing data to a new list
        for i in range(sheet.nrows):
            row_list = sheet.row_values(i)
            inventory_list.append(row_list)
        # converting float numbers to integers(except from the price column)
        for i in range(1, len(inventory_list)):
            num = inventory_list[i][0]
            inventory_list[i][0] = int(num)
            num = inventory_list[i][3]
            inventory_list[i][3] = int(num)
        # prints the current inventory
        print(tabulate(inventory_list, tablefmt="fancy_grid"))
        picks = []
        # input from the manager which item to send back
        pick = int(input("enter product code of item you want to send back?:"))
        try:  # if the input is not valid exeption is trown.
            if pick > (len(inventory_list) - 1):
                raise ValueError('Not valid choice.')
        except ValueError:
            print('Value Error.Pick again')
        # searching for the index of the item that the manager entered
        picks.append(pick)
        for i in range(len(inventory_list)):
            for j in range(len(inventory_list[i])):
                for k in range(len(picks)):
                    if picks[k] == inventory_list[i][j]:
                        l = i + 1
        # making a new list without the deleted item
        for i in range(0, l - 1, 1):
            updated_stock_list.append(inventory_list[i])
        for j in range(l, len(inventory_list), 1):
            updated_stock_list.append(inventory_list[j])
        try:
            pick2 = input("would you like to remove another item from the inventory?:[yes/no]")
            if pick2 == 'yes':
                flag = 0
            elif pick2 == 'no':
                flag = 1
            else:
                raise ValueError('Not valid choice.')
        except ValueError:
            print('Value Error.Pick again')

        # after the manager decides what to delete the updated inventory is being written back to the excel file.
        inventory_workbook = xlsxwriter.Workbook('inventory.xlsx')
        worksheet = inventory_workbook.add_worksheet('inventory1')
        for i in range(len(updated_stock_list)):
            for j in range(len(updated_stock_list[i])):
                worksheet.write(i, j, updated_stock_list[i][j])

        inventory_workbook.close()
    Open_Menu(access)
####################################################################################


# gets name of worker and add to data base of the presence
def arrival_to_work(access):
    password_loc = r'C:\Users\micha\PycharmProjects\yesodotFinish\passwarde.xlsx'
    password_file = xlrd.open_workbook(password_loc)
    password_sheet = password_file.sheet_by_index(0)
    name = input('enter your first name: ')
    last = input('enter your last name: ')
    flag = 0
    while True:
        for i in range(password_sheet.nrows):
            if password_sheet.cell_value(i, 0) == name and password_sheet.cell_value(i, 3) == last:
                flag = 1
                break
        if flag:
            break
        print('wrong name')
        name = input('enter your first name: ')
        last = input('enter your last name: ')

    date_now = time.localtime()
    presence_list = []
    row_list = []
    presence_loc = r'C:\Users\micha\PycharmProjects\yesodotFinish\presence1.xlsx'
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
    presence_workbook = xlsxwriter.Workbook('presence1.xlsx')
    worksheet = presence_workbook.add_worksheet('presence')

    for i in range(len(presence_list)):
        for j in range(len(presence_list[i])):
            worksheet.write(i, j, presence_list[i][j])
    presence_workbook.close()
    Open_Menu(access)


####################################################################################


# gets name of worker and add to data base of the presence the time of leaving the work
def departure(access):
    password_loc = r'C:\Users\micha\PycharmProjects\yesodotFinish\passwarde.xlsx'
    password_file = xlrd.open_workbook(password_loc)
    password_sheet = password_file.sheet_by_index(0)
    name = input('enter your first name: ')
    last = input('enter your last name: ')
    flag = 0
    while True:
        for i in range(password_sheet.nrows):
            if password_sheet.cell_value(i, 0) == name and password_sheet.cell_value(i, 3) == last:
                flag = 1
                break
        if flag:
            break
        print('wrong name')
        name = input('enter your first name: ')
        last = input('enter your last name: ')
    presence_list = []
    row_list = []
    presence_loc = r'C:\Users\micha\PycharmProjects\yesodotFinish\presence1.xlsx'
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

    # calculates the second of each time
    arrival_time = (int(worker_arrival[17]) * 10 + int(worker_arrival[18])) * 3600 + \
                   (int(worker_arrival[20]) * 10 + int(worker_arrival[21])) * 60 + \
                   (int(worker_arrival[23]) * 10 + int(worker_arrival[24]))
    departure_time = (int(worker_departure[17]) * 10 + int(worker_departure[18])) * 3600 + \
                     (int(worker_departure[20]) * 10 + int(worker_departure[21])) * 60 + \
                     (int(worker_departure[23]) * 10 + int(worker_departure[24]))
    delta = departure_time - arrival_time
    presence_list[worker][8] = delta

    presence_workbook = xlsxwriter.Workbook('presence1.xlsx')
    worksheet = presence_workbook.add_worksheet('presence')

    for i in range(len(presence_list)):
        for j in range(len(presence_list[i])):
            worksheet.write(i, j, presence_list[i][j])
    presence_workbook.close()
    Open_Menu(access)


####################################################################################


# this func recieves the meassage from the shift manager
def MessageForManager(access):
    messages_list = []
    row_list = []
    message_loc = r'C:\Users\micha\PycharmProjects\yesodotFinish\messages.xlsx'
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
####################################################################################


# find a custumer in the members club
def find_custumer(access):
    id = input('Enter ID: ')
    while id.isnumeric() == False:
        id = input('invalid id, try again:')

    file_loc = r'C:\Users\micha\PycharmProjects\yesodotFinish\membership.xlsx'
    workbook = xlrd.open_workbook(file_loc)
    worksheet = workbook.sheet_by_index(0)
    worksheet.cell_value(0, 0)
    for i in range(worksheet.nrows):
        if worksheet.cell_value(i, 2) == id:
            return True
    return False
####################################################################################

def check_if_customer_is_member_club(id):
    location = r'C:\Users\micha\PycharmProjects\yesodotFinish\membership.xlsx'
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

    # find the id in members list:
    for i in range(len(members_list)):
        if members_list[i][2] == id:
            return True

    return False


# ask the worker for 2 shifts he cant work and add them to the data base
def add_worker_Constraints(access):
    constraints_list = []
    row_list = []
    constraints_loc = r'C:\Users\micha\PycharmProjects\yesodotFinish\Constraints1.xlsx'
    constraints_file = xlrd.open_workbook(constraints_loc)
    amount_sheets = constraints_file.nsheets

    name = input('Enter your name: ')
    # if shift manager try to input constraints before all the workers
    if amount_sheets < 7 and (name == 'michal' or name == 'emilia'):
        print('You still can not submit constraints, wait for all employees to submit')
        Open_Menu(access)

    # if the worker already added his Constraints
    for i in range(amount_sheets):
        sheet = constraints_file.sheet_by_index(i)
        if sheet.name == name:
            print('You have already submitted constraints')
            Open_Menu(access)

    flag = 0
    if name == 'emilia':
        for i in range(amount_sheets):
            if sheet.name == 'michal':
                flag = 1
        if flag == 0:
            print('michal is not submitted her constraints')
            Open_Menu(access)
    for i in range(amount_sheets):
        sheet = constraints_file.sheet_by_index(i)
        if sheet.name == 'Sheet1':
            continue
        sheet_list = [sheet.name]
        for j in range(sheet.nrows):
            row_list = sheet.row_values(j)
            sheet_list.append(row_list)
        constraints_list.append(sheet_list)

    workbook = xlsxwriter.Workbook('Constraints1.xlsx')

    for i in range(len(constraints_list)):  # runs on 2 sheets - michal and shir
        worksheet = workbook.add_worksheet(constraints_list[i][0])  # constraints_list[i][0]- sheet name
        for j in range(1, len(constraints_list[i])):  # number of rows in sheet
            for k in range(len(constraints_list[i][j])):
                worksheet.write(j - 1, k, constraints_list[i][j][k])

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

    days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
    shift = ['Morning', 'Evening']
    print('enter constraints in this form: \nday- {0}\nshift- {1}'.format(days, shift))
    flag1, flag2 = 0, 0
    while True:
        constraint1_day = input('enter your first constraint-> day: ')
        if constraint1_day in days:
            flag1 = 1
        else:
            print('wrong input, try again:')
        constraint1_shift = input('enter your first constraint-> shift: ')
        if constraint1_shift in shift:
            flag2 = 1
        else:
            print('wrong input, try again:')
        if flag1 and flag2:
            break

    flag1, flag2 = 0, 0
    while True:
        constraint2_day = input('enter your second constraint-> day: ')
        if constraint2_day in days:
            flag1 = 1
        else:
            print('wrong input, try again:')
            continue
        constraint2_shift = input('enter your second constraint-> shift: ')
        if constraint2_shift in shift:
            flag2 = 1
        else:
            print('wrong input, try again:')
            continue
        if flag1 and flag2:
            break

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


####################################################################################


# order new stock
def add_new_inventory(access):
    flag = 0
    while flag == 0:
        inventory_list = []
        users_list = []
        inventory_loc = r'C:\Users\micha\PycharmProjects\yesodotFinish\inventory.xlsx'
        inventory_file = xlrd.open_workbook(inventory_loc)
        sheet = inventory_file.sheet_by_index(0)
        # coping the existing data to a new list
        for i in range(sheet.nrows):
            row_list = sheet.row_values(i)
            inventory_list.append(row_list)
        # converting float numbers to integers(exept from the price column)
        for i in range(1, len(inventory_list)):
            num = inventory_list[i][0]
            inventory_list[i][0] = int(num)
            num = inventory_list[i][3]
            inventory_list[i][3] = int(num)
        # information from the manager
        users_list.append(len(inventory_list))
        input_string = input('enter the name of the product: ')
        users_list.append(input_string)
        input_string = input('enter the size of the product: ')
        users_list.append(input_string)
        input_string = input('enter the amount of the product: ')
        users_list.append(input_string)
        input_string = input('enter the color of the product: ')
        users_list.append(input_string)
        input_string = input('enter the price of the product: ')
        users_list.append(input_string)
        inventory_list.append(users_list)
        print('Would you like to order new stock?[yes/no]')
        try:  # if the input is not valid exeption is trown.
            answer = input()
            if answer == 'yes':
                flag = 0
            elif answer == 'no':
                flag = 1
            else:
                raise ValueError('Not valid choice.')
        except ValueError:
            print('Value Error.Pick again')
        # writing the new inventory to the excel file
        inventory_workbook = xlsxwriter.Workbook('inventory.xlsx')
        worksheet = inventory_workbook.add_worksheet('inventory1')
        for i in range(len(inventory_list)):
            for j in range(len(inventory_list[i])):
                worksheet.write(i, j, inventory_list[i][j])

        inventory_workbook.close()
    print(tabulate(inventory_list, tablefmt="fancy_grid"))
    Open_Menu(access)
####################################################################################


def Add_custumer(access):
    location = r'C:\Users\micha\PycharmProjects\yesodotFinish\membership.xlsx'
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


####################################################################################


def Delete_customer(access):
    # saving location file
    location = r'C:\Users\micha\PycharmProjects\yesodotFinish\membership.xlsx'
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
    while ID.isnumeric() == False:
        ID = input('id is not valid, enter costumer id: ')

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


####################################################################################


# return the price of a product given its product code(used in sell function)
def GetPrice(product_code):
    inventory_list = []
    inventory_loc = r'C:\Users\micha\PycharmProjects\yesodotFinish\inventory.xlsx'
    inventory_file = xlrd.open_workbook(inventory_loc)
    sheet = inventory_file.sheet_by_index(0)
    price_index = 0
    # coping the existing data to a new list
    for i in range(sheet.nrows):
        row_list = sheet.row_values(i)
        inventory_list.append(row_list)
    # converting the float numbers to integetrs
    for i in range(1, len(inventory_list)):
        num = inventory_list[i][0]
        inventory_list[i][0] = int(num)
        num = inventory_list[i][3]
        inventory_list[i][3] = int(num)
    # searching for the given product code and returning its price
    for i in range(len(inventory_list)):
        for j in range(len(inventory_list[i])):
            if product_code == inventory_list[i][j]:
                price_index = inventory_list[i][j + 5]
                return price_index
####################################################################################


def check_validation_of_product_code(code):
    inventory_list = []
    inventory_loc = r'C:\Users\micha\PycharmProjects\yesodotFinish\inventory.xlsx'
    inventory_file = xlrd.open_workbook(inventory_loc)
    sheet = inventory_file.sheet_by_index(0)

    # copy the file to list:
    for i in range(0, sheet.nrows):
        row_list = sheet.row_values(i)
        inventory_list.append(row_list)

    # find the product in list:
    for i in range(len(inventory_list)):
        if inventory_list[i][0] == code:
            return True

    return False


####################################################################################


# when making a sell this function updates the amount of the products
def update_stock_with_sale(code_product, amount):
    inventory_list = []
    inventory_loc = r'C:\Users\micha\PycharmProjects\yesodotFinish\inventory.xlsx'
    inventory_file = xlrd.open_workbook(inventory_loc)
    sheet1 = inventory_file.sheet_by_index(0)

    amount_index = 0
    # copy the sheets into lists
    for i in range(sheet1.nrows):
        row_list = sheet1.row_values(i)
        inventory_list.append(row_list)

    # update the specipic product amount
    for i in range(len(inventory_list)):
        if inventory_list[i][0] == code_product:
            inventory_list[i][3] -= amount

    # copy updated list of stock to file
    inventory_workbook = xlsxwriter.Workbook('inventory.xlsx')
    worksheet = inventory_workbook.add_worksheet('inventory1')
    for i in range(len(inventory_list)):
        for j in range(len(inventory_list[i])):
            worksheet.write(i, j, inventory_list[i][j])

    inventory_workbook.close()


# ============ function for create recipt and save her at recipects data=======================
def make_recipect(date, price):
    # saving location file
    location = r'C:\Users\micha\PycharmProjects\yesodotFinish\recipects.xlsx'
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

    workbook.close()


# prints the current sales report
def get_sales_report(access):
    sales_loc = r'C:\Users\micha\PycharmProjects\yesodotFinish\sales.xlsx'
    sales_file = xlrd.open_workbook(sales_loc)
    sheet = sales_file.sheet_by_index(0)

    date_now = time.localtime()
    current_year, current_month, current_day = date_now[0], date_now[1], date_now[2]
    temp_list = []
    # copying the existing data to a new list
    for i in range(1, sheet.nrows):
        row_list = sheet.row_values(i)
        temp_list.append(row_list)
    # converting the float numbers to integers(not the price column)
    for i in range(len(temp_list)):
        temp_list[i][0] = int(temp_list[i][0])
        temp_list[i][1] = int(temp_list[i][1])
        temp_list[i][2] = int(temp_list[i][2])
        temp_list[i][3] = int(temp_list[i][3])
        temp_list[i][5] = int(temp_list[i][5])
    sales_list = temp_list
    printed_list = []
    printed_list.append(['Year', 'Month', 'Day', 'Code Product', 'Name', 'Amount', 'Price'])
    pick = input('[1]Daily report\n[2]Monthly report\nyour choice: ')
    try:
        # the manager is given a choice if he wants a daily report or monthly
        if pick == '1':  # in case the manager wants a daily report he can choose - current day or different day.
            date_choice = input('[1]current day\n[2]Enter different date: ')
            try:
                if date_choice == '1':  # print report for the current day
                    print('*****  Sales Report For Manager  *****')
                    # add to new list all the rows that relevant to this month
                    for j in range(len(sales_list)):
                        if current_year == sales_list[j][0] and current_month == sales_list[j][1] and current_day == \
                                sales_list[j][2]:
                            temp_list = [sales_list[j][0], sales_list[j][1], sales_list[j][2], sales_list[j][3],
                                         sales_list[j][4], sales_list[j][5], sales_list[j][6]]
                            printed_list.append(temp_list)
                    print(tabulate(printed_list, tablefmt="fancy_grid"))

                elif date_choice == '2':  # print report for a different day.
                    print('Enter date[dd/mm/yyyy]:')
                    day_choice = int(input('DAY:'))
                    month_choice = int(input('MONTH(number):'))
                    year_choice = int(input('YEAR:'))
                    for j in range(len(sales_list)):
                        if year_choice == sales_list[j][0] and month_choice == sales_list[j][1] and day_choice == \
                                sales_list[j][2]:
                            temp_list = [sales_list[j][0], sales_list[j][1], sales_list[j][2], sales_list[j][3],
                                         sales_list[j][4], sales_list[j][5], sales_list[j][6]]
                            printed_list.append(temp_list)
                    print(tabulate(printed_list, tablefmt="fancy_grid"))

                else:  # in case something else is pressed
                    raise ValueError('Not valid choice.')
            except ValueError:
                print('Value Error.Pick again')

        elif pick == '2':  # in case the manager wants a monthly report he can choose - current month or different month.
            date_choice = input('[1]current month\n[2]Enter different month: ')
            try:
                if date_choice == '1':  # print report for the current month
                    print('*****  Sales Report For Manager  *****')
                    # add to new list all the rows that relevant to this month
                    for j in range(len(sales_list)):
                        if current_year == sales_list[j][0] and current_month == sales_list[j][1]:
                            temp_list = [sales_list[j][0], sales_list[j][1], sales_list[j][2], sales_list[j][3],
                                         sales_list[j][4],
                                         sales_list[j][5], sales_list[j][6]]
                            printed_list.append(temp_list)
                    print(tabulate(printed_list, tablefmt="fancy_grid"))
                elif date_choice == '2':  # print report for different month
                    year_choice = int(input('YEAR:'))
                    month_choice = int(input('MONTH(number):'))
                    for j in range(len(sales_list)):
                        if year_choice == sales_list[j][0] and month_choice == sales_list[j][1]:
                            temp_list = [sales_list[j][0], sales_list[j][1], sales_list[j][2], sales_list[j][3],
                                         sales_list[j][4], sales_list[j][5], sales_list[j][6]]
                            printed_list.append(temp_list)
                    print(tabulate(printed_list, tablefmt="fancy_grid"))
                else:
                    raise ValueError('Not valid choice.')
            except ValueError:
                print('Value Error.Pick again')

        # in case something else is pressed
        else:
            raise ValueError('Not valid choice.')
    except ValueError:
        print('Value Error.Pick again')
    Open_Menu(access)


# return name of the product
def GetName(product_code):
    inventory_list = []
    inventory_loc = r'C:\Users\micha\PycharmProjects\yesodotFinish\inventory.xlsx'
    inventory_file = xlrd.open_workbook(inventory_loc)
    sheet = inventory_file.sheet_by_index(0)
    name_index = 0
    # coping the exidting data to new list.
    for i in range(sheet.nrows):
        row_list = sheet.row_values(i)
        inventory_list.append(row_list)
    # converting the float numbers into integers(not the price column).
    for i in range(1, len(inventory_list)):
        num = inventory_list[i][0]
        inventory_list[i][0] = int(num)
        num = inventory_list[i][3]
        inventory_list[i][3] = int(num)
    # searching for the name of the item with the given product code
    for i in range(len(inventory_list)):
        for j in range(len(inventory_list[i])):
            if product_code == inventory_list[i][j]:
                name_index = inventory_list[i][j + 1]
                return name_index


####################################################################################


def update_sales(list_1):
    updated_sales_list = []
    cancelled_sales_list = []
    temp_sales_list = []
    sales_loc = r'C:\Users\micha\PycharmProjects\yesodotFinish\sales.xlsx'
    sales_file = xlrd.open_workbook(sales_loc)
    sheet1 = sales_file.sheet_by_index(0)  # sheet of sales
    sheet2 = sales_file.sheet_by_index(1)  # sheet of cancelled sales

    date_now = time.localtime()
    current_year, current_month, current_day = date_now[0], date_now[1], date_now[2]

    # copy first sheet- sales sheet:
    for i in range(sheet1.nrows):
        row_list = sheet1.row_values(i)
        updated_sales_list.append(row_list)

    # copy second sheet- cancelled sales list:
    for i in range(sheet2.nrows):
        row_list = sheet2.row_values(i)
        cancelled_sales_list.append(row_list)

    # update sales list:
    temp_sales_list.append(current_year)
    temp_sales_list.append(current_month)
    temp_sales_list.append(current_day)
    temp_sales_list.extend(list_1)
    updated_sales_list.append(temp_sales_list)
    # converting the float numbers into integers(not the price column).
    for i in range(1, len(updated_sales_list)):
        for k in range(len(updated_sales_list[i])):
            updated_sales_list[i][0] = int(updated_sales_list[i][0])
            updated_sales_list[i][1] = int(updated_sales_list[i][1])
            updated_sales_list[i][2] = int(updated_sales_list[i][2])
            updated_sales_list[i][3] = int(updated_sales_list[i][3])
            updated_sales_list[i][5] = int(updated_sales_list[i][5])

    sales_workbook = xlsxwriter.Workbook('sales.xlsx')
    # add sheet number one- sheet for sales:
    worksheet1 = sales_workbook.add_worksheet('Sales01')
    for i in range(len(updated_sales_list)):
        for j in range(len(updated_sales_list[i])):
            worksheet1.write(i, j, updated_sales_list[i][j])

    # add sheet number two- sheet for cancelled sales:
    worksheet2 = sales_workbook.add_worksheet('cancelled sales')
    for i in range(len(cancelled_sales_list)):
        for j in range(len(cancelled_sales_list[i])):
            worksheet2.write(i, j, cancelled_sales_list[i][j])

    sales_workbook.close()


def The_number_of_next_recipct():
    """function that return the next number in recipects list, we need it for sell process"""
    # saving location file
    location = r'C:\Users\micha\PycharmProjects\yesodotFinish\recipects.xlsx'
    # variable that present the file we will work with
    recipects_file = xlrd.open_workbook(location)
    # the specific sheet we need from the file:
    sheet = recipects_file.sheet_by_index(0)

    number = sheet.nrows
    return number


def sell_items(access):
    # ============== check if customer is a friend in members club ==================

    print('\n\n------------------- sell page-------------------\n')
    answer = int(input("Is the customer a member of the customer club?\nIf he doe's enter 1, otherwise enter 0"))
    while answer != 1 and answer != 0:
        answer = int(input("invalid answer, try again"))

    if answer == 1:
        id = input('Enter customer ID: ')
        while id.isnumeric() == False:
            id = input('invalid id, Enter customer ID again: ')

        assumption = 0
        if check_if_customer_is_member_club(id):
            print('The customer is a member in the customer club, so he won 15% off the sale')
            assumption = 0.15
        else:
            print("The customer isn't a member in the customer club")
            assumption = 0
    else:
        assumption = 0

    # ==================================== Account Execution =========================================

    total_price = 0
    item_list = []
    item = []

    # get recipet number
    recipet_num = The_number_of_next_recipct()
    flag = 1

    while flag == 1:

        date1 = str(date.today())
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
        item = [product_code, product_name, Quantity, price, recipet_num]
        item_list.append(item)

        # if there's more items:
        flag = int(input('for add more items press 1, else press 0:'))
        # checking validation of input:
        while flag != 0 and flag != 1:
            flag = int(input('invalid answer, try again- for add more items press 1, else press 0'))

    # =========================== delete item from list during sale ====================================
    flag = 1
    while flag == 1:
        # show items list:
        print('************** products list is:**************\n')
        for i in range(len(item_list)):
            print('{0}) {1}'.format(i + 1, item_list[i]))
        flag = int(input('To delete items from list press 1, to continue press 0:'))
        # check validation of answer:
        while flag != 0 and flag != 1:
            flag = int(input('invalid answer, try again- To delete items from list press 1, to continue press 0'))

        if flag == 1:
            index = input('Enter index of item you want to remove')
            # check validation of index input:
            while index.isnumeric() == False:
                index = input('invalid index, try again. Enter index of item you want to remove')
            index = int(index)
            while index < 1 or index > (len(item_list)):
                index = int(input('invalid index, try again. Enter index of item you want to remove'))

            # if theres more the one item to same index, remove only 1 item:
            if item_list[index - 1][2] > 1:
                item_list[index - 1][2] -= 1
            # if there is only one item, remve all item line from list
            else:
                del (item_list[index - 1])

            print('item removed')

    # =========================== calculate total price =============================================
    total_price = 0
    for i in range(len(item_list)):
        item_amount = item_list[i][2]
        item_price = item_list[i][3]
        total_price += (item_price * item_amount)

    # ============================= update sell of items=============================================
    for i in range(len(item_list)):
        update_sales(item_list[i])

    # ============================= update stock with bought items====================================
    for i in range(len(item_list)):
        update_stock_with_sale(item_list[i][0], item_list[i][2])

    # =============================== print recipect ==================================================

    print('\n\n****costumer recipect****')
    print('--------recepict number:{0}--------'.format(recipet_num))
    print('Date:{0}\n'.format(date1))
    for i in range(len(item_list)):
        for j in range(1, 4):
            print(item_list[i][j], end=" ")
        print()

    print('\ntotal price is:{0}₪'.format(total_price))
    print('Members club assumption is: {0}%'.format(assumption))
    total_price = total_price - (total_price * assumption)
    total_price = round(total_price, 2)
    print('update price: {0}₪'.format(total_price))
    tax = total_price * 0.18
    tax = round(tax, 2)
    print('Tax=18%: {0}₪'.format(tax))
    total_price = total_price + tax
    print('===========================')
    total_price = total_price = round(total_price, 2)
    print('Sum is:{0}₪'.format(total_price))

    # ===================================== make recipect =======================================
    make_recipect(date1, total_price)

    # ====================================== back to main page ===================================
    Open_Menu(access)


def check_recipect_number_validation(rec_num):
    # ======================== get lists of recipect =====================================
    # saving location file
    location = r'C:\Users\micha\PycharmProjects\yesodotFinish\recipects.xlsx'
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

    # search for recipect number:
    for i in range(len(recipects_list)):
        if recipects_list[i][0] == rec_num:
            return True
    return False


####################################################################################


def get_recipect_date(number):
    # ======================== get lists of recipect =====================================
    # saving location file
    location = r'C:\Users\micha\PycharmProjects\yesodotFinish\recipects.xlsx'
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

    for i in range(len(recipects_list)):
        if recipects_list[i][0] == number:
            return recipects_list[i][1]
    return


####################################################################################


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
    file_loc = r'C:\Users\micha\PycharmProjects\yesodotFinish\messages.xlsx'

    workbook = xlrd.open_workbook(file_loc)
    worksheet = workbook.sheet_by_index(0)
    print('-----------------------------------------------')
    print("*****Dear manager,you have new alert*****")
    print('*****{0}*****'.format(worksheet.cell_value(worksheet.nrows - 1, 1)))
    print('-----------------------------------------------')
    print('** manager menu **')
    print('1- Cash Desk')
    print('2- Cancelling a transaction/ Refund')
    print('3- Reports')
    print('4- Constraints & Shifts')
    print('5- Customers Club')
    print('6- Inventory')
    print('7- Change User')
    print('-----------------------------------------------')

    choice = int(input('your choice: '))
    while True:
        if choice > 0 and choice < 15:
            break
        choice = int(input('wrong choice, try again: '))

    # *********************************************************
    if choice == 1:
        sell_items(access)
    # *********************************************************
    if choice == 2:
        cancel_sell(access)
    # *********************************************************
    if choice == 3:
        print('1- Presence report')
        print('2- Inventory report')
        print('3- Shift report')
        print('4- sales report')

        choice = int(input('your choice: '))
        while True:
            if choice > 0 and choice < 5:
                break
            choice = int(input('wrong choice, try again: '))

        if choice == 1:
            get_manager_presence_report(access)
        if choice == 2:
            get_inventory_report(access)
        if choice == 3:
            shifts_report(access)
        if choice == 4:
            get_sales_report(access)
    # *********************************************************
    if choice == 4:
        print('1- Submission of constraints')
        print('2- Build Shifts')
        print('3- make changes in shifts')
        choice = int(input('your choice: '))
        while True:
            if choice > 0 and choice < 4:
                break
            choice = int(input('wrong choice, try again: '))

        if choice == 1:
            clear_constraints(access)
        if choice == 2:
            build_shifts(access)
        if choice == 3:
            make_changes_in_shifts(access)
    # *********************************************************
    if choice == 5:
        print('1- add customer')
        print('2- remove customer')
        choice = int(input('your choice: '))
        while True:
            if choice > 0 and choice < 3:
                break
            choice = int(input('wrong choice, try again: '))

        if choice == 1:
            Add_custumer(access)
        if choice == 2:
            Delete_customer(access)
    # *********************************************************
    if choice == 6:
        print('1- Order new stock')
        print('2- Remove items')
        choice = int(input('your choice: '))
        while True:
            if choice > 0 and choice < 3:
                break
            choice = int(input('wrong choice, try again: '))

        if choice == 1:
            add_new_inventory(access)
        if choice == 2:
            return_inventory(access)
    # *********************************************************
    if choice == 7:
        Log_In()
    # *********************************************************


def Responsible_menu(access):
    print('-----------------------------------------------')
    print('** Responsible menu **')
    print('1- Cash Desk')
    print('2- Cancelling a transaction/ Refund')
    print('3- Reports')
    print('4- Submission of constraints')
    print('5- Customers Club')
    print('6- Messages to Manager')
    print('7- Entry & Departing')
    print('8- Change User')
    print('-----------------------------------------------')

    choice = int(input('your choice: '))
    while True:
        if choice > 0 and choice < 15:
            break
        choice = int(input('wrong choice, try again: '))

    # *********************************************************
    if choice == 1:
        sell_items(access)
    # *********************************************************
    if choice == 2:
        cancel_sell(access)
    # *********************************************************
    if choice == 3:
        print('1- Presence report')
        print('2- Inventory report')
        print('3- Shift report')

        choice = int(input('your choice: '))
        while True:
            if choice > 0 and choice < 4:
                break
            choice = int(input('wrong choice, try again: '))

        if choice == 1:
            get_monthly_presence_report(access)
        if choice == 2:
            get_inventory_report(access)
        if choice == 3:
            shifts_report(access)
    # *********************************************************
    if choice == 4:
        add_worker_Constraints(access)
    # *********************************************************
    if choice == 5:
        print('1- add customer')
        print('2- remove customer')
        choice = int(input('your choice: '))
        while True:
            if choice > 0 and choice < 3:
                break
            choice = int(input('wrong choice, try again: '))

        if choice == 1:
            Add_custumer(access)
        if choice == 2:
            Delete_customer(access)
    # *********************************************************
    if choice == 6:
        MessageForManager(access)
    # *********************************************************
    if choice == 7:
        print('1- arrival')
        print('2- departure')
        choice = int(input('your choice: '))
        while True:
            if choice > 0 and choice < 3:
                break
            choice = int(input('wrong choice, try again: '))
        if choice == 1:
            arrival_to_work(access)
        if choice == 2:
            departure(access)
    # *********************************************************
    if choice == 8:
        Log_In()
    # *********************************************************


def worker_menu(access):
    print('-----------------------------------------------')
    print('** worker menu **')
    print('1- Cash Desk')
    print('2- Reports')
    print('3- Submission of constraints')
    print('4- Customers Club')
    print('5- Entry & Departing')
    print('6- Closing The Register')
    print('7- Change User')
    print('-----------------------------------------------')

    choice = int(input('your choice: '))
    while True:
        if choice > 0 and choice < 15:
            break
        choice = int(input('wrong choice, try again: '))

    # *********************************************************
    if choice == 1:
        sell_items(access)
    # *********************************************************
    if choice == 2:
        print('1- Presence report')
        print('2- End of the day report')
        print('3- Shift report')

        choice = int(input('your choice: '))
        while True:
            if choice > 0 and choice < 4:
                break
            choice = int(input('wrong choice, try again: '))

        if choice == 1:
            get_monthly_presence_report(access)
        if choice == 2:
            total = EOD_report(access)
            print('The final amount of money for today is: {0} NIS'.format(total))
        if choice == 3:
            shifts_report(access)
    # *********************************************************
    if choice == 3:
        add_worker_Constraints(access)
    # *********************************************************
    if choice == 4:
        print('1- add customer')
        print('2- find customer')
        choice = int(input('your choice: '))
        while True:
            if choice > 0 and choice < 3:
                break
            choice = int(input('wrong choice, try again: '))

        if choice == 1:
            Add_custumer(access)
        if choice == 2:
            result = find_custumer(access)
            if result:
                print('Customer exists')
            else:
                print('Customer does not exist')
        Open_Menu(access)
    # *********************************************************
    if choice == 5:
        print('1- arrival')
        print('2- departure')
        choice = int(input('your choice: '))
        while True:
            if choice > 0 and choice < 3:
                break
            choice = int(input('wrong choice, try again: '))
        if choice == 1:
            arrival_to_work(access)
        if choice == 2:
            departure(access)
    # *********************************************************
    if choice == 6:
        Closing_The_Register(access)
    # *********************************************************
    if choice == 7:
        Log_In()
    # *********************************************************


def Error_page():
    exit(0)


def Log_In():
    file_loc = r'C:\Users\micha\PycharmProjects\yesodotFinish\passwarde.xlsx'
    pas_file = xlrd.open_workbook(file_loc)
    sheet = pas_file.sheet_by_index(0)
    sheet.cell_value(0, 0)

    flag = 0

    def check_name(flag):
        name = input('enter user name-english letters only: ')
        for i in range(1, sheet.ncols + 1):
            check = sheet.cell_value(i, 0)
            if check == name:
                flag = 1
                Password = int(input('Enter a 6-digit password-'))
                index = i
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

    ans = check_name(0)
    if ans == 0:
        for i in range(2):
            print("Name does not exist on the system, Try again")
            if check_name(0) != 0:
                break

#
# Log_In()






