# Credit Card payment and deposit report generator
#    
# Copyright (C) 2018  Mark
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.


from xlrd import open_workbook, xldate_as_tuple
import xlwt
from xlwt import Workbook, XFStyle

from datetime import date, datetime
from collections import OrderedDict
import os, sys


def get_paymentmode(filename):
    """
    Get the payment mode from file
    """
    with open(filename, 'r') as f:
        for line in f:
            line = line.strip()
            yield line


def checkPayment(payment):
    """
    Check if payment mode is for the list
    """
    return payment in PaymentMode
 

def getCreditCard(sheet, paymentcol, datecol):
    """
    Extract all credit card transaction from excel file
    """
    creditCard = []
    for i in range(sheet.nrows):
        a = []
        for j in range(sheet.ncols):
            if checkPayment(sheet.cell_value(i, paymentcol)):
                if j == datecol:
                    datevalue = xldate_as_tuple(sheet.cell(i, datecol).value,
                                                book.datemode)
                    a.append(datevalue)
                else:
                    a.append(sheet.cell_value(i, j))
            else:
                break

        if a != []:
            creditCard.append(a)
    return creditCard


def get_collect(ccard, ccdate):
    """
    Get collected credit card deposit from excel file
    """
    collected = []
    for i in range(len(ccard)):
        rows = []
        for j in range(len(ccard[i])):
            date_val = ccard[i][7]
            dateval = date(*date_val[:3])
            if ccard[i][12] < 0 and dateval == ccdate:
                rows.append(ccard[i][j])
        if rows != []:
            collected.append(rows)

    return collected


def get_refund(ccard, ccdate):
    """
    Get refunded credit card deposti from excel file
    """
    refunded = []
    for i in range(len(ccard)):
        rows = []
        for j in range(len(ccard[i])):
            date_val = ccard[i][7]
            dateval = date(*date_val[:3])
            if ccard[i][12] > 0 and dateval == ccdate:
                rows.append(ccard[i][j])
        if rows != []:
            refunded.append(rows)
    return refunded


def get_cif(ccard, ccdate):
    """
    Get credit card transaction from Cash In Flow excel file
    """
    cif = []
    for i in range(len(ccard)):
        rows = []
        for j in range(len(ccard[i])):
            date_val = ccard[i][2]
            dateval = date(*date_val[:3])
            if dateval == ccdate:
                rows.append(ccard[i][j])
        if rows != []:
            cif.append(rows)
    return cif


def remove_unwanted(refund, collect):
    """Remove refunded deposit from deposit list"""
    refunded = []

    for item in refund:
        for x in range(len(collect)):
            if collect[x] == []:
                continue
            elif item[16] == collect[x][16] and \
                    item[12] == abs(collect[x][12]) or \
                    item[5] == collect[x][5] and \
                    item[12] == abs(collect[x][12]):
                collect[x] = []
                refunded.append(item)
                break

    # remove empty list from collected list
    remove_me = []
    while remove_me in collect:
        collect.remove(remove_me)

    # remove refunded credit card
    for item in refunded:
        if item in refund:
            refund.remove(item)

    return collect, refund


if __name__ == '__main__':
    # Change directory to C:\Users\reception\Desktop\symphony\
    os.chdir('..')
    
    # Get Cash in flow name from user
    cashinflow_filename = input("Enter file name (CIF): ")
    cashinflow_filename += ".xls"

    # Get Deposit list name from user
    deposit_filename = input("Enter file name (Deposit): ")
    deposit_filename += ".xls"

    # Open cash in flow workbook
    book = open_workbook(cashinflow_filename)
    sheet = book.sheet_by_index(0)

    # Open deposit list workbook
    book = open_workbook(deposit_filename)
    sheet2 = book.sheet_by_index(0)

    # Get payment mode from payment.txt file
    payment_file = 'payment.txt'
    PaymentMode = list(get_paymentmode(payment_file))

    # Get the list of credit card from workbook
    credit_card_cif = getCreditCard(sheet, 6, 2)
    credit_card_deposit = getCreditCard(sheet2, 8, 7)

    # Test code to print creditcard & deposit list
    #print(len(credit_card_cif))
    #print(len(credit_card_deposit))
    #print(credit_card_deposit)

    # Get the date list from credit card list
    datelist = []

    if credit_card_cif: # get date from cash in flow 
        for i in range(len(credit_card_cif)):
            date_val = credit_card_cif[i][2]
            dateval = date(*date_val[:3])
            if dateval not in datelist:
                datelist.append(dateval)
        if credit_card_deposit: # get date from deposit
            for i in range(len(credit_card_deposit)):
                date_val = credit_card_deposit[i][7]
                dateval = date(*date_val[:3])
                if dateval not in datelist:
                    datelist.append(dateval)
    elif credit_card_deposit: # get date from deposit if cash in flow is empty
        for i in range(len(credit_card_deposit)):
            date_val = credit_card_deposit[i][7]
            dateval = date(*date_val[:3])
            if dateval not in datelist:
                datelist.append(dateval)
    
    else: # exit if no date from both cash in flow and deposit
        sys.exit("No data found")

    datelist.sort()

    first = datelist[0].strftime("%d%m%y")
    if len(datelist) > 1:
        last = datelist[len(datelist)-1].strftime("%d%m%y")
    else:
        last = ""

    # Separate credit card cif by date
    ccard_cif = {}

    for i in range(len(datelist)):
        dateStr = datelist[i].strftime("%y%m%d")
        ccard_cif[dateStr + "Revenue"] = get_cif(credit_card_cif, datelist[i])


    # create dictionaries of the credit card
    ccard_list = {}

    for i in range(len(datelist)):
        dateStr = datelist[i].strftime("%y%m%d")
        ccard_list[dateStr + "Collect"] = get_collect(credit_card_deposit,
                                                    datelist[i])
        ccard_list[dateStr + "Refund"] = get_refund(credit_card_deposit,
                                                    datelist[i])


    # sort the dictionary
    od = OrderedDict(sorted(ccard_list.items()))

    #SORT_ORDER = {"AMEX": 0, "CREDIT CARD": 1, "MASTER": 2, "VISA": 3, \
    #              "UNI": 4, "UNIONPAY": 5}

    SORT_ORDER = {}

    # sorting PaymentMode
    PaymentMode.sort()

    for i in range(len(PaymentMode)):
        SORT_ORDER[PaymentMode[i]] = i

    for i in range(len(datelist)):
        dateStr = datelist[i].strftime("%y%m%d")
        temp_depo, refund = remove_unwanted(od[dateStr + "Refund"],
                                            od[dateStr + "Collect"])
        print(refund)
        if refund != []:
            temp_depo += refund
        temp_depo.sort(key=lambda val: SORT_ORDER[val[8]])
        od[dateStr + "Collect"] = temp_depo
        del od[dateStr + "Refund"]


    for key in ccard_cif:
        for i in range(len(ccard_cif[key])):
            ccard_cif[key][i] = ["R", ccard_cif[key][i][2], ccard_cif[key][i][4],
                                ccard_cif[key][i][5], ccard_cif[key][i][9],
                                ccard_cif[key][i][10]]


    for key in od:
        for i in range(len(od[key])):
            if od[key][i][12] > 0:
                od[key][i][12] = od[key][i][12] - (od[key][i][12]*2)
            elif od[key][i][12] < 0:
                od[key][i][12] = od[key][i][12] - (od[key][i][12]*2)
            od[key][i] = ["D", od[key][i][7], od[key][i][4], od[key][i][5],
                        od[key][i][16], od[key][i][12]]


    # concate cif and deposit
    for key in ccard_cif:
        mydate = key[:6]
        for item in od:
            if mydate in item:
                ccard_cif[key] = ccard_cif[key] + od[item]


    # Sort the credit card listing
    ccard_temp = ccard_cif
    ccard_cif = OrderedDict(sorted(ccard_temp.items()))

    # Create workbook for writing
    wkbook = Workbook()
    sheet1 = []

    for i in range(len(datelist)):
        test = datelist[i].strftime("%d%m%y")
        sheet1.append(wkbook.add_sheet('Collect-' + test))
        # Sytling the date
    style = XFStyle()
    style.num_format_str = 'DD/MM/YYYY'

    # set to 2 decimal point
    style2 = XFStyle()
    style2.num_format_str = "#,##0.00"

    style3 = XFStyle()


    font = xlwt.Font()
    font.height = 180
    style.font = font
    style2.font = font
    style3.font = font

    h = -1
    for key in ccard_cif:
        h += 1
        #print(key)
        for i in range(len(ccard_cif[key])):
            for j in range(1, 7):
                if j == 1:  # type
                    sheet1[h].write(i, j, ccard_cif[key][i][0], style3)
                if j == 2:  # date
                    sheet1[h].write(i, j, datetime(*ccard_cif[key][i][1]), style)
                elif j == 3:  # room
                    sheet1[h].write(i, j, ccard_cif[key][i][2], style3)
                elif j == 4:  # name
                    sheet1[h].write(i, j, ccard_cif[key][i][3], style3)
                elif j == 5:  # folio
                    sheet1[h].write(i, j, ccard_cif[key][i][4], style3)
                elif j == 6:  # amount
                    sheet1[h].write(i, j, ccard_cif[key][i][5], style2)

    print("Exporting to Excel....")
    if last:
        filename = "CCard_" + first + "-" + last + ".xls"
    else:
        filename = "CCard_" + first + ".xls"

    wkbook.save(filename)

    print("Opening the newly created file ....")
    if os.name == 'nt':
        os.startfile(filename)