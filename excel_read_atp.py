import xlrd
import sys
import os
import time
import traceback
import atp_def
import Junipter
import Mitac
import Fabrinet
import re

#----------------------------------------------------------------------
def check_ATP_rule(first_sheet,key_row):

    if Junipter.search_junipter_rule(first_sheet,key_row) == 1:
        print "Juniper rule match"
        return 1
    elif Mitac.search_mitac_rule(first_sheet,key_row) == 1:
        print "Mitac rule match"
        return 1
    elif Fabrinet.search_fabrinet_rule(first_sheet,key_row) == 1:
        print "fabrinet rule match"
        return 1
    else:
        print "No rule found and need to verify manually"
        return 0

#---------------------------------------------------------------------- 
def search_txt_CM():
    search_txt_CM.txt_file_countj = 0
    search_txt_CM.key_sheet_name = []

    for file in os.listdir(atp_def.LIST_PATH):
        if file.endswith(atp_def.FILE_TYPE3):
            print "============> Txt Files = %s/%s" %(atp_def.LIST_PATH,file)
            search_txt_CM.key_sheet_name.append(file)
            search_txt_CM.xls_file_countj += 1

    print "txt file count = %d" %(search_txt_CM.txt_file_countj)

#----------------------------------------------------------------------
def search_key_in_txt_file(key,filename):
    """
    Open ic require file with part number
    """
    #print "key = " + key
    i = 1
    p = re.compile(key)
    fp = open(filename,"r")
    zops = fp.readlines()
    for lineStr in zops:
        if(p.search(lineStr))
            print "================> Matched at row[%d] at [%s]" %(i,filename)
            print "================> Matched line: [%s]" %(lineStr)

#---------------------------------------------------------------------- 
def search_xls_CM():
    search_xls_CM.xls_file_countj = 0
    search_xls_CM.key_sheet = []
    search_xls_CM.key_sheet_name = []

    for file in os.listdir(atp_def.LIST_PATH):
        if file.endswith(atp_def.FILE_TYPE1) | file.endswith(atp_def.FILE_TYPE2):
            print "============> Excel Files = %s/%s" %(atp_def.LIST_PATH,file)
            # Get file sheet
            book = xlrd.open_workbook(atp_def.LIST_PATH + "/" + file)
            print "============> number of sheets = %d" %(book.nsheets)
            # get the first worksheet
            search_xls_CM.key_sheet_name.append(file)
            search_xls_CM.key_sheet.append(book.sheet_by_index(0))
            search_xls_CM.xls_file_countj += 1

    print "xls file count = %d" %(search_xls_CM.xls_file_countj)

#----------------------------------------------------------------------
def search_key_in_file(first_sheet,key,filename):
    """
    Open ic require file with part number
    """
    #print "key = " + key
    for i in range(first_sheet.nrows):
        row = first_sheet.row_values(i)
        for j in range(len(row)):
            if row[j] == key:
                print "================> Matched at row[%d] at [%s]" %(i,filename)
                check_ATP_rule(first_sheet,i)
#----------------------------------------------------------------------
def search_key_in_CM_support(key):
    for i in range(0,search_xls_CM.xls_file_countj):
        search_key_in_file(search_xls_CM.key_sheet[i],key,search_xls_CM.key_sheet_name[i])

    for i in range(0,search_txt_CM.xls_file_countj):
        search_key_in_file(key,search_txt_CM.key_sheet_name[i])
#----------------------------------------------------------------------
def check_icq_row_range(ic_sheet,index):
    """
    Open ic require file with part number
    """
    end_count = 1

    while 1:
        if (index + end_count) >= ic_sheet.nrows:
            return 0
        row = ic_sheet.row_values(index + end_count)
        #print "IC row = " + row[0]
        if row[0] != "":
            break
        else:
            end_count += 1
            print " ===== PN = %s, eip forest qty = %d, %d SUM[%d]" %(row[3], row[4],row[5],row[4]+row[5])
            search_key_in_CM_support(row[3])


    #print "end count = ",
    #print end_count - 1
    return end_count - 1

#----------------------------------------------------------------------
def open_icreq_file(first_sheet,key):
    """
    Open ic require file with part number
    """
    #print "key = " + key
    for i in range(first_sheet.nrows):
        row = first_sheet.row_values(i)
        for j in range(len(row)):
            if row[j] == key:
                #print "IC require row = ",
                #print i
                product_count = check_icq_row_range(first_sheet,i)

#----------------------------------------------------------------------
def open_index_file(path):
    """
    Open and read an index Excel file
    """
    book = xlrd.open_workbook(path)
    # print number of sheets
    #print book.nsheets
    # print sheet names
    #print book.sheet_names()
    # get the first worksheet
    first_sheet = book.sheet_by_index(0)
    # read a row
    #print first_sheet.row_values(0)
    # read a cell
    cell = first_sheet.cell(1,0)
    #print cell
    #print cell.value
    # read a row slice
    #print first_sheet.row_slice(rowx=0,start_colx=0,end_colx=2)

    i = 0
    j = 0
    nrows = first_sheet.nrows
    ncols = first_sheet.ncols

    # Get file sheet
    book = xlrd.open_workbook(atp_def.ICREQ_FILE)
    # get the first worksheet
    icq_sheet = book.sheet_by_index(0)
    # ===============

    for i in range(2,nrows):
        print "[=========================================================================================================]"
    	print "[Index = %d] IC part num = %s" %(first_sheet.cell(i,0).value, first_sheet.cell(i,1).value)
        open_icreq_file(icq_sheet, first_sheet.cell(i,1).value)

    #print "xls file count = %d" %(search_xls_CM.xls_file_countj)
    #print first_sheet.row_values(i)
#----------------------------------------------------------------------
if __name__ == "__main__":
    path = atp_def.INDEX_FILE
    reload(sys)
    sys.setdefaultencoding('utf-8')
    ## date and time representation
    print "Setup Path = %s" %(path)
    print "Current date & time " + time.strftime("%c")
    search_xls_CM()
    search_txt_CM()
    open_index_file(path)