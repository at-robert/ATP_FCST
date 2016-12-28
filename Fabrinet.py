import xlrd
import os
import re
import atp_def

#----------------------------------------------------------------------
def check_if_string_num_add(str):
    num = 0
    rule = re.compile(r'\([0-9]*\+[0-9]*.*\)', re.IGNORECASE)

    try:
        val = int(str)
        return val
    except:
        val = 0

    if rule.search(str):
        str = str.replace('(','')
        str = str.replace(')','')
        strs = str.split('+')
    else:
        return 0

    for sr in strs:
        val = val + int(sr)

    return val
#----------------------------------------------------------------------
def check_fabrinet_month(first_sheet, index):

    if atp_def.MONTH_SET == "JAN_FEB":
        month1 = re.compile(r'[0-9][0-9]-JAN-2017', re.IGNORECASE)
        month2 = re.compile(r'[0-9][0-9]-FEB-2017', re.IGNORECASE)
        #print "Fabrinet Rule JAN_FEB"
    elif atp_def.MONTH_SET == "FEB_MARCH":
        month1 = re.compile(r'[0-9][0-9]-FEB-2017', re.IGNORECASE)
        month2 = re.compile(r'[0-9][0-9]-MAR-2017', re.IGNORECASE)
        #print "Fabrinet Rule FEB_MARCH"
    elif atp_def.MONTH_SET == "MARCH_APL":
        month1 = re.compile(r'[0-9][0-9]-MAR-2017', re.IGNORECASE)
        month2 = re.compile(r'[0-9][0-9]-APR-2017', re.IGNORECASE)
        #print "Fabrinet Rule MARCH_APL"
    elif atp_def.MONTH_SET == "MAY_JUNE":
        month1 = re.compile(r'[0-9][0-9]-MAY-2017', re.IGNORECASE)
        month2 = re.compile(r'[0-9][0-9]-JUN-2017', re.IGNORECASE)
        #print "Fabrinet Rule MAY JUNE"
    elif atp_def.MONTH_SET == "JUNE_JULY":
        month1 = re.compile(r'[0-9][0-9]-JUN-2017', re.IGNORECASE)
        month2 = re.compile(r'[0-9][0-9]-JUL-2017', re.IGNORECASE)
    elif atp_def.MONTH_SET == "JULY_AUGUST":
        month1 = re.compile(r'[0-9][0-9]-JUL-2017', re.IGNORECASE)
        month2 = re.compile(r'[0-9][0-9]-AUG-2017', re.IGNORECASE)
        #print "Fabrinet Rule JULY AUGUST"
    elif atp_def.MONTH_SET == "SEP_OCT":
        month1 = re.compile(r'[0-9][0-9]-SEP-2017', re.IGNORECASE)
        month2 = re.compile(r'[0-9][0-9]-OCT-2017', re.IGNORECASE)
        #print "Fabrinet Rule SEP_OCT"
    
    elif atp_def.MONTH_SET == "NOV_DEC":
        month1 = re.compile(r'[0-9][0-9]-NOV-2017', re.IGNORECASE)
        month2 = re.compile(r'[0-9][0-9]-DEC-2017', re.IGNORECASE)
        #print "Fabrinet Rule NOV_DEC"
    else:
        print "Fabrinet wrong Month"
        return 0

    row = first_sheet.row_values(search_fabrinet_rule.first_row)
    s = repr(row[search_fabrinet_rule.month_colm[index]])

    if month1.search(s):
        return 1
    if month2.search(s):
        return 2

    print "Fabrinet wrong Month"
    return 0

#----------------------------------------------------------------------
def print_spec_word(first_sheet, key_row, word, word_rule):
    spec_word = re.compile(word_rule, re.IGNORECASE)
    row = first_sheet.row_values(key_row)
    for j in range(len(row)):
        s = repr(row[j])
        if spec_word.search(s):
            print word
#----------------------------------------------------------------------
def print_rule(first_sheet, key_row):

    month1_sum = 0
    month2_sum = 0

    print_spec_word(first_sheet,key_row,"[PO Qty]",r'PO Qty')
    print_spec_word(first_sheet,key_row,"[FCST Allocate]",r'FCST \%Allocate')

    row = first_sheet.row_values(search_fabrinet_rule.first_row)
    for j in range(search_fabrinet_rule.month_colm_cout):
        print "%s " %(row[search_fabrinet_rule.month_colm[j]]),
    print ""

    row = first_sheet.row_values(key_row)
    for j in range(search_fabrinet_rule.month_colm_cout):
        print "%s         " %(row[search_fabrinet_rule.month_colm[j]]),
        if(check_fabrinet_month(first_sheet,j) == 1):
            month1_sum = month1_sum + check_if_string_num_add(row[search_fabrinet_rule.month_colm[j]])
        elif(check_fabrinet_month(first_sheet,j) == 2):
            month2_sum = month2_sum + check_if_string_num_add(row[search_fabrinet_rule.month_colm[j]])
    print ""

    print "Month 1 sum = %d, Month 2 sum = %d" %(month1_sum,month2_sum)
#----------------------------------------------------------------------
def search_fabrinet_rule(first_sheet, key_row):
    """
    Open ic require file with part number
    """
    
    if atp_def.MONTH_SET == "JAN_FEB":
        month1 = re.compile(r'[0-9][0-9]-JAN-2017', re.IGNORECASE)
        month2 = re.compile(r'[0-9][0-9]-FEB-2017', re.IGNORECASE)
        #print "Fabrinet Rule JAN_FEB"
    elif atp_def.MONTH_SET == "FEB_MARCH":
        month1 = re.compile(r'[0-9][0-9]-FEB-2017', re.IGNORECASE)
        month2 = re.compile(r'[0-9][0-9]-MAR-2017', re.IGNORECASE)
        #print "Fabrinet Rule FEB_MARCH"
    elif atp_def.MONTH_SET == "MARCH_APL":
        month1 = re.compile(r'[0-9][0-9]-MAR-2017', re.IGNORECASE)
        month2 = re.compile(r'[0-9][0-9]-APR-2017', re.IGNORECASE)
        #print "Fabrinet Rule MARCH_APL"
    elif atp_def.MONTH_SET == "MAY_JUNE":
        month1 = re.compile(r'[0-9][0-9]-MAY-2017', re.IGNORECASE)
        month2 = re.compile(r'[0-9][0-9]-JUN-2017', re.IGNORECASE)
        #print "Fabrinet Rule MAY JUNE"
    elif atp_def.MONTH_SET == "JUNE_JULY":
        month1 = re.compile(r'[0-9][0-9]-JUN-2017', re.IGNORECASE)
        month2 = re.compile(r'[0-9][0-9]-JUL-2017', re.IGNORECASE)
    elif atp_def.MONTH_SET == "JULY_AUGUST":
        month1 = re.compile(r'[0-9][0-9]-JUL-2017', re.IGNORECASE)
        month2 = re.compile(r'[0-9][0-9]-AUG-2017', re.IGNORECASE)
        #print "Fabrinet Rule JULY AUGUST"
    elif atp_def.MONTH_SET == "SEP_OCT":
        month1 = re.compile(r'[0-9][0-9]-SEP-2017', re.IGNORECASE)
        month2 = re.compile(r'[0-9][0-9]-OCT-2017', re.IGNORECASE)
        #print "Fabrinet Rule SEP_OCT"
    
    elif atp_def.MONTH_SET == "NOV_DEC":
        month1 = re.compile(r'[0-9][0-9]-NOV-2017', re.IGNORECASE)
        month2 = re.compile(r'[0-9][0-9]-DEC-2017', re.IGNORECASE)
        #print "Fabrinet Rule NOV_DEC"
    else:
        print "Fabrinet wrong Month"
    

    search_fabrinet_rule.month_colm = []
    search_fabrinet_rule.month_colm_cout = 0
    search_fabrinet_rule.first_row = 0


    for i in range(first_sheet.nrows):
        row = first_sheet.row_values(i)
        for j in range(len(row)):
            s = repr(row[j])
            if month1.search(s):
                #print "May Matched at colm[%d]" %(j)
                search_fabrinet_rule.month_colm.append(j)
                search_fabrinet_rule.month_colm_cout += 1
                search_fabrinet_rule.first_row = i
            if month2.search(s):
                #print "June Matched at colm[%d]" %(j)
                search_fabrinet_rule.month_colm.append(j)
                search_fabrinet_rule.month_colm_cout += 1
                search_fabrinet_rule.first_row = i

    if(search_fabrinet_rule.month_colm_cout < 1):
        return 0

    for k in range(0, search_fabrinet_rule.month_colm_cout-1):
        if(search_fabrinet_rule.month_colm[k+1] - search_fabrinet_rule.month_colm[k] == 1):
            print_rule(first_sheet,key_row)
            return 1
    return 0        
