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
def check_Mitac_month(first_sheet, index):

    if atp_def.MONTH_SET == "JAN_FEB":
        month1 = re.compile(r'2017\/01\/[0-9][0-9]~01\/[0-9][0-9]', re.IGNORECASE)
        month2 = re.compile(r'2017\/02\/[0-9][0-9]~02\/[0-9][0-9]', re.IGNORECASE)
        #print "Mitac Rule JAN_FEB"
    elif atp_def.MONTH_SET == "FEB_MARCH":
        month1 = re.compile(r'2017\/02\/[0-9][0-9]~02\/[0-9][0-9]', re.IGNORECASE)
        month2 = re.compile(r'2017\/03\/[0-9][0-9]~03\/[0-9][0-9]', re.IGNORECASE)
        #print "Mitac Rule FEB_MARCH"
    
    elif atp_def.MONTH_SET == "MARCH_APL":
        month1 = re.compile(r'2017\/03\/[0-9][0-9]~03\/[0-9][0-9]', re.IGNORECASE)
        month2 = re.compile(r'2017\/04\/[0-9][0-9]~04\/[0-9][0-9]', re.IGNORECASE)
        #print "Mitac Rule MARCH_APL"
    elif atp_def.MONTH_SET == "MAY_JUNE":
        month1 = re.compile(r'2017\/05\/[0-9][0-9]~05\/[0-9][0-9]', re.IGNORECASE)
        month2 = re.compile(r'2017\/06\/[0-9][0-9]~06\/[0-9][0-9]', re.IGNORECASE)
        #print "Mitac Rule MAY JUNE"
    elif atp_def.MONTH_SET == "JUNE_JULY":
        month1 = re.compile(r'2017\/06\/[0-9][0-9]~06\/[0-9][0-9]', re.IGNORECASE)
        month2 = re.compile(r'2017\/07\/[0-9][0-9]~07\/[0-9][0-9]', re.IGNORECASE)
    elif atp_def.MONTH_SET == "JULY_AUGUST":
        month1 = re.compile(r'2017\/07\/[0-9][0-9]~07\/[0-9][0-9]', re.IGNORECASE)
        month2 = re.compile(r'2017\/08\/[0-9][0-9]~08\/[0-9][0-9]', re.IGNORECASE)
        #print "Mitac Rule JULY AUGUST"
    elif atp_def.MONTH_SET == "SEP_OCT":
        month1 = re.compile(r'2017\/09\/[0-9][0-9]~09\/[0-9][0-9]', re.IGNORECASE)
        month2 = re.compile(r'2017\/10\/[0-9][0-9]~10\/[0-9][0-9]', re.IGNORECASE)
        #print "Mitac Rule SEP_OCT"
    
    elif atp_def.MONTH_SET == "NOV_DEC":
        month1 = re.compile(r'2017\/11\/[0-9][0-9]~11\/[0-9][0-9]|2017\/10\/[0-9][0-9]~11\/[0-9][0-9]|2017\/11\/[0-9][0-9]~12\/[0-9][0-9]', re.IGNORECASE)
        month2 = re.compile(r'2017\/12\/[0-9][0-9]~12\/[0-9][0-9]|2017\/12\/[0-9][0-9]~01\/[0-9][0-9]', re.IGNORECASE)
        #print "Mitac Rule NOV_DEC"
    else:
        print "Mitac wrong Month"
        return 0

    row = first_sheet.row_values(search_mitac_rule.first_row)
    s = repr(row[search_mitac_rule.month_colm[index]])

    if month1.search(s):
        return 1
    if month2.search(s):
        return 2

    print "Mitac wrong Month"
    return 0

#----------------------------------------------------------------------
def print_junipter_rule(first_sheet, key_row):

    month1_sum = 0
    month2_sum = 0

    row = first_sheet.row_values(search_mitac_rule.first_row)
    for j in range(search_mitac_rule.month_colm_cout):
        print "%s " %(row[search_mitac_rule.month_colm[j]]),
    print ""

#Sum Calculation
    row = first_sheet.row_values(key_row)
    for j in range(search_mitac_rule.month_colm_cout):
        print "%s                " %(row[search_mitac_rule.month_colm[j]]),
        if(check_Mitac_month(first_sheet,j) == 1):
            # month1_sum = month1_sum + int(row[search_mitac_rule.month_colm[j]])
            month1_sum = month1_sum + check_if_string_num_add(row[search_mitac_rule.month_colm[j]])
        elif(check_Mitac_month(first_sheet,j) == 2):
            # month2_sum = month2_sum + int(row[search_mitac_rule.month_colm[j]])
            month2_sum = month2_sum + check_if_string_num_add(row[search_mitac_rule.month_colm[j]])
    print ""

    print "Month 1 sum = %d, Month 2 sum = %d" %(month1_sum,month2_sum)
#----------------------------------------------------------------------
def search_mitac_rule(first_sheet, key_row):
    """
    Open ic require file with part number
    """
    
    if atp_def.MONTH_SET == "JAN_FEB":
        month1 = re.compile(r'2017\/01\/[0-9][0-9]~01\/[0-9][0-9]', re.IGNORECASE)
        month2 = re.compile(r'2017\/02\/[0-9][0-9]~02\/[0-9][0-9]', re.IGNORECASE)
        #print "Mitac Rule JAN_FEB"

    elif atp_def.MONTH_SET == "FEB_MARCH":
        month1 = re.compile(r'2017\/02\/[0-9][0-9]~02\/[0-9][0-9]', re.IGNORECASE)
        month2 = re.compile(r'2017\/03\/[0-9][0-9]~03\/[0-9][0-9]', re.IGNORECASE)
        #print "Mitac Rule FEB_MARCH"
    
    elif atp_def.MONTH_SET == "MARCH_APL":
        month1 = re.compile(r'2017\/03\/[0-9][0-9]~03\/[0-9][0-9]', re.IGNORECASE)
        month2 = re.compile(r'2017\/04\/[0-9][0-9]~04\/[0-9][0-9]', re.IGNORECASE)
        #print "Mitac Rule MARCH_APL"
    elif atp_def.MONTH_SET == "MAY_JUNE":
        month1 = re.compile(r'2017\/05\/[0-9][0-9]~05\/[0-9][0-9]', re.IGNORECASE)
        month2 = re.compile(r'2017\/06\/[0-9][0-9]~06\/[0-9][0-9]', re.IGNORECASE)
        #print "Mitac Rule MAY JUNE"
    elif atp_def.MONTH_SET == "JUNE_JULY":
        month1 = re.compile(r'2017\/06\/[0-9][0-9]~06\/[0-9][0-9]', re.IGNORECASE)
        month2 = re.compile(r'2017\/07\/[0-9][0-9]~07\/[0-9][0-9]', re.IGNORECASE)
    elif atp_def.MONTH_SET == "JULY_AUGUST":
        month1 = re.compile(r'2017\/07\/[0-9][0-9]~07\/[0-9][0-9]', re.IGNORECASE)
        month2 = re.compile(r'2017\/08\/[0-9][0-9]~08\/[0-9][0-9]', re.IGNORECASE)
        #print "Mitac Rule JULY AUGUST"
    elif atp_def.MONTH_SET == "SEP_OCT":
        month1 = re.compile(r'2017\/09\/[0-9][0-9]~09\/[0-9][0-9]', re.IGNORECASE)
        month2 = re.compile(r'2017\/10\/[0-9][0-9]~10\/[0-9][0-9]', re.IGNORECASE)
        #print "Mitac Rule SEP_OCT"
    
    elif atp_def.MONTH_SET == "NOV_DEC":
        month1 = re.compile(r'2017\/11\/[0-9][0-9]~11\/[0-9][0-9]|2017\/10\/[0-9][0-9]~11\/[0-9][0-9]|2017\/11\/[0-9][0-9]~12\/[0-9][0-9]', re.IGNORECASE)
        month2 = re.compile(r'2017\/12\/[0-9][0-9]~12\/[0-9][0-9]|2017\/12\/[0-9][0-9]~01\/[0-9][0-9]', re.IGNORECASE)
        #print "Mitac Rule NOV_DEC"
    else:
        print "Mitac wrong Month"
 

    search_mitac_rule.month_colm = []
    search_mitac_rule.month_colm_cout = 0
    search_mitac_rule.first_row = 0


    for i in range(first_sheet.nrows):
        row = first_sheet.row_values(i)
        for j in range(len(row)):
            s = repr(row[j])
            if month1.search(s):
                #print "May Matched at colm[%d]" %(j)
                search_mitac_rule.month_colm.append(j)
                search_mitac_rule.month_colm_cout += 1
                search_mitac_rule.first_row = i
            if month2.search(s):
                #print "June Matched at colm[%d]" %(j)
                search_mitac_rule.month_colm.append(j)
                search_mitac_rule.month_colm_cout += 1
                search_mitac_rule.first_row = i

    if(search_mitac_rule.month_colm_cout < 1):
        return 0

    for k in range(0, search_mitac_rule.month_colm_cout-1):
        if(search_mitac_rule.month_colm[k+1] - search_mitac_rule.month_colm[k] == 1):
            print_junipter_rule(first_sheet,key_row)
            return 1
    return 0        
