import xlrd
import win32com.client as win32
import sqlite3
import csv
import os
import ConfigParser


headers1 = ['bcode', 'prog_name', 'division', 'rgt', 'cost_center', 'account', 'acct_desc', 'jan', 'feb', 'mar', 'apr',
            'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec', 'year_total']
headers2 = ['bcode', 'prog_name', 'division', 'rgt', 'cost_center', 'jan', 'feb', 'mar', 'apr', 'may', 'jun',
            'jul', 'aug', 'sep', 'oct', 'nov', 'dec', 'year_total']
rptFl1 = open('Proj_Forecast.csv','wb')
rptFl2 = open('Proj_Actuals.csv','wb')
writer1 = csv.writer(rptFl1,delimiter=',',quotechar='"',quoting=csv.QUOTE_ALL)
writer2 = csv.writer(rptFl2,delimiter=',',quotechar='"',quoting=csv.QUOTE_ALL)
writer1.writerow(headers1)
writer2.writerow(headers2)

# setup for using config file
config = ConfigParser.ConfigParser()
config.readfp(open(r'Fcst_vs_Act_by_Program_Configs.txt'))

# set configs to those defined in the config file
str_db_name = config.get('User Defined Vars', 'strDbName')
sv_fl_path = config.get('User Defined Vars', 'svflPath')
qry18_fl_path = config.get('User Defined Vars', 'qry18flPath')
user_rpt_nm = config.get('User Defined Vars', 'userRptNm')
user_mon = config.get('User Defined Vars', 'userMon')
user_year = config.get('User Defined Vars', 'userYear')
user_forecast = config.get('User Defined Vars', 'userForecast')


# function to identify any missing data from the report so that they can be easily identified
def check_missing_vals(ws):
    if ws.Name == 'Total':
        wsname = 'Total'
    elif ws.Name == 'IT':
        wsname = 'IT'
    elif ws.Name == 'Bus':
        wsname = 'Bus'
    missing_vals = {'division': [], 'program': [], 'rgt': []}
    counter = 0
    div = 'true'
    while div != '' and div is not None:
        counter += 1
        div = ws.Cells(2, 1).Offset(counter + 1, 1).Value
        bcode = ws.Cells(2, 2).Offset(counter + 1, 1).Value
        prog = ws.Cells(2, 3).Offset(counter + 1, 1).Value
        rgt = ws.Cells(2, 4).Offset(counter + 1, 1).Value
        if div == 'TBD - Need to Add Mapping':
            missing_vals['division'].append(bcode + ' missing division in worksheet ' + wsname + ', need to add '
                                            + 'mapping to code in function get_division')
        if prog == 'Bcode Not In QRY18':
            missing_vals['program'].append(bcode + ' missing program name in worksheet ' + wsname + ', may need to '
                                           + 'refresh QRY18 input data from PeopleSoft')
        if rgt == 'TBD - Need RGT':
            missing_vals['rgt'].append(bcode + ' missing rgt in worksheet ' + wsname + ', may need to add mapping '
                                       + 'to code in function get_rgt')
    return missing_vals


# function to cleanup blank cells in cells to default to 0
def clean_up_blanks(val):
    if val is None or val == '':
        clean_val = 0
    else:
        clean_val = val
    return clean_val


# function to get current month and ytd values from the Forecast Check worksheet in smartview for use in automated
# reconciliation against the sqllite database and final report
def get_fcst_check_vals(ws, user_mon):
    fcst_check_vals_split_d = {'cur_tot': 0, 'ytd_tot': 0, 'cur_it': 0, 'ytd_it': 0, 'cur_bus': 0, 'ytd_bus': 0}
    for row in range(ws.nrows):
        if ws.row(row)[2].value != '':
            var_txt1 = ws.row(row)[2].value.strip()
            var_l = var_txt1.split('-')
            acct = var_l[0]
            jan = clean_up_blanks(ws.row(row)[3].value)
            feb = clean_up_blanks(ws.row(row)[4].value)
            mar = clean_up_blanks(ws.row(row)[5].value)
            apr = clean_up_blanks(ws.row(row)[6].value)
            may = clean_up_blanks(ws.row(row)[7].value)
            jun = clean_up_blanks(ws.row(row)[8].value)
            jul = clean_up_blanks(ws.row(row)[9].value)
            aug = clean_up_blanks(ws.row(row)[10].value)
            sep = clean_up_blanks(ws.row(row)[11].value)
            oct = clean_up_blanks(ws.row(row)[12].value)
            nov = clean_up_blanks(ws.row(row)[13].value)
            dec = clean_up_blanks(ws.row(row)[14].value)
            if acct == '845009':
                if user_mon == 'jan':
                    fcst_check_vals_split_d['cur_tot'] = jan
                    fcst_check_vals_split_d['ytd_tot'] = jan
                elif user_mon == 'feb':
                    fcst_check_vals_split_d['cur_tot'] = feb
                    fcst_check_vals_split_d['ytd_tot'] = jan + feb
                elif user_mon == 'mar':
                    fcst_check_vals_split_d['cur_tot'] = mar
                    fcst_check_vals_split_d['ytd_tot'] = jan + feb + mar
                elif user_mon == 'apr':
                    fcst_check_vals_split_d['cur_tot'] = apr
                    fcst_check_vals_split_d['ytd_tot'] = jan + feb + mar + apr
                elif user_mon == 'may':
                    fcst_check_vals_split_d['cur_tot'] = may
                    fcst_check_vals_split_d['ytd_tot'] = jan + feb + mar + apr + may
                elif user_mon == 'jun':
                    fcst_check_vals_split_d['cur_tot'] = jun
                    fcst_check_vals_split_d['ytd_tot'] = jan + feb + mar + apr + may + jun
                elif user_mon == 'jul':
                    fcst_check_vals_split_d['cur_tot'] = jul
                    fcst_check_vals_split_d['ytd_tot'] = jan + feb + mar + apr + may + jun + jul
                elif user_mon == 'aug':
                    fcst_check_vals_split_d['cur_tot'] = aug
                    fcst_check_vals_split_d['ytd_tot'] = jan + feb + mar + apr + may + jun + jul + aug
                elif user_mon == 'sep':
                    fcst_check_vals_split_d['cur_tot'] = sep
                    fcst_check_vals_split_d['ytd_tot'] = jan + feb + mar + apr + may + jun + jul + aug + sep
                elif user_mon == 'oct':
                    fcst_check_vals_split_d['cur_tot'] = oct
                    fcst_check_vals_split_d['ytd_tot'] = jan + feb + mar + apr + may + jun + jul + aug + sep + oct
                elif user_mon == 'nov':
                    fcst_check_vals_split_d['cur_tot'] = nov
                    fcst_check_vals_split_d['ytd_tot'] = jan + feb + mar + apr + may + jun + jul + aug + sep + oct + nov
                elif user_mon == 'dec':
                    fcst_check_vals_split_d['cur_tot'] = dec
                    fcst_check_vals_split_d['ytd_tot'] = (jan + feb + mar + apr + may + jun + jul + aug + sep + oct +
                                                          nov + dec)
            elif acct == '845033' or acct == '845034' or acct == '845036' or acct == '845039':
                if user_mon == 'jan':
                    fcst_check_vals_split_d['cur_it'] += jan
                    fcst_check_vals_split_d['ytd_it'] += jan
                elif user_mon == 'feb':
                    fcst_check_vals_split_d['cur_it'] += feb
                    fcst_check_vals_split_d['ytd_it'] += jan + feb
                elif user_mon == 'mar':
                    fcst_check_vals_split_d['cur_it'] += mar
                    fcst_check_vals_split_d['ytd_it'] += jan + feb + mar
                elif user_mon == 'apr':
                    fcst_check_vals_split_d['cur_it'] += apr
                    fcst_check_vals_split_d['ytd_it'] += jan + feb + mar + apr
                elif user_mon == 'may':
                    fcst_check_vals_split_d['cur_it'] += may
                    fcst_check_vals_split_d['ytd_it'] += jan + feb + mar + apr + may
                elif user_mon == 'jun':
                    fcst_check_vals_split_d['cur_it'] += jun
                    fcst_check_vals_split_d['ytd_it'] += jan + feb + mar + apr + may + jun
                elif user_mon == 'jul':
                    fcst_check_vals_split_d['cur_it'] += jul
                    fcst_check_vals_split_d['ytd_it'] += jan + feb + mar + apr + may + jun + jul
                elif user_mon == 'aug':
                    fcst_check_vals_split_d['cur_it'] += aug
                    fcst_check_vals_split_d['ytd_it'] += jan + feb + mar + apr + may + jun + jul + aug
                elif user_mon == 'sep':
                    fcst_check_vals_split_d['cur_it'] += sep
                    fcst_check_vals_split_d['ytd_it'] += jan + feb + mar + apr + may + jun + jul + aug + sep
                elif user_mon == 'oct':
                    fcst_check_vals_split_d['cur_it'] += oct
                    fcst_check_vals_split_d['ytd_it'] += jan + feb + mar + apr + may + jun + jul + aug + sep + oct
                elif user_mon == 'nov':
                    fcst_check_vals_split_d['cur_it'] += nov
                    fcst_check_vals_split_d['ytd_it'] += jan + feb + mar + apr + may + jun + jul + aug + sep + oct + nov
                elif user_mon == 'dec':
                    fcst_check_vals_split_d['cur_it'] += dec
                    fcst_check_vals_split_d['ytd_it'] += (jan + feb + mar + apr + may + jun + jul + aug + sep + oct +
                                                          nov + dec)
            elif acct == '815003':
                if user_mon == 'jan':
                    fcst_check_vals_split_d['cur_bus'] = jan
                    fcst_check_vals_split_d['ytd_bus'] = jan
                elif user_mon == 'feb':
                    fcst_check_vals_split_d['cur_bus'] = feb
                    fcst_check_vals_split_d['ytd_bus'] = jan + feb
                elif user_mon == 'mar':
                    fcst_check_vals_split_d['cur_bus'] = mar
                    fcst_check_vals_split_d['ytd_bus'] = jan + feb + mar
                elif user_mon == 'apr':
                    fcst_check_vals_split_d['cur_bus'] = apr
                    fcst_check_vals_split_d['ytd_bus'] = jan + feb + mar + apr
                elif user_mon == 'may':
                    fcst_check_vals_split_d['cur_bus'] = may
                    fcst_check_vals_split_d['ytd_bus'] = jan + feb + mar + apr + may
                elif user_mon == 'jun':
                    fcst_check_vals_split_d['cur_bus'] = jun
                    fcst_check_vals_split_d['ytd_bus'] = jan + feb + mar + apr + may + jun
                elif user_mon == 'jul':
                    fcst_check_vals_split_d['cur_bus'] = jul
                    fcst_check_vals_split_d['ytd_bus'] = jan + feb + mar + apr + may + jun + jul
                elif user_mon == 'aug':
                    fcst_check_vals_split_d['cur_bus'] = aug
                    fcst_check_vals_split_d['ytd_bus'] = jan + feb + mar + apr + may + jun + jul + aug
                elif user_mon == 'sep':
                    fcst_check_vals_split_d['cur_bus'] = sep
                    fcst_check_vals_split_d['ytd_bus'] = jan + feb + mar + apr + may + jun + jul + aug + sep
                elif user_mon == 'oct':
                    fcst_check_vals_split_d['cur_bus'] = oct
                    fcst_check_vals_split_d['ytd_bus'] = jan + feb + mar + apr + may + jun + jul + aug + sep + oct
                elif user_mon == 'nov':
                    fcst_check_vals_split_d['cur_bus'] = nov
                    fcst_check_vals_split_d['ytd_bus'] = jan + feb + mar + apr + may + jun + jul + aug + sep + oct + nov
                elif user_mon == 'dec':
                    fcst_check_vals_split_d['cur_bus'] = dec
                    fcst_check_vals_split_d['ytd_bus'] = (jan + feb + mar + apr + may + jun + jul + aug + sep + oct +
                                                          nov + dec)
    return fcst_check_vals_split_d


# function to get current month and ytd values from the Actual Check worksheet in smartview for use in automated
# reconciliation against the sqllite database and final report
def get_act_check_vals(ws, user_mon):
    act_check_vals_split_d = {'cur_tot': 0, 'ytd_tot': 0, 'cur_it': 0, 'ytd_it': 0, 'cur_bus': 0, 'ytd_bus': 0}
    for row in range(ws.nrows):
        if ws.row(row)[0].value != '':
            div = ws.row(row)[0].value.strip()
            jan = clean_up_blanks(ws.row(row)[3].value)
            feb = clean_up_blanks(ws.row(row)[4].value)
            mar = clean_up_blanks(ws.row(row)[5].value)
            apr = clean_up_blanks(ws.row(row)[6].value)
            may = clean_up_blanks(ws.row(row)[7].value)
            jun = clean_up_blanks(ws.row(row)[8].value)
            jul = clean_up_blanks(ws.row(row)[9].value)
            aug = clean_up_blanks(ws.row(row)[10].value)
            sep = clean_up_blanks(ws.row(row)[11].value)
            oct = clean_up_blanks(ws.row(row)[12].value)
            nov = clean_up_blanks(ws.row(row)[13].value)
            dec = clean_up_blanks(ws.row(row)[14].value)
            if div == 'All Divisions':
                if user_mon == 'jan':
                    act_check_vals_split_d['cur_tot'] = jan
                    act_check_vals_split_d['ytd_tot'] = jan
                elif user_mon == 'feb':
                    act_check_vals_split_d['cur_tot'] = feb
                    act_check_vals_split_d['ytd_tot'] = jan + feb
                elif user_mon == 'mar':
                    act_check_vals_split_d['cur_tot'] = mar
                    act_check_vals_split_d['ytd_tot'] = jan + feb + mar
                elif user_mon == 'apr':
                    act_check_vals_split_d['cur_tot'] = apr
                    act_check_vals_split_d['ytd_tot'] = jan + feb + mar + apr
                elif user_mon == 'may':
                    act_check_vals_split_d['cur_tot'] = may
                    act_check_vals_split_d['ytd_tot'] = jan + feb + mar + apr + may
                elif user_mon == 'jun':
                    act_check_vals_split_d['cur_tot'] = jun
                    act_check_vals_split_d['ytd_tot'] = jan + feb + mar + apr + may + jun
                elif user_mon == 'jul':
                    act_check_vals_split_d['cur_tot'] = jul
                    act_check_vals_split_d['ytd_tot'] = jan + feb + mar + apr + may + jun + jul
                elif user_mon == 'aug':
                    act_check_vals_split_d['cur_tot'] = aug
                    act_check_vals_split_d['ytd_tot'] = jan + feb + mar + apr + may + jun + jul + aug
                elif user_mon == 'sep':
                    act_check_vals_split_d['cur_tot'] = sep
                    act_check_vals_split_d['ytd_tot'] = jan + feb + mar + apr + may + jun + jul + aug + sep
                elif user_mon == 'oct':
                    act_check_vals_split_d['cur_tot'] = oct
                    act_check_vals_split_d['ytd_tot'] = jan + feb + mar + apr + may + jun + jul + aug + sep + oct
                elif user_mon == 'nov':
                    act_check_vals_split_d['cur_tot'] = nov
                    act_check_vals_split_d['ytd_tot'] = jan + feb + mar + apr + may + jun + jul + aug + sep + oct + nov
                elif user_mon == 'dec':
                    act_check_vals_split_d['cur_tot'] = dec
                    act_check_vals_split_d['ytd_tot'] = (jan + feb + mar + apr + may + jun + jul + aug + sep + oct + nov
                                                         + dec)
            elif div == 'Info Tech':
                if user_mon == 'jan':
                    act_check_vals_split_d['cur_it'] = jan
                    act_check_vals_split_d['ytd_it'] = jan
                elif user_mon == 'feb':
                    act_check_vals_split_d['cur_it'] = feb
                    act_check_vals_split_d['ytd_it'] = jan + feb
                elif user_mon == 'mar':
                    act_check_vals_split_d['cur_it'] = mar
                    act_check_vals_split_d['ytd_it'] = jan + feb + mar
                elif user_mon == 'apr':
                    act_check_vals_split_d['cur_it'] = apr
                    act_check_vals_split_d['ytd_it'] = jan + feb + mar + apr
                elif user_mon == 'may':
                    act_check_vals_split_d['cur_it'] = may
                    act_check_vals_split_d['ytd_it'] = jan + feb + mar + apr + may
                elif user_mon == 'jun':
                    act_check_vals_split_d['cur_it'] = jun
                    act_check_vals_split_d['ytd_it'] = jan + feb + mar + apr + may + jun
                elif user_mon == 'jul':
                    act_check_vals_split_d['cur_it'] = jul
                    act_check_vals_split_d['ytd_it'] = jan + feb + mar + apr + may + jun + jul
                elif user_mon == 'aug':
                    act_check_vals_split_d['cur_it'] = aug
                    act_check_vals_split_d['ytd_it'] = jan + feb + mar + apr + may + jun + jul + aug
                elif user_mon == 'sep':
                    act_check_vals_split_d['cur_it'] = sep
                    act_check_vals_split_d['ytd_it'] = jan + feb + mar + apr + may + jun + jul + aug + sep
                elif user_mon == 'oct':
                    act_check_vals_split_d['cur_it'] = oct
                    act_check_vals_split_d['ytd_it'] = jan + feb + mar + apr + may + jun + jul + aug + sep + oct
                elif user_mon == 'nov':
                    act_check_vals_split_d['cur_it'] = nov
                    act_check_vals_split_d['ytd_it'] = jan + feb + mar + apr + may + jun + jul + aug + sep + oct + nov
                elif user_mon == 'dec':
                    act_check_vals_split_d['cur_it'] = dec
                    act_check_vals_split_d['ytd_it'] = (jan + feb + mar + apr + may + jun + jul + aug + sep + oct + nov
                                                        + dec)
            else:
                if user_mon == 'jan':
                    act_check_vals_split_d['cur_bus'] += jan
                    act_check_vals_split_d['ytd_bus'] += jan
                elif user_mon == 'feb':
                    act_check_vals_split_d['cur_bus'] += feb
                    act_check_vals_split_d['ytd_bus'] += jan + feb
                elif user_mon == 'mar':
                    act_check_vals_split_d['cur_bus'] += mar
                    act_check_vals_split_d['ytd_bus'] += jan + feb + mar
                elif user_mon == 'apr':
                    act_check_vals_split_d['cur_bus'] += apr
                    act_check_vals_split_d['ytd_bus'] += jan + feb + mar + apr
                elif user_mon == 'may':
                    act_check_vals_split_d['cur_bus'] += may
                    act_check_vals_split_d['ytd_bus'] += jan + feb + mar + apr + may
                elif user_mon == 'jun':
                    act_check_vals_split_d['cur_bus'] += jun
                    act_check_vals_split_d['ytd_bus'] += jan + feb + mar + apr + may + jun
                elif user_mon == 'jul':
                    act_check_vals_split_d['cur_bus'] += jul
                    act_check_vals_split_d['ytd_bus'] += jan + feb + mar + apr + may + jun + jul
                elif user_mon == 'aug':
                    act_check_vals_split_d['cur_bus'] += aug
                    act_check_vals_split_d['ytd_bus'] += jan + feb + mar + apr + may + jun + jul + aug
                elif user_mon == 'sep':
                    act_check_vals_split_d['cur_bus'] += sep
                    act_check_vals_split_d['ytd_bus'] += jan + feb + mar + apr + may + jun + jul + aug + sep
                elif user_mon == 'oct':
                    act_check_vals_split_d['cur_bus'] += oct
                    act_check_vals_split_d['ytd_bus'] += jan + feb + mar + apr + may + jun + jul + aug + sep + oct
                elif user_mon == 'nov':
                    act_check_vals_split_d['cur_bus'] += nov
                    act_check_vals_split_d['ytd_bus'] += jan + feb + mar + apr + may + jun + jul + aug + sep + oct + nov
                elif user_mon == 'dec':
                    act_check_vals_split_d['cur_bus'] += dec
                    act_check_vals_split_d['ytd_bus'] += (jan + feb + mar + apr + may + jun + jul + aug + sep + oct +
                                                          nov + dec)
    return act_check_vals_split_d


# function to determine the correct division associated with a project; the mappings included here may need to be
# updated each time the report is run depending on changes to the project tree
def get_division(cursor, bcode):
    sql = "SELECT DISTINCT parent_initiative FROM qry18 WHERE bcode = ?;"
    cursor.execute(sql, (bcode,))
    row = cursor.fetchone()
    if row is not None:
        division = row[0]
    else:
        if bcode.find(' Run ') != -1:
            division = bcode[:bcode.find(' Run ')]
        elif bcode.find(' Grow ') != -1:
            division = bcode[:bcode.find(' Grow ')]
        elif bcode.find(' Transform ') != -1:
            division = bcode[:bcode.find(' Transform ')]
        else:
            division = bcode[:bcode.find(' Project TBD ')]
    if division[:1] == 'B':
        sql = "SELECT DISTINCT bcode FROM qry18 WHERE rcode = ?;"
        cursor.execute(sql, (bcode,))
        row = cursor.fetchone()
        if row is not None:
            division = row[0]
        else:
            division = 'Cannot Find Bcode'
    if division == 'INI_SF':
        division2 = 'SF'
    elif division == 'INI_LEGAL':
        division2 = 'Legal'
    elif division == 'INI_MULTIFAMILY':
        division2 = 'MF'
    elif division == 'INI_ICM':
        division2 = 'ICM'
    elif division == 'INI_HR':
        division2 = 'HR'
    elif division == 'INI_FINANCE':
        division2 = 'Finance'
    elif division == 'INI_ERM':
        division2 = 'ERM'
    elif division == 'INI_COMPLIANCE':
        division2 = 'Compliance'
    elif division == 'INI_SS':
        division2 = 'SF - SS'
    elif division == 'INI_LAS':
        division2 = 'SF - LAS'
    elif division == 'INI_ADMIN':
        division2 = 'Admin'
    elif division == 'INI_CSP':
        division2 = 'SF - CSP'
    elif division == 'INI_IR':
        division2 = 'SF - IR'
    elif division == 'INI_CRT':
        division2 = 'SF - CRT'
    elif division == 'INI_PECASH':
        division2 = 'SF - PECASH'
    elif division == 'INI_IT_OPTIMIZE':
        division2 = 'IT Optimization'
    elif division == 'INI_IT_INFOSEC':
        division2 = 'IT Info Security'
    elif division == 'Multifamily':
        division2 = 'MF'
    elif division == 'CSP':
        division2 = 'SF - CSP'
    elif division == 'SS':
        division2 = 'SF - SS'
    elif division == 'LAS':
        division2 = 'SF - LAS'
    elif division == 'IR':
        division2 = 'SF - IR'
    elif division == 'CRT':
        division2 = 'SF - CRT'
    elif division == 'Admin':
        division2 = 'Admin'
    elif division == 'SF':
        division2 = 'SF'
    elif division == 'Compliance':
        division2 = 'Compliance'
    elif division == 'IT Optimization':
        division2 = 'IT Optimization'
    elif division == 'INI_CORPORATE':
        division2 = 'Corporate Contingency'
    else:
        division2 = 'TBD - Need to Add Mapping'
    return division2


# function to get the YTD string dynamically for the report headers in Excel
def get_ytd_year_for_report(user_year):
    ytd_str = 'YTD ' + user_year + ' ($Ks)'
    return ytd_str


# function to get the month and year dynamically for the report headers in Excel
def get_mon_yr_for_report(user_mon, user_year):
    if user_mon == 'jan':
        rpt_mon = 'Jan ' + user_year + ' ($Ks)'
    elif user_mon == 'feb':
        rpt_mon = 'Feb ' + user_year + ' ($Ks)'
    elif user_mon == 'mar':
        rpt_mon = 'Mar ' + user_year + ' ($Ks)'
    elif user_mon == 'apr':
        rpt_mon = 'Apr ' + user_year + ' ($Ks)'
    elif user_mon == 'may':
        rpt_mon = 'May ' + user_year + ' ($Ks)'
    elif user_mon == 'jun':
        rpt_mon = 'Jun ' + user_year + ' ($Ks)'
    elif user_mon == 'jul':
        rpt_mon = 'Jul ' + user_year + ' ($Ks)'
    elif user_mon == 'aug':
        rpt_mon = 'Aug ' + user_year + ' ($Ks)'
    elif user_mon == 'sep':
        rpt_mon = 'Sep ' + user_year + ' ($Ks)'
    elif user_mon == 'oct':
        rpt_mon = 'Oct ' + user_year + ' ($Ks)'
    elif user_mon == 'nov':
        rpt_mon = 'Nov ' + user_year + ' ($Ks)'
    elif user_mon == 'dec':
        rpt_mon = 'Dec ' + user_year + ' ($Ks)'
    return rpt_mon


# function to make SQL for creating the temp tables dynamic
def get_cur_month_for_sql(user_mon, type):
    if type == 'fcst':
        tbl = 'f'
        mon_str = ', SUM(' + tbl + '.' + user_mon + ') AS tot_cur_mon'
    elif type == 'act':
        tbl = 'a'
        mon_str = ', SUM(' + tbl + '.' + user_mon + ') AS tot_cur_mon'
    return mon_str


# function to make SQL for creating the temp tables dynamic
def get_ytd_for_sql(user_mon, type):
    if type == 'fcst':
        tbl = 'f'
    elif type == 'act':
        tbl = 'a'
    if user_mon == 'jan':
        ytd_str = ', SUM(' + tbl + '.jan) AS tot_ytd'
    elif user_mon == 'feb':
        ytd_str = ', SUM(' + tbl + '.jan + ' + tbl + '.feb) AS tot_ytd'
    elif user_mon == 'mar':
        ytd_str = ', SUM(' + tbl + '.jan + ' + tbl + '.feb + ' + tbl + '.mar) AS tot_ytd'
    elif user_mon == 'apr':
        ytd_str = ', SUM(' + tbl + '.jan + ' + tbl + '.feb + ' + tbl + '.mar + ' + tbl + '.apr) AS tot_ytd'
    elif user_mon == 'may':
        ytd_str = ', SUM(' + tbl + '.jan + ' + tbl + '.feb + ' + tbl + '.mar + ' + tbl + '.apr + ' + tbl + \
                  '.may) AS tot_ytd'
    elif user_mon == 'jun':
        ytd_str = ', SUM(' + tbl + '.jan + ' + tbl + '.feb + ' + tbl + '.mar + ' + tbl + '.apr + ' + tbl + '.may + ' + \
                  tbl + '.jun) AS tot_ytd'
    elif user_mon == 'jul':
        ytd_str = ', SUM(' + tbl + '.jan + ' + tbl + '.feb + ' + tbl + '.mar + ' + tbl + '.apr + ' + tbl + '.may + ' + \
                  tbl + '.jun + ' + tbl + '.jul) AS tot_ytd'
    elif user_mon == 'aug':
        ytd_str = ', SUM(' + tbl + '.jan + ' + tbl + '.feb + ' + tbl + '.mar + ' + tbl + '.apr + ' + tbl + '.may + ' + \
                  tbl + '.jun + ' + tbl + '.jul + ' + tbl + '.aug) AS tot_ytd'
    elif user_mon == 'sep':
        ytd_str = ', SUM(' + tbl + '.jan + ' + tbl + '.feb + ' + tbl + '.mar + ' + tbl + '.apr + ' + tbl + '.may + ' + \
                  tbl + '.jun + ' + tbl + '.jul + ' + tbl + '.aug + ' + tbl + '.sep) AS tot_ytd'
    elif user_mon == 'oct':
        ytd_str = ', SUM(' + tbl + '.jan + ' + tbl + '.feb + ' + tbl + '.mar + ' + tbl + '.apr + ' + tbl + '.may + ' + \
                  tbl + '.jun + ' + tbl + '.jul + ' + tbl + '.aug + ' + tbl + '.sep + ' + tbl + '.oct) AS tot_ytd'
    elif user_mon == 'nov':
        ytd_str = ', SUM(' + tbl + '.jan + ' + tbl + '.feb + ' + tbl + '.mar + ' + tbl + '.apr + ' + tbl + '.may + ' + \
                  tbl + '.jun + ' + tbl + '.jul + ' + tbl + '.aug + ' + tbl + '.sep + ' + tbl + '.oct + ' + tbl + \
                  '.nov) AS tot_ytd'
    elif user_mon == 'dec':
        ytd_str = ', SUM(' + tbl + '.jan + ' + tbl + '.feb + ' + tbl + '.mar + ' + tbl + '.apr + ' + tbl + '.may + ' + \
                  tbl + '.jun + ' + tbl + '.jul + ' + tbl + '.aug + ' + tbl + '.sep + ' + tbl + '.oct + ' + tbl + \
                  '.nov + ' + tbl + '.dec) AS tot_ytd'
    return ytd_str


# function to build the temp table for forecast
def create_fcst_temp_tbl(cursor, user_mon):
    sql = '''INSERT INTO fcst_temp SELECT tot_fcst.division
            , tot_fcst.bcode
            , tot_fcst.prog_name
            , tot_fcst.rgt
            , tot_fcst.tot_cur_mon
            , CASE
                WHEN it_fcst.tot_cur_mon IS NULL THEN 0
                WHEN it_fcst.tot_cur_mon IS NOT NULL THEN it_fcst.tot_cur_mon
            END AS it_cur_mon
            , CASE
                WHEN bus_fcst.tot_cur_mon IS NULL THEN 0
                WHEN bus_fcst.tot_cur_mon IS NOT NULL THEN bus_fcst.tot_cur_mon
            END AS bus_cur_mon
            , tot_fcst.tot_ytd
            , CASE
                WHEN it_fcst.tot_ytd iS NULL THEN 0
                WHEN it_fcst.tot_ytd iS NOT NULL THEN it_fcst.tot_ytd
            END AS it_ytd
            , CASE
                WHEN bus_fcst.tot_ytd IS NULL THEN 0
                WHEN bus_fcst.tot_ytd IS NOT NULL THEN bus_fcst.tot_ytd
            END AS bus_ytd
            FROM
                (SELECT f.bcode
                , f.prog_name
                , f.division
                , f.rgt
                ''' + get_cur_month_for_sql(user_mon, 'fcst') + '''
                ''' + get_ytd_for_sql(user_mon, 'fcst') + '''
                FROM proj_forecast AS f
                WHERE f.account IN (845009)
                GROUP BY f.bcode, f.prog_name, f.division, f.rgt
                ) As tot_fcst
            LEFT OUTER JOIN
                (SELECT f.bcode
                , f.prog_name
                , f.division
                , f.rgt
                ''' + get_cur_month_for_sql(user_mon, 'fcst') + '''
                ''' + get_ytd_for_sql(user_mon, 'fcst') + '''
                FROM proj_forecast AS f
                WHERE f.account IN (845033,845034,845036,845039)
                GROUP BY f.bcode, f.prog_name, f.division, f.rgt
                ) AS it_fcst
            ON tot_fcst.bcode = it_fcst.bcode
            LEFT OUTER JOIN
                (SELECT f.bcode
                , f.prog_name
                , f.division
                , f.rgt
                ''' + get_cur_month_for_sql(user_mon, 'fcst') + '''
                ''' + get_ytd_for_sql(user_mon, 'fcst') + '''
                FROM proj_forecast AS f
                WHERE f.account in (815003)
                GROUP BY f.bcode, f.prog_name, f.division, f.rgt
                ) AS bus_fcst
            ON tot_fcst.bcode = bus_fcst.bcode
            ORDER by tot_fcst.division, tot_fcst.bcode
            '''
    cursor.execute(sql)


# function to build the temp table for actuals
def create_act_temp_tbl(cursor, user_mon):
    sql = '''INSERT INTO act_temp SELECT tot_act.division
            , tot_act.bcode
            , tot_act.prog_name
            , tot_act.rgt
            , tot_act.tot_cur_mon
            , CASE
                WHEN it_act.tot_cur_mon IS NULL THEN 0
                WHEN it_act.tot_cur_mon IS NOT NULL THEN it_act.tot_cur_mon
            END AS it_cur_mon
            , CASE
                WHEN bus_act.tot_cur_mon IS NULL THEN 0
                WHEN bus_act.tot_cur_mon IS NOT NULL THEN bus_act.tot_cur_mon
            END AS bus_cur_mon
            , tot_act.tot_ytd
            , CASE
                WHEN it_act.tot_ytd IS NULL THEN 0
                WHEN it_act.tot_ytd IS NOT NULL THEN it_act.tot_ytd
            END AS it_ytd
            , CASE
                WHEN bus_act.tot_ytd IS NULL THEN 0
                WHEN bus_act.tot_ytd IS NOT NULL THEN bus_act.tot_ytd
            END AS bus_ytd
            FROM
                (SELECT a.bcode
                , a.prog_name
                , a.division
                , a.rgt
                ''' + get_cur_month_for_sql(user_mon, 'act') + '''
                ''' + get_ytd_for_sql(user_mon, 'act') + '''
                FROM proj_actuals AS a
                GROUP BY a.bcode, a.prog_name, a.division, a.rgt
                ) AS tot_act
            LEFT OUTER JOIN
                (SELECT a.bcode
                , a.prog_name
                , a.division
                , a.rgt
                ''' + get_cur_month_for_sql(user_mon, 'act') + '''
                ''' + get_ytd_for_sql(user_mon, 'act') + '''
                FROM proj_actuals AS a
                WHERE a.cost_center IN ('Info Tech')
                GROUP BY a.bcode, a.prog_name, a.division, a.rgt
                ) AS it_act
            ON tot_act.bcode = it_act.bcode
            LEFT OUTER JOIN
                (SELECT a.bcode
                , a.prog_name
                , a.division
                , a.rgt
                ''' + get_cur_month_for_sql(user_mon, 'act') + '''
                ''' + get_ytd_for_sql(user_mon, 'act') + '''
                FROM proj_actuals AS a
                WHERE a.cost_center NOT IN ('Info Tech')
                GROUP BY a.bcode, a.prog_name, a.division, a.rgt
                ) AS bus_act
            ON tot_act.bcode = bus_act.bcode
            ORDER BY tot_act.division, tot_act.bcode
            '''
    cursor.execute(sql)


# function to determine and get the Run, Grow, or Transform expense classification
def get_rgt(cursor, prog):
    if prog[:1] == 'R':
        rgt = 'Run'
    elif prog[:1] == 'G':
        rgt = 'Grow'
    elif prog[:1] == 'T':
        rgt = 'Transform'
    else:
        if prog.find('Run') != -1:
            rgt = 'Run'
        elif prog.find('Grow') != -1:
            rgt = 'Grow'
        elif prog.find('Transform') != -1:
            rgt = 'Transform'
        elif prog.find('IT') != -1:
            rgt = 'Run'
        elif prog == 'Initiative Holdback':
            rgt = 'N/A'
        else:
            rgt = 'TBD - Need RGT'
    return rgt


# function to get the program name from qry18 instead of smartview
def get_prog_name(cursor, bcode):
    sql = "SELECT DISTINCT prog_name FROM qry18 WHERE bcode = ?;"
    cursor.execute(sql, (bcode,))
    row = cursor.fetchone()
    if row is not None:
        prog_nm = row[0]
    else:
        sql = "SELECT DISTINCT proj_name FROM qry18 WHERE rcode = ?;"
        cursor.execute(sql, (bcode,))
        row = cursor.fetchone()
        if row is not None:
            prog_nm = row[0]
        else:
            prog_nm = 'Bcode Not In QRY18'
    return prog_nm


# create new sqlite database
cnn = sqlite3.connect(str_db_name)
cursor = cnn.cursor()

# create table to load data into and delete all records if it exists
cursor.execute('''CREATE TABLE if not exists proj_forecast (bcode TEXT, prog_name TEXT, division TEXT, rgt TEXT,
                cost_center TEXT, account INTEGER, acct_desc TEXT, jan REAL, feb REAL, mar REAL, apr REAL, may REAL, jun REAL,
                jul REAL, aug REAL, sep REAL, oct REAL, nov REAL, dec REAL, year_total REAL);
                ''')
cursor.execute('DELETE FROM proj_forecast;')
cursor.execute('''CREATE TABLE if not exists proj_actuals (bcode TEXT, prog_name TEXT, division TEXT, rgt TEXT,
                cost_center TEXT, jan REAL, feb REAL, mar REAL, apr REAL, may REAL, jun REAL,
                jul REAL, aug REAL, sep REAL, oct REAL, nov REAL, dec REAL, year_total REAL);
                ''')
cursor.execute('DELETE FROM proj_actuals;')
cursor.execute('''CREATE TABLE if not exists qry18 (parent_initiative TEXT, initiative_desc TEXT, bcode TEXT,
                prog_name TEXT, rcode TEXT, proj_name TEXT, cc TEXT, proj_type TEXT, proj_status TEXT,
                proj_st_dt NUMERIC, proj_end_dt NUMERIC);
                ''')
cursor.execute('DELETE FROM qry18;')
cursor.execute('''CREATE TABLE if not exists fcst_temp (division TEXT, bcode TEXT, prog_name TEXT, rgt TEXT,
                tot_cur_mon REAL, it_cur_mon REAL, bus_cur_mon REAL, tot_ytd REAL, it_ytd REAL, bus_ytd REAL);
                ''')
cursor.execute('DELETE FROM fcst_temp;')
cursor.execute('''CREATE TABLE if not exists act_temp (division TEXT, bcode TEXT, prog_name TEXT, rgt TEXT,
                tot_cur_mon REAL, it_cur_mon REAL, bus_cur_mon REAL, tot_ytd REAL, it_ytd REAL, bus_ytd REAL);
                ''')
cursor.execute('DELETE FROM act_temp;')

counter = 0

# load qry18 data into a table in the sqllite database
with xlrd.open_workbook(qry18_fl_path) as wb:
    for ws in wb.sheets():
        if ws.name == 'sheet1':
            for row in range(ws.nrows):
                counter += 1
                if counter >= 5:
                    parinit = ws.row(row)[0].value
                    initdesc = ws.row(row)[1].value
                    bcode = ws.row(row)[2].value
                    prognm = ws.row(row)[3].value
                    rcode = ws.row(row)[4].value
                    projnm = ws.row(row)[5].value
                    cc = ws.row(row)[6].value
                    projtype = ws.row(row)[7].value
                    projstat = ws.row(row)[8].value
                    stdt = ws.row(row)[9].value
                    enddt = ws.row(row)[10].value
                    cursor.execute('''INSERT INTO qry18 (parent_initiative, initiative_desc, bcode, prog_name, rcode,
                                    proj_name, cc, proj_type, proj_status, proj_st_dt, proj_end_dt) VALUES
                                    (?,?,?,?,?,?,?,?,?,?,?);'''
                                   , (parinit, initdesc, bcode, prognm, rcode, projnm, cc, projtype, projstat, stdt,
                                      enddt))

counter = 0
# define blank dictionary or list for data storage
check_vals_tot_fcst = {}
check_vals_tot_act = []

# read through the Hyperion smartview query data and load into the sqllite database; includes both forecast and actuals
with xlrd.open_workbook(sv_fl_path) as wb:
    for ws in wb.sheets():
        if ws.name == 'Proj Forecast':
            for row in range(ws.nrows):
                if ws.row(row)[0].value != "":
                    vartxt1 = ws.row(row)[1].value.strip()
                    if vartxt1[0] == 'B':
                        varL = []
                        varL = vartxt1.split('-')
                        bcode = varL[0]
                        prog = get_prog_name(cursor, bcode)
                    else:
                        bcode = vartxt1
                        prog = bcode
                    division = get_division(cursor, bcode)
                    rgt = get_rgt(cursor, prog)
                    vartxt1 = ws.row(row)[2].value.strip()
                    varL = vartxt1.split('-')
                    acct = varL[0]
                    acctnm = varL[1]
                    cc = ws.row(row)[0].value
                    jan = ws.row(row)[3].value
                    feb = ws.row(row)[4].value
                    mar = ws.row(row)[5].value
                    apr = ws.row(row)[6].value
                    may = ws.row(row)[7].value
                    jun = ws.row(row)[8].value
                    jul = ws.row(row)[9].value
                    aug = ws.row(row)[10].value
                    sep = ws.row(row)[11].value
                    oct = ws.row(row)[12].value
                    nov = ws.row(row)[13].value
                    dec = ws.row(row)[14].value
                    yeartot = ws.row(row)[15].value
                    cursor.execute('''INSERT INTO proj_forecast (bcode, prog_name, division, rgt, cost_center, account,
                                    acct_desc, jan, feb, mar, apr, may, jun, jul, aug, sep, oct, nov, dec, year_total)
                                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?);'''
                                   , (bcode, prog, division, rgt, cc, acct, acctnm, jan, feb, mar, apr, may, jun, jul,
                                      aug, sep, oct, nov, dec, yeartot))
        elif ws.name == 'Proj Actuals':
            for row in range(ws.nrows):
                if ws.row(row)[1].value != "" and ws.row(row)[1].value != 'Iniative Alternte Project':
                    cc = ws.row(row)[0].value.strip()
                    varTxt1 = ws.row(row)[1].value.strip()
                    varL = []
                    varL = varTxt1.split("-")
                    bcode = varL[0]
                    prog = get_prog_name(cursor, bcode)
                    division = get_division(cursor, bcode)
                    rgt = get_rgt(cursor, prog)
                    jan = ws.row(row)[3].value
                    feb = ws.row(row)[4].value
                    mar = ws.row(row)[5].value
                    apr = ws.row(row)[6].value
                    may = ws.row(row)[7].value
                    jun = ws.row(row)[8].value
                    jul = ws.row(row)[9].value
                    aug = ws.row(row)[10].value
                    sep = ws.row(row)[11].value
                    oct = ws.row(row)[12].value
                    nov = ws.row(row)[13].value
                    dec = ws.row(row)[14].value
                    yeartot = ws.row(row)[15].value
                    cursor.execute('''INSERT INTO proj_actuals (bcode, prog_name, division, rgt, cost_center, jan, feb,
                                    mar, apr, may, jun, jul, aug, sep, oct, nov, dec, year_total) VALUES
                                    (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?);'''
                                   , (bcode, prog, division, rgt, cc, jan, feb, mar, apr, may, jun, jul, aug, sep, oct,
                                      nov, dec, yeartot))
        elif ws.name == 'Forecast Check':
            fcst_check_vals_splits = get_fcst_check_vals(ws, user_mon)
            for row in range(ws.nrows):
                if ws.row(row)[1].value != "":
                    vartxt1 = ws.row(row)[2].value.strip()
                    varL = []
                    varL = vartxt1.split('-')
                    acct = varL[0]
                    if acct not in check_vals_tot_fcst:
                        check_vals_tot_fcst[acct] = []
                        for x in range(0, 13):
                            check_vals_tot_fcst[acct].append(ws.row(row)[3 + x].value)
        elif ws.name == 'Actuals Check':
            act_check_vals_splits = get_act_check_vals(ws, user_mon)
            for row in range(ws.nrows):
                if ws.row(row)[0].value == 'All Divisions':
                    for x in range(0, 13):
                        check_vals_tot_act.append(ws.row(row)[3 + x].value)

create_fcst_temp_tbl(cursor, user_mon)
create_act_temp_tbl(cursor, user_mon)

# create excel file and format it for writing the report
xl = win32.gencache.EnsureDispatch('Excel.Application')
xl.DisplayAlerts = False
xl.Visible = False
wb = xl.Workbooks.Add()
ws3 = xl.Worksheets('Sheet3')
ws2 = xl.Worksheets('Sheet2')
ws1 = xl.Worksheets('Sheet1')
ws4 = wb.Worksheets.Add()
ws5 = wb.Worksheets.Add()
ws6 = wb.Worksheets.Add()
ws1.Name = 'Total'
ws2.Name = 'IT'
ws3.Name = 'Bus'
ws4.Name = 'Fcst Checks'
ws5.Name = 'Act Checks'
ws6.Name = 'Missing Data'

# this is to create the worksheet for the forecast vs actuals report for total by program
headers1 = ['', '', '', '', '', get_mon_yr_for_report(user_mon, user_year), '', '', get_ytd_year_for_report(user_year)
    , '', '']
counter = 0
for val in headers1:
    counter += 1
    ws1.Cells(1, 1).Offset(1, counter).Value = val
    ws1.Cells(1, 1).Offset(1, counter).Font.Bold = True
    ws2.Cells(1, 1).Offset(1, counter).Value = val
    ws2.Cells(1, 1).Offset(1, counter).Font.Bold = True
    ws3.Cells(1, 1).Offset(1, counter).Value = val
    ws3.Cells(1, 1).Offset(1, counter).Font.Bold = True
headers2 = ['Division', 'Bcode', 'Program', 'RGT', 'Category', user_forecast, 'Actual', 'Variance', user_forecast,
            'Actual', 'Variance']
counter = 0
for val in headers2:
    counter += 1
    ws1.Cells(2, 1).Offset(1, counter).Value = val
    ws1.Cells(2, 1).Offset(1, counter).Font.Bold = True
    ws1.Cells(2, 1).Offset(1, counter).Font.Underline = win32.constants.xlUnderlineStyleSingle
    ws2.Cells(2, 1).Offset(1, counter).Value = val
    ws2.Cells(2, 1).Offset(1, counter).Font.Bold = True
    ws2.Cells(2, 1).Offset(1, counter).Font.Underline = win32.constants.xlUnderlineStyleSingle
    ws3.Cells(2, 1).Offset(1, counter).Value = val
    ws3.Cells(2, 1).Offset(1, counter).Font.Bold = True
    ws3.Cells(2, 1).Offset(1, counter).Font.Underline = win32.constants.xlUnderlineStyleSingle
    if counter <= 5:
        ws1.Cells(2, 1).Offset(1, counter).Interior.ThemeColor = win32.constants.xlThemeColorDark1
        ws1.Cells(2, 1).Offset(1, counter).Interior.TintAndShade = -0.249977111117893
        ws2.Cells(2, 1).Offset(1, counter).Interior.ThemeColor = win32.constants.xlThemeColorDark1
        ws2.Cells(2, 1).Offset(1, counter).Interior.TintAndShade = -0.249977111117893
        ws3.Cells(2, 1).Offset(1, counter).Interior.ThemeColor = win32.constants.xlThemeColorDark1
        ws3.Cells(2, 1).Offset(1, counter).Interior.TintAndShade = -0.249977111117893
    if val == user_forecast:
        ws1.Cells(2, 1).Offset(1, counter).Interior.Color = 10092441
        ws2.Cells(2, 1).Offset(1, counter).Interior.Color = 10092441
        ws3.Cells(2, 1).Offset(1, counter).Interior.Color = 10092441
    elif val == 'Actual':
        ws1.Cells(2, 1).Offset(1, counter).Interior.Color = 16772300
        ws2.Cells(2, 1).Offset(1, counter).Interior.Color = 16772300
        ws3.Cells(2, 1).Offset(1, counter).Interior.Color = 16772300
    elif val == 'Variance':
        ws1.Cells(2, 1).Offset(1, counter).Interior.Color = 16764159
        ws2.Cells(2, 1).Offset(1, counter).Interior.Color = 16764159
        ws3.Cells(2, 1).Offset(1, counter).Interior.Color = 16764159
ws1.Range("F1:H1").MergeCells = True
ws1.Range("F1:H1").HorizontalAlignment = win32.constants.xlCenter
ws1.Range("I1:K1").MergeCells = True
ws1.Range("I1:K1").HorizontalAlignment = win32.constants.xlCenter
ws2.Range("F1:H1").MergeCells = True
ws2.Range("F1:H1").HorizontalAlignment = win32.constants.xlCenter
ws2.Range("I1:K1").MergeCells = True
ws2.Range("I1:K1").HorizontalAlignment = win32.constants.xlCenter
ws3.Range("F1:H1").MergeCells = True
ws3.Range("F1:H1").HorizontalAlignment = win32.constants.xlCenter
ws3.Range("I1:K1").MergeCells = True
ws3.Range("I1:K1").HorizontalAlignment = win32.constants.xlCenter

sql = '''SELECT *
    FROM
    (SELECT a.division
    , a.bcode
    , a.prog_name
    , a.rgt
    , 'Forecasted w/Actuals' AS category
    , f.tot_cur_mon AS fcst_tot_cur_mon
    , a.tot_cur_mon AS act_tot_cur_mon
    , (f.tot_cur_mon - a.tot_cur_mon) AS var_tot_cur_mon
    , f.tot_ytd AS fcst_tot_ytd
    , a.tot_ytd AS act_tot_ytd
    , (f.tot_ytd - a.tot_ytd) AS var_tot_ytd
    FROM fcst_temp AS f
    INNER JOIN act_temp As a
    ON f.bcode = a.bcode
    UNION ALL
    SELECT f.division
    , f.bcode
    , f.prog_name
    , f.rgt
    , 'Forecasted w/o Actuals' AS category
    , f.tot_cur_mon AS fcst_tot_cur_mon
    , 0 AS act_tot_cur_mon
    , (f.tot_cur_mon - 0) AS var_tot_cur_mon
    , f.tot_ytd AS fcst_tot_ytd
    , 0 AS act_tot_ytd
    , (f.tot_ytd - 0) AS var_tot_ytd
    FROM fcst_temp AS f
    LEFT OUTER JOIN act_temp As a
    ON f.bcode = a.bcode
    WHERE a.bcode IS NULL
    UNION ALL
    SELECT a.division
    , a.bcode
    , a.prog_name
    , a.rgt
    , 'Actuals w/o Forecast' AS category
    , 0 AS fcst_tot_cur_mon
    , a.tot_cur_mon AS act_tot_cur_mon
    , (0 - a.tot_cur_mon) AS var_tot_cur_mon
    , 0 AS fcst_tot_ytd
    , a.tot_ytd AS act_tot_ytd
    , (0 - a.tot_ytd) AS var_tot_ytd
    FROM act_temp AS a
    LEFT OUTER JOIN fcst_temp As f
    ON f.bcode = a.bcode
    WHERE f.bcode IS NULL
    ) AS subq1
    ORDER BY subq1.division, subq1.bcode, subq1.prog_name, subq1.category
    '''

fcst_check_vals_rpt = {'cur_tot': 0, 'ytd_tot': 0, 'cur_it': 0, 'ytd_it': 0, 'cur_bus': 0, 'ytd_bus': 0}
act_check_vals_rpt = {'cur_tot': 0, 'ytd_tot': 0, 'cur_it': 0, 'ytd_it': 0, 'cur_bus': 0, 'ytd_bus': 0}

rowCounter = 0
# write the output from sql to the Excel file in the Totals worksheet
for row in cursor.execute(sql):
    rowCounter += 1
    valCounter = 0
    for val in row:
        valCounter += 1
        ws1.Cells(2 + rowCounter, valCounter).Value = val
        if valCounter == 6:
            fcst_check_vals_rpt['cur_tot'] += val
        elif valCounter == 7:
            act_check_vals_rpt['cur_tot'] += val
        elif valCounter == 9:
            fcst_check_vals_rpt['ytd_tot'] += val
        elif valCounter == 10:
            act_check_vals_rpt['ytd_tot'] += val
        ws1.Cells(2 + rowCounter, valCounter).NumberFormat = "$#,###,;($#,###,)"
ws1.Columns.AutoFit()

sql = '''SELECT *
    FROM
    (SELECT a.division
    , a.bcode
    , a.prog_name
    , a.rgt
    , 'Forecasted w/Actuals' AS category
    , f.it_cur_mon AS fcst_it_cur_mon
    , a.it_cur_mon AS act_it_cur_mon
    , (f.it_cur_mon - a.it_cur_mon) AS var_it_cur_mon
    , f.it_ytd AS fcst_it_ytd
    , a.it_ytd AS act_it_ytd
    , (f.it_ytd - a.it_ytd) AS var_it_ytd
    FROM fcst_temp AS f
    INNER JOIN act_temp As a
    ON f.bcode = a.bcode
    UNION ALL
    SELECT f.division
    , f.bcode
    , f.prog_name
    , f.rgt
    , 'Forecasted w/o Actuals' AS category
    , f.it_cur_mon AS fcst_it_cur_mon
    , 0 AS act_it_cur_mon
    , (f.it_cur_mon - 0) AS var_it_cur_mon
    , f.it_ytd AS fcst_it_ytd
    , 0 AS act_it_ytd
    , (f.it_ytd - 0) AS var_it_ytd
    FROM fcst_temp AS f
    LEFT OUTER JOIN act_temp As a
    ON f.bcode = a.bcode
    WHERE a.bcode IS NULL
    UNION ALL
    SELECT a.division
    , a.bcode
    , a.prog_name
    , a.rgt
    , 'Actuals w/o Forecast' AS category
    , 0 AS fcst_it_cur_mon
    , a.it_cur_mon AS act_it_cur_mon
    , (0 - a.it_cur_mon) AS var_it_cur_mon
    , 0 AS fcst_it_ytd
    , a.it_ytd AS act_it_ytd
    , (0 - a.it_ytd) AS var_it_ytd
    FROM act_temp AS a
    LEFT OUTER JOIN fcst_temp As f
    ON f.bcode = a.bcode
    WHERE f.bcode IS NULL
    ) AS subq1
    ORDER BY subq1.division, subq1.bcode, subq1.prog_name, subq1.category
    '''

rowCounter = 0
# write the output from sql to the Excel file in the Totals worksheet
for row in cursor.execute(sql):
    rowCounter += 1
    valCounter = 0
    for val in row:
        valCounter += 1
        ws2.Cells(2 + rowCounter, valCounter).Value = val
        if valCounter == 6:
            fcst_check_vals_rpt['cur_it'] += val
        elif valCounter == 7:
            act_check_vals_rpt['cur_it'] += val
        elif valCounter == 9:
            fcst_check_vals_rpt['ytd_it'] += val
        elif valCounter == 10:
            act_check_vals_rpt['ytd_it'] += val
        ws2.Cells(2 + rowCounter, valCounter).NumberFormat = "$#,###,;($#,###,)"
ws2.Columns.AutoFit()

sql = '''SELECT *
    FROM
    (SELECT a.division
    , a.bcode
    , a.prog_name
    , a.rgt
    , 'Forecasted w/Actuals' AS category
    , f.bus_cur_mon AS fcst_bus_cur_mon
    , a.bus_cur_mon As act_bus_cur_mon
    , (f.bus_cur_mon - a.bus_cur_mon) AS var_bus_cur_mon
    , f.bus_ytd AS fcst_bus_ytd
    , a.bus_ytd AS act_bus_ytd
    , (f.bus_ytd - a.bus_ytd) AS var_bus_ytd
    FROM fcst_temp AS f
    INNER JOIN act_temp As a
    ON f.bcode = a.bcode
    UNION ALL
    SELECT f.division
    , f.bcode
    , f.prog_name
    , f.rgt
    , 'Forecasted w/o Actuals' AS category
    , f.bus_cur_mon AS fcst_bus_cur_mon
    , 0 As act_bus_cur_mon
    , (f.bus_cur_mon - 0) AS var_bus_cur_mon
    , f.bus_ytd AS fcst_bus_ytd
    , 0 AS act_bus_ytd
    , (f.bus_ytd - 0) AS var_bus_ytd
    FROM fcst_temp AS f
    LEFT OUTER JOIN act_temp As a
    ON f.bcode = a.bcode
    WHERE a.bcode IS NULL
    UNION ALL
    SELECT a.division
    , a.bcode
    , a.prog_name
    , a.rgt
    , 'Actuals w/o Forecast' AS category
    , 0 AS fcst_bus_cur_mon
    , a.bus_cur_mon As act_bus_cur_mon
    , (0 - a.bus_cur_mon) AS var_bus_cur_mon
    , 0 AS fcst_bus_ytd
    , a.bus_ytd AS act_bus_ytd
    , (0 - a.bus_ytd) AS var_bus_ytd
    FROM act_temp AS a
    LEFT OUTER JOIN fcst_temp As f
    ON f.bcode = a.bcode
    WHERE f.bcode IS NULL
    ) AS subq1
    ORDER BY subq1.division, subq1.bcode, subq1.prog_name, subq1.category
    '''

rowCounter = 0
# write the output from sql to the Excel file in the Totals worksheet
for row in cursor.execute(sql):
    rowCounter += 1
    valCounter = 0
    for val in row:
        valCounter += 1
        ws3.Cells(2 + rowCounter, valCounter).Value = val
        if valCounter == 6:
            fcst_check_vals_rpt['cur_bus'] += val
        elif valCounter == 7:
            act_check_vals_rpt['cur_bus'] += val
        elif valCounter == 9:
            fcst_check_vals_rpt['ytd_bus'] += val
        elif valCounter == 10:
            act_check_vals_rpt['ytd_bus'] += val
        ws3.Cells(2 + rowCounter, valCounter).NumberFormat = "$#,###,;($#,###,)"
ws3.Columns.AutoFit()

counter = 0
# this is the check whether there are any missing values in the report and output to the missing data worksheet
missing_vals = check_missing_vals(ws1)
for key in missing_vals:
    for val in missing_vals[key]:
        counter += 1
        ws6.Cells(1, 1).Offset(counter, 1).Value = val
missing_vals = check_missing_vals(ws2)
for key in missing_vals:
    for val in missing_vals[key]:
        counter += 1
        ws6.Cells(1, 1).Offset(counter, 1).Value = val
missing_vals = check_missing_vals(ws3)
for key in missing_vals:
    for val in missing_vals[key]:
        counter += 1
        ws6.Cells(1, 1).Offset(counter, 1).Value = val
if counter == 0:
    ws6.Cells(1, 1).Offset(1, 1).Value = 'No Missing Data'

# this is to create the check worksheet for automatic reconciliation of the forecast data between the staging table in
# the sqllite database and an independent check run in Smartview
headers1 = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'FY 17']
counter = 0
for val in headers1:
    counter += 1
    ws4.Cells(1, 1).Offset(1, counter + 1).Value = val
    ws4.Cells(1, 1).Offset(4, counter + 1).FormulaR1C1 = "=R[-2]C-R[-1]C"
    ws4.Cells(1, 1).Offset(8, counter + 1).FormulaR1C1 = "=R[-2]C-R[-1]C"
    ws4.Cells(1, 1).Offset(12, counter + 1).FormulaR1C1 = "=R[-2]C-R[-1]C"
    ws4.Cells(1, 1).Offset(16, counter + 1).FormulaR1C1 = "=R[-2]C-R[-1]C"
    ws4.Cells(1, 1).Offset(20, counter + 1).FormulaR1C1 = "=R[-2]C-R[-1]C"
    ws4.Cells(1, 1).Offset(24, counter + 1).FormulaR1C1 = "=R[-2]C-R[-1]C"

# this is to create the headers for the automatic reocnciliation between the temp table and the final report to check
# total, business, and IT split
headers2 = ['Tot-Cur Mon', 'IT-Cur Mon', 'Bus-Cur Mon', 'Tot-YTD', 'IT-YTD', 'Bus-YTD']
counter = 0
for val in headers2:
    counter += 1
    ws4.Cells(1, 1).Offset(26, counter + 1).Value = val
    ws5.Cells(1, 1).Offset(6, counter + 1).Value = val
    ws4.Cells(1, 1).Offset(30, counter + 1).FormulaR1C1 = "=R[-3]C-R[-2]C"
    ws4.Cells(1, 1).Offset(31, counter + 1).FormulaR1C1 = "=R[-4]C-R[-2]C"
    ws4.Cells(1, 1).Offset(32, counter + 1).FormulaR1C1 = "=R[-4]C-R[-3]C"
    ws5.Cells(1, 1).Offset(10, counter + 1).FormulaR1C1 = "=R[-3]C-R[-2]C"
    ws5.Cells(1, 1).Offset(11, counter + 1).FormulaR1C1 = "=R[-4]C-R[-2]C"
    ws5.Cells(1, 1).Offset(12, counter + 1).FormulaR1C1 = "=R[-4]C-R[-3]C"

for key in fcst_check_vals_splits:
    if key == 'cur_tot':
        ws4.Cells(1, 1).Offset(29, 1 + 1).Value = fcst_check_vals_splits['cur_tot']
    elif key == 'cur_it':
        ws4.Cells(1, 1).Offset(29, 2 + 1).Value = fcst_check_vals_splits['cur_it']
    elif key == 'cur_bus':
        ws4.Cells(1, 1).Offset(29, 3 + 1).Value = fcst_check_vals_splits['cur_bus']
    elif key == 'ytd_tot':
        ws4.Cells(1, 1).Offset(29, 4 + 1).Value = fcst_check_vals_splits['ytd_tot']
    elif key == 'ytd_it':
        ws4.Cells(1, 1).Offset(29, 5 + 1).Value = fcst_check_vals_splits['ytd_it']
    elif key == 'ytd_bus':
        ws4.Cells(1, 1).Offset(29, 6 + 1).Value = fcst_check_vals_splits['ytd_bus']

for key in fcst_check_vals_rpt:
    if key == 'cur_tot':
        ws4.Cells(1, 1).Offset(27, 1 + 1).Value = fcst_check_vals_rpt['cur_tot']
    elif key == 'cur_it':
        ws4.Cells(1, 1).Offset(27, 2 + 1).Value = fcst_check_vals_rpt['cur_it']
    elif key == 'cur_bus':
        ws4.Cells(1, 1).Offset(27, 3 + 1).Value = fcst_check_vals_rpt['cur_bus']
    elif key == 'ytd_tot':
        ws4.Cells(1, 1).Offset(27, 4 + 1).Value = fcst_check_vals_rpt['ytd_tot']
    elif key == 'ytd_it':
        ws4.Cells(1, 1).Offset(27, 5 + 1).Value = fcst_check_vals_rpt['ytd_it']
    elif key == 'ytd_bus':
        ws4.Cells(1, 1).Offset(27, 6 + 1).Value = fcst_check_vals_rpt['ytd_bus']

# checks for account 845009
ws4.Cells(2, 1).Value = '845009-Db Staging Tbl'
ws4.Cells(3, 1).Value = '845009-SV Check'
ws4.Cells(4, 1).Value = '845009-Var'

counter = 0
for v in check_vals_tot_fcst['845009']:
    counter += 1
    ws4.Cells(1, 1).Offset(3, 1 + counter).Value = v

# checks for account 845033
ws4.Cells(6, 1).Value = '845033-Db Staging Tbl'
ws4.Cells(7, 1).Value = '845033-SV Check'
ws4.Cells(8, 1).Value = '845033-Var'

counter = 0
for v in check_vals_tot_fcst['845033']:
    counter += 1
    ws4.Cells(1, 1).Offset(7, 1 + counter).Value = v

# checks for account 845034
ws4.Cells(10, 1).Value = '845034-Db Staging Tbl'
ws4.Cells(11, 1).Value = '845034-SV Check'
ws4.Cells(12, 1).Value = '845034-Var'

counter = 0
for v in check_vals_tot_fcst['845034']:
    counter += 1
    ws4.Cells(1, 1).Offset(11, 1 + counter).Value = v

# checks for account 845036
ws4.Cells(14, 1).Value = '845036-Db Staging Tbl'
ws4.Cells(15, 1).Value = '845036-SV Check'
ws4.Cells(16, 1).Value = '845036-Var'

counter = 0
for v in check_vals_tot_fcst['845036']:
    counter += 1
    ws4.Cells(1, 1).Offset(15, 1 + counter).Value = v

# checks for account 845039
ws4.Cells(18, 1).Value = '845039-Db Staging Tbl'
ws4.Cells(19, 1).Value = '845039-SV Check'
ws4.Cells(20, 1).Value = '845039-Var'

counter = 0
for v in check_vals_tot_fcst['845039']:
    counter += 1
    ws4.Cells(1, 1).Offset(19, 1 + counter).Value = v

# checks for account 815003
ws4.Cells(22, 1).Value = '815003-Db Staging Tbl'
ws4.Cells(23, 1).Value = '815003-SV Check'
ws4.Cells(24, 1).Value = '815003-Var'

counter = 0
for v in check_vals_tot_fcst['815003']:
    counter += 1
    ws4.Cells(1, 1).Offset(23, 1 + counter).Value = v

sql = '''SELECT f.account
    , SUM(f.jan) AS jan
    , SUM(f.feb) AS feb
    , SUM(f.mar) AS mar
    , SUM(f.apr) AS apr
    , SUM(f.may) AS may
    , SUM(f.jun) AS jun
    , SUM(f.jul) AS jul
    , SUM(f.aug) AS aug
    , SUM(f.sep) AS sep
    , SUM(f.oct) AS oct
    , SUM(f.nov) AS nov
    , SUM(f.dec) AS dec
    , SUM(f.year_total) AS fy
    FROM proj_forecast f
    GROUP BY f.account
    ORDER BY f.account;
    '''

for row in cursor.execute(sql):
    valCounter = 0
    for val in row:
        valCounter += 1
        if valCounter == 1:
            acct = val
        else:
            if acct == 845009:
                ws4.Cells(1, 1).Offset(2, valCounter).Value = val
            elif acct == 845033:
                ws4.Cells(1, 1).Offset(6, valCounter).Value = val
            elif acct == 845034:
                ws4.Cells(1, 1).Offset(10, valCounter).Value = val
            elif acct == 845036:
                ws4.Cells(1, 1).Offset(14, valCounter).Value = val
            elif acct == 845039:
                ws4.Cells(1, 1).Offset(18, valCounter).Value = val
            elif acct == 815003:
                ws4.Cells(1, 1).Offset(22, valCounter).Value = val

# checks for reconciling report values against the aggregation temp table in the sqllite database
ws4.Cells(27, 1).Value = 'Report'
ws4.Cells(28, 1).Value = 'Db Temp Tbl'
ws4.Cells(29, 1).Value = 'SmartView'
ws4.Cells(30, 1).Value = 'Var-Rpt vs Db'
ws4.Cells(31, 1).Value = 'Var-Rpt vs SV'
ws4.Cells(32, 1).Value = 'Var-Db vs SV'

ws5.Cells(7, 1).Value = 'Report'
ws5.Cells(8, 1).Value = 'Db Temp Tbl'
ws5.Cells(9, 1).Value = 'SmartView'
ws5.Cells(10, 1).Value = 'Var-Rpt vs Db'
ws5.Cells(11, 1).Value = 'Var-Rpt vs SV'
ws5.Cells(12, 1).Value = 'Var-Db vs SV'

counter = 0
for key in act_check_vals_splits:
    if key == 'cur_tot':
        ws5.Cells(1, 1).Offset(9, 1 + 1).Value = act_check_vals_splits['cur_tot']
    elif key == 'cur_it':
        ws5.Cells(1, 1).Offset(9, 2 + 1).Value = act_check_vals_splits['cur_it']
    elif key == 'cur_bus':
        ws5.Cells(1, 1).Offset(9, 3 + 1).Value = act_check_vals_splits['cur_bus']
    elif key == 'ytd_tot':
        ws5.Cells(1, 1).Offset(9, 4 + 1).Value = act_check_vals_splits['ytd_tot']
    elif key == 'ytd_it':
        ws5.Cells(1, 1).Offset(9, 5 + 1).Value = act_check_vals_splits['ytd_it']
    elif key == 'ytd_bus':
        ws5.Cells(1, 1).Offset(9, 6 + 1).Value = act_check_vals_splits['ytd_bus']

counter = 0
for key in act_check_vals_rpt:
    if key == 'cur_tot':
        ws5.Cells(1, 1).Offset(7, 1 + 1).Value = act_check_vals_rpt['cur_tot']
    elif key == 'cur_it':
        ws5.Cells(1, 1).Offset(7, 2 + 1).Value = act_check_vals_rpt['cur_it']
    elif key == 'cur_bus':
        ws5.Cells(1, 1).Offset(7, 3 + 1).Value = act_check_vals_rpt['cur_bus']
    elif key == 'ytd_tot':
        ws5.Cells(1, 1).Offset(7, 4 + 1).Value = act_check_vals_rpt['ytd_tot']
    elif key == 'ytd_it':
        ws5.Cells(1, 1).Offset(7, 5 + 1).Value = act_check_vals_rpt['ytd_it']
    elif key == 'ytd_bus':
        ws5.Cells(1, 1).Offset(7, 6 + 1).Value = act_check_vals_rpt['ytd_bus']

sql = '''SELECT SUM(f.tot_cur_mon) AS tot_cur_mon
    , SUM(f.it_cur_mon) AS it_cur_mon
    , SUM(f.bus_cur_mon) AS bus_cur_mon
    , SUM(f.tot_ytd) AS tot_ytd
    , SUM(f.it_ytd) AS it_ytd
    , SUM(f.bus_ytd) AS bus_ytd
    FROM fcst_temp f;
    '''

counter = 0
for row in cursor.execute(sql):
    for val in row:
        counter += 1
        ws4.Cells(1, 1).Offset(28, counter + 1).Value = val

sql = '''SELECT SUM(a.tot_cur_mon) AS tot_cur_mon
    , SUM(a.it_cur_mon) AS it_cur_mon
    , SUM(a.bus_cur_mon) AS bus_cur_mon
    , SUM(a.tot_ytd) AS tot_ytd
    , SUM(a.it_ytd) AS it_ytd
    , SUM(a.bus_ytd) AS bus_ytd
    FROM act_temp a;
    '''

counter = 0
for row in cursor.execute(sql):
    for val in row:
        counter += 1
        ws5.Cells(1, 1).Offset(8, counter + 1).Value = val

ws4.Columns.AutoFit()
ws5.Columns.AutoFit()

# this is to create the check worksheet for automatic reconciliation of the actuals data
headers1 = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'FY 17']
counter = 0
for val in headers1:
    counter += 1
    ws5.Cells(1, 1).Offset(1, counter + 1).Value = val
    ws5.Cells(1, 1).Offset(4, counter + 1).FormulaR1C1 = "=R[-2]C-R[-1]C"

ws5.Cells(2, 1).Value = 'Db Staging Tbl'
ws5.Cells(3, 1).Value = 'SV Check'
ws5.Cells(4, 1).Value = 'Var'

counter = 0
for v in check_vals_tot_act:
    counter += 1
    ws5.Cells(1, 1).Offset(3, 1 + counter).Value = v

sql = '''SELECT SUM(a.jan) AS jan
    , SUM(a.feb) AS feb
    , SUM(a.mar) AS mar
    , SUM(a.apr) AS apr
    , SUM(a.may) AS may
    , SUM(a.jun) AS jun
    , SUM(a.jul) AS jul
    , SUM(a.aug) AS aug
    , SUM(a.sep) AS sep
    , SUM(a.oct) AS oct
    , SUM(a.nov) AS nov
    , SUM(a.dec) AS dec
    , SUM(a.year_total) AS fy
    FROM proj_actuals a;
    '''

for row in cursor.execute(sql):
    valCounter = 0
    for val in row:
        valCounter += 1
        ws5.Cells(1, 1).Offset(2, valCounter + 1).Value = val

ws5.Columns.AutoFit()

wb.SaveAs(user_rpt_nm)
xl.Application.Quit()

#create feeder table reports
sql = "SELECT * FROM proj_forecast;"
for row in cursor.execute(sql):
    writer1.writerow(row)
sql = "SELECT * FROM proj_actuals;"
for row in cursor.execute(sql):
    writer2.writerow(row)
rptFl1.close()
rptFl2.close()

# cleanup
cursor.close()
cnn.commit()
cnn.close()
os.remove(str_db_name)

print "------------------------------------Script Successfully Run-----------------------------------------"
print "\n"
