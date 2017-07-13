import xlrd
import win32com.client as win32
import os
import sqlite3
import ConfigParser

# setup for using config file
config = ConfigParser.ConfigParser()
config.readfp(open(r'Fcst_by_Program_Configs.txt'))

# set configs to those defined in the config file
strDbName = config.get('User Defined Vars','strDbName')
svflPath = config.get('User Defined Vars','svflPath')
qry18flPath = config.get('User Defined Vars','qry18flPath')
userRptNm = config.get('User Defined Vars','userRptNm')
userYear = config.get('User Defined Vars','userYear')


# function to identify any missing data from the report so that they can be easily identified
def check_missing_vals(ws):
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
            missing_vals['division'].append(bcode + ' missing division in report, need to add '
                                            + 'mapping to code in function get_division')
        if prog == 'Bcode Not In QRY18':
            missing_vals['program'].append(bcode + ' missing program name in report, may need to '
                                           + 'refresh QRY18 input data from PeopleSoft')
        if rgt == 'TBD - Need RGT':
            missing_vals['rgt'].append(bcode + ' missing rgt in report, may need to add mapping '
                                       + 'to code in function get_rgt')
    return missing_vals


# function to determine and get the Run, Grow, or Transform expense classification
def getRGT(cursor, prog):
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
def getProgName(cursor, bcode):
    sql = "SELECT DISTINCT prog_name FROM qry18 WHERE bcode = ?;"
    cursor.execute(sql, (bcode,))
    row = cursor.fetchone()
    if row != None:
        progNm = row[0]
    else:
        sql = "SELECT DISTINCT proj_name FROM qry18 WHERE rcode = ?;"
        cursor.execute(sql, (bcode,))
        row = cursor.fetchone()
        if row != None:
            progNm = row[0]
        else:
            progNm = 'Bcode Not In QRY18'
    return progNm

# create new sqlite database
cnn = sqlite3.connect(strDbName)
cursor = cnn.cursor()

# create tables to load forecast data and qry18 data and delete all records if they exist
cursor.execute('''CREATE TABLE if not exists proj_forecast (bcode TEXT, prog_name TEXT, division TEXT, rgt TEXT, account INTEGER, acct_desc TEXT, jan REAL, feb REAL, mar REAL, apr REAL, may REAL, jun REAL,
                jul REAL, aug REAL, sep REAL, oct REAL, nov REAL, dec REAL, year_total REAL);
                ''')
cursor.execute('DELETE FROM proj_forecast;')
cursor.execute('''CREATE TABLE if not exists qry18 (parent_initiative TEXT, initiative_desc TEXT, bcode TEXT, prog_name TEXT, rcode TEXT, proj_name TEXT, cc TEXT, proj_type TEXT, proj_status TEXT, proj_st_dt NUMERIC, proj_end_dt NUMERIC);
                ''')
cursor.execute('DELETE FROM qry18;')

# set counter to 0 used to keep track of which row the script is currently working with
counter = 0

# load qry18 data into a table in the sqllite database
with xlrd.open_workbook(qry18flPath) as wb:
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
                    cursor.execute("INSERT INTO qry18 (parent_initiative, initiative_desc, bcode, prog_name, rcode, proj_name, cc, proj_type, proj_status, proj_st_dt, proj_end_dt) VALUES (?,?,?,?,?,?,?,?,?,?,?);"
                                   , (parinit, initdesc, bcode, prognm, rcode, projnm, cc, projtype, projstat, stdt, enddt))

counter = 0
checkVals = {}

# read through the Hyperion smartview query data and load into the sqllite database; includes ONLY forecast
with xlrd.open_workbook(svflPath) as wb:
    for ws in wb.sheets():
        if ws.name == 'Proj_Forecast ' + userYear:
            for row in range(ws.nrows):
                if ws.row(row)[1].value != "":
                    vartxt1 = ws.row(row)[1].value.strip()
                    if vartxt1[0] == 'B':
                        varL = []
                        varL = vartxt1.split('-')
                        bcode = varL[0]
                        prog = getProgName(cursor, bcode)
                    else:
                        bcode = vartxt1
                        prog = bcode
                    division = ws.row(row)[0].value
                    rgt = getRGT(cursor, prog)
                    vartxt1 = ws.row(row)[2].value.strip()
                    varL = vartxt1.split('-')
                    acct = varL[0]
                    acctnm = varL[1]
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
                    cursor.execute("INSERT INTO proj_forecast (bcode, prog_name, division, rgt, account, acct_desc, jan, feb, mar, apr, may, jun, jul, aug, sep, oct, nov, dec, year_total) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?);",
                                   (bcode, prog, division, rgt, acct, acctnm, jan, feb, mar, apr, may, jun, jul, aug, sep, oct, nov, dec, yeartot))
        elif ws.name == 'Check ' + userYear:
            for row in range(ws.nrows):
                if ws.row(row)[1].value != "":
                    vartxt1 = ws.row(row)[2].value.strip()
                    varL = []
                    varL = vartxt1.split('-')
                    acct = varL[0]
                    if acct not in checkVals:
                        checkVals[acct] = []
                        for x in range(0,13):
                            checkVals[acct].append(ws.row(row)[3 + x].value)

# query used to aggregate the data and write the report
sql = '''SELECT tot_fcst.division
        , tot_fcst.bcode
        , tot_fcst.prog_name
        , tot_fcst.rgt
        , CASE
            WHEN tot_fcst.tot_jan IS NULL THEN 0
            WHEN tot_fcst.tot_jan IS NOT NULL THEN tot_fcst.tot_jan
        END AS tot_jan
        , CASE
            WHEN it_fcst.it_jan IS NULL THEN 0
            WHEN it_fcst.it_jan IS NOT NULL THEN it_fcst.it_jan
        END AS it_jan
        , CASE
            WHEN bus_fcst.bus_jan IS NULL THEN 0
            WHEN bus_fcst.bus_jan IS NOT NULL THEN bus_fcst.bus_jan
        END AS bus_jan
        , CASE
            WHEN tot_fcst.tot_feb IS NULL THEN 0
            WHEN tot_fcst.tot_feb IS NOT NULL THEN tot_fcst.tot_feb
        END AS tot_feb
        , CASE
            WHEN it_fcst.it_feb IS NULL THEN 0
            WHEN it_fcst.it_feb IS NOT NULL THEN it_fcst.it_feb
        END AS it_feb
        , CASE
            WHEN bus_fcst.bus_feb IS NULL THEN 0
            WHEN bus_fcst.bus_feb IS NOT NULL THEN bus_fcst.bus_feb
        END AS bus_feb
        , CASE
            WHEN tot_fcst.tot_mar IS NULL THEN 0
            WHEN tot_fcst.tot_mar IS NOT NULL THEN tot_fcst.tot_mar
        END AS tot_mar
        , CASE
            WHEN it_fcst.it_mar IS NULL THEN 0
            WHEN it_fcst.it_mar IS NOT NULL THEN it_fcst.it_mar
        END AS it_mar
        , CASE
            WHEN bus_fcst.bus_mar IS NULL THEN 0
            WHEN bus_fcst.bus_mar IS NOT NULL THEN bus_fcst.bus_mar
        END AS bus_mar
        , CASE
            WHEN tot_fcst.tot_apr IS NULL THEN 0
            WHEN tot_fcst.tot_apr IS NOT NULL THEN tot_fcst.tot_apr
        END AS tot_apr
        , CASE
            WHEN it_fcst.it_apr IS NULL THEN 0
            WHEN it_fcst.it_apr IS NOT NULL THEN it_fcst.it_apr
        END AS it_apr
        , CASE
            WHEN bus_fcst.bus_apr IS NULL THEN 0
            WHEN bus_fcst.bus_apr IS NOT NULL THEN bus_fcst.bus_apr
        END AS bus_apr
        , CASE
            WHEN tot_fcst.tot_may IS NULL THEN 0
            WHEN tot_fcst.tot_may IS NOT NULL THEN tot_fcst.tot_may
        END AS tot_may
        , CASE
            WHEN it_fcst.it_may IS NULL THEN 0
            WHEN it_fcst.it_may IS NOT NULL THEN it_fcst.it_may
        END AS it_may
        , CASE
            WHEN bus_fcst.bus_may IS NULL THEN 0
            WHEN bus_fcst.bus_may IS NOT NULL THEN bus_fcst.bus_may
        END AS bus_may
        , CASE
            WHEN tot_fcst.tot_jun IS NULL THEN 0
            WHEN tot_fcst.tot_jun IS NOT NULL THEN tot_fcst.tot_jun
        END AS tot_jun
        , CASE
            WHEN it_fcst.it_jun IS NULL THEN 0
            WHEN it_fcst.it_jun IS NOT NULL THEN it_fcst.it_jun
        END AS it_jun
        , CASE
            WHEN bus_fcst.bus_jun IS NULL THEN 0
            WHEN bus_fcst.bus_jun IS NOT NULL THEN bus_fcst.bus_jun
        END AS bus_jun
        , CASE
            WHEN tot_fcst.tot_jul IS NULL THEN 0
            WHEN tot_fcst.tot_jul IS NOT NULL THEN tot_fcst.tot_jul
        END AS tot_jul
        , CASE
            WHEN it_fcst.it_jul IS NULL THEN 0
            WHEN it_fcst.it_jul IS NOT NULL THEN it_fcst.it_jul
        END AS it_jul
        , CASE
            WHEN bus_fcst.bus_jul IS NULL THEN 0
            WHEN bus_fcst.bus_jul IS NOT NULL THEN bus_fcst.bus_jul
        END AS bus_jul
        , CASE
            WHEN tot_fcst.tot_aug IS NULL THEN 0
            WHEN tot_fcst.tot_aug IS NOT NULL THEN tot_fcst.tot_aug
        END AS tot_aug
        , CASE
            WHEN it_fcst.it_aug IS NULL THEN 0
            WHEN it_fcst.it_aug IS NOT NULL THEN it_fcst.it_aug
        END AS it_aug
        , CASE
            WHEN bus_fcst.bus_aug IS NULL THEN 0
            WHEN bus_fcst.bus_aug IS NOT NULL THEN bus_fcst.bus_aug
        END AS bus_aug
        , CASE
            WHEN tot_fcst.tot_sep IS NULL THEN 0
            WHEN tot_fcst.tot_sep IS NOT NULL THEN tot_fcst.tot_sep
        END AS tot_sep
        , CASE
            WHEN it_fcst.it_sep IS NULL THEN 0
            WHEN it_fcst.it_sep IS NOT NULL THEN it_fcst.it_sep
        END AS it_sep
        , CASE
            WHEN bus_fcst.bus_sep IS NULL THEN 0
            WHEN bus_fcst.bus_sep IS NOT NULL THEN bus_fcst.bus_sep
        END AS bus_sep
        , CASE
            WHEN tot_fcst.tot_oct IS NULL THEN 0
            WHEN tot_fcst.tot_oct IS NOT NULL THEN tot_fcst.tot_oct
        END AS tot_oct
        , CASE
            WHEN it_fcst.it_oct IS NULL THEN 0
            WHEN it_fcst.it_oct IS NOT NULL THEN it_fcst.it_oct
        END AS it_oct
        , CASE
            WHEN bus_fcst.bus_oct IS NULL THEN 0
            WHEN bus_fcst.bus_oct IS NOT NULL THEN bus_fcst.bus_oct
        END AS bus_oct
        , CASE
            WHEN tot_fcst.tot_nov IS NULL THEN 0
            WHEN tot_fcst.tot_nov IS NOT NULL THEN tot_fcst.tot_nov
        END AS tot_nov
        , CASE
            WHEN it_fcst.it_nov IS NULL THEN 0
            WHEN it_fcst.it_nov IS NOT NULL THEN it_fcst.it_nov
        END AS it_nov
        , CASE
            WHEN bus_fcst.bus_nov IS NULL THEN 0
            WHEN bus_fcst.bus_nov IS NOT NULL THEN bus_fcst.bus_nov
        END AS bus_nov
        , CASE
            WHEN tot_fcst.tot_dec IS NULL THEN 0
            WHEN tot_fcst.tot_dec IS NOT NULL THEN tot_fcst.tot_dec
        END AS tot_dec
        , CASE
            WHEN it_fcst.it_dec IS NULL THEN 0
            WHEN it_fcst.it_dec IS NOT NULL THEN it_fcst.it_dec
        END AS it_dec
        , CASE
            WHEN bus_fcst.bus_dec IS NULL THEN 0
            WHEN bus_fcst.bus_dec IS NOT NULL THEN bus_fcst.bus_dec
        END AS bus_dec
        , CASE
            WHEN tot_fcst.tot_fy IS NULL THEN 0
            WHEN tot_fcst.tot_fy IS NOT NULL THEN tot_fcst.tot_fy
        END AS tot_fy
        , CASE
            WHEN it_fcst.it_fy IS NULL THEN 0
            WHEN it_fcst.it_fy IS NOT NULL THEN it_fcst.it_fy
        END AS it_fy
        , CASE
            WHEN bus_fcst.bus_fy IS NULL THEN 0
            WHEN bus_fcst.bus_fy IS NOT NULL THEN bus_fcst.bus_fy
        END AS bus_fy
        FROM
            (SELECT f.bcode
            , f.prog_name
            , f.division
            , f.rgt
            , SUM(f.jan) AS tot_jan
            , SUM(f.feb) AS tot_feb
            , SUM(f.mar) AS tot_mar
            , SUM(f.apr) AS tot_apr
            , SUM(f.may) AS tot_may
            , SUM(f.jun) AS tot_jun
            , SUM(f.jul) AS tot_jul
            , SUM(f.aug) AS tot_aug
            , SUM(f.sep) AS tot_sep
            , SUM(f.oct) AS tot_oct
            , SUM(f.nov) AS tot_nov
            , SUM(f.dec) AS tot_dec
            , SUM(f.year_total) AS tot_fy
            FROM proj_forecast AS f
            WHERE f.account IN (845009)
            GROUP BY f.bcode, f.prog_name, f.division, f.rgt
            ) As tot_fcst
        LEFT OUTER JOIN
            (SELECT f.bcode
            , f.prog_name
            , f.division
            , f.rgt
            , SUM(f.jan) AS it_jan
            , SUM(f.feb) AS it_feb
            , SUM(f.mar) AS it_mar
            , SUM(f.apr) AS it_apr
            , SUM(f.may) AS it_may
            , SUM(f.jun) AS it_jun
            , SUM(f.jul) AS it_jul
            , SUM(f.aug) AS it_aug
            , SUM(f.sep) AS it_sep
            , SUM(f.oct) AS it_oct
            , SUM(f.nov) AS it_nov
            , SUM(f.dec) AS it_dec
            , SUM(f.year_total) AS it_fy
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
            , SUM(f.jan) AS bus_jan
            , SUM(f.feb) AS bus_feb
            , SUM(f.mar) AS bus_mar
            , SUM(f.apr) AS bus_apr
            , SUM(f.may) AS bus_may
            , SUM(f.jun) AS bus_jun
            , SUM(f.jul) AS bus_jul
            , SUM(f.aug) AS bus_aug
            , SUM(f.sep) AS bus_sep
            , SUM(f.oct) AS bus_oct
            , SUM(f.nov) AS bus_nov
            , SUM(f.dec) AS bus_dec
            , SUM(f.year_total) AS bus_fy
            FROM proj_forecast AS f
            WHERE f.account IN (815003)
            GROUP BY f.bcode, f.prog_name, f.division, f.rgt
            ) AS bus_fcst
        ON tot_fcst.bcode = bus_fcst.bcode
        ORDER by tot_fcst.division, tot_fcst.bcode
        '''

# create excel file and format it for writing the report
xl = win32.gencache.EnsureDispatch('Excel.Application')
xl.DisplayAlerts = False
wb = xl.Workbooks.Add()
ws2 = xl.Worksheets('Sheet3')
ws1 = xl.Worksheets('Sheet2')
ws3 = xl.Worksheets('Sheet1')

# this is to create the worksheet for the forecast report by program
ws2.Name = "Forecast by Program"
headers1 = ['','','','','Jan','','','Feb','','','Mar','','','Apr','','','May','','','Jun','','','Jul','','','Aug','','','Sep','','','Oct','','','Nov','','','Dec','','','FY 17','','']
counter = 0
for val in headers1:
    counter += 1
    ws2.Cells(1,1).Offset(1,counter).Value = val
    ws2.Cells(1,1).Offset(1,counter).Font.Bold = True
headers2 = ['Division','Bcode','Program Name','RGT','Total','IT','Bus','Total','IT','Bus','Total','IT','Bus','Total','IT','Bus','Total','IT','Bus','Total','IT','Bus',
            'Total','IT','Bus','Total','IT','Bus','Total','IT','Bus','Total','IT','Bus','Total','IT','Bus','Total','IT','Bus','Total','IT','Bus']
counter = 0
for val in headers2:
    counter += 1
    ws2.Cells(2,1).Offset(1,counter).Value = val
    if counter < 5:
        ws2.Cells(2, 1).Offset(1, counter).Font.Underline = win32.constants.xlUnderlineStyleSingle
        ws2.Cells(2, 1).Offset(1, counter).Font.Bold = True
    elif counter >= 5:
        ws2.Cells(2, 1).Offset(1, counter).HorizontalAlignment = win32.constants.xlCenter
        ws2.Cells(2, 1).Offset(1, counter).Font.Underline = win32.constants.xlUnderlineStyleSingle
        ws2.Cells(2, 1).Offset(1, counter).Font.Bold = True
    if val == 'Total':
        ws2.Cells(2, 1).Offset(1, counter).Interior.Color = 10092441
    elif val == 'IT':
        ws2.Cells(2, 1).Offset(1, counter).Interior.Color = 16772300
    elif val == 'Bus':
        ws2.Cells(2, 1).Offset(1, counter).Interior.Color = 16764159
    else:
        ws2.Cells(2, 1).Offset(1, counter).Interior.ThemeColor = win32.constants.xlThemeColorDark1
        ws2.Cells(2, 1).Offset(1, counter).Interior.TintAndShade = -0.249977111117893
ws2.Range("E1:G1").MergeCells = True
ws2.Range("E1:G1").HorizontalAlignment = win32.constants.xlCenter
ws2.Range("H1:J1").MergeCells = True
ws2.Range("H1:J1").HorizontalAlignment = win32.constants.xlCenter
ws2.Range("K1:M1").MergeCells = True
ws2.Range("K1:M1").HorizontalAlignment = win32.constants.xlCenter
ws2.Range("N1:P1").MergeCells = True
ws2.Range("N1:P1").HorizontalAlignment = win32.constants.xlCenter
ws2.Range("Q1:S1").MergeCells = True
ws2.Range("Q1:S1").HorizontalAlignment = win32.constants.xlCenter
ws2.Range("T1:V1").MergeCells = True
ws2.Range("T1:V1").HorizontalAlignment = win32.constants.xlCenter
ws2.Range("W1:Y1").MergeCells = True
ws2.Range("W1:Y1").HorizontalAlignment = win32.constants.xlCenter
ws2.Range("Z1:AB1").MergeCells = True
ws2.Range("Z1:AB1").HorizontalAlignment = win32.constants.xlCenter
ws2.Range("AC1:AE1").MergeCells = True
ws2.Range("AC1:AE1").HorizontalAlignment = win32.constants.xlCenter
ws2.Range("AF1:AH1").MergeCells = True
ws2.Range("AF1:AH1").HorizontalAlignment = win32.constants.xlCenter
ws2.Range("AI1:AK1").MergeCells = True
ws2.Range("AI1:AK1").HorizontalAlignment = win32.constants.xlCenter
ws2.Range("AL1:AN1").MergeCells = True
ws2.Range("AL1:AN1").HorizontalAlignment = win32.constants.xlCenter
ws2.Range("AO1:AQ1").MergeCells = True
ws2.Range("AO1:AQ1").HorizontalAlignment = win32.constants.xlCenter

rowCounter = 0
# write the output from sql to the Excel file
for row in cursor.execute(sql):
    rowCounter += 1
    valCounter = 0
    for val in row:
        valCounter += 1
        ws2.Cells(2 + rowCounter, valCounter).Value = val
        ws2.Cells(2 + rowCounter, valCounter).NumberFormat = "$#,###,;($#,###,)"
ws2.Columns.AutoFit()

ws3.Name = 'Missing Data'

counter = 0
# this is the check whether there are any missing values in the report and output to the missing data worksheet
missing_vals = check_missing_vals(ws2)
for key in missing_vals:
    for val in missing_vals[key]:
        counter += 1
        ws3.Cells(1, 1).Offset(counter, 1).Value = val
if counter == 0:
    ws3.Cells(1, 1).Offset(1, 1).Value = 'No Missing Data'

# this is to create the check worksheet for automatic reconciliation
ws1.Name = "Checks"
headers1 = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec','FY 17']
counter = 0
for val in headers1:
    counter += 1
    ws1.Cells(1,1).Offset(1,counter + 1).Value = val
    ws1.Cells(1,1).Offset(4,counter + 1).FormulaR1C1 = "=R[-2]C-R[-1]C"
    ws1.Cells(1, 1).Offset(8, counter + 1).FormulaR1C1 = "=R[-2]C-R[-1]C"
    ws1.Cells(1, 1).Offset(12, counter + 1).FormulaR1C1 = "=R[-2]C-R[-1]C"
    ws1.Cells(1, 1).Offset(16, counter + 1).FormulaR1C1 = "=R[-2]C-R[-1]C"
    ws1.Cells(1, 1).Offset(20, counter + 1).FormulaR1C1 = "=R[-2]C-R[-1]C"
    ws1.Cells(1, 1).Offset(24, counter + 1).FormulaR1C1 = "=R[-2]C-R[-1]C"
# checks for account 845009
ws1.Cells(2,1).Value = '845009-Db Staging Tbl'
ws1.Cells(3,1).Value = '845009-Check'
ws1.Cells(4,1).Value = '845009-Var'

counter = 0
for v in checkVals['845009']:
    counter += 1
    ws1.Cells(1,1).Offset(3,1 + counter).Value = v

# checks for account 845033
ws1.Cells(6,1).Value = '845033-Db Staging Tbl'
ws1.Cells(7,1).Value = '845033-Check'
ws1.Cells(8,1).Value = '845033-Var'

counter = 0
for v in checkVals['845033']:
    counter += 1
    ws1.Cells(1,1).Offset(7,1 + counter).Value = v

# checks for account 845034
ws1.Cells(10,1).Value = '845034-Db Staging Tbl'
ws1.Cells(11,1).Value = '845034-Check'
ws1.Cells(12,1).Value = '845034-Var'

counter = 0
for v in checkVals['845034']:
    counter += 1
    ws1.Cells(1,1).Offset(11,1 + counter).Value = v

# checks for account 845036
ws1.Cells(14,1).Value = '845036-Db Staging Tbl'
ws1.Cells(15,1).Value = '845036-Check'
ws1.Cells(16,1).Value = '845036-Var'

counter = 0
for v in checkVals['845036']:
    counter += 1
    ws1.Cells(1,1).Offset(15,1 + counter).Value = v

# checks for account 845039
ws1.Cells(18,1).Value = '845039-Db Staging Tbl'
ws1.Cells(19,1).Value = '845039-Check'
ws1.Cells(20,1).Value = '845039-Var'

counter = 0
for v in checkVals['845039']:
    counter += 1
    ws1.Cells(1,1).Offset(19,1 + counter).Value = v

# checks for account 815003
ws1.Cells(22,1).Value = '815003-Db Staging Tbl'
ws1.Cells(23,1).Value = '815003-Check'
ws1.Cells(24,1).Value = '815003-Var'

counter = 0
for v in checkVals['815003']:
    counter += 1
    ws1.Cells(1,1).Offset(23,1 + counter).Value = v

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
                ws1.Cells(1,1).Offset(2, valCounter).Value = val
            elif acct == 845033:
                ws1.Cells(1,1).Offset(6, valCounter).Value = val
            elif acct == 845034:
                ws1.Cells(1,1).Offset(10, valCounter).Value = val
            elif acct == 845036:
                ws1.Cells(1,1).Offset(14,valCounter).Value = val
            elif acct == 845039:
                ws1.Cells(1,1).Offset(18, valCounter).Value = val
            elif acct == 815003:
                ws1.Cells(1,1).Offset(22,valCounter).Value = val

ws1.Columns.AutoFit()
ws3.Columns.AutoFit()
wb.SaveAs(userRptNm)
xl.Application.Quit()

# cleanup
cursor.close()
cnn.commit()
cnn.close()
os.remove(strDbName)

print "------------------------------------Script Successfully Run-----------------------------------------"
print "\n"