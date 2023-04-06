#!/usr/bin/python3

import csv

import datetime, sys

today=datetime.datetime.now()
#print (today.strftime('%Y-%m-%d'))
ptpdate=[]
for d in range(-2,10):
    day=today+datetime.timedelta(days=d)
    ptpdate.append(day.strftime('%Y-%m-%d'))

#print(ptpdate)

def ptpcheck(string):
    check=1
    for day in ptpdate:
        if (string==day):
#            print ('!--(',day,')--!')
            check=0
            continue
    return(check)

stats=open('statuses.cfg')
statuses=stats.read().split('\n')
#print (statuses)

def statuscheck(row):
    check=1
    for status in statuses:
        if (status==row):
            check=0
            continue
    return(check)

#import openpyxl
#report='../Загрузки/ru_skip_tracing'+today.strftime('%Y-%m-%d')+'.xlsx'
#wb = openpyxl.load_workbook(filename = report)
#sheet = wb['data']
#val=sheet['A1'].value
#print (val)
#
print('<table border=1>')
print('<tr><td><b>loan id</td><td><b>product</td><td><b>client name</td><td><b>curent delay</td><td><b>delayed ammount</b></td></tr>')

tab_file=""
for i in range(1,len(sys.argv)):
    tab_file=tab_file+' '+sys.argv[i]
tab_file = tab_file.lstrip()

with open(tab_file) as f:
    reader = csv.DictReader(f)
    for row in reader:
        if (row['LAST TOUCH DATE']=='---'):
            continue
        if (statuscheck(row['LAST TOUCH STATUS'])==0):
            continue
        if (ptpcheck(row['LAST PTP DATE'].split()[0])==0):
            continue
        print('<tr><td><a href="https://ecocrm.ya.ecofin.io/ru/collection/show/'+row['LOAN ID']+'/installment_pdl">'+row['LOAN ID']+'</a></td>'\
                '<td>',row['MINI PRODUCT NAME'],'</td>'\
                '<td>',row['CLIENT NAME'],'</td>'\
                '<td>',row['CURRENT DELAY'],'</td>'\
                '<td>',row['DELAYED AMOUNT'].split('.')[0],'</td></tr>')
print('</table>')
