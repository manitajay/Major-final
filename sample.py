import gspread

sa=gspread.service_account(filename="service_account.json")
sh=sa.open("EC431")

wks=sh.worksheet("Sheet1")

print('Rows: ',wks.row_count)
print('Cols: ',wks.col_count)

print(wks.acell('D84').value)

# wks.update('A3','ant')




from datetime import datetime

import time_table

import openpyxl



today = datetime.today()

# dd/mm/YY
d1 = today.strftime("%d/%m/%Y")
d1=str(d1)


# print(d1)
alfa=['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']


# for 

def printString(n):
 
    # To store result (Excel column name)
    string = ""
 
    # To store current index in str which is result
    i = 0
 
    while n > 0:
        # Find remainder
        rem = n % 26
        rem=int(rem)
        # if remainder is 0, then a
        # 'Z' must be there in output
        if rem == 0:
            string += 'Z'
            i += 1
            n = (n / 26) - 1
            n=int(n)
        else:

            string+=alfa[rem-1]
            i += 1
            n = n / 26
            n=int(n)
    # string[i] = '\0'
 
    # Reverse the string and print result
    # string = string[::-1]
    reversed(string)
    # print(string)
    return string



now=datetime.now()

day=now.strftime('%A')
day=str(day)






def attend_update(scholar_no,subject_code):
    # D prev_date;
    # E last coloum

    file=subject_code
    sa=gspread.service_account(filename="service_account.json")
    sh=sa.open("EC431")
    wks=sh.worksheet("Sheet1")
    
    ss1='D'+str(410)
    prev_date=wks.acell(ss1).value


    if(prev_date==d1):
        ss1='E'+str(410)
        col=wks.acell(ss1).value
        col=int(col)
        cola=printString(col)
        for i in range(2,400):
            ss1=cola+str(i)
            if(int(wks.acell(ss1).value)==0 and scholar_no==(i-1)):
                # lf[ss1]=1
                wks.update(ss1,1)
                ss1='B'+str(i)
                # lf[ss1]=lf[ss1].value+1
                val=(wks.acell(ss1)).value
                val=int(val)
                wks.update(ss1,val+1)
        
        # df.save(filename=file)


        return 
    ss1='E'+str(410)
    col=(wks.acell(ss1)).value
    col=int(col)
    col=col+1
    cola=printString(col)
    ss1=cola+str(1)

    # lf[ss1]=d1
    wks.update(ss1,d1)
    for i in range(2,400):
        
        ss1=cola+str(i)
        wks.update(ss1,0)
        # lf[ss1]=0

        if(i-1==scholar_no):
            # lf[ss1]=1
            wks.update(ss1,1)

            ss1='B'+str(i)
            # lf[ss1]=lf[ss1].value+1
            val=wks.acell(ss1).value
            val=int(val)
            val+=1
            wks.update(ss1,val)

    

    ss1='B'+str(410)
    val=wks.acell(ss1).value
    val=int(val)
    # lf[ss1]=lf[ss1].value+1
    wks.update(ss1,val+1)

    ss1='E'+str(410)
    # lf[ss1]=col
    wks.update(ss1,col)
    # df.save(filename=file)


    return


attend_update(96,'EC431')

