import constants as keys
from telegram.ext import *
import login_detail
# import pandas as pd
import openpyxl
# import responses as R

import send_mail as sm



print("JAI BAJARANG BALI")



print("Bot Started....")

s="191114083ajaybhak"
print(login_detail.scholar_pass[s])




#


def attend(scholar_no,subject_code):
    file=subject_code+'.xlsx'
    df=openpyxl.load_workbook(file)
    lf=df.active
    ss='B'+str(scholar_no+1)
    cl=lf[ss]
    ss='B'+str(410)
    cl2=lf[ss]
    # print(cl.value)
    if(cl2.value==0):
        return f"{subject_code} : No class Held Yet"


    per= (cl.value/cl2.value)*100.0
    per=per*100
    per=int(per)
    per=per/100.0
        
    return f"{subject_code} : {per}%"




# attend(83,'EC443')


def marks(scholar_no,subject_code):

    file=subject_code+'.xlsx'
    df=openpyxl.load_workbook(file)
    lf=df.active
    ss='C'+str(scholar_no+1)
    cl=lf[ss]
    ss='C'+str(410)
    cl2=lf[ss]

    if(cl2.value==0):
        return f"No exam held yet : {subject_code}"
    


    return f"{subject_code} : marks scored {cl.value} out of {cl2.value} \n"












# def 



def mail_sent(subject_code):

    file=subject_code+'.xlsx'

    subject= f"Warning mail for attendance of subject {subject_code} "
    print(subject)

    df=openpyxl.load_workbook(file)
    ls=[]
    lf=df.active
    for i in range(1,400):
        
        ss='B'+str(i+1)
        cl=lf[ss]
        ss='B'+str(410)
        cl2=lf[ss]
    # print(cl.value)
        if(cl2.value==0):
            continue
        per= (cl.value/cl2.value)*100.0

        per= (cl.value/cl2.value)*100.0
        per=per*100
        per=int(per)
        per=per/100.0
        print(i)
        if(per<75.0 and per>0.0):
            ss='D'+str(i+1)
            cl=lf[ss]
            ml=cl.value
            ml=str(ml)
            body=f"Hey, scholar no {i}, your attendance is {per}% please try to maintain it above 75% otherwise you will be detained from upcoming semester exam! meet me in contact hour."
            sm.send_mail(subject,body,ml)
            print(f"mail sent successfully to scholar no {i}")

    

        



    








def response(input_text,ask_login,login,user_name,subject_code,scholar_no,login_teacher):
    user_msg=str(input_text)
    user_message=str(input_text).lower()




    # print(ask_login)
    if(ask_login):
        print("YES")
        print(user_message)
        print(len(user_msg))
        l=len(user_msg)
        if(l==13):
            print(user_msg)
            subject_code=login_detail.subject_pass[user_msg]

            if(subject_code is None):
                return f"wrong credidantls",ask_login,login,user_name,subject_code,scholar_no,login_teacher
            login=True
            login_teacher =True
            user_name=user_msg
            return f"login successfully your subject code is {subject_code}",ask_login,login,user_name,subject_code,scholar_no,login_teacher
        elif(l==17):
            print(scholar_no)
            user_msg=str(user_msg)
            scholar_no= login_detail.scholar_pass[user_message]
            if(scholar_no is None):
                return f"wrong credidantls",ask_login,login,user_name,subject_code,scholar_no,login_teacher
            
            user_name=user_msg
            login=True
            return f"login successfully your scholar no is {scholar_no} ",ask_login,login,user_name,subject_code,scholar_no,login_teacher



    if(user_message in ("send warning mails")):
        if(login_teacher==False):
            return f"please login as a teacher",ask_login,login,user_name,subject_code,scholar_no,login_teacher
        else:
            print(subject_code)
            mail_sent(subject_code)
            return f"All mail sent",ask_login,login,user_name,subject_code,scholar_no,login_teacher
        
    if(user_message in ("marks")):
        if(login==True):
            if(login_teacher==False):
                return f"{marks(scholar_no,'EC431')}\n{marks(scholar_no,'EC437')}\n{marks(scholar_no,'EC443')}\n",ask_login,login,user_name,subject_code,scholar_no,login_teacher
            else:
                return f"please login as student",ask_login,login,user_name,subject_code,scholar_no,login_teacher
        else:
            return f"please login as student",ask_login,login,user_name,subject_code,scholar_no,login_teacher

    

    




    if(user_message in ("exit","quit","logout")):
        login=False
        login_teacher=False
        subject_code="EC"
        scholar_no=191114083
        return f"logout successfully",ask_login,login,user_name,subject_code,scholar_no,login_teacher
    

    if(user_message in ("attendance")):
        if(login == False):
            return f"please login first",ask_login,login,user_name,subject_code,scholar_no,login_teacher
        else:
            if(login_teacher==False):
                return f"your attendance is as follows:\n {attend(scholar_no,'EC431')}\n {attend(scholar_no,'EC437')}\n {attend(scholar_no,'EC443')}",ask_login,login,user_name,subject_code,scholar_no,login_teacher
            else:
                return f"you logged in as a teacher your subject code is : {subject_code}",ask_login,login,user_name,subject_code,scholar_no,login_teacher


    ask_login=False

    if user_message in ("hello","hi","hii"):
        print("YES")
        return "Hii manitian, type login to login",ask_login,login,user_name,subject_code,scholar_no,login_teacher
    
    if(user_message in ("login")):
        if(login):
            return f"You have already logged in with user name {user_name}",ask_login,login,user_name,subject_code,scholar_no,login_teacher 
        else:
            ask_login=True
            print(ask_login)
            return f"To login <USER_NAME><PASSWORD>, here for students <USER_NAME> is scholar no and password must be of 8 letters \n for teachers <USER_NAME> is subject code ans password must be of 8 letters (there must not be any space between scholar no or subject code and password)",ask_login,login,user_name,subject_code,scholar_no,login_teacher



#


def start_command(update, context):
    update.message.reply_text('Type anything to get started_')

def help_command(update,context):
    update.message.reply_text('help!')




def handle_message(update,context):
    # print(login)

    
    global login
    global user_name
    global ask_login

    global subject_code
    global scholar_no
    global login_teacher

    
    # print("Y")

    

    text=str(update.message.text).lower()
    print(ask_login)
    respons,ask_login,login,user_name,subject_code,scholar_no,login_teacher = response(text,ask_login,login,user_name,subject_code,scholar_no,login_teacher)
    print(ask_login)
    update.message.reply_text(respons)




login=False
user_name="EC"
ask_login=False
login_teacher=False

subject_code="EC"
scholar_no=191114083

def error(update,context):
    print(f"Update {update} caused error {context.error}")



def main():

    
    

    updater =Updater(keys.API_KEY,use_context=True)
    dp=updater.dispatcher
    # dp.add_handler(CommandHandler("start",start_command))
    # dp.add_handler(CommandHandler("start",help_command))
    # print("before")

    # print(ask_login)
    dp.add_handler(MessageHandler(Filters.text,handle_message))
    # print("after")
    # print(ask_login)

    dp.add_error_handler(error)
    updater.start_polling()
    updater.idle()



main()
    


