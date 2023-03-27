import login_detail


login=False
user_name="manit"
ask_login=False

subject_code="EC"
scholar_no=191114083


def response(input_text):
    user_message=str(input_text).lower()
    user_msg=str(input_text)
    # if(ask_login):
    #     if(len(user_msg)==13):
    #         subject_code=login_detail.subject_pass[user_msg]
    #         if(subject_code is None):
    #             return "wrong credidantls"
    #         login=True
    #         return "login successfully your subject code is ", subject_code
    #     elif(len(user_msg)==17):
    #         scholar_no=login_detail.subject_pass[user_msg]
    #         if(scholar_no is None):
    #             return "wrong credidantls"
    #         login=True
    #         return "login successfully your scholar no is ", scholar_no

    if(user_message in ("exit","quit","logout")):
        login=False
        subject_code="EC"
        scholar_no=191114083
        return "logout successfully"
    
    # ask_login=False

    if user_message in ("hello","hi","hii"):
        return "Hii manitian, type login to login"
    
    if(user_message in ("login")):
        if(login):
            return "You have already logined with user name ", user_name
        else:
            ask_login=True
            return "To login <USER_NAME><PASSWORD>, here for students <USER_NAME> is scholar no and password must be of 8 letters"



            
    
