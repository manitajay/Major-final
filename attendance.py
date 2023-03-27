import face_recognition
import cv2
import numpy as np
import csv 
import os
import glob
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

    file=subject_code+'.xlsx'
    df=openpyxl.load_workbook(file)
    lf=df.active

    ss1='D'+str(410)
    prev_date=lf[ss1].value

    if(prev_date==d1):
        ss1='E'+str(410)
        col=lf[ss1].value
        cola=printString(col)
        for i in range(2,400):
            ss1=cola+str(i)
            if(lf[ss1]==0 and scholar_no==(i-1)):
                lf[ss1]=1
                ss1='B'+str(i)
                lf[ss1]=lf[ss1].value+1
        
        df.save(filename=file)


        print("completed")

        return 
    ss1='E'+str(410)
    col=lf[ss1].value
    col=col+1
    cola=printString(col)
    ss1=cola+str(1)

    lf[ss1]=d1
    for i in range(2,400):
        
        ss1=cola+str(i)
        lf[ss1]=0
        if(i-1==scholar_no):
            lf[ss1]=1
            ss1='B'+str(i)
            lf[ss1]=lf[ss1].value+1
    

    ss1='B'+str(410)
    lf[ss1]=lf[ss1].value+1

    ss1='E'+str(410)
    lf[ss1]=col
    ss1='D'+str(410)
    lf[ss1]=d1

    df.save(filename=file)


    print("completed")


    return



        

    



# def attend_update(scholar_no,subject_code):
#     file=subject_code+'.xlsx'
#     df=openpyxl.load_workbook(file)
#     lf=df.active
#     ss1='B'+str(scholar_no+1)
#     cl=lf[ss1]
#     ss2='B'+str(410)
#     cl2=lf[ss2]
#     print(cl.value)
#     val1=cl.value+1
#     print(val1)
#     lf[ss1]=val1
#     val2=cl2.value+1
#     lf[ss2]=val2
#     df.save(filename=file)

# EC437 SR      satellite communication 0 
# EC431 DKR     mobile communication 1
# EC443 LXM     vlsi digital design 2
# 0 attendance 
# 1 marks


# arr=arr.reshape(2,200)
# arr[0][0]--> total held classes;
# arr[1][0]-->total marks



video_capture=cv2.VideoCapture(0,cv2.CAP_DSHOW)

ajay_image=face_recognition.load_image_file("photos/83.jpeg")
ajay_encoding=face_recognition.face_encodings(ajay_image)[0]



arun_image=face_recognition.load_image_file("photos/49.jpg")
arun_encoding=face_recognition.face_encodings(arun_image)[0]

known_face_encoding=[ajay_encoding,arun_encoding]

known_face_scholar=["83","49"]

students=known_face_scholar.copy()

face_locations=[]
face_encodings=[]
face_scholar=[]
s=True


now=datetime.now()

current_date=now.strftime("%Y-%m-%d")


f=open(current_date+'.csv','w+',newline='')

lnwriter=csv.writer(f)

while True:
    _,frame= video_capture.read()
    small_frame=cv2.resize(frame,(0,0),fx=0.25,fy=0.25)
    rgb_small_frame=small_frame[:,:,::-1]
    if s:
        face_locations=face_recognition.face_locations(rgb_small_frame)
        face_encodings=face_recognition.face_encodings(rgb_small_frame,face_locations)
        face_scholar=[]

        for face_encoding in face_encodings:
            matches=face_recognition.compare_faces(known_face_encoding,face_encoding)
            scholar=""
            face_distance=face_recognition.face_distance(known_face_encoding,face_encoding)
            best_match_index=np.argmin(face_distance)
            if(matches[best_match_index]):
                scholar=known_face_scholar[best_match_index]
            
            face_scholar.append(scholar)
            # print(scholar)
            if(scholar in known_face_scholar):
                if(scholar in students):
                    students.remove(scholar)
                    print(students)
                    print(f"scholar {scholar}")
                    current_time=now.strftime("%H-%M-%S")
                    day=now.strftime('%A')
                    print(day)
                    day=str(day)
                    ls=time_table.schedule[day]
                    sc=int(scholar)

                    if(ls is None):
                        print("holiday!")
                    else:
                        for subject_code in ls:
                            attend_update(sc,subject_code)

                    
                    # lnwriter.writerow[scholar,current_time]
            
    cv2.imshow("attendance system",frame)
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

video_capture.release()
cv2.destroyAllWindows()
f.close()


