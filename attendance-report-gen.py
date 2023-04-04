# Imported libraries required for programme (math, pandas, numpy, os, csv, datetime) and cleared screen using cls command
from logging.config import dictConfig
import math
import os
from time import strftime
import pandas as pd
import numpy as np
import csv
import shutil
from datetime import datetime
import imghdr
from email.message import EmailMessage
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import pandas.io.formats.excel
from datetime import timedelta, date

start_time = datetime.now()

os.system('cls')

# os.chdir() for changing directory for evaluation and smooth running, avoid confusions while writing code
os.chdir(r'C:\Users\Chandan\OneDrive\Documents\GitHub\2001CB21_2022\tut06')

# Function for date range
def daterange(date1, date2):
    for n in range(int ((date2 - date1).days)+1):
        yield date1 + timedelta(n)

# Function defination
def attendance_report():
    # Try block for file copying
    try:
        # Copied content of input file into output file using shutil library and copyfile function
        # File is copied in target file, it is created if does not exist and overwrites if exists
        original = r'input_registered_students.csv'
        target = r'output\attendance_report_consolidated.csv'
        shutil.copyfile(original, target)
    # Except statement if error occured in try block like input file does not exist or Correct Directory is not open
    except EnvironmentError:
        print("Oops! Error Occured. Please Check Input file present in Directory and Correct name given to it.")
        print("Or Check if any file is open anywhere, close it and retry")
        print("Correct input file name: input_registered_students.csv")
        exit()
    
    # Try block for creating dataframes
    try:
        # Created dataframe of input file using pandas library and read_csv() function 
        df1= pd.read_csv(r'input_attendance.csv')
        df2= pd.read_csv(r'output\attendance_report_consolidated.csv')

    # Except block if Input file is not found
    except:
        print("Input file or Consolidated file not found.")
        exit()
    
    # Try block to avoid any exception
    try:
        size_a = df1['Timestamp'].size          # size of dataframe

        # Converting Timestamp to datetime format for application simplicity
        df1['Timestamp']=pd.to_datetime(df1['Timestamp'], format='%d-%m-%Y %H:%M')
        # Creating DayofWeek column for storing day of week for understanding class days
        df1['Dayofweek']=df1['Timestamp'].dt.day_name()

        # Creating empty list for storing valid dates of lecture
        # Lectures were taken on Monday and Thursdays
        ls=[]

        # Start date and end dates extracting from Attendance
        start_dt = df1.loc[0,'Timestamp'].date()
        end_dt = df1.loc[size_a-1,'Timestamp'].date()

        # Storing dates in list ls that are Monday and Thursday
        for dt in daterange(start_dt, end_dt):
            if dt.strftime('%A') == 'Monday' or dt.strftime('%A')=='Thursday':
                ls.append(dt.strftime('%d-%m-%Y'))

    
        # Creating variables for valid days, holidays, total lectures taken and no. of students
        total_lecture_taken= len(ls)
        
        no_of_stud = df2["Roll No"].size
        
        # Creating empty dictionary for storing information about each roll no.
        dict_batch={}
        # Initializing dictionary with appropriate keys and values
        for roll in df2['Roll No']:
            dict_batch[roll]={
                date:{'Total Attendance Count':0,
                    'Real':0,
                    'Duplicate':0,
                    'Invalid':0,
                    'Absent':1,
                    'taken' : False} for date in ls
            }
        # Iterating over the attendance list and counting different values
        for i in range(size_a):
            # Condition for valid attendances, i.e. for ignoring empty cells
            if not(pd.isnull(df1.loc[i,'Attendance'])) == True:
                # Extracting roll No by slicing in st2
                st1 = df1.loc[i,'Attendance']
                st2= st1[0:8]
                
                # Condition for counting attendance in valid dates
                if df1.loc[i,'Timestamp'].strftime('%d-%m-%Y') in ls:
                    # Updating attendance
                    dict_batch[st2][df1.loc[i,'Timestamp'].strftime('%d-%m-%Y')]['Total Attendance Count'] += 1
                    # Condition for Valid attendance in 14:00:00 to 15:00:00(inclusive)
                    if df1.loc[i,'Timestamp'].time() >= pd.to_datetime('14:00').time() and df1.loc[i,'Timestamp'].time()<= pd.to_datetime('15:00').time():
                        # Condition for valid Roll no in registered students
                        if st2 in list(df2['Roll No']):
                            # Condition for counting only one attendance in one valid day
                            if dict_batch[st2][df1.loc[i,'Timestamp'].strftime('%d-%m-%Y')]['taken']==False:
                                # updating flag for attendance taken to True
                                dict_batch[st2][df1.loc[i,'Timestamp'].strftime('%d-%m-%Y')]['taken']=True
                                # Updating real attendance and Absent attendance
                                dict_batch[st2][df1.loc[i,'Timestamp'].strftime('%d-%m-%Y')]['Real']= 1
                                dict_batch[st2][df1.loc[i,'Timestamp'].strftime('%d-%m-%Y')]['Absent']= 0
                            else: 
                                # Updating duplicate attendance
                                dict_batch[st2][df1.loc[i,'Timestamp'].strftime('%d-%m-%Y')]['Duplicate'] += 1
                    # Else statement for Invalid attendances marked outside 14:00:00 to 15:00:00
                    else:
                        # Updating Invalid attendance of valid rolls
                        if st2 in list(df2['Roll No']):
                            dict_batch[st2][df1.loc[i,'Timestamp'].strftime('%d-%m-%Y')]['Invalid'] += 1
                    
        # Creating consolidated report file by iterating through each roll no
        for i in range(no_of_stud):
            real =0         # Variable for real attendance
            # Iterating through dates in ls
            for date in ls:
                # If dictionary value is 1 then writing 'P' in document
                if dict_batch[df2.loc[i,'Roll No']][date]['Real']==1:
                    df2.loc[i,date]= 'P'
                    real += 1
                # Else writing 'A' in the document
                else:
                    df2.loc[i,date]= 'A'
            # Creating structure 
            df2.loc[i,'Actual Lecture Taken']= total_lecture_taken
            df2.loc[i, 'Total Real']= real
            df2.loc[i,'% Attendance (Real/Actual Lecture Taken)']= round(real*100 / total_lecture_taken, 2)
        # Creating output excel file 
        pandas.io.formats.excel.ExcelFormatter.header_style = None
        df2.to_excel(r'output\attendance_report_consolidated.xlsx', index= False)
        os.remove('output\\attendance_report_consolidated.csv')

        # Creating individual roll no csv file and appending info
        for i in range(no_of_stud):
            df = pd.DataFrame()                     # empty dataframe
            # Creating structure 
            df['Date']=np.nan
            # Writing name and roll no
            df.loc[0,'Roll']= df2.loc[i,'Roll No']
            df.loc[0,'Name']= df2.loc[i, 'Name']
            k=1
            # updating structure of individual student file
            for date in ls:
                df.loc[k,'Date']= date
                df.loc[k,'Total Attendance Count']= dict_batch[df2.loc[i,'Roll No']][date]['Total Attendance Count']
                df.loc[k,'Real']= dict_batch[df2.loc[i,'Roll No']][date]['Real']
                df.loc[k,'Duplicate']= dict_batch[df2.loc[i,'Roll No']][date]['Duplicate']
                df.loc[k,'Invalid']= dict_batch[df2.loc[i,'Roll No']][date]['Invalid']
                df.loc[k,'Absent']= dict_batch[df2.loc[i,'Roll No']][date]['Absent']
                k = k+1
            roll = df2.loc[i,'Roll No']
            pandas.io.formats.excel.ExcelFormatter.header_style = None
            # Finally creating output files
            df.to_excel('output\{n}.xlsx'.format(n=roll), index=False)
    
    # except block will be executed if there is any exception in try block
    except:
        print("Code is not running properly due to some exception occured.")
        exit()

    # Try block for email utility
    try:
        # Add your email address
        FROM_ADDR = "mayank265@iitp.ac.in"
        FROM_PASSWD = "changeme"


        toaddr = "cs3842022@gmail.com"
        
        # instance of MIMEMultipart
        msg = MIMEMultipart()
        
        # storing the senders email address  
        msg['From'] = FROM_ADDR
        
        # storing the receivers email address 
        msg['To'] = toaddr
        
        # storing the subject 
        msg['Subject'] = "Consolidated Attendance Report"
        
        # string to store the body of the mail
        body = '''Dear Sir,
        PFA for Consolidated Attendance Report.

Thanks,
Chandan Gaikwad
Roll No. 2001CB21
        '''
        
        # attach the body with the msg instance
        msg.attach(MIMEText(body, 'plain'))
        
        # open the file to be sent 
        filename = "attendance_report_consolidated.xlsx"
        attachment = open("output\\attendance_report_consolidated.xlsx", "rb")
        
        # instance of MIMEBase and named as p
        p = MIMEBase('application', 'octet-stream')
        
        # To change the payload into encoded form
        p.set_payload((attachment).read())
        
        # encode into base64
        encoders.encode_base64(p)
        
        p.add_header('Content-Disposition', "attachment; filename= %s" % filename)
        
        # attach the instance 'p' to instance 'msg'
        msg.attach(p)
        
        # creates SMTP session
        s = smtplib.SMTP('stud.iitp.ac.in', 587)
        
        # start TLS for security
        s.starttls()
        
        # Authentication
        s.login(FROM_ADDR, FROM_PASSWD)
        
        # Converts the Multipart msg into a string
        text = msg.as_string()
        
        # sending the mail
        s.sendmail(FROM_ADDR, toaddr, text)
        
        # terminating the session
        s.quit()
        print("The mail has been sent to cs3842022@gmail.com with 'attendance_report_consolidated.xlsx' as attachment. Please Check inbox!")
    # Except block if there is exception in email utility
    except:
        print("Email could not be sent due to incorrect password or server information.")
        print("Check mail server and port information. Or password or email is incorrect.")
        print("Please check Line no. 184, 225, 233")
    print("Output files are ready. Please Check")


# Call of Function
attendance_report()

#This shall be the last lines of the code.
end_time = datetime.now()
print('Duration of Program Execution: {}'.format(end_time - start_time))