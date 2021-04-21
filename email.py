#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Nov  5 09:10:03 2020

@author: mahara
"""
import smtplib
import pandas as pd
import numpy as np
# from fpdf import FPDF 
# from time import sleep
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


def mysendmail(toaddrs,final_msg,**kwargs):
    roll=kwargs.get('roll',-1)
    name=kwargs.get("name","")
    fromaddr = kwargs.get('fromaddr','@iitdh.ac.in')
    password=kwargs.get('password', 1234)
    toaddr = toaddrs
    msg = MIMEMultipart()
    msg['From'] = fromaddr
    msg['To'] = ",".join(toaddr)
    msg['Cc']=",".join(cc_addresses)
    msg['Subject'] = subject
    msg.attach(MIMEText(final_msg, 'plain'))
    filename = kwargs.get('attachment',None)
    if type(filename)!=type(None):
        attachment = open(filename, "rb")
        part = MIMEBase('application', 'octet-stream')
        part.set_payload((attachment).read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
        msg.attach(part)
    
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.ehlo()
    try:
        if not test_mode:
            server.login(fromaddr, password)
        text = msg.as_string()
        toadd=[toaddr]+cc_addresses+bcc_addresses
        if not test_mode:
            server.sendmail(fromaddr, toaddr, text)
            server.quit()
        if not test_mode:
            print(" Sent mail to "+name, toaddr)
        else:
            print(f" Test mode enabled. Did not send email intended for {name} {toaddr}")
    except Exception as e:
        if e.smtp_code==535:
            print('\n Couldn\'t send mail, try disabling secure authentication for your account here \n https://myaccount.google.com/lesssecureapps \n Ensure that the correct Gmail account is selected while you turn this off. \n You can turn it back on after sending emails')
        else:
            print("\n Couldn't send mail. Got error Code ", e.smtp_code)
#%% Imports
import getpass
import argparse
#%% Getting user info
parser = argparse.ArgumentParser(description='Emailing in Bulk.')
parser.add_argument("--cc",required=False,type=lambda x: (str(x).lower() == 'true'),default=False, help="If you want to cc individuals")
parser.add_argument("--bcc",required=False,type=lambda x: (str(x).lower() == 'true'),default=False, help="If you want to bcc individuals. Note that this is not tested, so use at your own risk.")
parser.add_argument("--attachment",required=False,type=lambda x: (str(x).lower() == 'true'),default=False, help="If you want to attach files.")
parser.add_argument("--test_mode",required=False,type=lambda x: (str(x).lower() == 'true'),default=False, help="If you want run in test mode but not send a random individual's email to anyone.")
parser.add_argument("--test_mode_and_send",required=False,type=lambda x: (str(x).lower() == 'true'),default=False, help="If you want run in test mode and send a random individual's email to some predesignated test emails.")
parser.add_argument("--use_xlsx",required=False,type=lambda x: (str(x).lower() == 'false'),default=True, help="If you want to convert the program to read the xlsx file and convert it to a .txt attachment. By default it is True.")
parser.add_argument("--attachment_extention",required=False,type=str,default=' ', help="The attachment extention.")


args = parser.parse_args()
cc=args.cc
bcc=args.bcc
use_attachment=args.attachment
use_xlsx=args.use_xlsx
test_mode=args.test_mode
attachment_extention=args.attachment_extention
test_mode_and_send=args.test_mode_and_send

if test_mode_and_send and test_mode:
    import sys
    sys.exit("test_mode_and_send and test_mode arguments are both true. The script will not run.")

subject=input("Enter the email subject. Type -1 to use the default text 'Testing' \n")
if subject=="-1":
    subject='testing'
#CC part
cc_counter=0
cc_addresses=[]
starting_cc=True
while cc:
    if starting_cc:
        address=input(f"Enter the address cc recipient {cc_counter}. When you are done typing, press enter twice. \n")
        starting_cc=False
    else:
        address=input()
    if address:
        cc_addresses.append(address)
    else:
        break
    cc_counter+=1
#BCC part
bcc_counter=0
bcc_addresses=[]
starting_bcc=True
while bcc:
    if starting_bcc:
        address=input(f"Enter the address bcc recipient {bcc_counter}. When you are done typing, press enter twice. \n")
        starting_bcc=False
    else:
        address=input()
    if address:
        bcc_addresses.append(address)
    else:
        break
    bcc_counter+=1

# TEST EMAIL
test_counter=0
test_addresses=[]
starting_test=True
while test_mode_and_send:
    if starting_test:
        address=input(f"Enter test email address of recipient {test_counter}. When you are done typing, press enter twice. \n")
        starting_test=False
    else:
        address=input()
    if address:
        test_addresses.append(address)
    else:
        break
    test_counter+=1


if not test_mode:
    your_email_id=input(" Enter Your Email Address\n")
    your_password=getpass.getpass(" Enter your password. Whatever you enter will not be visible.\n")
else:
    your_email_id="xyz@email.com"
    your_password=1234 #If running in test mode, the dummy password is 1234
lines = []

#%%Greeting
greeting=input("Type in the greeting to use for everyone. Their name will be added to this, followed by a comma.\n")
#%% MAIN BODY
starting_main_body=True
while True:
    if starting_main_body:
        line=input(" Type in the main body of text that you would like to give everyone, excluding the greetings and salutations.\n When you are done typing, press enter twice. \n")
        starting_main_body=False
    else:
        line=input()
    if line:
        lines.append(line)
    else:
        break
common_body = '\n'.join(lines)
#%% ENDING SALUTATIONS
starting_salutations=True
while True:
    if starting_salutations:
        line=input(" Type in ending salutations you would like to give everyone. \n When you are done typing, press enter twice. \n")
        starting_salutations=False
    else:
        line=input()
    if line:
        lines.append(line)
    else:
        break
common_body_with_salutation= '\n'.join(lines)


#%% XLSX Sheet Code
Email_names=np.loadtxt('email_names.csv',dtype=str,delimiter=',')
if use_xlsx:
    sheet1=pd.read_excel("Summary.xlsx",sheet_name=0)
    sheet2=pd.read_excel("Summary.xlsx",sheet_name=1)
    sheets=[sheet1,sheet2]
    # na_filler=-12345678
    for sheet in sheets:
        sheet.dropna()
        # sheet.fillna(na_filler,inplace=True)
    #     sheet.replace("Not Applicable",na_filler,inplace=True)
#%% MAIN
pd.set_option('display.max_colwidth', None)
# test_mail=['191081003@iitdh.ac.in','191081002@iitdh.ac.in']

for individual in [np.random.randint(len(Email_names))] if test_mode_and_send else range(len(Email_names)):
# for individual in [4]:
    if use_xlsx:
        name=Email_names[individual,1]
        roll=str((Email_names[individual,0]))
        email=test_addresses if test_mode_and_send else roll+"@iitdh.ac.in"
        """xlsx to Text File generation """
        #From the first sheet, extract cells A1 and B1
        mssg1=sheets[0].iloc[individual,0:2].to_string()+"\n \n \n \n" 
        itr=0
        for work in ["Quiz 1", "Quiz 2"]:
            a=sheets[itr].iloc[individual,2:] #Consider the cells from column C onward, row 'itr'
            a=a.to_string()+'\n \n \n'
            mssg1=mssg1+work.center(60,'-')+'\n'+a
            itr=itr+1
        textfile = open("Text Files/"+roll+".txt", 'w')
        textfile.write(mssg1)
        textfile.close()
    f=1
    body=f"{greeting} {name},\n {common_body_with_salutation}"
    mysendmail(email,
                    body,
                    name=name,
                    fromaddr=your_email_id,
                    password=your_password,
                    attachment=None if attachment_extention==' ' else f"Attachments/{roll}{attachment_extention}"
            )