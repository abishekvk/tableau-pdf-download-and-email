# -*- coding: utf-8 -*-
"""
Created on Thu Apr  4 12:17:03 2019

@author: AVariyan
"""

import os
import smtplib 
from email.mime.multipart import MIMEMultipart 
from email.mime.text import MIMEText 
from email.mime.base import MIMEBase 
from email import encoders 
import pandas as pd
import subprocess
from PyPDF2 import PdfFileMerger
import datetime
class email_pdf:       
    def email_pdf(self,filepath,fromaddr,user):
        list=pd.read_excel('\\sample.xlsx',index=False)
        #a=list['Country'].unique
        list_user=list[list['EMAIL']==user]
        #filelist=[]
        subprocess.call(r'\\login.bat')
        for index,row in  list_user.iterrows():
            #print(row['URL'],row['FILE'])
            str='tabcmd export "'+row['URL']+'" --fullpdf --pagelayout landscape --pagesize tabloid --width 1200 -f "'+row['FILE']+'"'
            file = open("\\download_pdf.bat","w")
            #with open('helloworld.txt', 'w') as filehandle:
            file.write(str) 
            file.close()
            
            try:
                #print(row['URL'],row['FILE'])
                
                subprocess.call(r'download_pdf.bat')
                #filelist.append(row['FILE'])
            except Exception as e:
                print(str(e))
            
        fromaddr = fromaddr
        toaddr = user
        print(toaddr)
        #file=row['FILE']
        #filen=file+str(datetime.datetime.now)
        #print(file)
        # instance of MIMEMultipart 
        msg = MIMEMultipart() 
          
        # storing the senders email address   
        msg['From'] = fromaddr 
          
        # storing the receivers email address  
        msg['To'] = toaddr 
          
        # storing the subject  
        msg['Subject'] = "WORKSHEET"
          
        # string to store the body of the mail 
        body = "Hi, Please find the attached"
          
        # attach the body with the msg instance 
        msg.attach(MIMEText(body, 'plain')) 
          
        # open the file to be sent  
        #filename = "abi.pdf"
        #attachment = open(file, "rb") 
        
        for a_file in list_user['FILE']:
            print(a_file)
            attachment = open(a_file, 'rb')
            file_name = os.path.basename(a_file)
            part = MIMEBase('application','octet-stream')
            part.set_payload(attachment.read())
            part.add_header('Content-Disposition',
                            'attachment',
                            filename=file_name)
            encoders.encode_base64(part)
            msg.attach(part)
        
          
        # instance of MIMEBase and named as p 
        #p = MIMEBase('application', 'octet-stream') 
          
        # To change the payload into encoded form 
        #p.set_payload((attachment).read()) 
          
        # encode into base64 
        #encoders.encode_base64(p) 
           
        #p.add_header('Content-Disposition', "attachment; filename= %s" % file) 
          
        # attach the instance 'p' to instance 'msg' 
        #msg.attach(p) 
          
        # creates SMTP session 
        s = smtplib.SMTP('smtp.office365.com',587) 
          
        # start TLS for security 
        s.starttls() 
          
        # Authentication 
        try:
            s.login(fromaddr, "password") 
        except Exception as e:
            print(str(e))
          
        # Converts the Multipart msg into a string 
        text = msg.as_string() 
          
        # sending the mail 
        s.sendmail(fromaddr, toaddr, text) 
          
        # terminating the session 
        s.quit() 

if __name__ == '__main__':
    filepath='\\list.xlsx'
    fromaddr='abc@company.com'
    user='xyz@company.com'
    email_pdf=email_pdf()
    email_pdf.email_pdf(filepath,fromaddr,user)