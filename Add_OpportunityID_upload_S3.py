username="**********"
password="*****************"

'''
aws_access_key_id= '**************'
aws_secret_access_key= '************'
service_name = 's3'
region_name = 'us-west-2'
bucket_outbound = '*****'
bucket_inbound_folder = '**********'

import tkinter as tk
from tkinter import filedialog, messagebox
import openpyxl
import numpy as np
import pandas as pd
import time
import shutil 
from tkinter import *
from openpyxl import load_workbook
import sys
import openpyxl
from tkinter import Frame, Tk, Button
import os
import logging
import boto3
from botocore.exceptions import ClientError
import Reference_Part3
import requests
from requests.auth import HTTPBasicAuth
from bs4 import BeautifulSoup
import pandas as pd
from urllib import parse
import urllib

class Record(tk.Frame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.pack()
        self.createWidgets()
        self.parent= parent
        self.initUI()
        self.add_label()        

    def initUI(self):
        self.parent.title("Validate_File_Application")
        self.pack(fill =tk.BOTH, expand =1)
        self.configure(background='#cdf7f0')

    def add_label(self):
        self.label = tk.Label(self.parent,
                    text='Please select reference number of the incident',
                    font=('helvetica', 12, 'bold'),
                    fg="white",
                    bg="#eb67ad",  
                    relief=RAISED                 
                )
        self.label.pack(anchor=CENTER, expand=1)       

    def run_process(self):   
        dialog = Reference_Part3.MyDialog3(self)
        self.reference_id = dialog.reference_ID
        self.OpportunityID = dialog.OpportunityID
        reference_no = str(self.reference_id)
        self.file_path = None            
        self.label['text'] = "Thanks for adding reference number"   
        self.account_id = None
        self.seqNum = '1'

        self.label['text'] = "Searching for the file"
        url = "https://ecenter.custhelp.com/services/rest/connect/v1.3/incidents/"

        find_lookupName= "https://ecenter.custhelp.com/services/rest/connect/v1.3/incidents?q=lookupName='" + reference_no + "'"
        #print(find_lookupName)
        find_lookupName_Page = requests.get(find_lookupName,auth=HTTPBasicAuth(username, password))
        incident_page = BeautifulSoup(find_lookupName_Page.content, 'html.parser')
        str_incident_page = str(incident_page)
        Incident = str_incident_page[45:(str_incident_page.find(','))]
        Incident_number = str(Incident)
        #print(Incident_number)

        file_path_url = url + '/'+ Incident_number + '/fileAttachments' 
        file_attachment =  requests.get(file_path_url,auth=HTTPBasicAuth(username, password))
        file_attachment_links = BeautifulSoup(file_attachment.content, 'html.parser')
        file_number= str(file_attachment_links)
        #print (file_number)

        fileAttachments_link = url +  Incident_number + '/fileAttachments'  + '/'
        first_index = file_number.rindex(fileAttachments_link)
        last_index = file_number.find('"',first_index +1)
        file_number_url =file_number[first_index:last_index]
        file_download_url = file_number_url + "?download"
        str_file_download_url = str(file_download_url)

        self.label['text'] = "Opening the file and validating the data " 
        response = requests.get(str_file_download_url,auth=HTTPBasicAuth(username, password))
        with open('Report.xlsx', 'wb') as output:
            output.write(response.content)

        print("Done!")       

    
        for file in os.listdir():
            if '.xlsx' in file:
                self.file_path  = file
          
        self.label['text'] = "Collecting the account number of incident" 

        primaryContact = url +  '/'+  Incident_number+  '/primaryContact'
        primaryContact_Page = requests.get(primaryContact, auth=HTTPBasicAuth(username, password))
        primaryContact_url = BeautifulSoup(primaryContact_Page.content, 'html.parser')
        new_primaryContact_url = str(primaryContact_url)
        link= new_primaryContact_url[115: 187]+ "/customFields/c/account_id"
        account_url  = requests.get(link, auth=HTTPBasicAuth(username, password))
        account = BeautifulSoup(account_url.content, 'html.parser')
        str_account = str(account)
        self.AccountID = str_account[21:(str_account.find('"',21))]
        print(self.AccountID)  

        data  = pd.read_excel(self.file_path, sheet_name= "TO ODS", engine = "openpyxl",dtype={'EmployeeZip':str} )
        df1 = pd.DataFrame(data)  
        df1.loc[(df1['EmployerApprovalEmployeeIndicator'] == 'Approved'),'OpportunityID' ] = self.OpportunityID
        df2 = df1[(df1['EligibilityReason']== 'Employee is eligible for ordering.')]   
        df2['EmployeeWFHID'] = df2['EmployeeWFHID'].astype(np.int64)  
        df2.drop(df2.columns[df2.columns.str.contains('unnamed',case = False)],axis = 1, inplace = True)
        df2 = df2.replace(0, "")
        new_name = "WFHEligibilityResults" + '_'+ self.AccountID + "_"+time.strftime("%m%d%Y%H%M%S")+ '_' + self.seqNum
                     
        df2.to_csv(new_name +".csv",sep =',', index= False, header = True, date_format='%Y-%m-%d %H:%M:%S', na_rep='' )
        self.label['text'] = "File has downloaded to desktop. Please press Upload to upload on s3 bucket."
        self.process.config(state = 'disabled')
        self.Upload.config(state = 'active')

        restURL = "https://ecenter.custhelp.com/services/rest/connect/v1.3/incidents/"+Incident_number+"/statusWithType/status/"
        params = {'lookupName':"Orders in Progress/Request completed"}       
        headers = {'X-HTTP-Method-Override':'PATCH'}
        resp = requests.post(restURL, json=params, auth=(username, password), headers=headers)
        print ("Status: ", resp)
       

    def Upload_S3_Bucket(self):
        self.label['text'] = "Uploading on S3 bucket... "
        s3 = boto3.resource(service_name = service_name,region_name = region_name ,aws_access_key_id = aws_access_key_id , aws_secret_access_key = aws_secret_access_key)
        try: 
            for file in os.listdir():
                if '.csv' in file:
                    file = file 
                    s3.Bucket(bucket_outbound).upload_file(Filename = file ,Key = bucket_inbound_folder+file)
        except FileNotFoundError:  
            self.label['text'] = "There is no file..."
            tk.messagebox.showerror("Information", "There is no file to upload.. ")
            return None 

        self.label['text'] = "Process completed. Please press done! "  
        self.Upload.config(state = 'disabled')     

        
    def createWidgets(self):        
        self.process = tk.Button(self)
        self.process["text"] = "Process",
        self.process["font"]=('helvetica', 12, 'bold')
        self.process['width'] = 30,
        #self.process['state'] = 'disabled',
        self.process["command"] = self.run_process
        self.process['bg']='#eb67ad',
        self.process.pack({"side": "top"},pady =10, padx = 5)

        self.Upload = tk.Button(self)
        self.Upload["text"] = "Upload",
        self.Upload["font"]=('helvetica', 12, 'bold')
        self.Upload['state'] = 'disabled',
        self.Upload['width'] = 30,
        self.Upload["command"] = self.Upload_S3_Bucket
        self.Upload['bg']='#eb67ad',
        self.Upload.pack({"side": "top"},pady =10, padx = 5) 

        self.QUIT = tk.Button(self)
        self.QUIT["text"] = "Done"
        self.QUIT["fg"]   = "white",
        self.QUIT["font"]=('helvetica', 12, 'bold')
        self.QUIT['width'] = 30,
        self.QUIT["command"] =  self.quit
        self.QUIT['bg']='#CA1F7B',
        self.QUIT.pack({"side": "top"},pady =10, padx = 5)


def main():
    root = tk.Tk()
    root.configure(bg='#cdf7f0')
    root.geometry("700x250")
    app = Record(parent=root)
    root.lift()
    root.call('wm', 'attributes', '.', '-topmost', True)
    root.after_idle(root.call, 'wm', 'attributes', '.', '-topmost', False)

    app.mainloop()

main()
