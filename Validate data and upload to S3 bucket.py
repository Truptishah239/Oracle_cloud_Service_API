username="*********"
password="**********"
aws_access_key_id= '****************'
aws_secret_access_key= '********************'
service_name = 's3'
region_name = 'us-west-2'
bucket_outbound = '***********'
bucket_inbound_folder = '**********'

import requests
import tkinter as tk
import requests
from requests.auth import HTTPBasicAuth
from bs4 import BeautifulSoup
from tkinter import filedialog, messagebox
import openpyxl
import collections
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from pandas.api.types import is_string_dtype
from pandas.api.types import is_numeric_dtype
from tkinter import *
import webbrowser
import datetime
import sys
from tkinter import simpledialog
from tkinter import Frame, Tk, Button
import tkinter.messagebox as msg
import os
import logging
import boto3
from botocore.exceptions import ClientError
import Reference_Part1
import wget
import time


class file_validate(tk.Frame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.pack()
        self.createWidgets()
        self.parent= parent
        self.initUI()
        self.add_label()
        

    def initUI(self):
        self.parent.title("Validate_File and_Upload_to_S3")
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

        self.label['text'] = "Thanks for adding reference number" 

        dialog = Reference_Part1.MyDialog3(self)
        self.reference_id = dialog.reference_ID
        reference_no = str(self.reference_id)
        self.file_path = None


        self.label['text'] = "Searching for the file FILENAME.xlsx " 

        find_lookupName= "https://ecenter.custhelp.com/services/rest/connect/v1.3/incidents?q=lookupName='" + reference_no + "'"
        #print(find_lookupName)
        import requests
        find_lookupName_Page = requests.get(find_lookupName,auth=HTTPBasicAuth(username, password))
        incident_page = BeautifulSoup(find_lookupName_Page.content, 'html.parser')
        str_incident_page = str(incident_page)
        Incident = str_incident_page[45:(str_incident_page.find(','))]
        Incident_number = str(Incident)
        #print(Incident_number)

        url = "https://ecenter.custhelp.com/services/rest/connect/v1.3/incidents/"
        
        file_path_url = url + '/'+ Incident_number + '/fileAttachments' 
        file_attachment =  requests.get(file_path_url,auth=HTTPBasicAuth(username, password))
        file_attachment_links = BeautifulSoup(file_attachment.content, 'html.parser')
        file_number= str(file_attachment_links)
        self.label['text'] = "Downloading File FILENAME.xlsx" 
        '''
        file_number_url = file_number[(file_number.find('https:')):(file_number.find('"', (file_number.find('https:'))))]
        file_download_url = file_number_url + "?download"
        #print(file_download_url)
        str_file_download_url = str(file_download_url)
        #print(str_file_download_url)
        '''
        fileAttachments_link = url +  Incident_number + '/fileAttachments'  + '/'
        first_index = file_number.rindex(fileAttachments_link)
        last_index = file_number.find('"',first_index +1)
        file_number_url =file_number[first_index:last_index]
        file_download_url = file_number_url + "?download"
        str_file_download_url = str(file_download_url)


        self.label['text'] = "Opening the file and validating the data " 
        response = requests.get(str_file_download_url,auth=HTTPBasicAuth(username, password))
        with open('FILENAME.xlsx', 'wb') as output:
            output.write(response.content)
        print("Done!")       

    
        for file in os.listdir():
            if '.xlsx' in file:
                self.file_path  = file
                workbook = openpyxl.load_workbook(self.file_path,read_only=False, keep_vba=True) 
                std= workbook['SHEET1']
                workbook.remove(std)
                workbook.save('FILENAME.xlsx')
                workbook.close()            
        try:
            excel_filename = r"{}".format(self.file_path)
            if excel_filename[-4:] == ".csv":
                check_file = pd.read_csv(excel_filename)
            else:
                check_file = pd.read_excel(excel_filename, engine= 'openpyxl')
            
        except ValueError:
            self.label['text'] = "Invalid data"
            tk.messagebox.showerror("Information", "The file you have chosen is invalid")
            return None
        except FileNotFoundError:
            self.label['text'] = "No such file"
            tk.messagebox.showerror("Information", f"No such file as {self.file_path}")
            return None  
           
        def return_file(df, name):
            df.to_excel(name + ".xlsx")
            
        if ((len(check_file.columns)== 10) != True) and  (len(check_file.select_dtypes(['object']).dtypes) >= 8):
            tk.messagebox.showinfo("Valid File", "The file you have chosen is valid so far.") 
        elif(check_file.select_dtypes(['object']).apply(lambda x : x.str.len().gt(255)).any(axis =1).any() == True):
            #print ('Data with more than 255 char... ')
            df1 = check_file.loc[check_file.select_dtypes(['object']).apply(lambda x : x.str.len().gt(255)).any(axis =1)]
            return_file(df1, "character_More_Than_250")
            tk.messagebox.showerror("Invalid File", "Hello there! \nThe file has data with more than 255 char.Downloaded file on your desktop with incorrect data. Please check the file..")
            sys.exit() 
        elif (check_file.select_dtypes(['object']).apply(lambda x : x.str.contains(r'[&$%+*!+:=?><[~}{/)(^]')).any(axis =1).any()== True):
            #print ('Record with specific characters... ')
            df2 = check_file.loc[check_file.select_dtypes(['object']).apply(lambda x : x.str.contains(r'[&$%+*!+:=?><[~}{)(^/]')).any(axis =1)]
            return_file(df2, "Special_character")
            tk.messagebox.showerror("Invalid File", "Hello there! \nThe file has data with specific char.Downloaded file on your desktop with incorrect data. Please check the file..") 
            sys.exit() 
        else:  
            tk.messagebox.showinfo("Valid File", "Hello there! \nThis is a valid file. You can proceed now.\n" ) 
            self.label['text'] = "The file you have chosen is valid. Please press 'Process' to proceed."


        #Get Account id, reference ID, seq_num, Customer Type
        self.label['text'] = "File has valid data"
          
        self.account_id = None
        self.seqNum = '1'
        self.CustomerType =None   
        

        self.label['text'] = "Collecting the account number of incident" 

        primaryContact = url +  '/'+  Incident_number+  '/primaryContact'
        primaryContact_Page = requests.get(primaryContact, auth=HTTPBasicAuth(username, password))
        primaryContact_url = BeautifulSoup(primaryContact_Page.content, 'html.parser')
        new_primaryContact_url = str(primaryContact_url)
        link = new_primaryContact_url[(new_primaryContact_url.find("href")+8): (new_primaryContact_url.find('"',(new_primaryContact_url.find("href")+8)))] + "/customFields/c/account_id"
        #link= new_primaryContact_url[115: 187]+ "/customFields/c/account_id"
        account_url  = requests.get(link, auth=HTTPBasicAuth(username, password))
        account = BeautifulSoup(account_url.content, 'html.parser')
        str_account = str(account)
        self.account_id = str_account[21:(str_account.find('"',21))]
        print(self.account_id)    
        

        self.label['text'] = "Collecting the Customer Type of  incident"


        form_interface_url = url + '/'+ Incident_number + '/customFields/c/form_interface'
        Customer_Type_Page = requests.get(form_interface_url,auth=HTTPBasicAuth(username, password))
        Customer_Type = BeautifulSoup(Customer_Type_Page.content, 'html.parser')
        Customer = (str(Customer_Type).replace("{","").replace("}", ""))
        self.CustomerType = Customer[24:27]
        print(self.CustomerType)      

        link2= new_primaryContact_url[115: 187]+ "/customFields/c/sales_rep_email"
        print(link2)
        sales_rep_email_url  = requests.get(link2, auth=HTTPBasicAuth(username, password))
        sales_rep_email = BeautifulSoup(sales_rep_email_url.content, 'html.parser')
        str_sales_rep_email = str(sales_rep_email)
        #print(str_sales_rep_email)
        first_position = str_sales_rep_email.find('"',23)
        last_position = str_sales_rep_email.find('"',first_position+1)
        sales_rep = str_sales_rep_email[first_position+1: last_position ]
        #print(sales_rep) 
        name = sales_rep[0: (sales_rep.find('.'))]
        print(name)

        link3= new_primaryContact_url[115: 187]+ "/customFields/c/se_email_address"
        print(link3)
        se_email_url  = requests.get(link3, auth=HTTPBasicAuth(username, password))
        se_email = BeautifulSoup(se_email_url.content, 'html.parser')
        str_se_email = str(se_email)
        #print(str_se_email)
        first_position = str_se_email.find('"',24)
        last_position = str_se_email.find('"',first_position+1)
        se_email = str_se_email[first_position+1: last_position ]
        #print(se_email)             

      
        self.label['text'] = "Adding new column EmployeeReference3"
        data = pd.read_excel(self.file_path,sheet_name ="Employee Addresses", engine = "openpyxl",dtype={'EmployeeZip':str,'EmployeeEmail':str})
        df = pd.DataFrame(data)  
        print(df)
        df.dropna(how='all',inplace=True)
        #df = df.applymap(str)
        df.drop(df.columns[df.columns.str.contains('unnamed',case = False)],axis = 1, inplace = True)
        df['EmployeeZip'] = df['EmployeeZip'].astype(str).replace('\.0', '', regex=True)
        df['EmployeeReference1'] = df['EmployeeReference1'].astype(str).replace('\.0', '', regex=True)
             
        #print (df)
        var = self.CustomerType + reference_no
        df['EmployeeReference3'] = var
        print (df)

        self.label['text'] = "Adding new column Account ID"
        df['AccountID'] = self.account_id
        
        cols = df.columns.tolist()
        cols = cols[-1:]+cols[:-1]
        df = df[cols]
        #print(df)
      
        self.label['text'] = "Renaming the file and converting it into CSV" 
        df = df.fillna("")
        df = df.replace('nan', "")
        #df = df.applymap(str)
        new_name = "WFHEmployee" + '_' +self.account_id + "_"+time.strftime("%m%d%Y%H%M%S")+ '_' + self.seqNum
        df.to_csv(new_name +".csv",sep ='|', index= False, header = True, mode = 'a')

             
        restURL = "https://ecenter.custhelp.com/services/rest/connect/v1.3/incidents/"+Incident_number+"/statusWithType/status/"
        params = {'lookupName': "File Processing in Progress"}       
        headers = {'X-HTTP-Method-Override':'PATCH'}
        resp = requests.post(restURL, json=params, auth=(username, password), headers=headers)
        print ("Status: ", resp)
        #print(resp.json())

        restURL = "https://ecenter.custhelp.com/services/rest/connect/v1.3/incidents/"+Incident_number+"/customFields/c/"
        params = {'alternateemail':sales_rep + ","+ se_email}       
        headers = {'X-HTTP-Method-Override':'PATCH'}
        resp = requests.post(restURL, json=params, auth=(username, password), headers=headers)
        print ("alternateemail", resp)
   
        self.label['text'] = "File has downloaded to desktop. Please press Upload to upload on s3 bucket."
        self.Upload.config(state = 'active')
        self.process.config(state = 'disabled')  
      
        

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
        self.QUIT.config(state = 'active')  

    def createWidgets(self):
     
        self.process = tk.Button(self)
        self.process["text"] = "Select_reference_Number",
        self.process["font"]=('helvetica', 12, 'bold')
        self.process['fg'] = 'white'
        self.process['width'] = 30,
        #self.process['state'] = 'disabled',
        #self.process["font"] = "bold"
        self.process["command"] = self.run_process
        self.process['bg']='#eb67ad',
        self.process.pack({"side": "top"},pady =10, padx = 5)

        self.Upload = tk.Button(self)
        self.Upload["text"] = "Upload_file_to_S3",
        self.Upload["font"]=('helvetica', 12, 'bold')
        self.Upload['fg'] = 'white'
        self.Upload['state'] = 'disabled',
        #self.Upload["font"] = "bold"
        self.Upload['width'] = 30,
        self.Upload["command"] = self.Upload_S3_Bucket
        self.Upload['bg']='#eb67ad',
        self.Upload.pack({"side": "top"},pady =10, padx = 5) 

        self.QUIT = tk.Button(self)
        self.QUIT["text"] = "Done/Quit"
        self.QUIT["fg"]   = "white",
        self.QUIT["font"]=('helvetica', 12, 'bold')
        self.QUIT['width'] = 30,
        self.QUIT["command"] =  self.quit
        self.QUIT['bg']='#CA1F7B',
        self.QUIT.pack({"side": "top"},pady =10, padx = 5)

def main():
    root = tk.Tk()
    #root.configure(background='black')
    root.configure(bg='#cdf7f0')
    root.geometry("700x250")
    app = file_validate(parent=root)
    root.lift()
    root.call('wm', 'attributes', '.', '-topmost', True)
    root.after_idle(root.call, 'wm', 'attributes', '.', '-topmost', False)

    app.mainloop()

main()
