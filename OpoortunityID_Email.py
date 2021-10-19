
import smtplib
from tkinter import *
from tkinter import simpledialog
from tkinter import Frame, Tk, Button
import tkinter.messagebox as msg
import tkinter as tk
import requests
from requests.auth import HTTPBasicAuth
from bs4 import BeautifulSoup
from email.message import EmailMessage
username=""
password=""
import requests

class SendEmail(tk.Frame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.pack()
        self.createWidgets()
        self.parent= parent
        self.initUI()
        self.add_label()         

    def initUI(self):
        self.parent.title("Python Mail Send App")
        self.pack(fill =tk.BOTH, expand =1)
        self.configure(background='#cdf7f0')
        
    def add_label(self):
        self.label = tk.Label(self.parent,
                    text='',
                    font=('helvetica', 12, 'bold'),
                    fg="white",
                    bg="#eb67ad",  
                    relief=RAISED                 
                )
        self.label.pack(side = LEFT)  
      

    def send_message(self): 
        self.reference_ID = self.reference_ID.get()
        reference_no = str(self.reference_ID)

        self.No_of_Employees = self.No_of_Employees .get()
        No_of_Employees = str(self.No_of_Employees)

        
        import requests
        find_lookupName= "https://ecenter.custhelp.com/services/rest/connect/v1.3/incidents?q=lookupName='" + reference_no + "'"
        print(find_lookupName)
        find_lookupName_Page = requests.get(find_lookupName,auth=HTTPBasicAuth(username, password))
        incident_page = BeautifulSoup(find_lookupName_Page.content, 'html.parser')
        str_incident_page = str(incident_page)
        Incident = str_incident_page[45:(str_incident_page.find(','))]
        Incident_number = str(Incident)
        print(Incident_number)
        url = "https://ecenter.custhelp.com/services/rest/connect/v1.3/incidents/"

        primaryContact = url +  '/'+  Incident_number+  '/primaryContact'
        primaryContact_Page = requests.get(primaryContact, auth=HTTPBasicAuth(username, password))
        primaryContact_url = BeautifulSoup(primaryContact_Page.content, 'html.parser')
        new_primaryContact_url = str(primaryContact_url)
        link= new_primaryContact_url[115: 187]+ "/customFields/c/account_id"
        account_url  = requests.get(link, auth=HTTPBasicAuth(username, password))
        account = BeautifulSoup(account_url.content, 'html.parser')
        str_account = str(account)
        account_id = str_account[21:(str_account.find('"',21))]
        print(account_id)  

        form_interface_url = url + '/'+ Incident_number + '/customFields/c/form_interface'
        Customer_Type_Page = requests.get(form_interface_url,auth=HTTPBasicAuth(username, password))
        Customer_Type = BeautifulSoup(Customer_Type_Page.content, 'html.parser')
        Customer = (str(Customer_Type).replace("{","").replace("}", ""))
        CustomerType = Customer[24:27]
        #print(CustomerType)            
      
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

        link4= new_primaryContact_url[115: 187]+ "/customFields/c/account_name"
        #print(link3)
        account_name_url  = requests.get(link4, auth=HTTPBasicAuth(username, password))
        account_name_email = BeautifulSoup(account_name_url.content, 'html.parser')
        str_account_name = str(account_name_email)
        #print(str_account_name)
        first_position = str_account_name.find('"',21)
        last_position = str_account_name.find('"',first_position+1)
        account_name = str_account_name[first_position+1: last_position ]
        #print(account_name) 

        import requests
      
        restURL = "https://ecenter.custhelp.com/services/rest/connect/v1.3/incidents/"+ Incident_number +"/statusWithType/status/"
        params = {'lookupName':"AE Opportunity ID Requested"}       
        headers = {'X-HTTP-Method-Override':'PATCH'}
        resp = requests.post(restURL, json=params, auth=(username, password), headers=headers)
        print ("Status: ", resp)

     
        self.label['text'] = "Sending Email... "  
        To_address_field = sales_rep
        Cc_address_field = se_email
        msg = EmailMessage()
 
        status_msg = '''
        Hi {},
        Ref Home Office Internet Reference {}- {}
        Your Customer has provided their order approvals for their Home Office Internet Employee locations for the below account.  
   
        SalesAccount ID: {}
        Account Name: {}

        Total No. of Employees Approved for Ordering: {}
        
        Please reply back within 24 hours with the Opportunity ID number(s) from Salesforce and quantity of lines to be used for ordering. 
        The total quantity across opportunities (if more than one) should match the above number approved for ordering. 

        To locate Opportunity ID in Salesforce, go to the Opportunity tab, click Detail, then scroll down to system information. 
        
        OpportunityID(s):  
        No. Of Lines: 

        Thanks,
        Tier 2 Support
        Trupti  '''.format (name,CustomerType,reference_no,account_id,account_name,No_of_Employees)
        msg.set_content(status_msg)
        msg['subject'] = 'Home Office Internet Opportunity Details Needed for {} {}'.format(account_name,account_id)
        #print(address_field,reference_no)        
        sender_email = "" 
        sender_password = ""
        msg['From']=sender_email
        #TO_address = [sender_email , To_address_field]
        #msg ['To'] = ','.join(TO_address)
        msg ['To'] = To_address_field + ',' +"" 
        msg ['Cc'] = Cc_address_field
        server = smtplib.SMTP('smtp-mail.outlook.com',587)
        server.starttls()
        #server.ehlo()
        server.login(sender_email,sender_password)
        self.label['text'] ="Login successful."
        #Address = (To_address_field + ','+ Cc_address_field)
        #Address = (address_field)
        server.send_message(msg)
        self.label['text'] = "Message sent." 

        import requests
        restURL = "https://ecenter.custhelp.com/services/rest/connect/v1.3/incidents/"+Incident_number+"/statusWithType/status/"
        params = {'lookupName': "AE Opportunity ID Requested"}       
        headers = {'X-HTTP-Method-Override':'PATCH'}
        resp = requests.post(restURL, json=params, auth=(username, password), headers=headers)
        print ("Status: ", resp)
        server.quit()
        self.label['text'] = "Completed.Press Done!"     
     
    def createWidgets(self):
        self.reference_ID_title = tk.Label(self, text="reference_ID ",font=('helvetica', 12, 'bold'),
                    fg="white",
                    bg="#eb67ad", relief=RAISED)
        self.reference_ID_title.pack({"side": "top"},pady =10, padx = 5)

        # Simple Entry widget:
        self.reference_ID = tk.Entry(self, width=80)
        self.reference_ID.pack({"side": "top"},pady =10, padx = 5)
        self.reference_ID.delete(0, tk.END)   

        self.No_of_Employees_title = tk.Label(self, text="Total No. of Employees Approved for Ordering ",font=('helvetica', 12, 'bold'),
                fg="white",
                bg="#eb67ad", relief=RAISED)
        self.No_of_Employees_title.pack({"side": "top"},pady =10, padx = 5)

        # Simple Entry widget:
        self.No_of_Employees = tk.Entry(self, width=80)
        self.No_of_Employees.pack({"side": "top"},pady =10, padx = 5)
        self.No_of_Employees.delete(0, tk.END)   

        self.send = tk.Button(self)
        self.send["text"] = "Send_Message",
        self.send['width'] = 12,
        self.send["font"]=('helvetica', 12, 'bold')
        self.send["command"] = self.send_message
        self.send['bg']='#eb67ad',
        self.send.pack({"side": "top"},pady =10, padx = 5)

        self.cancel_button = tk.Button(self)
        self.cancel_button["text"] = "Done",
        self.cancel_button['width'] = 12,
        self.cancel_button["font"]=('helvetica', 12, 'bold')
        self.cancel_button["command"] = quit
        self.cancel_button['bg']='#eb67ad',
        self.cancel_button.pack({"side": "bottom"},pady =10, padx = 5)
       
def main():
    root = tk.Tk()
    root.configure(bg='#cdf7f0')
    app = SendEmail(parent=root)
    root.lift()
    root.call('wm', 'attributes', '.', '-topmost', True)
    root.after_idle(root.call, 'wm', 'attributes', '.', '-topmost', False)

    app.mainloop()

main()