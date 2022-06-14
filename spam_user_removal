from copy import copy
import imaplib
import email
from sys import flags
import traceback 
import re
import datetime

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
import pyperclip

from openpyxl import Workbook

ORG_EMAIL = "@gmail.com" 
FROM_EMAIL = "" + ORG_EMAIL 
FROM_PWD = "" 
SMTP_SERVER = "imap.gmail.com" 
SMTP_PORT = 993

def read_email_from_gmail():
    try:
        mail = imaplib.IMAP4_SSL(SMTP_SERVER)
        mail.login(FROM_EMAIL,FROM_PWD)
        mail.select('inbox')

        data = mail.search(None, 'ALL')

        mail_ids = data[1]
        id_list = mail_ids[0].split()   
        first_email_id = int(id_list[0])
        latest_email_id = int(id_list[-1])

        for i in range(latest_email_id,first_email_id-1, -1):
            data = mail.fetch(str(i), '(RFC822)' )
            for response_part in data:
                arr = response_part[0]
                if isinstance(arr, tuple):
                    msg = email.message_from_string(str(arr[1],'utf-8'))
                    email_subject = msg['subject']
                    email_from = msg['from']
                    print('From : ' + email_from + '\n')
                    print('Subject : ' + email_subject + '\n')

                    try:
                        email_address = re.search('\((.+?)\).', email_subject).group(1) #prints the email address within open/closed brackets
                        mail.store(str(i),'+FLAGS', '\\Deleted')
                        mail.expunge()
                        web(email_address)
                    except:
                        print("Error: Subject from was too long")
                        global flag
                        flag = flag + 1

    except Exception as e:
        traceback.print_exc() 
        print(str(e))
    
def web(email_address):
    driver.get("#website desired#")

    elem = driver.find_element_by_name("search")
    pyperclip.copy(email_address) 
    elem.send_keys(Keys.CONTROL, 'v')

    driver.find_element_by_css_selector("button ").click()
    driver.find_element_by_css_selector(".rowsEven").click()

    userList = {"ID_email": "D","dateAdded": "A", "organisation": "C", "reason": "E", "firstname": "Y", "Lastname": "Z"}
    
    global row
    row = row + 1

    excel(driver,userList)

    driver.find_element_by_partial_link_text("Delete this record").click()

def excel(driver, userList):

    for (i,j) in userList.items():
        
        testcopy = driver.find_element_by_name(i)
        testcopy.send_keys(Keys.CONTROL, 'a')
        testcopy.send_keys(Keys.CONTROL, 'c')
        sheet[j+str(row)] = pyperclip.paste()
        

        if i == "Lastname":
            sheet["B"+str(row)] = "=Y" + str(row) +"&"+ "\" \"" + "&Z" + str(row)
            sheet["F"+str(row)] = "DECLINED"
            sheet["G"+str(row)] = datetime.datetime.now()
            sheet["H"+str(row)] = "Spam - profile deleted"
            return


row = 0
flag = 0
workbook = Workbook()
sheet = workbook.active

username = ""
password = ""
driver = webdriver.Chrome(executable_path="#insert path#\\Documents\\Chrome_Python\\chromedriver.exe")
driver.get("#selected website#")
driver.find_element_by_name("username").send_keys(username)
driver.find_element_by_name("pwd").send_keys(password)
driver.find_element_by_css_selector("button ").click()

read_email_from_gmail()
workbook.save(filename="spamusers.xlsx")
print("There is " + str(flag) + " email that could not be completed")
