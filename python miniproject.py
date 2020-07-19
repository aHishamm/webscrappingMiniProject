#importing the relevant libraries 
import smtplib
import openpyxl 
import requests
from bs4 import BeautifulSoup
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders

url = "https://money.cnn.com/data/markets/"
def webScrappingNames(url): 
    html = requests.get(url)
    soup = BeautifulSoup(html.text , 'html.parser')
    stockNameList = soup.find_all(class_="column stock-name")
    stockNames = [] 
    for stockName in stockNameList: 
        stockNames.append(stockName.text)
    return stockNames
def webScrappingPrices(url):
    html = requests.get(url)
    soup = BeautifulSoup(html.text , 'html.parser')
    stockPriceList = soup.find_all(class_="column stock-price")
    stockPrices = [] 
    for stockPrice in stockPriceList: 
        stockPrices.append(stockPrice.text)
    return stockPrices
def openpyXLlisting(filename):
    wb = openpyxl.Workbook()
    dest_filename = input(filename)
    ws1 = wb.active
    ws1.title = "range names"
    for row in range(1, 40):
        ws1.append(range(600))
    wb.save(excelfilename=dest_filename)
    wb = openpyxl.load_workbook(excelfilename)
    sheet = wb.get_sheet_by_name('Sheet1')
    for i in range(1,len(stockNames)+1): 
        sheet["A"+str(i)] = stockNames[i-1]
    for i in range(1,len(stockPrices)+1):
        sheet["B"+str(i)] = stockPrices[i-1]
    wb.save(excelfilename)
def sendMail(email,password,recipientEmail,filename,subject):
    mail = smtplib.SMTP('smtp.gmail.com',587)
    mail.ehlo()
    mail.starttls()
    email = email
    password = password
    mail.login(email,password)
    msg = MIMEMultipart()
    msg['Subject']=subject
    msg['From']=email
    msg['To']=recipientEmail
    filename = filename
    attachment = open(filename, 'rb')
    xlsx = MIMEBase('application','vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    xlsx.set_payload(attachment.read())
    encoders.encode_base64(xlsx)
    xlsx.add_header('Content-Disposition', 'attachment', filename=filename)
    msg.attach(xlsx)
    mail.sendmail(email,recipientEmail, msg.as_string())

stockNames = webScrappingNames(url)
stockPrices = webScrappingPrices(url)
filename = input("Please enter a path for an xlsx file: ")
openpyXLlisting(filename)
email = input("Please enter your Gmail email address: ")
password = input("Please enter your Gmail email password: ")
recipientEmail = input("Please enter the recipient email address: ")
subjectEmail = input("Please enter the subject of the email: ")
sendMail(email,password,recipientEmail,filename)










