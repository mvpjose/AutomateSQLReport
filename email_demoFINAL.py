import smtplib
import cx_Oracle
import xlsxwriter
import creds_demoFINAL
import email
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders

# smtp credentials
EMAIL_ADDRESS = creds_demoFINAL.email_address
EMAIL_PASSWORD = creds_demoFINAL.email_password
RECEIVER = 'johndoe123@gmail.com'

# Oracle DB credentials
DB_USER = creds_demoFINAL.database_username
DB_PASSWORD = creds_demoFINAL.database_password
DB_DSN = creds_demoFINAL.database_dsn
sql_query = creds_demoFINAL.sql_query.replace('\n', ' ')

# Create excel workbook
today = datetime.today()
fileLoc = 'C:\\Users\\johndoe123\\Downloads\\POC Activity\\Excel\\outputWorksheet ' + today.strftime('%m %d %Y %H %M %S') + '.xlsx'
filename = fileLoc.rsplit('\\')[-1]
workbook = xlsxwriter.Workbook(fileLoc)
worksheet = workbook.add_worksheet()

# Creates email content
msg = MIMEMultipart()
msg['From'] = EMAIL_ADDRESS
msg['To'] = RECEIVER
msg['Date'] = formatdate(localtime = True)
msg['Subject'] = 'Cases With Missing Patient DOB'
msg.attach(MIMEText('Hello,\nFor cases submitted within the last week, the following are missing patient DOB:', "plain"))

# Create SMTP instance with hose and port parameters 
with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
    smtp.ehlo() # Identify to smtp server
    smtp.starttls() # Set SMTP connect to TLS
    smtp.ehlo()
        
    smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        
    # Make Connection to database
    con = cx_Oracle.connect(DB_USER, DB_PASSWORD, DB_DSN)
    
    # Executes query
    cur = con.cursor()
    cur.execute(sql_query)
    tableHeaders = cur.description
    raw_results = cur.fetchall()
        
    # Start from the first cell. Rows and columns are zero indexed.
    row = 0
    col = 0

    # Sets headers for the worksheet
    for listIn, tuple in enumerate(tableHeaders):
        worksheet.write(row, col, tableHeaders[listIn][0])
        col += 1

    row = 1
    col = 0

    # Enters data from SQL query to Excel
    for listIn, tuple in enumerate(raw_results):
        for tupleItem in tuple:
            worksheet.write(row, col, tupleItem)
            col +=1
        col = 0
        row += 1
    
    cur.close()
    con.close()
    workbook.close()
  
    with open(fileLoc, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
        
    encoders.encode_base64(part)

    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {filename}",
    )

    msg.attach(part)
      
    smtp.sendmail(EMAIL_ADDRESS, RECEIVER, msg.as_string())