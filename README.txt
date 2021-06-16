Proof of Concept: Automate SQL to Email Listing
================================================

	This project solves the problem of running a query on SQL Devloper, downloading the 
	results as an excel file, and attaching them to an email. 

Required Libraries
==================

	smtp
	cx_Oracle
	xlsxWriter
	email
	
How to Read
===========

Code will be highlighted as shown below: 

	/*****************************************************************************

		...code snippet...

	******************************************************************************\
	
Initializing variables and Instantiating Objects
================================================

	/*****************************************************************************

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
		fileLoc = 'C:\\Users\johndoe123\\Downloads\\POC Activity\\Excel\\outputWorksheet ' + today.strftime('%m %d %Y %H %M %S') + '.xlsx'
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

	******************************************************************************\

	The code above sets credentials to successfully login to gmail SMTP server, Oracle
	database from an external file (called creds_demoFINAL created in Python Script folder),
	sets excel workbook location, creates excel workbook, and creates body of the email to 
	be sent.

Create Connection to SMTP Server
================================

	/*****************************************************************************

		with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
			smtp.ehlo() # Identify to smtp server
			smtp.starttls() # Set SMTP connect to TLS
			smtp.ehlo()
				
			smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)

	******************************************************************************\
	
Oracle Database Actvities
==========================

	/*****************************************************************************

		with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
			
			...
			
			con = cx_Oracle.connect(DB_USER, DB_PASSWORD, DB_DSN)
    
			cur = con.cursor()
			cur.execute(sql_query)
			tableHeaders = cur.description
			raw_results = cur.fetchall()

	******************************************************************************\
	
	Logs into Oracle database, creates crusor, executes sql query (found in creds_demoFINAL),
	and returns data (as a list of tuples) from the query as well as the column 
	names (set as tableHeaders).
	
List of Tuples to Excel
=======================

	/*****************************************************************************

		with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
			
			...
			
		    row = 0
			col = 0

			for listIn, tuple in enumerate(tableHeaders):
				worksheet.write(row, col, tableHeaders[listIn][0])
				col += 1

			row = 1
			col = 0

			for listIn, tuple in enumerate(raw_results):
				for tupleItem in tuple:
					worksheet.write(row, col, tupleItem)
					col +=1
				col = 0
				row += 1

	******************************************************************************\
	
	Converts the output from cx_Oracle (list of tuples) to a readable format that 
	will be written onto the excel workbook. 
	
Attach Results to Email
=======================

	/*****************************************************************************
		
		with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
			
			...
			
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

	******************************************************************************\
	
	Converts excel workbook as an attachment, encodes it, sets title of the attachment,
	and sends the email.
	
																				FIN