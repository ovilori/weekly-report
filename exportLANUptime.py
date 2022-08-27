# This program runs at 6:05PM every Friday on a solarwind server (or any server that can communicate with the solarwind DB) and does the following:
# 1. Connects to a solarwind database server to query and export the LAN uptime data for all site/location for the week.
# 2. Sends the error message via email if connection to the database server fails.
# 3. Calls a module to arrange the data for each location per days of the week.
# 4. Sends the raw data extracted from the DB and the sorted report via email.
# 5. Remove the copies of the exported and sorted report from the server/PC.

#importing required modules
from re import S
import pandas as pd
import pyodbc, os, smtplib
from datetime import date, datetime, timedelta
from email.message import EmailMessage
#import the custom sortLANreport module
from sortLANUptime import sortUptime
#defining the database connection variables.
server = 'ip address' 
database = 'database name' 
username = 'username' 
password = 'password' 
#defining the email variables.
msg = EmailMessage()
msg['Subject'] = 'Email Subject'
msg['From'] = 'sender email address'
recipients = ['recipient email address']
msg['To'] = ', '.join(recipients)
mail_server = smtplib.SMTP('mail server address', port no)
#connection to the database using the above variables.
try:
    connectDB = pyodbc.connect(
        'DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+
        ';DATABASE='+database+
        ';UID='+username+
        ';PWD='+ password,)
except pyodbc.Error as e:
    print(f"The error '{e}' occured.")
    error_time = datetime.now().strftime('%I:%M:%S %p')
    error_message = f"The error: '{e}' occured at " + error_time + ', while connecting to the database server to export this week\'s report.'
    #send error message as an email
    msg.set_content(error_message)
    mail_server.send_message(msg)
    mail_server.quit()
#setting the first date (Monday) and end date (Friday).
end_date = date.today()
start_date = end_date - timedelta(days = 4)
start_date = str(start_date) + 'T08:00:00'
end_date = str(end_date) + 'T16:00:00'
#SQL command to read the data
sqlQuery = '''SELECT Convert(Date, Datetime) AS SummaryDate, 
    Nodes.Caption AS NodeName, Nodes.IP_Address AS IP_Address, 
    ROUND(AVG(ResponseTime.Availability),2) AS AVERAGE_of_Availability 
	FROM Nodes INNER JOIN ResponseTime ON ( Nodes.NodeID = ResponseTime.NodeID )  WHERE  
	((DatePart(Hour,DateTime) >= 8) AND (DatePart(Hour,DateTime) <= 16) AND ((Nodes.IPType = 'XXX') OR (Nodes.IPType = 'XXX')))
	AND (DateTime BETWEEN ? AND ?)   
	GROUP BY Convert(Date,Datetime), 
	Nodes.Caption, Nodes.IP_Address  ORDER BY SummaryDate ASC, 3 ASC'''
value = (start_date,end_date)	
#getting the data from SQL into pandas dataframe
Query_output = pd.read_sql(sql = sqlQuery, con = connectDB, params=value)
#storing the time exported in a variable
report_date = datetime.now().strftime('%Y_%m_%d_%I_%M_%S_%p')
#Convert exported report to excel format
Query_output.to_excel('Weekly_Report_exported_on_' + report_date + '.xlsx', index=False)
# Query_output.to_excel(os.environ['userprofile'] + '\\Documents\\Weekly Report\\' + 'Weekly_Report_exported_on_' + report_date + '.xlsx', index=False)
# calling the sort module to arrange the data for each provider on seperate sheets.
sortedReport = sortUptime('Weekly_Report_exported_on_' + report_date + '.xlsx')
# reading the generated report to get the filename
# with open(os.environ['userprofile'] + '\\Documents\\Weekly Report\\' + 'Weekly_Report_exported_on_' + report_date + '.xlsx', 'rb') as file:
with open('Weekly_Report_exported_on_' + report_date + '.xlsx', 'rb') as file:
    report = file.read() 
    reportExcel_Filename = os.path.basename(file.name)
#reading the sorted report and getting the filename
with open(sortedReport, 'rb') as f:
    sorted = f.read()
    sortedReport_Filename = os.path.basename(f.name)
#sending the generated & sorted report as an email.
message_body = 'Locations LAN uptime report for the week attached.'
msg.set_content(message_body)
#adding generated report file as an attachment
msg.add_attachment(report, maintype = 'application', subtype = 'xlsx', filename = reportExcel_Filename)
#adding the sorted report file as an attachment
msg.add_attachment(sorted, maintype = 'application', subtype = 'xlsx', filename = sortedReport_Filename)
mail_server.send_message(msg)
mail_server.quit()
#cleanup the files from the server after sending the email.
os.remove('Weekly_Report_exported_on_' + report_date + '.xlsx')
os.remove(sortedReport)


