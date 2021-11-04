#this program connects to the Solarwinds DB to query the weekly ISP performance report and sorts the data.
#the program runs on my computer at 6:15PM every Friday

#importing required modules
from email import message
import pandas as pd
import pyodbc, os, smtplib
from datetime import date, datetime, timedelta
from email.message import EmailMessage
import sort_module_v1 #the module that sorts the exported data

#defining the database connection parameters
server = 'x.x.x.x' #replace with the DB IP/hostname
database = 'database' #replace with the DB name
username = 'username' #replace with the username
password = 'password' #replace with the password

#defining the email parameters
msg = EmailMessage()
msg['Subject'] = 'Daily ISP Availability Report (8AM - 6PM) - WEEKLY'
msg['From'] = 'abcd@yxz.com' #replace with real email address
recipients = ['efgh@xyz.com', 'ijkl@xyz.com'] #replace with real email addresses
msg['To'] = ', '.join(recipients)
mail_server = smtplib.SMTP('mail_server', port_no) #replace mail_server & port_no with real mail server address # port no. Use smtplib.SMTP_SSL for secured connection.

#connection to the database using the above parameters
try:
    connectDB = pyodbc.connect(
        'DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+
        ';DATABASE='+database+
        ';UID='+username+
        ';PWD='+ password,)
except pyodbc.Error as e:
    print(f"The error '{e}' occured.")
    error_time = datetime.now().strftime('%I:%M:%S %p')
    error_message = f"The error: '{e}' occured at " + error_time + ', while connecting to the Solarwinds DB to export this week\'s report.'
    #send error message via email
    msg.set_content(error_message)
    mail_server.send_message(msg)
    mail_server.quit()

#setting the first date (Monday) and end date (Friday)
end_date = date.today()
diff = timedelta(days = 4)
start_date = end_date - diff
start_date = str(start_date) + 'T08:00:00'
end_date = str(end_date) + 'T18:00:00'
#SQL command to read the data
sqlQuery = 'SELECT Convert(Date,Datetime) AS SummaryDate, \
        Nodes.Caption AS NodeName, Nodes.VendorName AS VendorName,\
        ROUND(AVG(ResponseTime.Availability),2) AS AVERAGE_of_Availability \
        FROM Nodes INNER JOIN ResponseTime ON ( Nodes.NodeID = ResponseTime.NodeID ) WHERE \
        ((DatePart(Hour,DateTime) >= 8) AND (DatePart(Hour,DateTime) <= 18) AND ((Nodes.ISP_ = \'ABC_ISP\'))) \
        AND (DateTime BETWEEN ? AND ?) \
        GROUP BY Convert(Date,Datetime), \
        Nodes.Caption, Nodes.VendorName  ORDER BY SummaryDate ASC, 3 ASC'

value = (start_date,end_date)
#getting the data from SQL into pandas dataframe
Query_output = pd.read_sql(sql = sqlQuery, con = connectDB, params=value)

#storing the time generated in a variable
report_date = datetime.now().strftime('%Y_%m_%d_%I_%M_%S_%p')

#Export the data to weekly report folder
Query_output.to_excel(os.environ['userprofile'] + '\\Documents\\Weekly ISP Availability Report\\' + 'Weekly_Report_exported_on_' + report_date + '.xlsx', index=False)

#calling the sort module to sort the data for each provider
sorted_report = sort_module_v1.sort_report(os.environ['userprofile'] + '\\Documents\\Weekly ISP Availability Report\\' + 'Weekly_Report_exported_on_' + report_date + '.xlsx')

#reading the generated report and getting the filename
with open(os.environ['userprofile'] + '\\Documents\\Weekly ISP Availability Report\\' + 'Weekly_Report_exported_on_' + report_date + '.xlsx', 'rb') as file:
    report = file.read()
    report_name = os.path.basename(file.name)

#reading the sorted report and getting the filename
with open(sorted_report, 'rb') as f:
    sorted = f.read()
    sorted_name = os.path.basename(f.name)

#sending the generated report & sorted report as an email.
message_body = 'Daily ISP availability report for the week attached.'
msg.set_content(message_body)

#adding generated report file as an attachment
msg.add_attachment(report, maintype = 'application', subtype = 'xlsx', filename = report_name)

#adding sorted report file as an attachment
msg.add_attachment(sorted, maintype = 'application', subtype = 'xlsx', filename = sorted_name)
mail_server.send_message(msg)
mail_server.quit()