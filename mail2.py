#!/usr/bin/python3
import pandas as pd
import datetime
import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
#os.chdir("/root/Desktop/Birthdaywish_py")
fromaddr = "vismaygunda@gmail.com"
#toaddr = "vismaygunda@gmail.com"
#server.sendmail(fromaddr, toaddr, text)
def sendEmail(to, sub, msg):
	print(f"Email to {to} sent with subject: {sub} and message {msg}")
	msg = MIMEMultipart()
	msg['From'] = fromaddr
	msg['To'] = toaddr
	msg['Subject'] = "Python email"
	body = "Python test mail"
	msg.attach(MIMEText(body, 'plain'))
	server = smtplib.SMTP('smtp.gmail.com', 587)
	server.ehlo()
	server.starttls()
	server.ehlo()
	server.login("vismaygunda@gmail.com","******"#email password)
	text = msg.as_string()
	server.sendmail(fromaddr, to, f"Subject:{sub}\n\n{msg}")
	#server.sendmail(GMAIL_ID, to, f"Subject:{sub}\n\n{msg}")
	server.quit()

if __name__== "__main__":
	#sendEmail(fromaddr,"subject","test message")
	#exit()
	df=pd.read_excel("data.xlsx")
	#print(df)
	today=datetime.datetime.now().strftime("%d-%m-%y")
	yearNow=datetime.datetime.now().strftime("%Y")
	#print(today)
	writeInd=[]
    for index,item in df.iterrows():						
 		#print(index,item['Birthday'])
		#print(type(item['Year']))
		bday=item['Birthday'].strftime("%d-%m-%y")
		if(today==bday) and yearNow not in str(item['Year']):
 			sendEmail(item['Email'],"Happy Birthday", item['Dialogue'])
 			writeInd.append(index)
		#print(writeInd)
    for i in writeInd:
		yr=df.loc[i,'Year']
		df.loc[i,'Year']= str(yr)+ ',' +str(yearNow)
	    # print(df.loc[i,'Year'])
	# print(df)
	f.to_excel('data.xlsx',index=False)
