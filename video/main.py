import pandas as pd
import datetime
import smtplib

# email configuration
email_id = ""
password = ""

# send mail
def sendEmail(to, sub, msg):

    # server
    server = smtplib.SMTP('smtp.gmail.com', 587)

    # setting tls server
    server.starttls()

    # now logging
    server.login(email_id, password)

    # sending mail
    server.sendmail(email_id, to, msg)

# todays date
today = datetime.datetime.now().strftime("%d-%m")
# print(today)

# todays year
presentYear = datetime.datetime.now().strftime("%Y")
print(presentYear)
presentYear = int(presentYear)

# index
indexlist = []

if __name__ == "__main__":
    df = pd.read_excel('birthdaylist.xlsx')
    # print(df)
    for index, item in df.iterrows():
        if ((today == item["Birthday"].strftime("%d-%m")) and (presentYear == item["Year"])):
            sendEmail(item["EmailID"], "happy birthday champion", item["Message"])
            indexlist.append(index)
    
    for i in indexlist:
        df.at[i, "Year"] += 1
        print(df.at[i, "Year"])
    
    df.to_excel('birthdaylist.xlsx')
