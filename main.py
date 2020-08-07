import pandas as pd
import datetime
import smtplib
# setting the pipeline for the project

# global a
a = False
# this function would send the email to neccessary things with three parameters to, subject and message
def sendEmail(to, sub, msg):
    global a
    # print("SENT MAIL")
    # setting the server for sending the mail
    server = smtplib.SMTP('smtp.gmail.com')

    # starting a tls session
    server.starttls()

    # now logging in
    server.login('myline.dicksjohn@gmail.com', 'johndicks1992')

    # sending the mail
    server.sendmail('myline.dicksjohn@gmail.com', to, sub, msg)
    a = True
    return a

if __name__ == "__main__":

    # loading the data frame in python.
    df = pd.read_excel('data.xlsx')
    # print(df)

    # current month and day
    today = datetime.datetime.now().strftime("%d-%m")
    # print(today)

    # current year
    yearNow = datetime.datetime.now().strftime("%Y")
    y = int(yearNow)

    # index list
    writeAppend = []

    # now we are going to iterate the dataframe, we can use df.iterrows fucntion that provied us all the values with their indexes
    for index, item in df.iterrows():

        # we can take the birthday in form day-month by adding strftime funnction
        # print(index, item['Birthday'].strftime("%d-%m"))
        bday = item['Birthday'].strftime("%d-%m")
        # now we will be comparing these birthdays with today's date and also check if year is matching and if mail has been sent or not. once sent we will update the year and sent status and send email accordingly
        # print(type(item["Year"]))
        if (today == bday) and (y == item["Year"]):
            sendEmail(item["Email"], "Happy Birthday Champion", item["Dialogue"])
            writeAppend.append(index)

    if a:
        print("??True")

    for i in writeAppend:
        df.at[i, "Year"] += 1

    df.to_excel('data6.xlsx', index=False)
