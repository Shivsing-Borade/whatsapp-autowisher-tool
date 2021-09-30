# Required Modules
import pandas as pd
import datetime
import pyautogui
import keyboard as k
import time
import pywhatkit as kit




# TODO 1. Send Whatsapp Messages
def send_message(number, msg, time_hr, time_min, time_sec):
    """This is  a simple function to send messages and wish them on any particular occasion at required time"""
    try:
        kit.sendwhatmsg(f'91+{number}', msg, time_hr, time_min, time_sec)
    except:
        print("Enter Proper Timing in hr:min:sec format")

    pyautogui.click(1050, 950)
    time.sleep(2)
    k.press_and_release('enter')  # For pressing enter button


if __name__ == '__main__':
    """To Run this Sotware you have to create data.xlsx file to store data and column must be Name, Number,Birthdate,Year"""
    try:
        df = pd.read_excel('data.xlsx')  # Read xls file
    except:
        print("Create an excel file  data.xlsx")
    #  print(df)

    today = datetime.datetime.now().strftime("%d-%m")  # today's date
    yearNow = datetime.datetime.now().strftime("%Y")  # current year
    # print(type(today))
    writeInd = []
print("********** Whatsapp Autowisher Tool ***********")

try:
    for index, item in df.iterrows():

        bday = item['Birthday'].strftime("%d-%m") # Birthdate calculated
        #print(bday)
        #print(today)



        if (today == bday) and yearNow not in str(item['Year']):   # check today== bday if yes then send_message

            hr = int(input("Enter time in 24 hr format:"))
            minute = int(input("Enter minutes in which u want to send msg:"))
            sec = int(input("Enter seconds in which u want to send:"))
            send_message(item['Number'], item['Wish'] + ' ', hr, minute, sec)
            writeInd.append(index)


    for i in writeInd:
        yr = df.loc[i, 'Year']
        df.loc[i, 'Year'] = str(yr) + ', ' + str(yearNow)

    df.to_excel('data.xlsx', sheet_name='Sheet1', index=False)

except:
    print("Enter data in exel sheet and try again ")
