import time

import serial
import openpyxl
from datetime import datetime

# EXCEL SHEET INITIALIZER
wb = openpyxl.Workbook()


def getuserinput():
    inp_1 = input("Do yo want to proceed with default location (y/n): ")
    if inp_1 == "y":
        loc = "C:\\Users\\abhil\\OneDrive\\Documents\\RMC-ProJECT.xlsx"
        return loc
    elif inp_1 == "n":
        loc = input("Enter the file path you want to save in:")
        return loc


file = getuserinput()
ws1 = wb.active

# EXCEL LOCATION VARIABLES
i = 0
j = 0

# CARD_ID WITH NAME DICT
data = {"A72F7563": "White Card", "994C4EA2": "Blue_Tag", "CAC6BE2C": "Abhilash Francis"}
# CARD_ID LIST
IDs = ["A72F7563", "994C4EA2", "CAC6BE2C"]
# SERIAL COMMUNICATION INITIALIZER
ser = serial.Serial("COM4", 9600)
ser.timeout = 1


def check_duplicate(name):
    k = 1
    if ws1['A1'].value is None:
        return True
    else:
        col = "A" + str(k)
        k = 0
        while ws1[col].value is not None:

            k = k + 1
            col = "A" + str(k)
            if name == ws1[col].value:
                return False
        return True


def add_data(name, i, date):
    column_for_name = "A" + i
    column_for_time = "B" + i
    ws1[column_for_name] = name
    ws1[column_for_time] = date
    print("Data Entered")
    wb.save(filename=file)


print("\nPlease Place Your Tag")
while True:
    Id_read = ser.readline().decode('ascii')
    Id_read = str(Id_read).strip()
    print(Id_read)
    for ID in IDs:
        if Id_read == ID:
            print(data[ID])
            print("Valid Tag")
            now = datetime.now()
            date = now.strftime("%d/%m/%Y %H:%M:%S")
            j = 1
            if check_duplicate(data[ID]):
                i = i + 1
                add_data(data[ID], str(i), date)
            else:
                print("Attendee already Exists")
    if len(Id_read) != 0 and j == 0:
        j = 0
        print("invalid tag")
    # ser.write("n")
    # column = "A" + str(j)
    # ws2 = wb.create_sheet(title="invalid_tags")
    # ws2[column] = Id_read
    # j = j + 1
    # Developed By DIONEX11 (for more queris mail: code.testx11@gmail.com)
