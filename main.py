import tkinter as tk
from tkinter import filedialog as fd
from tkinter import PhotoImage
from tkinter.messagebox import showinfo
from csv import excel
import pandas as pd
import numpy as np
import time


# Global variables
path_array = []
column_data = []
primary_key = ""
present_column_name = ""
isRollNo = False
isEmail = False
isPhone = False



# info_btn = tk.PhotoImage(file='Information_icon.png')
primary_key_input = ""
def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.

def render_heading():
    text2 = tk.Label(root, text="Enter excel sheets path ", font=("Arial", 16), pady=10 )
    text2.pack()

def select_file():
    file_types = (
        ("All files","*.*")
    )
    filename = fd.askopenfilename(
        title="Open a file",
        initialdir="/",
        filetypes=file_types
    )


def handle_primary_key(primary_key_input):
    global primary_key
    primary_key = primary_key_input.get()
    print(primary_key)
    get_presenty_column_name()

def handle_presenty_column(present_column_input):
    global present_column_name
    present_column_name = present_column_input.get()
    print(present_column_name)
    prepare_column_data()




def prepare_column_data():
    for file in path_array:
        # file = file.replace("/", "//")
        print(file)
        dataset = pd.read_excel(file)
        if "," in primary_key:
            keys = primary_key.split(",")
            column_data.append({"Primary Key": keys, "Present_Column_Name": present_column_name, "dataset": dataset})
            print(primary_key.split(","))
        else:
            column_data.append({"Primary Key": primary_key, "Present_Column_Name": present_column_name, "dataset": dataset})
    print("column data")
    print(column_data)
    get_necessary_data()

def get_necessary_data():
    # Main Array
    # Here we fetch necessary data from dataset like roll no and status column
    actual_data = []
    is_composite_primary_key = False
    for item in column_data:
        data_arr = []
        if type(item.get("Primary Key")) is list:
            is_composite_primary_key = True
            # LOOP if primary key is list
            for li in item.get("Primary Key"):
                print("Item: " + li)
                print(type(item.get("dataset")))
                # Formatting key
                li1 = str(f"{li}").strip()
                # Fetching data from dataset and putting in data array
                data_arr.append(item.get("dataset").get(li1))
        else:
            # If key is only one
            li1 = item.get("Primary Key").strip()
            #         print(item.get("Primary Key"))
            data_arr.append(item.get("dataset").get(li1))
        data_arr.append(item.get("dataset").get(item.get("Present_Column_Name")))
        actual_data.append(data_arr)
        print(actual_data)


# Format Single Primary Key
# Get key and formatting according to it's type... e.g. key types
# rollno, emailid, phone no., etc

def identify_key_and_format(item_name, data):
    item_name = item_name.lower()

    rollno = ["roll", "rollno", "roll_no", "roll number", "roll no"]
    email = ["email", "emailid", "email_id"]
    phone = ["phone", "phoneno", "phone no", "phone_no"]
    global isRollNo
    global isEmail
    global isPhone

    #    Check if key is rollno
    for item in rollno:
        if (item_name in item):
            isRollNo = True
            break

    #    Check if key is email
    for item in email:
        if (item_name in item):
            isEmail = True
            break

    #     Check if key is phone
    for item in phone:
        if (item_name in item):
            isPhone = True
            break

    #    Check if key is email
    for item in email:
        if (item_name in item):
            isEmail = True
            break

    #     Check if key is phone
    for item in phone:
        if (item_name in item):
            isPhone = True
            break

            key_arr = []
    if isRollNo == True:
        print("Roll no. formatting")

        for item1 in data:
            item1 = str(item1)
            key = np.nan
            if len(item1) > 3:
                print(str(item1[len(item1) - 3:]))
                key = str(item1[len(item1) - 3:])
                key = int(key)
                key_arr.append(key)
            else:
                print(item1)
                key = item1
                key = int(key)
                key_arr.append(key)

    elif isEmail == True:
        print("Email formatting")
    elif isPhone == True:

        print("Phone Formatting")

    print("Called...", item_name)
    return key_arr

# Handle Composite Primary Key
def handle_composite_primary_key(data):
    #     print(data)
    global key_arr1
    global isCompRollNo
    global isCompEmail
    global isCompPhone
    key_arr1 = ["" for i in range(len(data[0]))]
    for item in data:
        rollno = ["roll", "rollno", "roll_no", "roll number", "roll no"]
        email = ["email", "emailid", "email_id"]
        phone = ["phone", "phoneno", "phone no", "phone_no"]
        column_name = item.name
        column_name = column_name.lower()
        print("HIII: ", column_name)
        #    Check if key is rollno
        for li in rollno:
            if (column_name in li):
                isCompRollNo = True
                break
        if isCompRollNo == True:
            print("Roll no. formatting")

            for i in range(len(item)):
                item1 = str(item[i])
                key = np.nan
                if len(item1) > 3:
                    key = str(item1[len(item1) - 3:])
                    key = int(key)
                    key_arr1[i] = key_arr1[i] + str(key)
                else:
                    #                     print(item1)
                    key = item1
                    key = int(key)
                    key_arr1[i] = key_arr1[i] + str(key)

        elif isCompEmail == True:
            print("Email formatting")
        elif isCompPhone == True:
            print("Phone Formatting")
        else:
            for i in range(len(item)):
                item1 = str(item[i])
                key = str(item1)
                key_arr1[i] = key_arr1[i] + key

        isCompRollNo = False
        isCompEmail = False
        isCompPhone = False

    print("Handle composite primary key...")
    return key_arr1


def show_info():
    showinfo("Info", "Primary key for all Excelsheets must be same")

def get_primary_key():
    text2 = tk.Label(root, text="Enter a primary key", font=("Arial", 16), padx=10, pady=20)
    text2.pack()
    label1 = tk.Label(primaryKeyInputFrame, text="Enter primary key for Excelsheets: ", font=("Arial", 10))
    label1.grid(row=0, column=0)
    primary_key_input = tk.Entry(primaryKeyInputFrame, font=("Arial", 16))
    primary_key_input.grid(row=0, column=1, sticky=tk.W)

    btn1 = tk.Button( primaryKeyInputFrame, text="I", font=("Arial italic bold", 10), padx=10, fg="white", bg="blue", command=show_info)
    btn1.grid(row=0, column=2, sticky=tk.W)
    primaryKeyInputFrame.pack(padx=10, pady=10)
    btn = tk.Button(root, text="Submit", font=("Arial", 10), command=lambda: handle_primary_key(primary_key_input))
    btn.pack()


def get_presenty_column_name():
    text2 = tk.Label(root, text="Enter a present column", font=("Arial", 16), padx=10, pady=20)
    text2.pack()
    label1 = tk.Label(presentyColumnInputFrame, text="Enter a presenty column name: ", font=("Arial", 10))
    label1.grid(row=0, column=0)
    present_column_input = tk.Entry(presentyColumnInputFrame, font=("Arial", 16))
    present_column_input.grid(row=0, column=1, sticky=tk.W)

    btn1 = tk.Button( presentyColumnInputFrame, text="I", font=("Arial italic bold", 10), padx=10, fg="white", bg="blue", command=show_info)
    btn1.grid(row=0, column=2, sticky=tk.W)
    presentyColumnInputFrame.pack(padx=10, pady=10)
    btn = tk.Button(root, text="Submit", font=("Arial", 10), command=lambda: handle_presenty_column(present_column_input))
    btn.pack()


# input excel files
def render_input_UI(no_of_excels):
    render_heading()
    inputFrame = tk.Frame(root)
    inputFrame.columnconfigure(0, weight=1)
    inputFrame.columnconfigure(1, weight=1)
    inputFrame.columnconfigure(2, weight=1)
    for i in range(no_of_excels):
        filename = fd.askopenfilename()
        # Add to file path array
        path_array.append(filename)
        input_field = tk.Entry(inputFrame, font=("Arial", 16))
        input_field.grid(row=i, column=0, pady=10)
        btn1 = tk.Button(inputFrame, text="File",  font=("Arial", 12))
        btn1.grid(row=i, column=1)
        # input_field.delete(0,END)
        input_field.insert(0,filename)
        # fill= tk.Y, expand=True, padx=2, pady=2
        inputFrame.pack()
        time.sleep(1)
    inputFrame.pack()
    print(path_array)
    get_primary_key()

def get_excel_no():
    no_of_excels = int(input1.get())
    print("Function called....")
    print(no_of_excels)
    render_input_UI(no_of_excels)

print_hi('PyCharm')
root = tk.Tk()
root.title("Excelizer")
root.geometry("500x500")
label1 = tk.Label(root,text="Excelizer",font=("Arial",18))
label1.pack(padx=10,pady=10)


# label2 = tk.Label(root,text="The only excel analyzer you need😎")
# label2.pack()
# label3 = tk.Label(root,text="Enter the number of Excelsheets: ")
# label3.pack()
# tetxbox = tk.Text(root,height=1, font=("Arial",16))
# tetxbox.pack()

excelInputFrame = tk.Frame(root)
excelInputFrame.columnconfigure(0, weight=1)
excelInputFrame.columnconfigure(1, weight=1)
excelInputFrame.columnconfigure(2, weight=1)

primaryKeyInputFrame = tk.Frame(root)
primaryKeyInputFrame.columnconfigure(0, weight=1)
primaryKeyInputFrame.columnconfigure(1, weight=1)
primaryKeyInputFrame.columnconfigure(2, weight=1)

presentyColumnInputFrame = tk.Frame(root)
presentyColumnInputFrame.columnconfigure(0, weight=1)
presentyColumnInputFrame.columnconfigure(1, weight=1)
presentyColumnInputFrame.columnconfigure(2, weight=1)

text1 = tk.Label(excelInputFrame, text="Enter the number of Excelsheets:", font=("Arial",10))
text1.grid(row=0, column=0)
input1 = tk.Entry(excelInputFrame, font=("Arial", 16))
input1.grid(row=0, column=1, sticky=tk.W)
excelInputFrame.pack(padx=10,pady=10)

btn = tk.Button(root, text="Submit", font=("Arial",10), command= get_excel_no)
btn.pack()




# for i in range(5):
#     input = tk.Entry(inputFrame, font=("Arial", 16))
#     input.grid(row=i, column=0, sticky= tk.W)
#     btn1 = tk.Button(inputFrame, text="Open File", font=("Arial", 12))
#     btn1.grid(row=i, column=1)
#
# input = tk.Entry(inputFrame, font=("Arial", 16))
# input.grid(row=1, column=0, sticky=tk.W)
# btn1 = tk.Button(inputFrame, text="Open File", font=("Arial", 12))
# btn1.grid(row=1, column=1)
#
# input = tk.Entry(inputFrame, font=("Arial", 16))
# input.grid(row=2, column=0, sticky=tk.W)
# btn1 = tk.Button(inputFrame, text="Open File", font=("Arial", 12))
# btn1.grid(row=2, column=1)




# entry = tk.Entry(root)
# entry.pack()
btn = tk.Button(root,text="Submit")
# label3.grid(row=0,column=0)
root.mainloop()















