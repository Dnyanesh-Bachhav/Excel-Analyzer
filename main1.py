import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter import PhotoImage
from tkinter.messagebox import showinfo
from csv import excel
import pandas as pd
import numpy as np
import time
from matplotlib import pyplot as plt
import customtkinter



# Global variables
path_array = []
column_data = []
actual_data = []
final_keys_dict = {}
primary_key = ""
present_column_name = ""
isRollNo = False
isEmail = False
isPhone = False
is_composite_primary_key = False



# info_btn = tk.PhotoImage(file='Information_icon.png')
primary_key_input = ""
def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.

def render_heading():
    text2 = customtkinter.CTkLabel(master=second_frame, text="Enter excel sheets path ", font=("Arial", 16), pady=10)
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
            column_data.append({"Primary Key": [primary_key], "Present_Column_Name": present_column_name, "dataset": dataset})
    print("column data")
    print(column_data)
    get_necessary_data()

def get_necessary_data():
    # Main Array
    # Here we fetch necessary data from dataset like roll no and status column
    global actual_data
    global is_composite_primary_key
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
        get_final_keys_arr()


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



def get_final_keys_arr():
    # Final Keys Array
    # Primary key set
    # key_set = {}
    final_key_arr = []

    for item in actual_data:
        data_arr = []

        for i in range(len(item) - 1):
            item_name = item[i].name
            #         print("HI",item_name)
            #         If primary is only one
            if is_composite_primary_key == False:
                keys = identify_key_and_format(item_name, item[i])
                for key in keys:
                    if key not in final_key_arr:
                        final_key_arr.append(key)
            #             print(keys)
            #         If it is composite primary key
            else:
                if i < len(item) - 2:
                    data_arr.append(item[i])
                else:
                    data_arr.append(item[i])
                    keys = handle_composite_primary_key(data_arr)
                    for key in keys:
                        if key not in final_key_arr:
                            final_key_arr.append(key)


        #         print("Hello World...!!!")

        # Roll Number Formatting
        #         for item1 in item[i]:
        #             item1 = str(item1)
        #             key = np.nan
        #             if len(item1) > 3:
        #                 print(str(item1[ len(item1)-3: ]))
        #                 key = str(item1[ len(item1)-3: ])
        #                 key_arr.append(key)
        #             else:
        #                 print(item1)
        #                 key = item1
        #                 key_arr.append(key)
        #     item1 = item[len(item)-1]
        print("---------------------------")


    print(final_key_arr)
    # Initialize empty dictionary with final keys
    for key in final_key_arr:
        final_keys_dict[key] = 0
    print(final_keys_dict)
    get_present_count()

def get_present_count():
    # Get presenty count of students
    for item in actual_data:
        #     If key is composite
        if is_composite_primary_key == True:
            print("Composite primary key...")
            key_arr = ["" for i in range(len(item[0]))]
            for i in range(len(item) - 1):
                is_Roll = False
                is_Email = False
                is_Phone = False
                rollno = ["roll", "rollno", "roll_no", "roll number", "roll no"]
                email = ["email", "emailid", "email_id"]
                phone = ["phone", "phoneno", "phone no", "phone_no"]
                column_name = item[i].name
                column_name = column_name.lower()
                #    Check if key is rollno
                for li in rollno:
                    if (column_name in li):
                        is_Roll = True
                        break
                if is_Roll == True:
                    print("Roll no. formatting")

                    for j in range(len(item[i])):
                        item1 = str(item[i][j])
                        key = np.nan
                        if len(item1) > 3:
                            key = str(item1[len(item1) - 3:])
                            key = int(key)
                            key_arr[j] = key_arr[j] + str(key)
                        else:
                            #                     print(item1)
                            key = item1
                            key = int(key)
                            key_arr[j] = key_arr[j] + str(key)


                elif is_Email == True:
                    print("Email formatting")
                elif is_Phone == True:
                    print("Phone Formatting")
                else:
                    for k in range(len(item[i])):
                        item1 = str(item[i][k])
                        key = str(item1)
                        key_arr[k] = key_arr[k] + key

            for i in range(len(key_arr)):
                if item[len(item) - 1][i] == 1:
                    final_keys_dict[key_arr[i]] = final_keys_dict.get(key_arr[i]) + 1

            key_arr = None





        else:
            print("Single Primary key...")
            if isRollNo == True:
                for i in range(len(item[0])):
                    item1 = str(item[0][i])
                    key = np.nan
                    if len(item1) > 3:
                        key = str(item1[len(item1) - 3:])
                        key = int(key)
                    else:
                        key = item1
                        key = int(key)
                    print(key)
                    if item[1][i] == 1:
                        dict[key] = dict.get(key) + 1
            if isEmail == True:
                print("It is Email...")
            if isPhone == True:
                print("It is Phone...")
    #                 column_name = item[i]
    #                 print(column_name)

    print(final_keys_dict)
    render_chart_buttons()

def bar_plot():
    present_distinct_values = []
    for key in final_keys_dict:
        if final_keys_dict[key] not in present_distinct_values:
            present_distinct_values.append(final_keys_dict[key])

        # print( key )
    print( present_distinct_values )
    dict = {}
    # Initialize to zero
    for key in present_distinct_values:
        dict[key] = 0

    for key in final_keys_dict:
        dict[ final_keys_dict[key] ] =  dict[ final_keys_dict[key] ] + 1

    print(dict)

    data = []
    labels = []
    for key in dict:
        labels.append(key)
        data.append(dict[key])

    total = sum(data)
    plt.pie(data, labels=labels, autopct=lambda p: '{:.0f}'.format(p * total / 100))
    plt.show()
def render_chart_buttons():
    btn1 = customtkinter.CTkButton(master=second_frame, text="Bar Plot", font=("Arial", 10), command=bar_plot)
    btn1.pack()
    btn2 = customtkinter.CTkButton(master=second_frame, text="Submit", font=("Arial", 10))
    btn2.pack()
    btn3 = customtkinter.CTkButton(master=second_frame, text="Submit", font=("Arial", 10))
    btn3.pack()


def show_info():
    showinfo("Info", "Primary key for all Excelsheets must be same")

def get_primary_key():
    text2 = customtkinter.CTkLabel(master=second_frame, text="Enter a primary key", font=("Arial", 16))
    text2.pack(padx=10, pady=10)
    label1 = customtkinter.CTkLabel(master=primaryKeyInputFrame, text="Enter primary key for Excelsheets: ", width=270, anchor=tk.CENTER, font=("Arial", 14))
    label1.grid(row=0, column=0)
    primary_key_input = customtkinter.CTkEntry(master=primaryKeyInputFrame, font=("Arial", 16))
    primary_key_input.grid(row=0, column=1)

    # btn1 = customtkinter.CTkButton(master=primaryKeyInputFrame, text="I", font=("Arial italic bold", 10), padx=10, fg="white", bg="blue", command=show_info)
    btn1 = customtkinter.CTkButton(master=primaryKeyInputFrame, text="I", font=("Arial italic bold", 16), width=70, command=show_info)
    btn1.grid(row=0, column=2, padx=10, pady=10, sticky=tk.W)

    btn = customtkinter.CTkButton(master=primaryKeyInputFrame, text="Submit", width=250, font=("Arial", 16), command=lambda: handle_primary_key(primary_key_input))
    btn.grid(row=1, column=1, padx=10, pady=10, sticky=tk.W)
    primaryKeyInputFrame.pack(padx=20, pady=20, fill="both", expand=False)


def get_presenty_column_name():
    text2 = customtkinter.CTkLabel(master=second_frame, text="Enter a present column", font=("Arial", 16))
    text2.pack(padx=10, pady=10)
    label1 = customtkinter.CTkLabel(master=presentyColumnInputFrame, text="Enter a presenty column name: ", width=270, anchor=tk.CENTER, font=("Arial", 16))
    label1.grid(row=0, column=0)
    present_column_input = customtkinter.CTkEntry(master=presentyColumnInputFrame, font=("Arial", 16))
    present_column_input.grid(row=0, column=1)

    btn1 = customtkinter.CTkButton(master=presentyColumnInputFrame, text="I", font=("Arial italic bold", 16), width=70, command=show_info)
    btn1.grid(row=0, column=2, padx=10, pady=10)
    btn = customtkinter.CTkButton(master=presentyColumnInputFrame, text="Submit", font=("Arial", 16), width=250, command=lambda: handle_presenty_column(present_column_input))
    btn.grid(row=1, column=1, padx=10, pady=10)
    presentyColumnInputFrame.pack(padx=20, pady=20, fill="both", expand=False)


# input excel files
def render_input_UI(no_of_excels):
    render_heading()
    inputFrame = customtkinter.CTkFrame(master=second_frame)
    inputFrame.columnconfigure(0, weight=1)
    inputFrame.columnconfigure(1, weight=1)
    inputFrame.columnconfigure(2, weight=1)
    for i in range(no_of_excels):
        filename = fd.askopenfilename()
        # Add to file path array
        path_array.append(filename)
        input_field = customtkinter.CTkEntry(master=inputFrame, width=240, font=("Arial", 16))
        input_field.grid(row=i, column=0, pady=10)
        btn1 = customtkinter.CTkButton(master=inputFrame, text="File",  font=("Arial", 16))
        btn1.grid(row=i, column=1)
        # input_field.delete(0,END)
        input_field.insert(0,filename)
        # fill= tk.Y, expand=True, padx=2, pady=2
        inputFrame.pack(padx=20, pady=20, fill="both", expand=False)
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
customtkinter.set_appearance_mode("Light")
customtkinter.set_default_color_theme("dark-blue")
root = customtkinter.CTk()
# root = tk.Tk()

root.title("Excelizer")
root.geometry("550x500")

# Create A Main Frame
main_frame = customtkinter.CTkFrame(master=root)
main_frame.pack(fill=tk.BOTH, expand=1)

# Create A Canvas
my_canvas = tk.Canvas(main_frame)
my_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

# Add A Scrollbar To The Canvas
my_scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=my_canvas.yview, style='Vertical.TScrollbar')
my_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

# Configure The Canvas
my_canvas.configure(yscrollcommand=my_scrollbar.set)
my_canvas.bind('<Configure>', lambda e: my_canvas.configure(scrollregion = my_canvas.bbox("all")))

# Create ANOTHER Frame INSIDE the Canvas
second_frame = tk.Frame(my_canvas)

# Add that New frame To a Window In The Canvas
my_canvas.create_window((0, 0), window=second_frame, anchor="nw")

# window = customtkinter.CTkScrollableFrame(master=root)
# scrollbar = customtkinter.CTkScrollbar(master=root)
# scrollbar.pack( side = customtkinter.RIGHT, fill = customtkinter.Y)
label1 = customtkinter.CTkLabel(master=second_frame, text="Excelizer", font=("Arial", 24))
label1.pack(padx=10, pady=10)


excelInputFrame = customtkinter.CTkFrame(master=second_frame)
excelInputFrame.columnconfigure(0, weight=1)
excelInputFrame.columnconfigure(1, weight=1)
excelInputFrame.columnconfigure(2, weight=1)

primaryKeyInputFrame = customtkinter.CTkFrame(master=second_frame)
primaryKeyInputFrame.columnconfigure(0, weight=1)
primaryKeyInputFrame.columnconfigure(1, weight=1)
primaryKeyInputFrame.columnconfigure(2, weight=1)

presentyColumnInputFrame = customtkinter.CTkFrame(master=second_frame)
presentyColumnInputFrame.columnconfigure(0, weight=1)
presentyColumnInputFrame.columnconfigure(1, weight=1)
presentyColumnInputFrame.columnconfigure(2, weight=1)

outputButtonFrame = customtkinter.CTkFrame(master=second_frame)
outputButtonFrame.columnconfigure(0, weight=1)
outputButtonFrame.columnconfigure(1, weight=1)
outputButtonFrame.columnconfigure(2, weight=1)

text1 = customtkinter.CTkLabel(master=excelInputFrame, text="Enter the number of Excelsheets:", font=("Arial", 16))
text1.grid(row=0, column=0, pady=10)
input1 = customtkinter.CTkEntry(master=excelInputFrame, font=("Arial", 16))
input1.grid(row=0, column=1)

btn = customtkinter.CTkButton(master=excelInputFrame, text="Submit", font=("Arial", 16), command= get_excel_no)
btn.grid(row=1, column=1, padx=10, pady=12)
excelInputFrame.pack(padx=20, pady=20, fill="both", expand=False)



root.mainloop()

