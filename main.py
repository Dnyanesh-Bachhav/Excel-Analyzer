import tkinter as tk
from tkinter import filedialog as fd
from tkinter import PhotoImage
from tkinter.messagebox import showinfo
from csv import excel

# Global variables
path_array = []
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
   data = primary_key_input.get()
   print(data)


def show_info():
    showinfo("Info","Primary key for all Excelsheets must be same")

def get_primary_key():
    text2 = tk.Label(root, text="Enter a primary key ",font=("Arial",16), padx=10, pady=20)
    text2.pack()
    label1 = tk.Label(primaryKeyInputFrame, text="Enter primary key for Excelsheets: ", font=("Arial", 10))
    label1.grid(row=0, column=0)
    primary_key_input = tk.Entry(primaryKeyInputFrame, font=("Arial", 16))
    primary_key_input.grid(row=0, column=1, sticky=tk.W)

    btn1 = tk.Button( primaryKeyInputFrame, text="I", font=("Arial bold", 10), padx=10, fg="white", bg="blue", command=show_info)
    btn1.grid(row=0, column=2, sticky=tk.W)
    primaryKeyInputFrame.pack(padx=10, pady=10)
    btn = tk.Button(root, text="Submit", font=("Arial", 10), command=lambda: handle_primary_key(primary_key_input))
    btn.pack()

    # img_label = tk.Label(image=info_btn)


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
    inputFrame.pack()
    print(path_array)
    get_primary_key()

def get_excel_no():
    no_of_excels = int(input1.get())
    print("Function called....")
    print(no_of_excels)
    render_input_UI(no_of_excels)
# Press the green button in the gutter to run the script.

print_hi('PyCharm')
root = tk.Tk()
root.title("Excelizer")
root.geometry("500x500")
label1 = tk.Label(root,text="Excelizer",font=("Arial",18))
label1.pack(padx=10,pady=10)
click_btn = tk.PhotoImage(file='info.png')


# Let us create a dummy button and pass the image
btn1 = tk.Button(root, image=click_btn, command=show_info, borderwidth=0)
btn1.pack()

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















