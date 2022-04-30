#this program will both generate a random password and allow you to either enter that random password into an excel sheet or one you create.

import tkinter as tk
import xlwt
import random
import string
#our applications we import

root = tk.Tk()
root.geometry("700x700")
root.title("Password Generator and Saver")
root.resizable(0,0)
root.configure(bg='lightgrey')
#our gui

#Excel sheet
#Labels for entry boxes
cpass_label = tk.Label(root, text = "Password", bg='lightgrey', font=('Arial', 10, 'bold') )
csite_label = tk.Label(root, text = "Site_File", bg='lightgrey', font=('Arial', 10, 'bold') )
curl_label = tk.Label(root, text = "URL", bg='lightgrey', font=('Arial', 10, 'bold') )
cdate_label = tk.Label(root, text = "Date Changed", bg='lightgrey', font=('Arial', 10, 'bold') )



#Declaring string or int variables for storing data in entry box
#Ill use c_... to match it but it can be anything and does not have to match
cpass_var = tk.StringVar()
csite_var = tk.StringVar()
curl_var = tk.StringVar()
cdate_var = tk.StringVar()


#entry boxes to allow people to type data

cpass_entry = tk.Entry(root, textvariable = cpass_var, bg='orange', font=('Arial', 10, 'bold') )
csite_entry = tk.Entry(root, textvariable = csite_var, bg='orange', font=('Arial', 10, 'bold') )
curl_entry = tk.Entry(root, textvariable = curl_var, bg='orange', font=('Arial', 10, 'bold') )
cdate_entry = tk.Entry(root, textvariable = cdate_var, bg='orange', font=('Arial', 10, 'bold') )


#placing the labels and entry boxes as required
cpass_label.grid(row=0, column=0)
cpass_entry.grid(row=0, column=1)

csite_label.grid(row=1, column=0)
csite_entry.grid(row=1, column=1)

curl_label.grid(row=2, column=0)
curl_entry.grid(row=2, column=1)

cdate_label.grid(row=3, column=0)
cdate_entry.grid(row=3, column=1)



#get or receive the data from the entrybox

def submit():
    cpass = cpass_var.get()
    csite = csite_var.get()
    curl = curl_var.get()
    cdate = cdate_var.get()


    workbook = xlwt.Workbook()
#naming columns of our work sheet
    sheet = workbook.add_sheet("Passwords")#name of the worksheet that excel
    sheet.write(0, 0, "Passwords")#table name at top
    sheet.write(1, 0, "Password")
    sheet.write(2, 0, "Site_File")
    sheet.write(3, 0, "URL")
    sheet.write(4, 0, "Date Changed")

    #adding data to the heading we created above

    sheet.write(1, 1, cpass)
    sheet.write(2, 1, csite)
    sheet.write(3, 1, curl)
    sheet.write(4, 1, cdate)


    #saving data to the excel
    workbook.save("passwords.xls")

    cpass_var.set("")
    csite_var.set("")
    curl_var.set("")
    cdate_var.set("")


#creating button
sub_button = tk.Button(root, text= 'submit the data', command = submit)
sub_button.grid(row=9, column=1)

#Random generator
def genie():
    password = []  #these lines of code is defining the data in the generator
    for i in range(3):
        lower=random.choice(string.ascii_lowercase) #defining the data set using string and making it randomly selected
        upper=random.choice(string.ascii_uppercase)#defining the data set using string and making it randomly selected
        num=random.choice(string.digits) #defining the data set using string digits
        password.append(lower) #storing the data lower case
        password.append(upper) #storing the data uppercase
        password.append(num) #storing the data numbers
        passw=" ".join(str(x)for x in password) # passing combined data with length and joining them
        label1.config(text=passw, bg='lightgrey') #configuration


label1 = tk.Label(root, font = ('arial', 40, 'bold'))
label1.place(x=200, y=500)#this line is the font and its size for the label
button1 = tk.Button(root,text="Generate", fg='white', bg='grey', font = ('arial', 40, 'bold'), command = genie) #this line is for labeling and font of the button on the gui
button1.place(x=200,y=300) #this line is for placement of the button

root.mainloop()
