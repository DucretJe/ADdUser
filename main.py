# coding: utf-8

import sender as s
import interface as i
import excel_reader as e
import sys

### 1. Creating interface
window = i.Tk()
interface = i.Interface(window)

### 2. Welcome message
interface.simple_message(window, 'Welcome in this script: \n It will create the Users you\'ll provide in an excel file\n You need to set:\n -Name\n -Surname\n -Title\n -Domain name\n -Service\n for each employee')
interface.wait_variable(interface.waiting)
# Cleaning interface
for widget in window.winfo_children():
    widget.pack_forget()

### 3. Set the excel file
interface.browse_message(window, 'Please load your excel file')
interface.wait_variable(interface.waiting)
excel = interface.file
# Cleaning interface
for widget in window.winfo_children():
    widget.pack_forget()

### 4. Set the AD Server name
interface.input_message(window, 'Please set up the AD server name')
interface.wait_variable(interface.waiting)
AD = interface.var
## 4.a Test the AD Server Upgrade <==================================
# Cleaning interface
for widget in window.winfo_children():
    widget.pack_forget()
### 5. Read the Excel File
users = e.file(excel)
dictionnary = users.dictionnary
## 5.a Testing validity of input_message0

### 6. Sending it to the AD
for u in dictionnary.values():
    print(u)
    dict = {'name' : u['name'], 'surname' : u['surname'], 'title' : u['title'], 'domain' : u['domain'], 'service' : u['service']}
    instance_user = s.Users(**dict)
    send = instance_user.SendAD(AD)

interface.mainloop()
interface.destroy()

sys.exit()
