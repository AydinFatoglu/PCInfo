import datetime
import socket
import tkinter as tk
from tkinter import ttk
import wmi
import pyperclip
import uuid
import time
import os
import sys


def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


# Connect to WMI namespace
c = wmi.WMI()

# Get computer system information
computer_system = c.Win32_ComputerSystem()[0]

# Get system enclosure information
system_enclosure = c.Win32_SystemEnclosure()[0]

# Get operating system information
operating_system = c.Win32_OperatingSystem()[0]

# Create root window
root = tk.Tk()
root.iconbitmap(resource_path("./help.ico"))
root.title("BT DESTEK")

root.geometry('280x210')

root.resizable(False, False)

# get the screen width and height
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# calculate the x and y coordinates to center the window
x = (screen_width // 2) - (280 // 2)
y = (screen_height // 2) - (210 // 2)

# set the position of the window to the center of the screen
root.geometry(f'+{x}+{y}')


# Get system information
hostname = socket.gethostname()
ip_address = socket.gethostbyname(hostname)
user_name = computer_system.UserName
domain = computer_system.Domain
model_name = computer_system.SystemFamily
model_no = computer_system.Model
system_serial_number = system_enclosure.SerialNumber.strip()
install_date_string = operating_system.InstallDate
install_date = datetime.datetime.strptime(install_date_string.split(".")[0], "%Y%m%d%H%M%S")
mac_address = ":".join([hex(uuid.getnode())[i:i + 2] for i in range(2, 8)][::-1])  # Use built-in uuid module to get MAC address

# Create labels
hostname_label = ttk.Label(root, text="Hostname:")
ip_address_label = ttk.Label(root, text="IP Address:")
mac_address_label = ttk.Label(root, text="MAC Address:")
user_name_label = ttk.Label(root, text="User Name:")
domain_label = ttk.Label(root, text="Computer Domain:")
model_name_label = ttk.Label(root, text="Model Name:")
model_no_label = ttk.Label(root, text="Model No:")
system_serial_number_label = ttk.Label(root, text="System Serial Number:")
install_date_label = ttk.Label(root, text="Windows Installation Date:")


# Create text variables for labels
hostname_var = tk.StringVar(value=hostname)
ip_address_var = tk.StringVar(value=ip_address)
mac_address_var = tk.StringVar(value=mac_address)
user_name_var = tk.StringVar(value=user_name)
domain_var = tk.StringVar(value=domain)
model_name_var = tk.StringVar(value=model_name)
model_no_var = tk.StringVar(value=model_no)
system_serial_number_var = tk.StringVar(value=system_serial_number)
install_date_var = tk.StringVar(value=install_date.strftime("%Y-%m-%d %H:%M:%S"))


# Create text boxes for labels
hostname_textbox = ttk.Entry(root, textvariable=hostname_var, state="readonly")
ip_address_textbox = ttk.Entry(root, textvariable=ip_address_var, state="readonly")
mac_address_textbox = ttk.Entry(root, textvariable=mac_address_var, state="readonly")
user_name_textbox = ttk.Entry(root, textvariable=user_name_var, state="readonly")
domain_textbox = ttk.Entry(root, textvariable=domain_var, state="readonly")
model_name_textbox = ttk.Entry(root, textvariable=model_name_var, state="readonly")
model_no_textbox = ttk.Entry(root, textvariable=model_no_var, state="readonly")
system_serial_number_textbox = ttk.Entry(root, textvariable=system_serial_number_var, state="readonly")


# Lay out labels and text boxes using grid layout
hostname_label.grid(row=0, column=0, sticky="w")
hostname_textbox.grid(row=0, column=1)
ip_address_label.grid(row=1, column=0, sticky="w")
ip_address_textbox.grid(row=1, column=1)
mac_address_label.grid(row=2, column=0, sticky="w")
mac_address_textbox.grid(row=2, column=1)
user_name_label.grid(row=3, column=0, sticky="w")
user_name_textbox.grid(row=3, column=1)
domain_label.grid(row=4, column=0, sticky="w")
domain_textbox.grid(row=4, column=1)
model_name_label.grid(row=5, column=0, sticky="w")
model_name_textbox.grid(row=5, column=1)
model_no_label.grid(row=6, column=0, sticky="w")
model_no_textbox.grid(row=6, column=1)
system_serial_number_label.grid(row=7, column=0, sticky="w")
system_serial_number_textbox.grid(row=7, column=1)


# Create Copy to Clipboard button
def copy_to_clipboard():
    data = f"Hostname: {hostname}\n" \
           f"IP Address: {ip_address}\n" \
           f"MAC Address: {mac_address}\n" \
           f"User Name: {user_name}\n" \
           f"Computer Domain: {domain}\n" \
           f"Model Name: {model_name}\n" \
           f"Model No: {model_no}\n" \
           f"System Serial Number: {system_serial_number}"

    pyperclip.copy(data)
    copy_button.configure(text="Copied!")
    root.after(400, lambda: copy_button.configure(text="Copy to Clipboard"))


copy_button = ttk.Button(root, text="Copy to Clipboard", command=copy_to_clipboard)
copy_button.grid(row=len(root.grid_slaves()), column=0, columnspan=2, pady=(10, 0))


# Add Copy and Select All options to the context menu of text boxes
def copy_context_menu(event):
    event.widget.event_generate("<<Copy>>")
    return "break"


def select_all_context_menu(event):
    event.widget.select_range(0, "end")
    return "break"


text_boxes = [hostname_textbox, ip_address_textbox, mac_address_textbox, user_name_textbox, domain_textbox,
              model_name_textbox, model_no_textbox, system_serial_number_textbox]

for textbox in text_boxes:
    textbox.bind("<Button-3>", lambda event: textbox.focus())
    context_menu = tk.Menu(textbox, tearoff=0)
    context_menu.add_command(label="Copy", command=copy_to_clipboard)
    context_menu.add_command(label="Select All", command=lambda textbox=textbox: textbox.select_range(0, "end"))
    textbox.bind("<Control-c>", lambda event, textbox=textbox: textbox.event_generate("<<Copy>>"))
    textbox.bind("<Control-a>", lambda event, textbox=textbox: textbox.select_range(0, "end"))
    textbox.bind("<Button-3>", lambda event, context_menu=context_menu: context_menu.post(event.x_root, event.y_root))
    textbox.bind("<Control-v>", lambda event: textbox.delete("sel.first", "sel.last"))
    textbox.bind("<Control-v>", lambda event: textbox.insert("insert", pyperclip.paste()), "+")

# Run Tkinter event loop
root.mainloop()

