import datetime
import customtkinter as tkinter
import docx2pdf
import sqlite3
from customtkinter import CTkLabel, CTkEntry, CTkButton, CTkFrame
from tkinter import ttk
from tkinter import messagebox
from docxtpl import DocxTemplate
from datetime import date
from docx2pdf import convert

# Variables -------------------------------------------------------------------------------------------------------------------------------------
global name
global email
global PO
global area
global zipCode
invoice_list = []

# Create Tables ----------------------------------------------------------------------------------------------------------------------------------
connProfile = sqlite3.connect("profile_data.db")
profile_table_create_query = '''CREATE TABLE IF NOT EXISTS profile_data (name TEXT, email TEXT, PO_box TEXT, area TEXT, zip_code TEXT)'''
connProfile.execute(profile_table_create_query)
connProfile.close()

connDesc = sqlite3.connect("description_data.db")
desc_table_create_query = '''CREATE TABLE IF NOT EXISTS description_data (description TEXT)'''
connDesc.execute(desc_table_create_query)
connDesc.close()

# New Recipient / New Description ---------------------------------------------------------------------------------------------------------------


def new_desc():
    new_win = tkinter.CTkToplevel()
    new_win.title("Add a New Description")
    new_win.attributes('-topmost', True)

    instructionsR_label = CTkLabel(
        new_win, text="Please enter the description you would like to add :", font=("Calibri", 16))
    instructionsR_label.grid(row=1, column=0, padx=10, pady=5)

    desc_label = CTkLabel(new_win, text="Description", font=("Calibri", 16))
    desc_label.grid(row=2, column=0, pady=5)
    desc_entry = CTkEntry(new_win)
    desc_entry.grid(row=2, column=1, padx=10, pady=5)

    def create_desc():
        desc = desc_entry.get()

        if desc:
            # SQL INSERT
            connDesc = sqlite3.connect("description_data.db")
            Pdata_insert_query = ''' INSERT INTO description_data (name, email, PO_box, area, zip_code) VALUES (?, ?, ?, ?, ?) '''
            Pdata_insert_tuple = (desc)
            cursor = connDesc.cursor()
            cursor.execute(Pdata_insert_query, Pdata_insert_tuple)
            connDesc.commit()
            connDesc.close()
            messagebox.showinfo("Creation Successful",
                                "A new description has been created!")
            desc_entry.delete(0, tkinter.END)

            desc_drop_menu.configure(values=description_option_box_values())
        else:
            messagebox.showwarning(
                "Empty Fields", "Please complete all the fields")

    add_desc_btn = CTkButton(
        new_win, text="Add Recipient", command=create_desc)
    add_item_btn.grid(row=7, column=0, padx=10, pady=10,
                      columnspan=2, sticky="news")

    close_btn = CTkButton(
        new_win, text="Close", command=new_win.destroy)
    close_btn.grid(row=8, column=0, padx=10, pady=10,
                   columnspan=2, sticky="news")


def new_recipient():
    new_win = tkinter.CTkToplevel()
    new_win.title("Add a New Recipient")
    new_win.attributes('-topmost', True)

    instructionsR_label = CTkLabel(
        new_win, text="Please enter all information regarding the recipient of the invoice :", font=("Calibri", 16))
    instructionsR_label.grid(row=1, column=0, padx=10, pady=5)

    name_label = CTkLabel(new_win, text="Recipient Name", font=("Calibri", 16))
    name_label.grid(row=2, column=0, pady=5)
    name_entry = CTkEntry(new_win)
    name_entry.grid(row=2, column=1, padx=10, pady=5)

    email_label = CTkLabel(
        new_win, text="Recipient Email Address", font=("Calibri", 16))
    email_label.grid(row=3, column=0, pady=5)
    email_entry = CTkEntry(new_win)
    email_entry.grid(row=3, column=1, padx=10, pady=5)

    PO_label = CTkLabel(new_win, text="P.O. Box", font=("Calibri", 16))
    PO_label.grid(row=4, column=0, pady=5)
    PO_entry = CTkEntry(new_win)
    PO_entry.grid(row=4, column=1, padx=10, pady=5)

    area_label = CTkLabel(new_win, text="Recipient Area", font=("Calibri", 16))
    area_label.grid(row=5, column=0, pady=5)
    area_entry = CTkEntry(new_win)
    area_entry.grid(row=5, column=1, padx=10, pady=5)

    zip_label = CTkLabel(new_win, text="Zip Code", font=("Calibri", 16))
    zip_label.grid(row=6, column=0, pady=5)
    zip_entry = CTkEntry(new_win)
    zip_entry.grid(row=6, column=1, padx=10, pady=5)

    def create_profile():
        name = name_entry.get()
        email = email_entry.get()
        PO = PO_entry.get()
        area = area_entry.get()
        zipCode = zip_entry.get()

        if name and email and PO and area and zipCode:
            # SQL INSERT
            connProfile = sqlite3.connect("description_data.db")
            Pdata_insert_query = ''' INSERT INTO profile_data (name, email, PO_box, area, zip_code) VALUES (?, ?, ?, ?, ?) '''
            Pdata_insert_tuple = (name, email, PO, area, zipCode)
            cursor = connProfile.cursor()
            cursor.execute(Pdata_insert_query, Pdata_insert_tuple)
            connProfile.commit()
            connProfile.close()
            messagebox.showinfo("Creation Successful",
                                "A new recipient has been created and added")
            name_entry.delete(0, tkinter.END)
            email_entry.delete(0, tkinter.END)
            area_entry.delete(0, tkinter.END)
            PO_entry.delete(0, tkinter.END)
            zip_entry.delete(0, tkinter.END)

            drop_menu.configure(values=recipient_option_box_values())
        else:
            messagebox.showwarning(
                "Empty Fields", "Please complete all the fields")

    add_item_btn = CTkButton(
        new_win, text="Add Recipient", command=create_profile)
    add_item_btn.grid(row=7, column=0, padx=10, pady=10,
                      columnspan=2, sticky="news")

    close_btn = CTkButton(
        new_win, text="Close", command=new_win.destroy)
    close_btn.grid(row=8, column=0, padx=10, pady=10,
                   columnspan=2, sticky="news")

# Functions -------------------------------------------------------------------------------------------------------------------------------------


def description_option_box_values():
    connDesc = sqlite3.connect("description_data.db")
    cursor = connDesc.cursor()
    query = cursor.execute('SELECT name FROM description_data')

    data = []
    for row in cursor.fetchall():
        data.append(row[0])
    return data


def recipient_option_box_values():
    connProfile = sqlite3.connect("profile_data.db")
    cursor = connProfile.cursor()
    query = cursor.execute('SELECT name FROM profile_data')

    data = []
    for row in cursor.fetchall():
        data.append(row[0])
    return data


def get_last_invoice_number():
    try:
        with open("last_invoice_number.txt", "r") as file:
            return int(file.read())
    except FileNotFoundError:
        return 0


def update_invoice_number(new_invoice_number):
    with open("last_invoice_number.txt", "w") as file:
        file.write(str(new_invoice_number))


last_invoice_number = get_last_invoice_number()
new_invoice_number = last_invoice_number + 1
update_invoice_number(new_invoice_number)


def clear_items():
    qty_spin.delete(0, tkinter.END)
    qty_spin.insert(0, "1")
    rate_spin.delete(0, tkinter.END)
    rate_spin.insert(0, "0.0")


def add_item():
    qty = int(qty_spin.get())
    rate = float(rate_spin.get())
    amount = qty*rate
    invoice_item = [qty, rate, amount]

    tree.insert("", 0, values=invoice_item)
    clear_items()

    invoice_list.append(invoice_item)


def new_invoice():
    clear_items()
    tree.delete(*tree.get_children())

    invoice_list.clear()


def gen_invoice_docx():
    doc = DocxTemplate("invoice_template.docx")
    # Variables = SELECT from db WHERE name = combo
    total = sum(item[3] for item in invoice_list)
    current_date = date.today()
    inNo = new_invoice_number

    doc.render({
        "date": current_date,
        "name": name,
        "PObox": PO,
        "area": area,
        "zip": zipCode,
        "invoice_list": invoice_list,
        "total": total,
        "inNo": inNo
    })

    doc_name = "new_invoice " + name + " " + \
        datetime.datetime.now().strftime("%Y-%m-%d-%H%M%S") + ".docx"
    doc.save(doc_name)

    messagebox.showinfo("Invoice Generation", "Invoice Completed")

    new_invoice()


def gen_invoice_pdf():
    doc = DocxTemplate("invoice_template.docx")
    # Variables = SELECT from db WHERE name = combo
    total = sum(item[3] for item in invoice_list)
    current_date = date.today()

    doc.render({
        "date": current_date,
        "name": name,
        "PObox": PO,
        "area": area,
        "zip": zipCode,
        "invoice_list": invoice_list,
        "total": total
    })

    doc_name = "new_invoice " + name + " " + \
        datetime.datetime.now().strftime("%Y-%m-%d-%H%M%S") + ".docx"
    doc.save(doc_name)

    convert(doc_name)

    messagebox.showinfo("Invoice Generation", "Invoice Completed")


# Main Window -----------------------------------------------------------------------------------------------------------------------------------
mainWindow = tkinter.CTk()
mainWindow.title("InvoiceAuto")

tkinter.set_default_color_theme("green")

title_label = CTkLabel(
    mainWindow, text="InvoiceAuto : Automatic Invoice Generator", font=("Calibri", 22))
title_label.grid(row=0, column=0, pady=10, columnspan=3)

recipient_frame = CTkFrame(mainWindow, border_width=3)
recipient_frame.grid(row=1, column=0, pady=10, padx=10)

invoice_items_frame = CTkFrame(mainWindow, border_width=3)
invoice_items_frame.grid(row=1, column=1, pady=10, padx=10)

export_frame = CTkFrame(mainWindow, border_width=3)
export_frame.grid(row=1, column=2, pady=10, padx=10)

# Recipeint Frame -------------------------------------------------------------------------------------------------------------------------------
title_label = CTkLabel(
    recipient_frame, text="Invoice Recipient", font=("Calibri", 18))
title_label.grid(row=0, column=0, pady=10, columnspan=2)

selectR_label = tkinter.CTkLabel(
    recipient_frame, font=("Calibri", 16))
selectR_label.grid(row=1, column=0, padx=10, pady=10)

drop_menu = tkinter.CTkOptionMenu(
    recipient_frame, values=recipient_option_box_values())
drop_menu.grid(row=1, column=1, padx=10, pady=10)

if drop_menu._values:
    drop_menu.configure(values=recipient_option_box_values())
else:
    drop_menu.set("Create a recipient")

new_recipient_btn = CTkButton(
    recipient_frame, text="Add a recipient", command=new_recipient)
new_recipient_btn.grid(row=2, column=0, columnspan=2, sticky="news",
                       padx=20, pady=10)

# Invoice Items Frame ---------------------------------------------------------------------------------------------------------------------------
instructionsI_label = CTkLabel(
    invoice_items_frame, text="Invoice Items", font=("Calibri", 18))
instructionsI_label.grid(row=0, column=0, pady=5, columnspan=2)

qty_label = CTkLabel(invoice_items_frame,
                     text="Select Quantity", font=("Calibri", 16))
qty_label.grid(row=1, column=0, padx=10, pady=10)
qty_spin = ttk.Spinbox(invoice_items_frame, from_=1, to=1000)
qty_spin.grid(row=1, column=1, padx=10, pady=10)
qty_spin.insert(0, "1")

rate_label = CTkLabel(invoice_items_frame,
                      text="Select Rate", font=("Calibri", 16))
rate_label.grid(row=2, column=0, padx=10, pady=10)
rate_spin = ttk.Spinbox(invoice_items_frame, from_=1,
                        to=1000000, increment=100)
rate_spin.grid(row=2, column=1, padx=10, pady=10)
rate_spin.insert(0, "0.0")

description_label = CTkLabel(
    invoice_items_frame, text="Select Description", font=("Calibri", 16))
description_label.grid(row=3, column=0, padx=10, pady=10)
desc_drop_menu = tkinter.CTkOptionMenu(invoice_items_frame, values=["", ""])
desc_drop_menu.grid(row=3, column=1, padx=10, pady=10)

if desc_drop_menu._values:
    desc_drop_menu.configure(values=description_option_box_values())
else:
    desc_drop_menu.set("Create a description")

add_item_btn = CTkButton(
    invoice_items_frame, text="Add Item", command=add_item)
add_item_btn.grid(row=4, column=0, columnspan=2,
                  sticky="news", padx=10, pady=10)

edit_desc_btn = CTkButton(
    invoice_items_frame, text="Edit Description List", command=new_desc)
edit_desc_btn.grid(row=5, column=0, columnspan=2,
                   sticky="news", padx=10, pady=10)

# Export Frame ----------------------------------------------------------------------------------------------------------------------------------
export_frame_label = CTkLabel(
    export_frame, text="Export Invoice", font=("Calibri", 18))
export_frame_label.grid(row=0, column=0, pady=5, columnspan=2)

save_invoicedocx_btn = CTkButton(
    export_frame, text="Generate Word Invoice", font=("Calibri", 16), command=gen_invoice_docx)
save_invoicedocx_btn.grid(row=1, column=0,
                          sticky="news", padx=10, pady=10)

save_invoicepdf_btn = CTkButton(
    export_frame, text="Generate PDF Invoice", font=("Calibri", 16), command=gen_invoice_pdf)
save_invoicepdf_btn.grid(row=2, column=0,
                         sticky="news", padx=10, pady=10)

email_btn = CTkButton(
    export_frame, text="Send invoice to recipient through email", font=("Calibri", 16))
email_btn.grid(row=3, column=0, sticky="news", padx=10, pady=10)

# Tree View -------------------------------------------------------------------------------------------------------------------------------------
columns = ("Quantity", "Description", "Rate", "Total")
tree = ttk.Treeview(mainWindow, columns=columns, show="headings")

tree.heading("Quantity", text="Quantity")
tree.heading("Description", text="Description")
tree.heading("Rate", text="Rate")
tree.heading("Total", text="Total")

tree.grid(row=2, column=0, rowspan=8, columnspan=3, padx=20, pady=10)

new_invoice_btn = CTkButton(
    mainWindow, text="New Invoice", font=("Calibri", 16), command=new_invoice)
new_invoice_btn.grid(row=10, column=0, columnspan=3,
                     sticky="news", padx=20, pady=10)

mainWindow.mainloop()
