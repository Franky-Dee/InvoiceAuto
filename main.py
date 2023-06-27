import datetime
import customtkinter as tkinter
import docx2pdf
from customtkinter import CTkLabel, CTkEntry, CTkButton, CTkFrame
from tkinter import ttk
from tkinter import messagebox
from docxtpl import DocxTemplate
from datetime import date
from docx2pdf import convert

# Functions
def new_recipient():
    new_win = tkinter.CTkToplevel()
    new_win.title("Add a New Recipient")
    new_win.attributes('-topmost', True)

    instructionsR_label = CTkLabel(
        new_win, text="Please enter all information regarding the recipient of the invoice :")
    instructionsR_label.grid(row=1, column=0, padx=10, pady=5)

    name_label = CTkLabel(new_win, text="Recipient Name")
    name_label.grid(row=2, column=0, pady=5)
    name_entry = CTkEntry(new_win)
    name_entry.grid(row=2, column=1, padx=10, pady=5)

    email_label = CTkLabel(new_win, text="Recipient Email Address")
    email_label.grid(row=3, column=0, pady=5)
    email_entry = CTkEntry(new_win)
    email_entry.grid(row=3, column=1, padx=10, pady=5)

    PO_label = CTkLabel(new_win, text="P.O. Box")
    PO_label.grid(row=4, column=0, pady=5)
    PO_entry = CTkEntry(new_win)
    PO_entry.grid(row=4, column=1, padx=10, pady=5)

    area_label = CTkLabel(new_win, text="Recipient Area")
    area_label.grid(row=5, column=0, pady=5)
    area_entry = CTkEntry(new_win)
    area_entry.grid(row=5, column=1, padx=10, pady=5)

    zip_label = CTkLabel(new_win, text="Zip Code")
    zip_label.grid(row=6, column=0, pady=5)
    zip_entry = CTkEntry(new_win)
    zip_entry.grid(row=6, column=1, padx=10, pady=5)

    add_item_btn = CTkButton(new_win, text="Add Recipient", command=create_profile)
    add_item_btn.grid(row=7, column=0, padx=10, pady=10,
                      columnspan=2, sticky="news")

def create_profile():
    print("created")

def get_last_invoice_number():
    try:
        with open("last_invoice_number.txt", "r") as file:
            return int(file.read())
    except FileNotFoundError:
        return 0


def update_invoice_number(new_invoice_number):
    with open("last_invoice_number.txt", "w") as file:
        file.write(str(new_invoice_number))


def update_invoice_number(new_invoice_number):
    with open("last_invoice_number.txt", "w") as file:
        file.write(str(new_invoice_number))


last_invoice_number = get_last_invoice_number()
new_invoice_number = last_invoice_number + 1
update_invoice_number(new_invoice_number)


def clear_items():
    qty_spin.delete(0, tkinter.END)
    qty_spin.insert(0, "1")
    description_entry.delete(0, tkinter.END)
    rate_spin.delete(0, tkinter.END)
    rate_spin.insert(0, "0.0")


invoice_list = []


def add_item():
    qty = int(qty_spin.get())
    desc = description_entry.get()
    rate = float(rate_spin.get())
    amount = qty*rate
    invoice_item = [qty, desc, rate, amount]

    tree.insert("", 0, values=invoice_item)
    clear_items()

    invoice_list.append(invoice_item)


def new_invoice():
    name_entry.delete(0, tkinter.END)
    PO_entry.delete(0, tkinter.END)
    area_entry.delete(0, tkinter.END)
    zip_entry.delete(0, tkinter.END)
    clear_items()
    tree.delete(*tree.get_children())

    invoice_list.clear()


def gen_invoice_docx():
    doc = DocxTemplate("invoice_template.docx")
    name = name_entry.get()
    PObox = PO_entry.get()
    area = area_entry.get()
    zip = zip_entry.get()
    total = sum(item[3] for item in invoice_list)
    current_date = date.today()
    inNo = new_invoice_number

    doc.render({
        "date": current_date,
        "name": name,
        "PObox": PObox,
        "area": area,
        "zip": zip,
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
    name = name_entry.get()
    PObox = PO_entry.get()
    area = area_entry.get()
    zip = zip_entry.get()
    total = sum(item[3] for item in invoice_list)
    current_date = date.today()

    doc.render({
        "date": current_date,
        "name": name,
        "PObox": PObox,
        "area": area,
        "zip": zip,
        "invoice_list": invoice_list,
        "total": total
    })

    doc_name = "new_invoice " + name + " " + \
        datetime.datetime.now().strftime("%Y-%m-%d-%H%M%S") + ".docx"
    doc.save(doc_name)

    convert(doc_name)

    messagebox.showinfo("Invoice Generation", "Invoice Completed")


# Main Window
mainWindow = tkinter.CTk()
mainWindow.title("Invoice Generator")

tkinter.set_appearance_mode("light")

frame = CTkFrame(mainWindow)
frame.pack(padx=20, pady=20)

title_label = CTkLabel(frame, text="Invoice Generator")
title_label.grid(row=0, column=0, pady=10, columnspan=3)

# Information to be entered by the user regarding recipient
selectR_label = tkinter.CTkLabel(
    frame, text="Select the invoice recipient or add a new recipient")
selectR_label.grid(row=1, column=0, padx=10, pady=10)

new_recipient_btn = CTkButton(
    frame, text="Add a recipient", command=new_recipient)
new_recipient_btn.grid(row=2, column=1,
                       padx=20, pady=10)

drop_menu = tkinter.CTkOptionMenu(frame, values=["1", "2", "3"])
drop_menu.grid(row=1, column=1)

# Information entered by user regarding invoice item
instructionsI_label = CTkLabel(
    frame, text="Please enter all information regarding items to be added to the invoice :")
instructionsI_label.grid(row=3, column=0, padx=10, pady=5)

qty_label = CTkLabel(frame, text="Quantity")
qty_label.grid(row=4, column=0, pady=5)
qty_spin = ttk.Spinbox(frame, from_=1, to=1000)
qty_spin.grid(row=4, column=1, pady=5)
qty_spin.insert(0, "1")

description_label = CTkLabel(frame, text="Description")
description_label.grid(row=5, column=0, pady=5)
description_entry = CTkEntry(frame)
description_entry.grid(row=5, column=1, pady=5)

rate_label = CTkLabel(frame, text="Rate")
rate_label.grid(row=6, column=0, pady=5)
rate_spin = ttk.Spinbox(frame, from_=1, to=1000000, increment=0.5)
rate_spin.grid(row=6, column=1, pady=5)
rate_spin.insert(0, "0.0")

add_item_btn = CTkButton(frame, text="Add Item", command=add_item)
add_item_btn.grid(row=6, column=0, pady=10)

# Tree View
columns = ("Quantity", "Description", "Rate", "Total")
tree = ttk.Treeview(frame, columns=columns, show="headings")

tree.heading("Quantity", text="Quantity")
tree.heading("Description", text="Description")
tree.heading("Rate", text="Rate")
tree.heading("Total", text="Total")

tree.grid(row=0, column=2, rowspan=8, padx=20, pady=10)

# Final Buttons
save_invoicedocx_btn = CTkButton(
    frame, text="Generate Word Invoice", command=gen_invoice_docx)
save_invoicedocx_btn.grid(row=7, column=0, columnspan=3,
                          sticky="news", padx=20, pady=5)

save_invoicepdf_btn = CTkButton(
    frame, text="Generate PDF Invoice", command=gen_invoice_pdf)
save_invoicepdf_btn.grid(row=8, column=0, columnspan=3,
                         sticky="news", padx=20, pady=5)

email_btn = CTkButton(
    frame, text="Send invoice to recipient through email")
email_btn.grid(row=9, column=0, columnspan=3, sticky="news", padx=20, pady=5)

new_invoice_btn = CTkButton(
    frame, text="New Invoice", command=new_invoice)
new_invoice_btn.grid(row=10, column=0, columnspan=3,
                     sticky="news", padx=20, pady=10)

mainWindow.mainloop()
