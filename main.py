import datetime
import customtkinter as tkinter
import aspose.words as aw
from customtkinter import CTkLabel, CTkEntry, CTkButton, CTkFrame
from tkinter import ttk
from tkinter import messagebox
from docxtpl import DocxTemplate
from datetime import date


# Functions


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

    doc_pdf = aw.Document(doc_name)
    doc.save("test.pdf")


# Main Window
mainWindow = tkinter.CTk()
mainWindow.title("Invoice Generator")

tkinter.set_appearance_mode("light")

frame = CTkFrame(mainWindow)
frame.pack(padx=20, pady=20)

# Information to be entered by the user regarding recipient
instructionsR_label = CTkLabel(
    frame, text="Please enter all information regarding the recipient of the invoice :")
instructionsR_label.grid(row=0, column=0, pady=5)

name_label = CTkLabel(frame, text="Recipient Name")
name_label.grid(row=1, column=0, pady=5)
name_entry = CTkEntry(frame)
name_entry.grid(row=1, column=1, pady=5)

email_label = CTkLabel(frame, text="Recipient Email Address")
email_label.grid(row=2, column=0, pady=5)
email_entry = CTkEntry(frame)
email_entry.grid(row=2, column=1, pady=5)

PO_label = CTkLabel(frame, text="P.O. Box")
PO_label.grid(row=3, column=0, pady=5)
PO_entry = CTkEntry(frame)
PO_entry.grid(row=3, column=1, pady=5)

area_label = CTkLabel(frame, text="Recipient Area")
area_label.grid(row=4, column=0, pady=5)
area_entry = CTkEntry(frame)
area_entry.grid(row=4, column=1, pady=5)

zip_label = CTkLabel(frame, text="Zip Code")
zip_label.grid(row=5, column=0, pady=5)
zip_entry = CTkEntry(frame)
zip_entry.grid(row=5, column=1, pady=5)

# Information entered by user regarding invoice item
instructionsI_label = CTkLabel(
    frame, text="Please enter all information regarding items to be added to the invoice :")
instructionsI_label.grid(row=6, column=0, pady=5)

qty_label = CTkLabel(frame, text="Quantity")
qty_label.grid(row=7, column=0, pady=5)
qty_spin = ttk.Spinbox(frame, from_=1, to=1000)
qty_spin.grid(row=7, column=1, pady=5)
qty_spin.insert(0, "1")

description_label = CTkLabel(frame, text="Description")
description_label.grid(row=8, column=0, pady=5)
description_entry = CTkEntry(frame)
description_entry.grid(row=8, column=1, pady=5)

rate_label = CTkLabel(frame, text="Rate")
rate_label.grid(row=10, column=0, pady=5)
rate_spin = ttk.Spinbox(frame, from_=1, to=1000000, increment=0.5)
rate_spin.grid(row=10, column=1, pady=5)
rate_spin.insert(0, "0.0")

add_item_btn = CTkButton(frame, text="Add Item", command=add_item)
add_item_btn.grid(row=11, column=0, pady=5)

# Tree View
columns = ("Quantity", "Description", "Rate", "Total")
tree = ttk.Treeview(frame, columns=columns, show="headings")

tree.heading("Quantity", text="Quantity")
tree.heading("Description", text="Description")
tree.heading("Rate", text="Rate")
tree.heading("Total", text="Total")

tree.grid(row=0, column=2, rowspan=9, padx=20, pady=10)

# Final Buttons
save_invoicedocx_btn = CTkButton(
    frame, text="Generate Word Invoice", command=gen_invoice_docx)
save_invoicedocx_btn.grid(row=10, column=0, columnspan=3, padx=20, pady=5)

save_invoicepdf_btn = CTkButton(
    frame, text="Generate PDF Invoice", command=gen_invoice_pdf)
save_invoicepdf_btn.grid(row=11, column=0, columnspan=3, padx=20, pady=5)

email_btn = CTkButton(
    frame, text="Send invoice to recipient through email")
email_btn.grid(row=12, column=0, columnspan=3, padx=20, pady=5)

new_invoice_btn = CTkButton(
    frame, text="New Invoice", command=new_invoice)
new_invoice_btn.grid(row=13, column=0, columnspan=3, padx=20, pady=10)

mainWindow.mainloop()
