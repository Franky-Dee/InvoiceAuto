import tkinter
from tkinter import ttk

mainWindow = tkinter.Tk()
mainWindow.title("Invoice Generator")

frame = tkinter.Frame(mainWindow)
frame.pack()

# Information to be entered by the user regarding recipient
instructionsR_label = tkinter.Label(frame, text="Please enter all information regarding the recipient of the invoice :")
instructionsR_label.grid(row=0, column=0)

name_label = tkinter.Label(frame, text="Recipient Name")
name_label.grid(row=1, column=0)
name_entry = tkinter.Entry(frame)
name_entry.grid(row=1, column=1)

PO_label = tkinter.Label(frame, text="P.O. Box")
PO_label.grid(row=2, column=0)
PO_entry = tkinter.Entry(frame)
PO_entry.grid(row=2, column=1)

area_label = tkinter.Label(frame, text="Recipient Area")
area_label.grid(row=3, column=0)
area_entry = tkinter.Entry(frame)
area_entry.grid(row=3, column=1)

zip_label = tkinter.Label(frame, text="Zip Code")
zip_label.grid(row=4, column=0)
zip_entry = tkinter.Entry(frame)
zip_entry.grid(row=4, column=1)

#Information entered by user regarding invoice item
instructionsI_label = tkinter.Label(frame, text="Please enter all information regarding items to be added to the invoice :")
instructionsI_label.grid(row=5, column=0)

qty_label = tkinter.Label(frame, text="Quantity")
qty_label.grid(row=6, column=0)
qty_spin = tkinter.Spinbox(frame, from_=1, to=1000)
qty_spin.grid(row=6, column=1)

description_label = tkinter.Label(frame, text="Description")
description_label.grid(row=7, column=0)
description_entry = tkinter.Entry(frame)
description_entry.grid(row=7, column=1)

rate_label = tkinter.Label(frame, text="Rate")
rate_label.grid(row=8, column=0)
rate_spin = tkinter.Spinbox(frame, from_=1, to=1000000, increment=0.5)
rate_spin.grid(row=8, column=1)
mainWindow.mainloop()
