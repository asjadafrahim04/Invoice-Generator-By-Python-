import tkinter
from tkinter import ttk
from docxtpl import DocxTemplate
import datetime
from tkinter import messagebox


def clear_item():
    qty_spinbox.delete(0, tkinter.END)
    qty_spinbox.insert(0, "1")
    desc_entry.delete(0, tkinter.END)
    price_spinbox.delete(0, tkinter.END)
    price_spinbox.insert(0, "0.0")


invoice_list = []


def add_item():
    qty = int(qty_spinbox.get())
    desc = desc_entry.get()
    price = float(price_spinbox.get())
    line_total = qty * price
    invoice_item = [qty, desc, price, line_total]
    tree.insert('', 0, values=invoice_item)
    clear_item()

    invoice_list.append(invoice_item)


def new_invoice():
    name_entry.delete(0, tkinter.END)
    phone_entry.delete(0, tkinter.END)
    member_entry.delete(0, tkinter.END)
    clear_item()
    tree.delete(*tree.get_children())

    invoice_list.clear()


def generate_invoice():
    doc = DocxTemplate("invoice_template.docx")
    name = name_entry.get()
    phone = phone_entry.get()
    member = member_entry.get()
    subtotal = sum(item[3] for item in invoice_list)
    salestax = 0.1
    total = subtotal * (1 - salestax)

    doc.render({"name": name,
                "phone": phone,
                "member": member,
                "invoice_list": invoice_list,
                "subtotal": subtotal,
                "salestax": str(salestax * 100) + "%",
                "total": total})

    doc_name = name + datetime.datetime.now().strftime("%Y-%m-%d-%H%M%S") + ".docx"
    doc.save(doc_name)

    messagebox.showinfo("Invoice Complete", "Invoice Complete")

    new_invoice()


window = tkinter.Tk()
window.title("Invoice Generator Form")


frame = tkinter.Frame(window)
frame.pack(padx = 30, pady = 20)

name_label = tkinter.Label(frame, text="Name")
name_label.grid(row = 0, column = 0)

name_entry = tkinter.Entry(frame)
name_entry.grid(row = 1, column = 0, columnspan = 1, sticky = "news")

phone_label = tkinter.Label(frame, text="Phone Number")
phone_label.grid(row = 2, column = 0)

phone_entry = tkinter.Entry(frame)
phone_entry.grid(row = 3, column = 0,columnspan = 1, sticky = "news")

member_label= tkinter.Label(frame, text="Membership Number")
member_label.grid(row = 2, column = 1)

member_entry = tkinter.Entry(frame)
member_entry.grid(row = 3, column = 1)

qty_label = tkinter.Label(frame, text="Qty")
qty_label.grid(row = 5, column = 0,sticky = "w")
qty_spinbox = tkinter.Spinbox(frame, from_=1, to=100)
qty_spinbox.grid(row = 6, column = 0,sticky = "w")

desc_label = tkinter.Label(frame, text="Description")
desc_label.grid(row = 5, column = 1,sticky = "nw")
desc_entry = tkinter.Entry(frame)
desc_entry.grid(row = 6, column = 1,sticky = "nw")

price_label = tkinter.Label(frame, text="Unit Price")
price_label.grid(row = 5, column = 2)
price_spinbox = tkinter.Spinbox(frame, from_=0.0, to=5000, increment = 1.00)
price_spinbox.grid(row = 6, column = 2)

add_item_button = tkinter.Button(frame, text="Add Item", command = add_item)
add_item_button.grid(row = 7, column = 2, pady = 6)

columns = ('qty','dese','price','total')
tree = ttk.Treeview(frame, columns=columns, show = "headings")
tree.heading('qty', text = "Quantity")
tree.heading('dese', text = "Description")
tree.heading('price', text = "Price")
tree.heading('total', text = "Total")


tree.grid(row =8, column = 0, columnspan=3, padx = 20, pady = 10)

save_invoice_button = tkinter.Button(frame, text ="Generate Invoice", command=generate_invoice)
save_invoice_button.grid(row = 9, column = 0, columnspan = 3, sticky = "news", padx = 20, pady = 5)
new_invoice_button = tkinter.Button(frame, text ="New Invoice", command=new_invoice)
new_invoice_button.grid(row = 10, column = 0, columnspan = 3, sticky = "news", padx = 20, pady = 5)







window.mainloop()
