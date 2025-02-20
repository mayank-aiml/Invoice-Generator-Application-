import tkinter
from tkinter import ttk
from docxtpl import DocxTemplate
import datetime
from tkinter import messagebox

# Initialize window
window = tkinter.Tk()
window.title("Invoice Generator Form")
frame = tkinter.Frame(window)
frame.pack(padx=20, pady=20)

# Function to clear input fields for a new item
def clear_item():
    Quantity_spinbox.delete(0, tkinter.END)
    Quantity_spinbox.insert(0, "1")
    Description_entry.delete(0, tkinter.END)
    Price_spinbox.delete(0, tkinter.END)
    Price_spinbox.insert(0, "0.0")

invoice_list = []

# Function to add an item to the invoice
def add_item():
    try:
        qty = int(Quantity_spinbox.get())
        desc = Description_entry.get().strip()
        price = float(Price_spinbox.get())
        
        if not desc:
            messagebox.showerror("Input Error", "Description cannot be empty.")
            return

        line_total = qty * price
        invoice_item = [qty, desc, price, line_total]
        tree.insert('', 0, values=invoice_item)
        invoice_list.append(invoice_item)

        clear_item()
    
    except ValueError:
        messagebox.showerror("Input Error", "Please enter valid numeric values for Quantity and Price.")

# Function to reset the invoice form
def New_invoice():
    first_name_entry.delete(0, tkinter.END)
    address_entry.delete(0, tkinter.END)
    bill_entry.delete(0, tkinter.END)
    clear_item()
    tree.delete(*tree.get_children())
    invoice_list.clear()

# Function to generate the invoice document
def generate_invoice():
    if not first_name_entry.get().strip():
        messagebox.showerror("Input Error", "Customer Name is required.")
        return

    doc = DocxTemplate("invoice_template.docx")
    name = first_name_entry.get().strip()
    address = address_entry.get().strip()
    bill_id = bill_entry.get().strip()
    
    # Automatically insert the current date
    current_date = datetime.datetime.now().strftime("%Y-%m-%d")

    subtotal = sum(item[3] for item in invoice_list)
    
    total = subtotal   # Corrected tax calculation

    # Render data into the Word template
    doc.render({
        "name": name,
        "address": address,
        "ID": bill_id,  # Maps Bill No. to ID variable in the template
        "Date": current_date,  # Automatically fills the date
        "invoice_list": invoice_list,
        "subtotal": subtotal,
        
        "total": total
    })

    doc_name = f"invoice_{name}_{current_date}.docx"
    doc.save(doc_name)

    messagebox.showinfo("Invoice Complete", f"Invoice saved as {doc_name}")
    New_invoice()

# Labels
first_name_label = tkinter.Label(frame, text="Name")
address_label = tkinter.Label(frame, text="Address")
bill_label = tkinter.Label(frame, text="Bill No.")
Quantity_label = tkinter.Label(frame, text="Quantity")
Description_label = tkinter.Label(frame, text="Description")
Price_label = tkinter.Label(frame, text="Unit Price")

first_name_label.grid(row=0, column=0)
address_label.grid(row=0, column=1)
bill_label.grid(row=0, column=2)
Quantity_label.grid(row=2, column=0)
Description_label.grid(row=2, column=1)
Price_label.grid(row=2, column=2)

# Entry Fields
first_name_entry = tkinter.Entry(frame)
address_entry = tkinter.Entry(frame)
bill_entry = tkinter.Entry(frame)
Quantity_spinbox = tkinter.Spinbox(frame, from_=1, to=1000)
Description_entry = tkinter.Entry(frame)
Price_spinbox = tkinter.Spinbox(frame, from_=0.0, to=1000.0, increment=0.5)

first_name_entry.grid(row=1, column=0, padx=20, pady=10)
address_entry.grid(row=1, column=1, padx=20, pady=10)
bill_entry.grid(row=1, column=2, padx=20, pady=10)
Quantity_spinbox.grid(row=3, column=0, padx=20, pady=10)
Description_entry.grid(row=3, column=1, padx=20, pady=10)
Price_spinbox.grid(row=3, column=2, padx=20, pady=10)

# Buttons
add_item_button = tkinter.Button(frame, text="Add Item", command=add_item)
add_item_button.grid(row=4, column=2, pady=5)

columns = ('Quantity', 'Description', 'Price', 'Total')
tree = ttk.Treeview(frame, columns=columns, show="headings")
tree.grid(row=5, column=0, columnspan=3, padx=20, pady=10)

for col in columns:
    tree.heading(col, text=col)
    tree.column(col, anchor="center")

generate_button = tkinter.Button(frame, text='Generate Invoice', command=generate_invoice)
new_invoice_button = tkinter.Button(frame, text='New Invoice', command=New_invoice)

generate_button.grid(row=6, column=0, columnspan=3, sticky="news", padx=20, pady=5)
new_invoice_button.grid(row=7, column=0, columnspan=3, sticky="news", padx=20, pady=5)

window.mainloop()
