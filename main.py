
import tkinter
from tkinter import ttk
from docxtpl import DocxTemplate
import datetime
from tkinter import messagebox
from tkinter import* 

def clear_item():
    qty_spinbox.delete(0, tkinter.END)
    qty_spinbox.insert(0, "1")
    desc_entry.delete(0, tkinter.END)
    price_spinbox.delete(0, tkinter.END)
    price_spinbox.insert(0, "0.0")

#Adaugare produse

invoice_list = []
i=0
def add_item():
    global i
    i=i+1
    crt=i
    qty = int(qty_spinbox.get())
    desc = desc_entry.get()
    price = float(price_spinbox.get())
    line_total = qty*price
    invoice_item = [crt, qty, desc, price, line_total]
    tree.insert('',0, values=invoice_item)
    clear_item()
    
    invoice_list.append(invoice_item)

#Stergere produse

def delete():
   selected_item = tree.selection()[0]
   tree.delete(selected_item)
   global i
   i=i-1

#Factura noua

def new_invoice():
    global i
    i=0
    name_entry.delete(0, tkinter.END)
    cif_entry.delete(0, tkinter.END)   
    reg_entry.delete(0, tkinter.END)
    judet_entry.delete(0, tkinter.END)
    cont_entry.delete(0, tkinter.END)
    clear_item()
    tree.delete(*tree.get_children())
    
    invoice_list.clear()
    
#Generare factura

def generate_invoice():
    doc = DocxTemplate("invoice_template.docx")
    name = name_entry.get()
    cif = cif_entry.get()
    reg = reg_entry.get()
    judet = judet_entry.get()
    cont = cont_entry.get()
    subtotal = sum(item[3] for item in invoice_list) 
    salestax = 0.19 * subtotal
    total = subtotal+salestax
    
#Introducere date in factura

    doc.render({"name":name, 
            "cif":cif,
            "reg":reg,
            "judet":judet,
            "cont":cont,
            "invoice_list": invoice_list,
            "subtotal":subtotal,
            "salestax":str(0.19*100)+"%",
            "total":total})
    
#Salvare factura

    doc_name = "./Facturi/Factura_" + name + datetime.datetime.now().strftime("_%d-%m-%Y_%H%M%S") + ".docx"
    doc.save(doc_name)
    
    messagebox.showinfo("Factura Generata", "Factura Generata")
    
    new_invoice()


#Creere interfata grafica

window = tkinter.Tk()
window.title("Generator Facturi")
window.iconbitmap("./Logo.ico")
window.tk.call("source", "azure.tcl")
window.tk.call("set_theme", "light")

frame = tkinter.Frame(window)
frame.grid(padx=20, pady=10)

name_label = tkinter.Label(frame, text="Nume")
name_label.grid(row=0, column=0)
name_entry = ttk.Entry(frame)
name_entry.grid(row=1, column=0)

reg_label = tkinter.Label(frame, text="Nr. reg")
reg_label.grid(row=0, column=1)
reg_entry = ttk.Entry(frame)
reg_entry.grid(row=1, column=1)

cif_label = tkinter.Label(frame, text="C.I.F.")
cif_label.grid(row=0, column=2)
cif_entry = ttk.Entry(frame)
cif_entry.grid(row=1, column=2)

cont_label = tkinter.Label(frame, text="Cont IBAN")
cont_label.grid(row=0, column=3)
cont_entry = ttk.Entry(frame)
cont_entry.grid(row=1, column=3)

judet_label = tkinter.Label(frame, text="Judet")
judet_label.grid(row=2, column=0)
judet_entry = ttk.Entry(frame)
judet_entry.grid(row=3, column=0)

qty_label = tkinter.Label(frame, text="Cantitate")
qty_label.grid(row=2, column=3)
qty_spinbox = ttk.Spinbox(frame, from_=1, to=100)
qty_spinbox.insert(0, "1")
qty_spinbox.grid(row=3, column=3)

desc_label = tkinter.Label(frame, text="Denumire Produs")
desc_label.grid(row=2, column=1)
desc_entry = ttk.Entry(frame)
desc_entry.grid(row=3, column=1)

price_label = tkinter.Label(frame, text="Pret Unitar")
price_label.grid(row=2, column=2)
price_spinbox = ttk.Spinbox(frame, from_=0.0, to=500, increment=0.5)
price_spinbox.insert(0, "0.0")
price_spinbox.grid(row=3, column=2)

add_item_button = ttk.Button(frame, text = "Adaugare Produs", style='TButton', command = add_item)
add_item_button.grid(row=5, column=2, pady=8)

del_item_button = ttk.Button(frame, text = "Stergere Produs", style='TButton', command = delete)
del_item_button.grid(row=5, column=3, pady=8)

columns = ('crt', 'qty', 'desc', 'price', 'total')
tree = ttk.Treeview(frame, columns=columns, show="headings")
tree.heading('crt', text='Nr. crt.')
tree.heading('qty', text='Cantitate')
tree.heading('desc', text='Denumire Produs')
tree.heading('price', text='Pret Unitar')
tree.heading('total', text="Total")
    
tree.grid(row=6, column=0, columnspan=4, padx=20, pady=10)

scrollbar = ttk.Scrollbar(window, orient=tkinter.VERTICAL, command=tree.yview)
tree.configure(yscroll=scrollbar.set)
scrollbar.grid(row=0, column=1, sticky='ns')



save_invoice_button = ttk.Button(frame, text="Generare Factura", style='TButton', command=generate_invoice)
save_invoice_button.grid(row=7, column=0, columnspan=4, sticky="news", padx=20, pady=5)
new_invoice_button = ttk.Button(window, text='Factura Noua', style='TButton', command=new_invoice)
new_invoice_button.grid(row=8, column=0, columnspan=4, sticky="news", padx=20, pady=5)

window.mainloop()