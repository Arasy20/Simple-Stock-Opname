import tkinter as tk
from tkinter import messagebox, filedialog
from tkinter import ttk
import json
import os
import pandas as pd

class Inventory:
    def __init__(self, filename="inventory.json"):
        self.filename = filename
        self.items = self.load_items()

    def load_items(self):
        if os.path.exists(self.filename):
            with open(self.filename, 'r') as f:
                return json.load(f)
        return []

    def save_items(self):
        with open(self.filename, 'w') as f:
            json.dump(self.items, f)

    def add_item(self, name, type_, brand, quantity):
        for item in self.items:
            if item['name'] == name and item['type'] == type_ and item['brand'] == brand:
                item['quantity'] += quantity
                self.save_items()
                return
        self.items.append({
            'name': name,
            'type': type_,
            'brand': brand,
            'quantity': quantity
        })
        self.save_items()

    def remove_item(self, name, type_, brand, quantity):
        for item in self.items:
            if item['name'] == name and item['type'] == type_ and item['brand'] == brand:
                if item['quantity'] >= quantity:
                    item['quantity'] -= quantity
                    if item['quantity'] == 0:
                        self.items.remove(item)
                    self.save_items()
                    return True
                else:
                    return False
        return False

    def search_by_type(self, type_):
        return [item for item in self.items if item['type'] == type_]

    def get_items(self):
        return self.items

    def import_from_excel(self, file_path):
        try:
            df = pd.read_excel(file_path)
            for _, row in df.iterrows():
                self.add_item(row['Nama'], row['Type'], row['Brand'], int(row['Quantity']))
        except Exception as e:
            messagebox.showerror("Error", f"Failed to import from Excel: {e}")

class InventoryApp:
    def __init__(self, root):
        self.inventory = Inventory()
        self.title_label = tk.Label(root, text="Pendataan Barang", font=('Arial', 16))
        self.title_label.pack(pady=10)

        self.root = root
        self.root.title("INVENTORY")
        self.root.geometry("800x600")  # Adjust the window size

        # Main frame for Add and Remove item sections
        self.main_frame = tk.Frame(root)
        self.main_frame.pack(pady=10)

        # Add item section
        self.add_frame = tk.Frame(self.main_frame)
        self.add_frame.pack(side=tk.LEFT, padx=10)

        tk.Label(self.add_frame, text="Add Item").grid(row=0, columnspan=2)
        tk.Label(self.add_frame, text="Name:").grid(row=1, column=0)
        tk.Label(self.add_frame, text="Type:").grid(row=2, column=0)
        tk.Label(self.add_frame, text="Brand:").grid(row=3, column=0)
        tk.Label(self.add_frame, text="Quantity:").grid(row=4, column=0)

        self.name_entry = tk.Entry(self.add_frame)
        self.name_entry.grid(row=1, column=1)
        self.type_entry = tk.Entry(self.add_frame)
        self.type_entry.grid(row=2, column=1)
        self.brand_entry = tk.Entry(self.add_frame)
        self.brand_entry.grid(row=3, column=1)
        self.quantity_entry = tk.Entry(self.add_frame)
        self.quantity_entry.grid(row=4, column=1)

        self.add_button = tk.Button(self.add_frame, text="Add Item", command=self.add_item)
        self.add_button.grid(row=5, columnspan=2, pady=5)

        # Remove item section
        self.remove_frame = tk.Frame(self.main_frame)
        self.remove_frame.pack(side=tk.LEFT, padx=10)

        tk.Label(self.remove_frame, text="Remove Item").grid(row=0, columnspan=2)
        tk.Label(self.remove_frame, text="Name:").grid(row=1, column=0)
        tk.Label(self.remove_frame, text="Type:").grid(row=2, column=0)
        tk.Label(self.remove_frame, text="Brand:").grid(row=3, column=0)
        tk.Label(self.remove_frame, text="Quantity:").grid(row=4, column=0)

        self.name_remove_entry = tk.Entry(self.remove_frame)
        self.name_remove_entry.grid(row=1, column=1)
        self.type_remove_entry = tk.Entry(self.remove_frame)
        self.type_remove_entry.grid(row=2, column=1)
        self.brand_remove_entry = tk.Entry(self.remove_frame)
        self.brand_remove_entry.grid(row=3, column=1)
        self.quantity_remove_entry = tk.Entry(self.remove_frame)
        self.quantity_remove_entry.grid(row=4, column=1)

        self.remove_button = tk.Button(self.remove_frame, text="Remove Item", command=self.remove_item)
        self.remove_button.grid(row=5, columnspan=2, pady=5)

        # Search by type section
        self.search_frame = tk.Frame(root)
        self.search_frame.pack(pady=10)

        tk.Label(self.search_frame, text="Type:").grid(row=0, column=0, padx=(0, 5), sticky='e')
        self.type_search_entry = tk.Entry(self.search_frame, width=40)
        self.type_search_entry.grid(row=0, column=1, padx=(5, 0), sticky='w')

        self.search_button = tk.Button(self.search_frame, text="Search Product by Type", command=self.search_by_type)
        self.search_button.grid(row=1, column=0, pady=5, padx=5, sticky='ew')

        self.view_button = tk.Button(self.search_frame, text="View Stock", command=self.view_stock)
        self.view_button.grid(row=1, column=1, pady=5, padx=5, sticky='ew')

        self.import_button = tk.Button(self.search_frame, text="Import from Excel", command=self.import_from_excel)
        self.import_button.grid(row=1, column=2, pady=5, padx=5, sticky='ew')

        self.tree = ttk.Treeview(self.search_frame, columns=("Name", "Type", "Brand", "Quantity"), show="headings")
        self.tree.heading("Name", text="Name")
        self.tree.heading("Type", text="Type")
        self.tree.heading("Brand", text="Brand")
        self.tree.heading("Quantity", text="Quantity")
        self.tree.grid(row=2, columnspan=3, pady=5)

        # Initial view of stock
        self.view_stock()

    def add_item(self):
        name = self.name_entry.get()
        type_ = self.type_entry.get()
        brand = self.brand_entry.get()
        quantity = int(self.quantity_entry.get())

        self.inventory.add_item(name, type_, brand, quantity)
        messagebox.showinfo("Success", "Item added successfully!")

        self.name_entry.delete(0, tk.END)
        self.type_entry.delete(0, tk.END)
        self.brand_entry.delete(0, tk.END)
        self.quantity_entry.delete(0, tk.END)

        # Update the stock view
        self.view_stock()

    def remove_item(self):
        name = self.name_remove_entry.get()
        type_ = self.type_remove_entry.get()
        brand = self.brand_remove_entry.get()
        quantity = int(self.quantity_remove_entry.get())

        if self.inventory.remove_item(name, type_, brand, quantity):
            messagebox.showinfo("Success", "Item removed successfully!")
        else:
            messagebox.showerror("Error", "Failed to remove item.")

        self.name_remove_entry.delete(0, tk.END)
        self.type_remove_entry.delete(0, tk.END)
        self.brand_remove_entry.delete(0, tk.END)
        self.quantity_remove_entry.delete(0, tk.END)

        # Update the stock view
        self.view_stock()

    def view_stock(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        stock = self.inventory.get_items()
        for item in stock:
            self.tree.insert('', tk.END, values=(item['name'], item['type'], item['brand'], item['quantity']))

    def search_by_type(self):
        type_ = self.type_search_entry.get()

        for item in self.tree.get_children():
            self.tree.delete(item)

        results = self.inventory.search_by_type(type_)
        for item in results:
            self.tree.insert('', tk.END, values=(item['name'], item['type'], item['brand'], item['quantity']))

    def import_from_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.inventory.import_from_excel(file_path)
            messagebox.showinfo("Success", "Items imported successfully!")
            self.view_stock()

if __name__ == "__main__":
    root = tk.Tk()
    app = InventoryApp(root)
    root.mainloop()