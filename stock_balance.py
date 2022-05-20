import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
import sqlite3
import pandas as pd
import os
import csv


class ScrollTree(ttk.Treeview):

    def __init__(self, window, **kwargs):
        super().__init__(window, **kwargs)

        self.scrollbar = ttk.Scrollbar(window, orient=tk.VERTICAL, command=self.yview)

    def grid(self, row, column, rowspan=1, columnspan=1, sticky='nse', **kwargs):
        super().grid(row=row, column=column, rowspan=rowspan, columnspan=columnspan, sticky=sticky, **kwargs)
        self.scrollbar.grid(row=row, column=column, sticky='nse', rowspan=rowspan)
        self['yscrollcommand'] = self.scrollbar.set


class Stocktree(ScrollTree):

    def __init__(self, window, connection, table, **kwargs):

        super().__init__(window, **kwargs)

        self.cursor = connection.cursor()
        self.table = table
        self.count = 0

        # Bindings
        self.bind('<ButtonRelease-1>', self.select_product)

    def query(self):

        # Clear Treeview
        for product in self.get_children():
            self.delete(product)

        self.cursor.execute("SELECT rowid, * FROM " + self.table)
        products = self.cursor.fetchall()

        self.count = 0
        for product in products:
            if self.count % 2 == 0:
                self.insert(parent='', index='end', iid=self.count,
                            values=(product[0], product[1], product[2], product[3], product[4], product[5]),
                            tags=('evenrow',))
            else:
                self.insert(parent='', index='end', iid=self.count,
                            values=(product[0], product[1], product[2], product[3], product[4], product[5]),
                            tags=('oddrow',))
            self.count += 1

    @staticmethod
    def data_b(*args):
        con = sqlite3.connect('warehouse_stock.db')
        cursor = con.cursor()
        cursor.execute(*args)
        con.commit()
        con.close()

    def remove_selected(self):
        ask_user = tk.messagebox.askyesno('Remove Selected', 'Are you sure you want to remove selected item(s)?')
        if ask_user:
            selected_products = self.selection()
            print(selected_products)
            # print(selected_products)
            # oid_numbers = [(int(product) + 1) for product in selected_products]
            # print(oid_numbers)
            barcodes = [self.item(selected_item, 'values')[2] for selected_item in selected_products]
            print(barcodes)  # TODO remove
            for product in selected_products:
                self.delete(product)

            con = sqlite3.connect('warehouse_stock.db')
            cu = con.cursor()
            cu.executemany("DELETE FROM " + self.table + " WHERE Barcode = ?", [(b,) for b in barcodes])

            con.commit()

            self.clear_entries()
            tk.messagebox.showinfo('Delete Item', 'Item(s) has been successfully deleted.')

    def remove_all(self):
        ask_user = tk.messagebox.askyesno('Remove All Products', 'Are you sure you want to remove all items?')
        if ask_user:
            for row in self.get_children():
                self.delete(row)

            self.data_b("DROP TABLE " + self.table)

            self.clear_entries()
            tk.messagebox.showinfo('Delete Items', 'Items have been successfully deleted.')

            self.create_table()

    @staticmethod
    def clear_entries():
        item_entry.delete(0, tk.END)
        barcode_entry.delete(0, tk.END)
        item_description_entry.delete(0, tk.END)
        size_entry.delete(0, tk.END)
        quantity_entry.delete(0, tk.END)

    def add_product(self):
        ask_user = tk.messagebox.askyesno('Add Item', 'Are you sure you want to add a new item?')
        if ask_user:
            self.insert('', 'end', values=(item_entry.get(), barcode_entry.get(), item_description_entry.get(),
                                           size_entry.get(), quantity_entry.get(),))

            self.data_b("INSERT into " + self.table + """ VALUES(
                                :item, 
                                :barcode, 
                                :item_d,
                                :size,
                                :quantity)""",

                        {
                            'item': item_entry.get(),
                            'barcode': barcode_entry.get(),
                            'item_d': item_description_entry.get(),
                            'size': size_entry.get(),
                            'quantity': quantity_entry.get(),
                        })

            self.clear_entries()
            self.delete(*self.get_children())  # asterisk makes it so that all rows get unpacked
            self.query()
            tk.messagebox.showinfo('Added Item', 'Item has successfully been added')

    def move_up(self):
        selected_product = self.selection()
        for row in selected_product:
            self.move(row, self.parent(row), self.index(row) - 1)

    def move_down(self):
        selected_product = self.selection()
        for row in reversed(selected_product):  # use reverse to keep correct index positions
            self.move(row, self.parent(row), self.index(row) + 1)

    def select_product(self, event):
        # Clear the entry boxes from data
        self.clear_entries()
        try:
            # Grab selected data
            selected_item = self.focus()
            values = self.item(selected_item, 'values')

            # Insert data to corresponding entry boxes
            item_entry.insert(0, values[1])
            barcode_entry.insert(0, values[2])
            item_description_entry.insert(0, values[3])
            size_entry.insert(0, values[4])
            quantity_entry.insert(0, values[5])
        except IndexError:
            pass

    def update_product(self):
        ask_user = tk.messagebox.askyesno('Update selected', 'Are you sure you want to update selected item?')
        if ask_user:
            selected_product = self.focus()
            oid_number = int(selected_product) + 1
            self.item(selected_product, values=(str(oid_number), item_entry.get(), barcode_entry.get(),
                                                item_description_entry.get(),
                                                size_entry.get(), quantity_entry.get(),))

            self.data_b("UPDATE " + self.table + """ SET
                                Item = :item, 
                                Barcode = :barcode, 
                                'Item Description' = :item_d,
                                Size = :size,
                                Quantity = :quantity 
                                
                                WHERE OID = :oid""",

                        {
                            'item': item_entry.get(),
                            'barcode': barcode_entry.get(),
                            'item_d': item_description_entry.get(),
                            'size': size_entry.get(),
                            'quantity': quantity_entry.get(),
                            'oid': str(oid_number)
                        })

            self.clear_entries()
            tk.messagebox.showinfo('Update Item', 'Item has successfully been updated.')

    def search_item(self):
        find_item = find_i_entry.get()
        find_i.destroy()

        for item in self.get_children():
            self.delete(item)

        con = sqlite3.connect('warehouse_stock.db')
        cu = con.cursor()
        cu.execute("SELECT rowid, * FROM " + self.table + " WHERE Item LIKE ?", (find_item,))
        products = cu.fetchall()
        print(products)

        self.count = 0
        for product in products:
            if self.count % 2 == 0:
                self.insert(parent='', index='end', iid=self.count,
                            values=(product[0], product[1], product[2], product[3], product[4], product[5]),
                            tags=('evenrow',))
            else:
                self.insert(parent='', index='end', iid=self.count,
                            values=(product[0], product[1], product[2], product[3], product[4], product[5]),
                            tags=('oddrow',))
            self.count += 1

        con.commit()
        con.close()

    def find_item(self):
        global find_i, find_i_entry

        find_i = tk.Toplevel(root)
        find_i.title = 'Find Item'
        find_i.geometry('400x200')

        find_i_frame = tk.LabelFrame(find_i, text='Item')
        find_i_frame.pack(padx=10, pady=10)

        find_i_entry = tk.Entry(find_i_frame)
        find_i_entry.pack(padx=20, pady=20)

        find_button = tk.Button(find_i, text='Search Item ', command=self.search_item)
        find_button.pack(padx=20, pady=20)

    def search_barcode(self):
        find_barcode = find_b_entry.get()
        find_b.destroy()

        for item in self.get_children():
            self.delete(item)

        con = sqlite3.connect('warehouse_stock.db')
        cu = con.cursor()
        cu.execute("SELECT rowid, * FROM " + self.table + " WHERE  Barcode = ?", (find_barcode,))
        products = cu.fetchall()

        self.count = 0
        for product in products:
            if self.count % 2 == 0:
                self.insert(parent='', index='end', iid=self.count,
                            values=(product[0], product[1], product[2], product[3], product[4], product[5]),
                            tags=('evenrow',))
            else:
                self.insert(parent='', index='end', iid=self.count,
                            values=(product[0], product[1], product[2], product[3], product[4], product[5]),
                            tags=('oddrow',))
            self.count += 1

        con.commit()
        con.close()

    def find_barcode(self):
        global find_b, find_b_entry

        find_b = tk.Toplevel(root)
        find_b.title = 'Find Barcode'
        find_b.geometry('400x200')

        find_b_frame = tk.LabelFrame(find_b, text='Barcode')
        find_b_frame.pack(padx=10, pady=10)

        find_b_entry = tk.Entry(find_b_frame)
        find_b_entry.pack(padx=20, pady=20)

        find_b_button = tk.Button(find_b, text='Find Barcode', command=self.search_barcode)
        find_b_button.pack(padx=20, pady=20)

    def create_table(self):
        self.data_b(
            """
            CREATE TABLE IF NOT EXISTS products (
                    Item TEXT,
                    Barcode TEXT,
                    'Item Description' TEXT,
                    Size TEXT,
                    Quantity TEXT
                    );
                """
        )

    def primary_color(self):
        primary_color = colorchooser.askcolor()[1]

        if primary_color:
            self.tag_configure('evenrow', background=primary_color)

    def secondary_color(self):
        secondary_color = colorchooser.askcolor()[1]

        if secondary_color:
            self.tag_configure('oddrow', background=secondary_color)

    @staticmethod
    def highlight_color():
        highlight_color = colorchooser.askcolor()[1]
        # Change Selected Color
        if highlight_color:
            style.map('Treeview', background=[('selected', highlight_color)])

    def save_session(self):
        cols = ['Item', 'Barcode', 'Item Description', 'Size', 'Quantity']
        path = 'read.csv'
        excel_name = filedialog.asksaveasfilename(title='Save to location',
                                                  initialfile='Untitled.xlsx',
                                                  initialdir=os.path.expanduser('~/Documents'),
                                                  defaultextension=[('Excel File', '*.xlsx')],
                                                  filetypes=[('Excel File', '*.xlsx')])
        if excel_name:

            lst = []
            with open(path, "w", newline='') as my_file:
                csv_writer = csv.writer(my_file, delimiter=',')
                for row_id in self.get_children():
                    row = self.item(row_id, str('values'))
                    lst.append(row)
                lst = list(map(list, lst))
                lst.insert(0, cols)
                for row in lst:
                    csv_writer.writerow(row)

            writer = pd.ExcelWriter(excel_name)
            scan_t = pd.read_csv(path)
            scan_t.to_excel(writer, 'sheetname')
            writer.save()


class Scantree(Stocktree):

    def __init__(self, tab, **kwargs):

        super().__init__(tab, **kwargs)

    def scan_check(self, event):
        scanned_item = scan_entry.get().strip()

        if scanned_item == '':
            # do nothing if empty string is input
            return

        # elif len(scanned_item) == 12:   # Some barcodes start with  leading 0 and barcode reader skips the 0 when scanning
        #     n = 1
        #     scanned_item = scanned_item.zfill(n + len(scanned_item))

        # Check if item is al ready in scantree
        for child in self.get_children():
            # row = self.set(child)
            row = self.item(child, 'values')
            if row[1] == scanned_item:
                # update quantity
                self.set(child, column='Scanned Quantity', value=int(row[4]) + 1)
                break

        else:
            con = sqlite3.connect('warehouse_stock.db')
            cu = con.cursor()
            cu.execute("SELECT rowid, * FROM " + self.table + " WHERE Barcode = ?", (scanned_item,))
            product = cu.fetchone()
            if product:
                self.insert('', 'end', values=(product[1], product[2], product[3], product[4], int(1)),
                            tag=('scanlist',))

            else:
                unknown_listbox.insert(tk.END, scanned_item)
        scan_entry.delete(0, 'end')

    def s_remove_selected(self):
        ask_user = tk.messagebox.askyesno('Delete Item(s)', 'Are you sure you want to delete item(s)?')
        if ask_user:
            selected_items = self.selection()
            for row in selected_items:
                self.delete(row)

            tk.messagebox.showinfo('Delete Item(s)', 'Item(s) have been successfully deleted.')

    def reset(self):
        ask_user = tk.messagebox.askyesno('Reset Scan list', 'Are you sure you want to reset current scanning session?')
        if ask_user:
            for child in self.get_children():
                self.delete(child)
            unknown_listbox.delete(0, tk.END)

    def sort(self):
        items = [self.item(child, 'values')[0] for child in self.get_children()]
        for child in self.get_children():
            row = self.item(child, 'values')

            self.data_b(""" 
            CREATE TABLE IF NOT EXISTS quantity (
                                                Item TEXT,
                                                Barcode TEXT,
                                                'Item Description' TEXT,
                                                Size TEXT,
                                                'Scanned Quantity' INT
                                                );
                        """)

            self.data_b(""" INSERT INTO quantity VALUES(
                                    :item,
                                    :barcode,
                                    :item_description,
                                    :size,
                                    :scanned_quantity); """,

                        {'item': row[0],
                         'barcode': row[1],
                         'item_description': row[2],
                         'size': row[3],
                         'scanned_quantity': row[4]
                         })
        self.delete(*self.get_children())
        for item in sorted(items):
            con = sqlite3.connect('warehouse_stock.db')
            cu = con.cursor()
            cu.execute("SELECT * FROM quantity WHERE Item = ?", (item,))
            products = cu.fetchall()

            for product in products:
                self.insert(parent='', index='end', values=(product[0], product[1], product[2], product[3], product[4]),
                            tag=('scanlist',))

        self.data_b("DROP TABLE IF EXISTS quantity")


class Balancetree(Scantree):

    def __init__(self, tab, **kwargs):

        super().__init__(tab, **kwargs)

        # Bindings
        self.bind('<ButtonRelease-1>', self.b_select_product)

    def b_select_product(self, event):
        # Clear the entry boxes from data
        self.b_clear_entries()

        try:
            # Grab selected data
            selected_item = self.focus()
            values = self.item(selected_item, 'values')

            # Insert data to corresponding entry boxes
            balance_item_entry.insert(0, values[0])
            balance_barcode_entry.insert(0, values[1])
            balance_item_description_entry.insert(0, values[2])
            balance_size_entry.insert(0, values[3])
            balance_quantity_entry.insert(0, values[4])
            balance_scanned_entry.insert(0, values[5])
            balance_quantity_difference_entry.insert(0, values[6])
        except IndexError:
            pass

    def balance(self):

        for child in scan_tree.get_children():
            row = scan_tree.item(child, 'values')
            barcode = row[1]

            con = sqlite3.connect('warehouse_stock.db')
            cu = con.cursor()
            cu.execute("SELECT rowid, * FROM " + self.table + " WHERE Barcode = ?", (barcode,))

            product = cu.fetchone()

            for ch in self.get_children():
                r = self.set(ch)
                if r['Barcode'] == barcode:
                    self.set(ch, column='Scanned Quantity', value=row[4])
                    self.set(ch, column='Quantity Difference', value=(int(row[4]) - int(product[5])))

                    break  # Prevent adding multiple lines of the same entry when pressing balance button

            else:
                con = sqlite3.connect('warehouse_stock.db')
                cu = con.cursor()
                cu.execute("SELECT rowid, * FROM " + self.table + " WHERE Barcode = ?", (barcode,))
                product = cu.fetchone()
                print(product)
                self.insert(parent='', index='end', values=(product[1], product[2], product[3], product[4],
                                                            product[5], row[4], (int(row[4]) - int(product[5]))),
                            tag=('balancelist',))

    @staticmethod
    def b_clear_entries():
        balance_item_entry.delete(0, tk.END)
        balance_barcode_entry.delete(0, tk.END)
        balance_item_description_entry.delete(0, tk.END)
        balance_size_entry.delete(0, tk.END)
        balance_quantity_entry.delete(0, tk.END)
        balance_scanned_entry.delete(0, tk.END)
        balance_quantity_difference_entry.delete(0, tk.END)

    def b_remove_all(self):
        ask_user = tk.messagebox.askyesno('Remove All Products', 'Are you sure you want to remove all items?')
        if ask_user:
            for row in self.get_children():
                self.delete(row)

            self.b_clear_entries()
            tk.messagebox.showinfo('Delete Items', 'Items have been successfully deleted.')

    def b_remove_selected(self):
        ask_user = tk.messagebox.askyesno('Remove Products', 'Are you sure you want to remove all items?')
        if ask_user:
            selected_items = self.selection()
            for row in selected_items:
                self.delete(row)

            self.b_clear_entries()
            tk.messagebox.showinfo('Delete Items', 'Item(s) have been successfully deleted.')

    def b_update_product(self):
        ask_user = tk.messagebox.askyesno('Update selected', 'Are you sure you want to update selected item?')
        if ask_user:
            selected_item = self.focus()
            self.item(selected_item, values=(balance_item_entry.get(), balance_barcode_entry.get(),
                                             balance_item_description_entry.get(), balance_size_entry.get(),
                                             balance_quantity_entry.get(), balance_scanned_entry.get(),
                                             balance_quantity_difference_entry.get(),))

        self.clear_entries()
        tk.messagebox.showinfo('Update Item', 'Item has successfully been updated.')

    def b_sort(self):
        items = [self.item(child, 'values')[0] for child in self.get_children()]
        for child in self.get_children():
            row = self.item(child, 'values')

            self.data_b("""CREATE TABLE IF NOT EXISTS quantity ( 
                                                                Item TEXT,
                                                                Barcode TEXT,
                                                                'Item Description' TEXT,
                                                                Size TEXT,
                                                                Quantity INT,
                                                                'Scanned Quantity' INT,
                                                                'Quantity Difference' INT
                                                                );
            
                        """)

            self.data_b(""" INSERT INTO quantity VALUES( 
                                                        :item,
                                                        :barcode,
                                                        :item_description,
                                                        :size,
                                                        :quantity,
                                                        :scanned_quantity,
                                                        :quantity_difference); """,

                        {'item': row[0],
                         'barcode': row[1],
                         'item_description': row[2],
                         'size': row[3],
                         'quantity': row[4],
                         'scanned_quantity': row[5],
                         'quantity_difference': row[6]
                         })

        self.delete(*self.get_children())
        for item in sorted(items):
            con = sqlite3.connect('warehouse_stock.db')
            cu = con.cursor()
            cu.execute("SELECT * FROM quantity WHERE Item = ?", (item,))
            products = cu.fetchall()

            for product in products:
                self.insert(parent='', index='end', values=(product[0], product[1], product[2], product[3], product[4],
                                                            product[5], product[6]), tag=('balancelist'),)

        self.data_b("DROP TABLE IF EXISTS quantity")


def remove_selected():
    selected_item = unknown_listbox.curselection()
    unknown_listbox.delete(selected_item)


def save_unknown():
    path = os.path.expanduser('~/Documents')
    file = filedialog.asksaveasfilename(
        title='Save file as',
        initialfile='Unnamed.txt',
        initialdir=path,
        defaultextension=[('txt File', '*.txt')],
        filetypes=[('txt File', '*.txt')])
    if file:  # When not canceled
        with open(file, 'w') as f:
            for row in unknown_listbox.get(0, tk.END):
                f.write(row)
                f.write('\n')
        f.close()


if __name__ == "__main__":
    # SQL #

    # Create a database or connect to one
    conn = sqlite3.connect('warehouse_stock.db')

    # Create cursor (Cursor gets a quest and completes it)
    c = conn.cursor()

    # Create table
    c.execute(
        """
        CREATE TABLE IF NOT EXISTS products (
                Item TEXT,
                Barcode TEXT,
                'Item Description' TEXT,
                Size TEXT,
                Quantity TEXT
                );
            """
    )

    # Add data to table
    # warehouse_stock = pd.read_excel('Hooijer Artikelen.xlsx', sheet_name='Products', header=0)
    # warehouse_stock.to_sql('products', con=conn, if_exists='append', index=False)

    # Commit Change
    conn.commit()

    # Setup screen
    root = tk.Tk()
    root.title('Stock')
    root.geometry('1024x768')
    root.state('zoomed')

    # Creating the frame to place the tabs in
    tabControl = ttk.Notebook(root)
    tabControl.pack(expand='True', fill="both")

    # Creating the Tabs
    tab1 = ttk.Frame(tabControl)
    tab2 = ttk.Frame(tabControl)
    tab3 = ttk.Frame(tabControl)
    tab4 = ttk.Frame(tabControl)

    # Adding the tabs
    tabControl.add(tab1, text='Current Stock')
    tabControl.add(tab2, text='Scan')
    tabControl.add(tab3, text='Unknown Products')
    tabControl.add(tab4, text='Balance')

    # Configuring tab 1
    tab1.columnconfigure(0, weight=1)

    tab1.rowconfigure(0, weight=10)
    tab1.rowconfigure(1, weight=1)
    tab1.rowconfigure(2, weight=1)

    # Configuring tab 2
    tab2.columnconfigure(0, weight=1)

    tab2.rowconfigure(0, weight=10)
    tab2.rowconfigure(1, weight=1)
    tab2.rowconfigure(2, weight=1)

    # Configuring tab 3
    tab3.columnconfigure(0, weight=1)

    tab3.rowconfigure(0, weight=10)
    tab3.rowconfigure(1, weight=1)

    # Configuring tab 4
    tab4.columnconfigure(0, weight=1)

    tab4.rowconfigure(0, weight=10)
    tab4.rowconfigure(1, weight=1)
    tab4.rowconfigure(2, weight=1)

    # AddStyle
    style = ttk.Style()

    # Pick A Theme
    style.theme_use('default')

    # Configure the Treeview Colors
    style.configure('Treeview',
                    background='#D3D3D3',
                    foreground='black',
                    rowheight=25,
                    fieldbackground='#D3D3D3')

    # Change Selected Color
    style.map('Treeview',
              background=[('selected', '#347083')])

    # Instantiate ScrollTree #
    stock_tree = Stocktree(tab1, connection=conn, table='products', selectmode=tk.EXTENDED)
    stock_tree.grid(row=0, column=0, sticky='nsew', pady=10)
    stock_tree['columns'] = ('ID', 'Item', 'Barcode', 'Item Description', 'Size', 'Quantity')

    # Create menu
    menu = tk.Menu(root)
    root.config(menu=menu)

    # files menu
    file_menu = tk.Menu(menu, tearoff=False)
    menu.add_cascade(label='File', menu=file_menu)
    file_menu.add_command(label='Save Scanned Item', command=stock_tree.save_session)
    file_menu.add_command(label='Exit', command=root.quit)

    # Search menu
    search_menu = tk.Menu(menu, tearoff=False)
    menu.add_cascade(label='Search', menu=search_menu)
    search_menu.add_command(label='Search Item', command=stock_tree.find_item)
    search_menu.add_command(label='Search Barcode', command=stock_tree.find_barcode)
    search_menu.add_separator()
    search_menu.add_command(label='Reset', command=stock_tree.query)

    # Color options
    color_menu = tk.Menu(menu, tearoff=False)
    menu.add_cascade(label='Color Options', menu=color_menu)
    color_menu.add_command(label='Primary Color', command=stock_tree.primary_color)
    color_menu.add_command(label='Secondary Color', command=stock_tree.secondary_color)
    color_menu.add_command(label='Highlight Color', command=stock_tree.highlight_color)

    # Configure Stocktree columns
    stock_tree.column("#0", width=0, stretch="no")  # Ghost column
    stock_tree.column("ID", anchor='w', width=50)
    stock_tree.column("Item", anchor="w", width=200)
    stock_tree.column("Barcode", anchor="w", width=150)
    stock_tree.column("Item Description", anchor="w", width=425)
    stock_tree.column("Size", anchor="center", width=75)
    stock_tree.column("Quantity", anchor="center", width=150)

    stock_tree.heading("ID", text="ID", anchor='w')
    stock_tree.heading("Item", text="Item", anchor="w")
    stock_tree.heading("Barcode", text="Barcode", anchor="w")
    stock_tree.heading("Item Description", text="Item Description", anchor="w")
    stock_tree.heading("Size", text="Size", anchor="center")
    stock_tree.heading("Quantity", text="Quantity", anchor="center")

    # Create Striped Row Tags
    stock_tree.tag_configure('oddrow', background='white')
    stock_tree.tag_configure('evenrow', background='lightblue')

    # Create products Labelframe
    products_labelframe = tk.LabelFrame(tab1, text='Product')
    products_labelframe.grid(row=1, column=0, sticky='nsew')

    # Create product entry boxes
    item_entry = tk.Entry(products_labelframe, width=25)
    item_entry.grid(row=1, column=0, padx=(10, 10), pady=(5, 10))
    barcode_entry = tk.Entry(products_labelframe, width=25)
    barcode_entry.grid(row=1, column=1, padx=(10, 10), pady=(5, 10))
    item_description_entry = tk.Entry(products_labelframe, width=70)
    item_description_entry.grid(row=1, column=2, padx=(10, 10), pady=(5, 10))
    size_entry = tk.Entry(products_labelframe, width=10)
    size_entry.grid(row=1, column=3, padx=(10, 10), pady=(5, 10))
    quantity_entry = tk.Entry(products_labelframe, width=20)
    quantity_entry.grid(row=1, column=4, padx=(10, 5), pady=(5, 10))

    # Create entry box labels
    item_label = tk.Label(products_labelframe, text='Item').grid(row=0, column=0, padx=(10, 10), pady=10)
    barcode_label = tk.Label(products_labelframe, text='Barcode').grid(row=0, column=1, padx=(10, 10), pady=10)
    item_description_label = tk.Label(products_labelframe, text='Item Description').grid(row=0, column=2, padx=(10, 10),
                                                                                         pady=10)
    size_label = tk.Label(products_labelframe, text='Size').grid(row=0, column=3, padx=(10, 10), pady=10)
    quantity_label = tk.Label(products_labelframe, text='Quantity').grid(row=0, column=4, padx=(10, 10), pady=10)

    # Create Command buttons labelframe
    command_label_frame = tk.LabelFrame(tab1, text='Commands')
    command_label_frame.grid(row=2, column=0, sticky='nsew')

    # Command Buttons
    update_product = tk.Button(command_label_frame, text='Update Product', command=stock_tree.update_product)
    update_product.grid(row=0, column=0, padx=(10, 10), pady=10)
    add_product = tk.Button(command_label_frame, text='Add Product', command=stock_tree.add_product)
    add_product.grid(row=0, column=1, padx=(10, 10), pady=10)
    remove_product = tk.Button(command_label_frame, text='Remove Product', command=stock_tree.remove_selected)
    remove_product.grid(row=0, column=2, padx=(10, 10), pady=10)
    remove_all_products = tk.Button(command_label_frame, text='Remove All Products', command=stock_tree.remove_all)
    remove_all_products.grid(row=0, column=3, padx=(10, 10), pady=10)
    move_up_product = tk.Button(command_label_frame, text='Move Up', command=stock_tree.move_up)
    move_up_product.grid(row=0, column=4, padx=(10, 10), pady=10)
    move_down_product = tk.Button(command_label_frame, text='Move Down', command=stock_tree.move_down)
    move_down_product.grid(row=0, column=5, padx=(10, 10), pady=10)
    clear_product = tk.Button(command_label_frame, text='Clear Entry Boxes', command=stock_tree.clear_entries)
    clear_product.grid(row=0, column=6, padx=(10, 10), pady=10)

    # Run at start to occupy treeview with data
    stock_tree.query()

    # Scan Tree #
    scan_tree = Scantree(tab2, connection=conn, table='products', selectmode=tk.EXTENDED)
    scan_tree.grid(row=0, column=0, sticky='nsew', pady=10)
    scan_tree['columns'] = ('Item', 'Barcode', 'Item Description', 'Size', 'Scanned Quantity')

    # Configure Stocktree columns
    scan_tree.column('#0', width=0, stretch='no')  # Ghost column
    scan_tree.column('Item', anchor='w', width=200)
    scan_tree.column('Barcode', anchor='w', width=150)
    scan_tree.column('Item Description', anchor='w', width=425)
    scan_tree.column('Size', anchor='center', width=75)
    scan_tree.column('Scanned Quantity', anchor='center', width=150)

    scan_tree.heading('Item', text='Item', anchor='w')
    scan_tree.heading('Barcode', text='Barcode', anchor='w')
    scan_tree.heading('Item Description', text='Item Description', anchor='w')
    scan_tree.heading('Size', text='Size', anchor='center')
    scan_tree.heading('Scanned Quantity', text='Scanned Quantity', anchor='center')

    # Create Striped Row Tags
    # scan_tree.tag_configure('oddrow', background='white')
    # scan_tree.tag_configure('evenrow', background='lightblue')
    scan_tree.tag_configure('scanlist', background='lightblue')

    # Scan entry frame
    scan_frame = tk.LabelFrame(tab2, text="Scan Barcode")
    scan_frame.grid(row=1, column=0, sticky='nsew')
    scan_entry = tk.Entry(scan_frame, width=30)
    scan_entry.grid(row=0, column=0, padx=10, pady=10)

    scan_entry.bind("<Return>", scan_tree.scan_check)

    # Commands frame + buttons
    scan_command_label_frame = tk.LabelFrame(tab2, text='Commands')
    scan_command_label_frame.grid(row=2, column=0, sticky='nsew')

    remove_scan = tk.Button(scan_command_label_frame, text='Remove Product', command=scan_tree.s_remove_selected)
    remove_scan.grid(row=0, column=0, padx=(10, 10), pady=(10, 10))
    reset_scan = tk.Button(scan_command_label_frame, text='Reset Scan Session', command=scan_tree.reset)
    reset_scan.grid(row=0, column=1, padx=(10, 10), pady=(10, 10))
    sort_scan = tk.Button(scan_command_label_frame, text='Sort', command=scan_tree.sort)
    sort_scan.grid(row=0, column=2, padx=(10, 10), pady=(10, 10))
    save_scan = tk.Button(scan_command_label_frame, text='Save Scanned Items', command=scan_tree.save_session)
    save_scan.grid(row=0, column=3, padx=(10, 10), pady=(10, 10))

    # Unknown items listbox #
    unknown_listbox = tk.Listbox(tab3, background='#D3D3D3', foreground='black',
                                 activestyle='none', font=20)
    unknown_listbox.grid(row=0, column=0, sticky='nsew')

    # Scrollbar
    unknown_scrollbar = tk.Scrollbar(tab3, orient=tk.VERTICAL, command=unknown_listbox.yview)
    unknown_scrollbar.grid(row=0, column=0, sticky='nse')
    unknown_listbox['yscrollcommand'] = unknown_scrollbar.set

    # Commands frame + buttons
    command_label_frame_tab3 = tk.LabelFrame(tab3, text='Commands')
    command_label_frame_tab3.grid(row=1, column=0, sticky='nsew')

    remove_unknown = tk.Button(command_label_frame_tab3, text='Remove Product', command=remove_selected)
    remove_unknown.grid(row=0, column=0, padx=(10, 10), pady=(10, 10))
    save_unknown = tk.Button(command_label_frame_tab3, text='Save Unknown Items', command=save_unknown)
    save_unknown.grid(row=0, column=1, padx=(10, 10), pady=(10, 10))

    # Balance Tree #
    balance_tree = Balancetree(tab4, connection=conn, table='products', selectmode=tk.EXTENDED)
    balance_tree.grid(row=0, column=0, sticky='nsew', pady=10)
    balance_tree['columns'] = ('Item', 'Barcode', 'Item Description', 'Size', 'Quantity', 'Scanned Quantity',
                               'Quantity Difference')

    balance_tree.tag_configure('balancelist', background='lightblue')

    # Configuring Columns
    balance_tree.column('#0', width=0, stretch='no')  # Ghost column
    balance_tree.column('Item', anchor='w', width=100)
    balance_tree.column('Barcode', anchor='w', width=100)
    balance_tree.column('Item Description', anchor='w', width=100)
    balance_tree.column('Size', anchor='center', width=100)
    balance_tree.column('Quantity', anchor='center', width=100)
    balance_tree.column('Scanned Quantity', anchor='center', width=100)
    balance_tree.column('Quantity Difference', anchor='center', width=100)

    balance_tree.heading('Item', text='Item', anchor='w')
    balance_tree.heading('Barcode', text='Barcode', anchor='w')
    balance_tree.heading('Item Description', text='Item Description', anchor='w')
    balance_tree.heading('Size', text='Size', anchor='center')
    balance_tree.heading('Quantity', text='Quantity', anchor='center')
    balance_tree.heading('Scanned Quantity', text='Scanned Quantity', anchor='center')
    balance_tree.heading('Quantity Difference', text='Quantity Difference', anchor='center')

    # Entry Frame + Boxes
    balance_products_labelframe = tk.LabelFrame(tab4, text='Product')
    balance_products_labelframe.grid(row=1, column=0, sticky='nsew')

    # Create product entry boxes
    balance_item_entry = tk.Entry(balance_products_labelframe, width=25)
    balance_item_entry.grid(row=1, column=0, padx=(10, 10), pady=(5, 10))
    balance_barcode_entry = tk.Entry(balance_products_labelframe, width=25)
    balance_barcode_entry.grid(row=1, column=1, padx=(10, 10), pady=(5, 10))
    balance_item_description_entry = tk.Entry(balance_products_labelframe, width=70)
    balance_item_description_entry.grid(row=1, column=2, padx=(10, 10), pady=(5, 10))
    balance_size_entry = tk.Entry(balance_products_labelframe, width=10)
    balance_size_entry.grid(row=1, column=3, padx=(10, 10), pady=(5, 10))
    balance_quantity_entry = tk.Entry(balance_products_labelframe, width=20)
    balance_quantity_entry.grid(row=1, column=4, padx=(10, 5), pady=(5, 10))
    balance_scanned_entry = tk.Entry(balance_products_labelframe, width=20)
    balance_scanned_entry.grid(row=1, column=5, padx=(10, 10), pady=(5, 10))
    balance_quantity_difference_entry = tk.Entry(balance_products_labelframe, width=20)
    balance_quantity_difference_entry.grid(row=1, column=6, padx=(10, 10), pady=(5, 10))

    # Create entry box labels
    balance_item_label = tk.Label(balance_products_labelframe, text='Item')
    balance_item_label.grid(row=0, column=0, padx=(10, 10), pady=10)
    balance_barcode_label = tk.Label(balance_products_labelframe, text='Barcode')
    balance_barcode_label.grid(row=0, column=1, padx=(10, 10), pady=10)
    balance_item_description_label = tk.Label(balance_products_labelframe, text='Item Description')
    balance_item_description_label.grid(row=0, column=2, padx=(10, 10), pady=10)
    balance_size_label = tk.Label(balance_products_labelframe, text='Size')
    balance_size_label.grid(row=0, column=3, padx=(10, 10), pady=10)
    balance_quantity_label = tk.Label(balance_products_labelframe, text='Quantity')
    balance_quantity_label.grid(row=0, column=4, padx=(10, 10), pady=10)
    balance_scanned_label = tk.Label(balance_products_labelframe, text='Scanned Quantity')
    balance_scanned_label.grid(row=0, column=5, padx=(10, 10), pady=10)
    balance_quantity_difference = tk.Label(balance_products_labelframe, text='Quantity Difference')
    balance_quantity_difference.grid(row=0, column=6, padx=(10, 10), pady=10)

    # Command Frame + Buttons
    balance_command_labelframe = tk.LabelFrame(tab4, text='Commands')
    balance_command_labelframe.grid(row=2, column=0, sticky='nsew')

    balance_button = tk.Button(balance_command_labelframe, text='Balance', width=30, command=balance_tree.balance)
    balance_button.grid(row=0, column=8, padx=(100, 10), pady=(10, 10))
    balance_update_product = tk.Button(balance_command_labelframe, text='Update Product',
                                       command=balance_tree.b_update_product)
    balance_update_product.grid(row=0, column=0, padx=(10, 10), pady=10)
    balance_remove_product = tk.Button(balance_command_labelframe, text='Remove Product',
                                       command=balance_tree.b_remove_selected)
    balance_remove_product.grid(row=0, column=2, padx=(10, 10), pady=10)
    balance_remove_all_products = tk.Button(balance_command_labelframe, text='Remove All Products',
                                            command=balance_tree.b_remove_all)
    balance_remove_all_products.grid(row=0, column=3, padx=(10, 10), pady=10)
    balance_move_up_product = tk.Button(balance_command_labelframe, text='Move Up', command=balance_tree.move_up)
    balance_move_up_product.grid(row=0, column=4, padx=(10, 10), pady=10)
    balance_move_down_product = tk.Button(balance_command_labelframe, text='Move Down', command=balance_tree.move_down)
    balance_move_down_product.grid(row=0, column=5, padx=(10, 10), pady=10)
    balance_clear_product = tk.Button(balance_command_labelframe, text='Clear Entry Boxes',
                                      command=balance_tree.b_clear_entries)
    balance_clear_product.grid(row=0, column=6, padx=(10, 10), pady=10)
    balance_sort = tk.Button(balance_command_labelframe, text='Sort', command=balance_tree.b_sort)
    balance_sort.grid(row=0, column=7, padx=(10, 10), pady=(10, 10))

    root.mainloop()
    conn.close()
