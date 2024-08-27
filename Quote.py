import os
import paper1
import paper2sty
import Fun
import tkinter as tk
from tkinter import *
import tkinter.ttk as ttk
import tkinter.font
from tkinter import filedialog
import pandas as pd
import pypinyin
import xlwings as xw
from tkinter import messagebox

# Project Name: Excel Quote Integrator
# File: Quote.py
# Author: HarryYoi
# Email: luhuaqing0@gmail.com
# Copyright (c) 2024 HarryYoi. All rights reserved.


# Load the Excel file
folder_path = 'data_folder'
file_product_data = 'Products_data.xlsx'
file_client_data = 'Client_data.xlsx'

# Create the folder if it doesn't exist
if not os.path.exists(folder_path):
    os.makedirs(folder_path)
    print(f"Folder '{folder_path}' created. Please place '{file_product_data}' and '{file_client_data}' in this folder.")

# Define the full file paths
product_data_path = os.path.join(folder_path, file_product_data)
client_data_path = os.path.join(folder_path, file_client_data)

# Check if the files exist in the folder
if not os.path.exists(product_data_path) or not os.path.exists(client_data_path):
    raise FileNotFoundError(f"Please place '{file_product_data}' and '{file_client_data}' in the '{folder_path}' folder.")

# Load the Excel files
data = pd.read_excel(product_data_path)
data_client = pd.read_excel(client_data_path)

folder_path2 = 'pastdata'
if not os.path.exists(folder_path2):
    os.makedirs(folder_path2)

# Function to convert Chinese characters to Pinyin


class AutoCompleteCombobox(ttk.Combobox):
    def set_completion_list(self, completion_list):
        self._completion_list = sorted(completion_list)

        self._hits = []
        self._hit_index = 0
        self.position = 0
        self.bind('<KeyRelease>', self.handle_keyrelease)
        self.bind('<Return>', self.on_enter)
        self['values'] = self._completion_list  # Set the combobox values

    def autocomplete(self, delta=0):
        if delta:
            self.delete(self.position, tk.END)
        else:
            self.position = len(self.get())
        _hits = []


        if _hits != self._hits:
            self._hit_index = 0
            self._hits = _hits
        if _hits:
            self._hit_index = (self._hit_index + delta) % len(_hits)
            self.delete(0, tk.END)
            self.insert(0, _hits[self._hit_index])
            self.select_range(self.position, tk.END)

    def handle_keyrelease(self, event):
        if event.keysym in ('BackSpace', 'Left', 'Right', 'Up', 'Down'):
            return
        self.autocomplete()

    def on_enter(self, event):
        self.event_generate("<<ComboboxSelected>>")

class Project1:
    def __init__(self, root, isTKroot=True, params=None):
        uiName = Fun.GetUIName(root, self.__class__.__name__)
        self.uiName = uiName
        Fun.Register(uiName, 'UIClass', self)
        self.root = root
        self.isTKroot = isTKroot
        self.firstRun = True

        Fun.G_UIParamsDictionary[uiName] = params
        Fun.G_UICommandDictionary[uiName] = paper1
        Fun.Register(uiName, 'root', root)
        style = paper2sty.SetupStyle(isTKroot)
        if isTKroot:
            Fun.SetTitleBar(root, titleText='Quote Information Panel', isDarkMode=False)
            Fun.CenterDlg(uiName, root, 1980, 1060)
            Fun.WindowDraggable(root, False, 0, '#ffffff')
            root['background'] = '#FFFFFF'
        Form_1 = tk.Frame(root)
        Form_1.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        Form_1.configure(bg="#FFFFFF")
        Fun.Register(uiName, 'Form_1', Form_1)
        Fun.SetUIRootSize(uiName, 1980, 1060)
        Group_1_Variable = Fun.AddTKVariable(uiName, 'Group_1')
        Group_1_Variable.set(1)

        self.excel_data = []

        footer_frame = tk.Frame(root, bd=1, relief=tk.SUNKEN)
        footer_frame.pack(side=tk.BOTTOM, fill=tk.X)

        # Add the project information label
        footer_label = tk.Label(
            footer_frame,
            text="Excel Quote Integrator | © 2024 HarryYoi | Contact: luhuaqing0@gmail.com",
            font=("Helvetica", 10),
            bd=1,
            relief=tk.FLAT
        )
        footer_label.pack(side=tk.LEFT, padx=10, pady=5)

        # Create the title label at the top
        Label_1 = tk.Label(Form_1, text="Title customization", anchor='center', justify='center')
        Fun.Register(uiName, 'PyMe_Label_1', Label_1, 'Label_1')
        Label_1.pack(side=tk.TOP, pady=20, fill=tk.X)
        Label_1.configure(bg="#EFEFEF")
        Label_1.configure(fg="#000000")
        Label_1_Ft = tk.font.Font(family='微软雅黑', size=25, weight='bold', slant='roman', underline=0, overstrike=0)
        Label_1.configure(font=Label_1_Ft)

        # Create and place the export button in the bottom right corner
        button_font2 = tk.font.Font(family='微软雅黑', size=20, weight='bold', slant='roman', underline=0, overstrike=0)
        export_button = tk.Button(Form_1, text="Export Final Excel",  command=self.export_combined_excel,font=button_font2, fg='blue')
        Fun.Register(uiName, 'ExportButton', export_button)
        export_button.place(x=1420, y=500, width=350, height=50)

        label_font = tk.font.Font(family='Microsoft YaHei UI', size=22, weight='normal', slant='roman', underline=0, overstrike=0)

        Label_2 = tk.Label(Form_1, text="Model")
        Fun.Register(uiName, 'PyMe_Label_2', Label_2, 'Label_18')
        Fun.SetControlPlace(uiName, 'PyMe_Label_2', 80, 100, 200, 50, 'nw', True, True)
        Label_2.configure(bg="#FFFFFF")
        Label_2.configure(fg="#000000")
        Label_2.configure(font=label_font)

        Label_3 = tk.Label(Form_1, text="Product Name")
        Fun.Register(uiName, 'PyMe_Label_3', Label_3, 'Label_19')
        Fun.SetControlPlace(uiName, 'PyMe_Label_3', 80, 150, 200, 50, 'nw', True, True)
        Label_3.configure(bg="#FFFFFF")
        Label_3.configure(fg="#000000")
        Label_3.configure(font=label_font)

        Label_4 = tk.Label(Form_1, text="Product Code")
        Fun.Register(uiName, 'PyMe_Label_4', Label_4, 'Label_20')
        Fun.SetControlPlace(uiName, 'PyMe_Label_4', 80, 200, 200, 50, 'nw', True, True)
        Label_4.configure(bg="#FFFFFF")
        Label_4.configure(fg="#000000")
        Label_4.configure(font=label_font)

        Label_5 = tk.Label(Form_1, text="Brand")
        Fun.Register(uiName, 'PyMe_Label_5', Label_5, 'Label_21')
        Fun.SetControlPlace(uiName, 'PyMe_Label_5', 80, 250, 200, 50, 'nw', True, True)
        Label_5.configure(bg="#FFFFFF")
        Label_5.configure(fg="#000000")
        Label_5.configure(font=label_font)

        Label_6 = tk.Label(Form_1, text="Unit")
        Fun.Register(uiName, 'PyMe_Label_6', Label_6, 'Label_21')
        Fun.SetControlPlace(uiName, 'PyMe_Label_6', 80, 300, 200, 50, 'nw', True, True)
        Label_6.configure(bg="#FFFFFF")
        Label_6.configure(fg="#000000")
        Label_6.configure(font=label_font)

        # 客户表格
        Label_7 = tk.Label(Form_1, text="Customer Full Name")
        Fun.Register(uiName, 'PyMe_Label_7', Label_7, 'Label_23')
        Fun.SetControlPlace(uiName, 'PyMe_Label_7', 1100, 100, 300, 50, 'nw', True, True)
        Label_7.configure(bg="#FFFFFF")
        Label_7.configure(fg="#000000")
        Label_7.configure(font=label_font)

        Label_8 = tk.Label(Form_1, text="Customer Number")
        Fun.Register(uiName, 'PyMe_Label_8', Label_8, 'Label_24')
        Fun.SetControlPlace(uiName, 'PyMe_Label_8', 1100, 150, 300, 50, 'nw', True, True)
        Label_8.configure(bg="#FFFFFF")
        Label_8.configure(fg="#000000")
        Label_8.configure(font=label_font)

        Label_9 = tk.Label(Form_1, text="Address")
        Fun.Register(uiName, 'PyMe_Label_9', Label_9, 'Label_25')
        Fun.SetControlPlace(uiName, 'PyMe_Label_9', 1100, 200, 300, 50, 'nw', True, True)
        Label_9.configure(bg="#FFFFFF")
        Label_9.configure(fg="#000000")
        Label_9.configure(font=label_font)

        Label_10 = tk.Label(Form_1, text="Telephone Number")
        Fun.Register(uiName, 'PyMe_Label_10', Label_10, 'Label_26')
        Fun.SetControlPlace(uiName, 'PyMe_Label_10', 1100, 250, 300, 50, 'nw', True, True)
        Label_10.configure(bg="#FFFFFF")
        Label_10.configure(fg="#000000")
        Label_10.configure(font=label_font)

        Label_11 = tk.Label(Form_1, text="Contact")
        Fun.Register(uiName, 'PyMe_Label_11', Label_11, 'Label_27')
        Fun.SetControlPlace(uiName, 'PyMe_Label_11', 1100, 300, 300, 50, 'nw', True, True)
        Label_11.configure(bg="#FFFFFF")
        Label_11.configure(fg="#000000")
        Label_11.configure(font=label_font)



        def on_select_client(event):
            selected_name = event.widget.get()
            selected_rows = data_client.loc[data_client['Customer Full Name'] == selected_name]

            if not selected_rows.empty:
                selected_row = selected_rows.iloc[0]
                self.client_code_var.set(selected_row['Customer Number'])
                self.address_entry_var.set(selected_row['Address'])
                self.telephone_entry_var.set(selected_row['Telephone Number'])
                self.name_entry_var.set(selected_row['Contact'])


        def on_select_product(event):
            selected_name = event.widget.get()
            selected_rows = data[data['Model'] == selected_name]

            if not selected_rows.empty:
                selected_row = selected_rows.iloc[0]
                self.brand_name_var.set(selected_row['Product Name'])
                self.model_entry_var.set(selected_row['Product Code'])
                self.brand_entry_var.set(selected_row['Brand'])
                self.unit_entry_var.set(selected_row['Unit'])

        # Define the monospace font


        # Adjust the height and width of the entries
        entry_height = 40
        entry_width = 650
        entry_width2 = 450

        # Autocomplete Entry for products
        self.product_entry_var = tk.StringVar()
        self.autocomplete_product_entry = AutoCompleteCombobox(Form_1, textvariable=self.product_entry_var, font=label_font)
        self.autocomplete_product_entry.set_completion_list(data['Model'].tolist())
        self.autocomplete_product_entry.place(x=300, y=100, width=entry_width, height=entry_height)

        self.autocomplete_product_entry.bind("<<ComboboxSelected>>", on_select_product)

        # 商品名称 Entry
        self.brand_name_var = tk.StringVar()
        self.brand_name = ttk.Entry(Form_1, textvariable=self.brand_name_var, font=label_font)
        self.brand_name.place(x=300, y=150, width=entry_width, height=entry_height)

        # 型号 Entry
        self.model_entry_var = tk.StringVar()
        self.model_entry = ttk.Entry(Form_1, textvariable=self.model_entry_var, font=label_font)
        self.model_entry.place(x=300, y=200, width=entry_width, height=entry_height)

        # 品牌 Entry
        self.brand_entry_var = tk.StringVar()
        self.brand_entry = ttk.Entry(Form_1, textvariable=self.brand_entry_var, font=label_font)
        self.brand_entry.place(x=300, y=250, width=entry_width, height=entry_height)

        # 单位 Entry
        self.unit_entry_var = tk.StringVar()
        self.unit_entry = ttk.Entry(Form_1, textvariable=self.unit_entry_var, font=label_font)
        self.unit_entry.place(x=300, y=300, width=entry_width, height=entry_height)

        # Autocomplete Entry for clients
        self.client_entry_var = tk.StringVar()
        self.autocomplete_client_entry = AutoCompleteCombobox(Form_1, textvariable=self.client_entry_var, font=label_font)
        self.autocomplete_client_entry.set_completion_list(data_client['Customer Full Name'].tolist())
        self.autocomplete_client_entry.place(x=1420, y=100, width=entry_width2, height=entry_height)
        self.autocomplete_client_entry.bind("<<ComboboxSelected>>", on_select_client)

        # 客户编号 Entry
        self.client_code_var = tk.StringVar()
        self.client_code = ttk.Entry(Form_1, textvariable=self.client_code_var, font=label_font)
        self.client_code.place(x=1420, y=150, width=entry_width2, height=entry_height)

        # 单位地址 Entry
        self.address_entry_var = tk.StringVar()
        self.address_entry = ttk.Entry(Form_1, textvariable=self.address_entry_var, font=label_font)
        self.address_entry.place(x=1420, y=200, width=entry_width2, height=entry_height)

        # 电话 Entry
        self.telephone_entry_var = tk.StringVar()
        self.telephone_entry = ttk.Entry(Form_1, textvariable=self.telephone_entry_var, font=label_font)
        self.telephone_entry.place(x=1420, y=250, width=entry_width2, height=entry_height)

        # 联系人 Entry
        self.name_entry_var = tk.StringVar()
        self.name_entry = ttk.Entry(Form_1, textvariable=self.name_entry_var, font=label_font)
        self.name_entry.place(x=1420, y=300, width=entry_width2, height=entry_height)


        button_font = tk.font.Font(family='微软雅黑', size=20, weight='bold', slant='roman', underline=0, overstrike=0)

        # Create the "单独导出客户信息" button and place it below the entries
        input_button2 = tk.Button(Form_1, text="Export customer information individually", command=self.record_data2, font=button_font)
        Fun.Register(uiName, 'InputButton2', input_button2)
        input_button2.place(x=1100, y=350, width=700, height=50)

        # "导入报价单模板" button
        import_button = tk.Button(Form_1, text="Import quotation template", command=self.import_excel, font=button_font)
        Fun.Register(uiName, 'ImportButton', import_button)
        import_button.place(x=300, y=490, width=300, height=50)

        # "价格输入" button
        price_input_button = tk.Button(Form_1, text="Price Input", command=self.open_price_input_window, font=button_font)
        Fun.Register(uiName, 'PriceInputButton', price_input_button)
        price_input_button.place(x=630, y=420, width=300, height=50)

        # Initialize a list to store the recorded data
        self.recorded_data = []

        # Listbox to show recorded data
        self.listbox = Listbox(Form_1, font=("微软雅黑", 18))
        self.listbox.place(x=70, y=560, width=1300, height=400)


        # "使用说明" button to record data
        confirm_button = tk.Button(Form_1, text="Instructions", command=self.read_reference)
        Fun.Register(uiName, 'ConfirmButton', confirm_button)
        confirm_button.place(x=50, y=30, width=70, height=40)

        # "确定" button to record data
        confirm_button = tk.Button(Form_1, text="Add to", command=self.record_data, font=button_font)
        Fun.Register(uiName, 'ConfirmButton', confirm_button)
        confirm_button.place(x=300, y=350, width=300, height=50)

        # "查询往期报价" button to record data
        confirm_button = tk.Button(Form_1, text="Check past quotations", command=self.window_review, font=button_font)
        Fun.Register(uiName, 'ConfirmButton', confirm_button)
        confirm_button.place(x=1420, y=450, width=350, height=50)

        # "删除" button to remove selected data
        delete_button = tk.Button(Form_1, text="Delete", command=self.delete_data, font=button_font)
        Fun.Register(uiName, 'DeleteButton', delete_button)
        delete_button.place(x=630, y=350, width=300, height=50)

        # "单独导出商品信息" button to export all recorded data to Excel
        input_button = tk.Button(Form_1, text="Export product information separately", command=self.export_recorded_data, font=button_font)
        input_button.place(x=50, y=420, width=550, height=50)

        # Initialize the checkboxes for selecting fields to include in the export
        self.include_item_code_var = tk.BooleanVar(value=True)
        self.include_item_name_var = tk.BooleanVar(value=True)
        self.include_model_var = tk.BooleanVar(value=True)
        self.include_brand_var = tk.BooleanVar(value=True)
        self.include_unit_var = tk.BooleanVar(value=True)

        self.include_item_code_checkbox = tk.Checkbutton(Form_1, variable=self.include_model_var)
        self.include_item_code_checkbox.place(x=1000, y=100, width=20, height=30)

        self.include_item_name_checkbox = tk.Checkbutton(Form_1, variable=self.include_item_name_var)
        self.include_item_name_checkbox.place(x=1000, y=150, width=20, height=30)

        self.include_model_checkbox = tk.Checkbutton(Form_1, variable=self.include_item_code_var)
        self.include_model_checkbox.place(x=1000, y=200, width=20, height=30)

        self.include_brand_checkbox = tk.Checkbutton(Form_1, variable=self.include_brand_var)
        self.include_brand_checkbox.place(x=1000, y=250, width=20, height=30)

        self.include_unit_checkbox = tk.Checkbutton(Form_1, variable=self.include_unit_var)
        self.include_unit_checkbox.place(x=1000, y=300, width=20, height=30)

        # "导入报价单模板" button
        import_button = tk.Button(Form_1, text="Import quotation template", command=self.import_excel, font=button_font)
        Fun.Register(uiName, 'ImportButton', import_button)
        import_button.place(x=300, y=490, width=400, height=50)

    def import_excel(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if self.file_path:
            try:
                self.imported_data = pd.read_excel(self.file_path)
                print(f"File imported successfully: {self.file_path}")
            except Exception as e:
                print(f"Error importing file: {e}")

    def export_recorded_data(self):
        if not self.recorded_data:
            messagebox.showinfo("Info", "No data to export")
            return

        columns_to_include = []
        if self.include_item_code_var.get():
            columns_to_include.append('Product Code')
        if self.include_item_name_var.get():
            columns_to_include.append('Product Name')
        if self.include_model_var.get():
            columns_to_include.append('Model')
        if self.include_brand_var.get():
            columns_to_include.append('Brand')
        if self.include_unit_var.get():
            columns_to_include.append('Unit')

        df_to_export = pd.DataFrame(self.recorded_data, columns=columns_to_include)
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                 filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if file_path:
            df_to_export.to_excel(file_path, index=False)
            messagebox.showinfo("Info", f"Data exported to {file_path}")

    def export_combined_excel(self):
        if not hasattr(self, 'imported_data'):
            print("No template imported.")
            return

        if not self.recorded_data:
            print("No data to export.")
            return


        # Initialize xlwings
        try:
            app = xw.App(visible=False)
            wb = app.books.open(self.file_path)
            sheet = wb.sheets[0]

            # Find the header row
            header_row_index = 14  # Assuming the header row is at row 14 (0-indexed)
            header_row = sheet.range(f'A{header_row_index + 1}:Z{header_row_index + 1}').value

            selected_name = self.client_entry_var.get()
            selected_rows = data_client[data_client['Customer Full Name'] == selected_name]

            if not selected_rows.empty:
                selected_row = selected_rows.iloc[0]
                app = xw.App(visible=False)
                wb = app.books.open(self.file_path)
                sheet = wb.sheets[0]

                # Assume the template has specific cells for each piece of client information
                sheet.range("C9").value = selected_row['Customer Full Name']
                sheet.range("C14").value = selected_row['Address']
                sheet.range("C11").value = selected_row['Telephone number']
                sheet.range("C10").value = selected_row['Contact']

            # Get the column indices based on header names
            columns = {
               # 'Product Code': header_row.index('Product Code'),
                'Product Name': header_row.index('Product Name'),
                'Model': header_row.index('Model'),
               # 'Brand': header_row.index('Brand'),
                'Unit': header_row.index('Unit'),
                'Quantity': header_row.index('Quantity') if 'Quantity' in header_row else None,
                'Unit price': header_row.index('Unit price') if 'Unit price' in header_row else None,
                'Period': header_row.index('Period') if 'Period' in header_row else None,
                'Total price': header_row.index('Total price') if 'Total price' in header_row else None,
            }

            # Insert data into the template while preserving formatting
            start_row = header_row_index + 2
            for i, row in enumerate(self.recorded_data):
                # sheet.range((start_row + i, columns['Product Code'] + 1)).value = row.get('Product Code', '')
                sheet.range((start_row + i, columns['Product Name'] + 1)).value = row.get('Product Name', '')
                sheet.range((start_row + i, columns['Model'] + 1)).value = row.get('Model', '')
                # sheet.range((start_row + i, columns['Brand'] + 1)).value = row.get('品牌', '')
                sheet.range((start_row + i, columns['Unit'] + 1)).value = row.get('Unit', '')

                # Insert additional price-related data if available
                if i in self.price_data:
                    price_data = self.price_data[i]
                    if columns['Quantity'] is not None:
                        sheet.range((start_row + i, columns['Quantity'] + 1)).value = price_data.get('数量', '')
                    if columns['Unit price'] is not None:
                        sheet.range((start_row + i, columns['Unit price'] + 1)).value = price_data.get('单价', '')
                    if columns['Period'] is not None:
                        sheet.range((start_row + i, columns['Period'] + 1)).value = price_data.get('周期', '')
                    if columns['Total price'] is not None:
                        sheet.range((start_row + i, columns['Total price'] + 1)).value = price_data.get('总价', '')

            # Save the modified file
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                     filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
            if save_path:
                wb.save(save_path)
                print(f"Modified file saved to: {save_path}")

            # Close the workbook and quit xlwings
            wb.close()
            app.quit()
        except Exception as e:
            print(f"Error while exporting: {e}")

    def record_data(self):
        data_to_save = {}
        display_text = []

        if self.include_item_code_var.get():
            data_to_save['Product Code'] = self.model_entry_var.get()
            display_text.append(f"Product Code: {self.model_entry_var.get()}")
        if self.include_item_name_var.get():
            data_to_save['Product Name'] = self.brand_name_var.get()
            display_text.append(f"Product Name: {self.brand_name_var.get()}")
        if self.include_model_var.get():
            data_to_save['Model'] = self.product_entry_var.get()
            display_text.append(f"Model: {self.product_entry_var.get()}")
        if self.include_brand_var.get():
            data_to_save['Brand'] = self.brand_entry_var.get()
            display_text.append(f"Brand: {self.brand_entry_var.get()}")
        if self.include_unit_var.get():
            data_to_save['Unit'] = self.unit_entry_var.get()
            display_text.append(f"Unit: {self.unit_entry_var.get()}")

        self.recorded_data.append(data_to_save)
        self.listbox.insert(END, ", ".join(display_text))
        print(f"Data recorded: {data_to_save}")

    def delete_data(self):
        selected_index = self.listbox.curselection()
        if selected_index:
            self.listbox.delete(selected_index)
            del self.recorded_data[selected_index[0]]
            print(f"Data at index {selected_index[0]} deleted")

    def record_data2(self):
        selected_name = self.autocomplete_client_entry.get()
        selected_rows = data_client[data_client['Customer Full Name'] == selected_name]

        if not selected_rows.empty:
            selected_row = selected_rows.iloc[0]
            data_to_save = {
                'Customer Full Name': [selected_row['Customer Full Name']],
                #'Customer Number': [selected_row['Customer Number']],
                'Address': [selected_row['Address']],
                'Telephone number': [selected_row['Telephone number']],
                'Contact': [selected_row['Contact']]
            }
            df_to_save = pd.DataFrame(data_to_save)
            file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                     filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
            if file_path:
                df_to_save.to_excel(file_path, index=False)
                print(f"Data recorded to {file_path}")

    def window_review(self): #Query function
        new_window = tk.Toplevel()
        new_window.title('Previous Quote Query')
        new_window.geometry('1980x1080')

        Form_2 = tk.Frame(new_window)
        Form_2.pack(fill="both", expand=True)

        label_font = tk.font.Font(family='Microsoft YaHei UI', size=22, weight='normal', slant='roman', underline=0,
                                  overstrike=0)

        self.Label_2 = tk.Label(Form_2, text="Customer Full Name:")
        self.Label_2.place(x=70, y=50, width=200, height=50)
        self.Label_2.configure(bg="#FFFFFF")
        self.Label_2.configure(fg="#000000")
        self.Label_2.configure(font=label_font)

        self.entry_2 = AutoCompleteCombobox(Form_2, font=label_font)
        self.entry_2.set_completion_list(data_client['Customer Full Name'].tolist())
        self.entry_2.place(x=300, y=50, width=300, height=50)
        self.entry_2.configure(font=label_font)

        self.Label_3 = tk.Label(Form_2, text="Model:")
        self.Label_3.place(x=700, y=50, width=300, height=50)
        self.Label_3.configure(bg="#FFFFFF")
        self.Label_3.configure(fg="#000000")
        self.Label_3.configure(font=label_font)

        self.entry_3 = AutoCompleteCombobox(Form_2, font=label_font)
        self.entry_3.set_completion_list(data['Model'].tolist())
        self.entry_3.place(x=1070, y=50, width=300, height=50)
        self.entry_3.configure(font=label_font)

        self.Label_4 = tk.Label(Form_2, text="Date:")
        self.Label_4.place(x=700, y=120, width=300, height=50)
        self.Label_4.configure(bg="#FFFFFF")
        self.Label_4.configure(fg="#000000")
        self.Label_4.configure(font=label_font)

        self.entry_4 = AutoCompleteCombobox(Form_2, font=label_font)
        self.entry_4.place(x=1070, y=120, width=300, height=50)
        self.entry_4.configure(font=label_font)

        read_button = tk.Button(Form_2, text="Reading Data", command=self.read_excel_files)
        read_button.place(x=70, y=150, width=200, height=50)
        read_button.configure(font=label_font)

        self.result_listbox = tk.Listbox(Form_2, font=label_font)
        self.result_listbox.place(x=70, y=250, width=1840, height=700)
        self.result_listbox.configure(font=label_font)

        # Add the export button
        export_button = tk.Button(Form_2, text="Export to excel", command=self.export_to_excel)
        export_button.place(x=300, y=150, width=200, height=50)
        export_button.configure(font=label_font)

        self.result_listbox.bind("<Double-Button-1>", self.open_excel_file)  # Bind double-click event

    def read_excel_files(self):
        directory = 'pastdata'
        customer_name = self.entry_2.get().strip()
        selected_date = self.entry_4.get().strip()
        model = self.entry_3.get().strip()
        self.excel_data = []
        self.dates = set()

        for filename in os.listdir(directory):
            if filename.startswith('~$'):
                continue
            if filename.endswith('.xlsx'):
                filepath = os.path.join(directory, filename)
                data = pd.read_excel(filepath, header=None)

                cell_value_c9 = str(data.iat[8, 2])  # Get value from cell C9 (index 8, 2)
                data_date = str(data.iat[9, 6])  # Get value from cell G11 (index 10, 6)
                self.dates.add(data_date)  # Collect unique dates

                # Check if the customer name matches
                if (not customer_name or cell_value_c9 == customer_name) and (
                        not selected_date or data_date == selected_date):
                    for row in range(15, len(data)):
                        cell_value_c = str(data.iat[row, 2])  # Get value from the C column
                        # Skip rows where cell in column C is empty
                        if pd.isna(cell_value_c) or cell_value_c.strip() == 'nan' or cell_value_c.strip() == '合计':
                            continue

                        data_e = str(data.iat[row, 4])  # Get value from cell E in the same row
                        data_f = str(data.iat[row, 5])  # Get value from cell F in the same row
                        data_g = str(data.iat[row, 6])  # Get value from cell G in the same row

                        # Check if model is empty or matches the cell value, and ignore rows with any 'nan' values
                        if (not model or cell_value_c == model) and all(
                                [data_date, cell_value_c, data_e, data_f, data_g]) and not any(
                                pd.isna([data_date, cell_value_c, data_e, data_f, data_g])):
                            self.excel_data.append((filepath, data_date, cell_value_c, data_e, data_f, data_g))

        self.result_listbox.delete(0, tk.END)  # Clear the listbox
        if self.excel_data:
            for filepath, data_date, cell_value_c, data_e, data_f, data_g in self.excel_data:
                self.result_listbox.insert(tk.END,
                                           f"'Date': {data_date}, 'Model': {cell_value_c}, 'Transaction Quantity': {data_e}, 'Quote of the day': {data_f}, 'Total price for the day': {data_g}")
        else:
            self.result_listbox.insert(tk.END, "No matching data found")

        self.update_date_combobox()

    def update_date_combobox(self):
        self.entry_4.set_completion_list(sorted(self.dates))

    def open_excel_file(self, event):
        selection = self.result_listbox.curselection()
        if selection:
            index = selection[0]
            filepath, _, _, _, _, _ = self.excel_data[index]
            if os.path.exists(filepath):
                os.startfile(filepath)
            else:
                messagebox.showerror("Error", f"File not found: {filepath}")

    def export_to_excel(self):
        items = self.result_listbox.get(0, tk.END)  # Get all items from the listbox
        if not items:
            messagebox.showinfo("Info", "No data to export")
            return

        # Convert listbox items to a dataframe
        data = []
        for item in items:
            parts = item.split(', ')
            row = {}
            for part in parts:
                key, value = part.split(': ')
                key = key.strip("'")
                row[key] = value
            data.append(row)

        df = pd.DataFrame(data)

        # Save the dataframe to an Excel file
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                 filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if file_path:
            df.to_excel(file_path, index=False)
            messagebox.showinfo("Info", f"Data exported to {file_path}")

    def read_reference(self):
        # Create a new window
        new_window = tk.Toplevel()
        new_window.title("Instructions")
        new_window.geometry("1000x500")

        # Add a label with the message
        label = tk.Label(new_window,
                         text="1.'Products_data' and 'Client_data' database support formats: Excel (xlsx, csv). If the database is updated later, just replace the original file in the data_folder folder.  2. Enter 'Product Number' to automatically find 'Product Name', 'Model', 'Brand', 'Unit'. After checking, click 'Add' to record it in the storage box on the right. Select the selected item in the box and click 'Delete' to remove the item.  3. Product number and customer full name (support pinyin first letter) can be manually searched.  4. If you want to generate a quotation, you need to click the 'Import Quotation Template' button to import the template. (The final Excel format is consistent with the template. If you adjust the format, such as centering or font, etc., you can change it in the original template)  5. When using the 'Export Merged Quotation Excel' button, you must only check 'Product Name', 'Model', and 'Unit'.  6. When using the 'Export Product Information Separately' button, just check any required data. (Pre-check before use, do not change it midway).  7. After 'enter the price', press Save to record the price. If you need to change the price, repeat this step to overwrite it.  8. 'Query past quotation function': After entering the window, filter a single specified customer, click Read Data, and you can see all past quotation products and prices sorted by date. After completing the above operations, if you need to filter a specific 'date' or 'model' (can be filtered at the same time), select it on the right and click 'Read Data' again (can be exported to Excel at any step).  9. You can double-click the product you want to view in the filter window, and the source Excel file to which the data belongs will be automatically opened.  10. When you open the 'Query Past Quotations' window for the first time, a pastdata folder will be automatically generated. Please store all past quotations (be sure to confirm that the date has been filled in the table) here for easy retrieval.",
                         wraplength=700, font=("微软雅黑", 14))
        label.pack(pady=20, padx=10)

        # Close window button


    def open_price_input_window(self):
        selected_index = self.listbox.curselection()
        if not selected_index:
            print("No data selected.")
            return

        selected_data = self.recorded_data[selected_index[0]]

        price_window = tk.Toplevel()
        price_window.title("Price input")
        price_window.geometry("600x700")

        # Define font for labels and entries
        label_font = tk.font.Font(family='Microsoft YaHei UI', size=14)
        entry_font = tk.font.Font(family='Microsoft YaHei UI', size=14)
        # Create labels and entries for each field
        row_index = 0
        for key, value in selected_data.items():
            label = tk.Label(price_window, text=key, font=label_font)
            label.grid(row=row_index, column=0, padx=10, pady=10, sticky='e')
            entry = ttk.Entry(price_window, font=entry_font)
            entry.insert(0, value)
            entry.grid(row=row_index, column=1, padx=10, pady=10)
            row_index += 1

        # Add a label and entry for 数量
        quantity_label = tk.Label(price_window, text="Quantity", font=label_font)
        quantity_label.grid(row=row_index, column=0, padx=10, pady=10, sticky='e')
        quantity_entry = ttk.Entry(price_window, font=entry_font)
        quantity_entry.grid(row=row_index, column=1, padx=10, pady=10)
        row_index += 1

        # Add a label and entry for 单价
        price_label = tk.Label(price_window, text="Unit price", font=label_font)
        price_label.grid(row=row_index, column=0, padx=10, pady=10, sticky='e')
        price_entry = ttk.Entry(price_window, font=entry_font)
        price_entry.grid(row=row_index, column=1, padx=10, pady=10)
        row_index += 1

        # Add a label and entry for 周期
        period_label = tk.Label(price_window, text="Period", font=label_font)
        period_label.grid(row=row_index, column=0, padx=10, pady=10, sticky='e')
        period_entry = ttk.Entry(price_window, font=entry_font)
        period_entry.grid(row=row_index, column=1, padx=10, pady=10)
        row_index += 1

        # Add a label and entry for 总价
        total_price_label = tk.Label(price_window, text="Total price", font=label_font)
        total_price_label.grid(row=row_index, column=0, padx=10, pady=10, sticky='e')
        total_price_entry = ttk.Entry(price_window, font=entry_font, state='readonly')
        total_price_entry.grid(row=row_index, column=1, padx=10, pady=10)
        row_index += 1

        def calculate_total_price(*args):
            try:
                quantity = float(quantity_entry.get())
                price = float(price_entry.get())
                total_price = quantity * price
                total_price_entry.config(state='normal')
                total_price_entry.delete(0, tk.END)
                total_price_entry.insert(0, f"{total_price:.2f}")
                total_price_entry.config(state='readonly')
            except ValueError:
                total_price_entry.config(state='normal')
                total_price_entry.delete(0, tk.END)
                total_price_entry.config(state='readonly')

        quantity_entry.bind('<KeyRelease>', calculate_total_price)
        price_entry.bind('<KeyRelease>', calculate_total_price)

        # Initialize a dictionary to store price data
        if not hasattr(self, 'price_data'):
            self.price_data = {}

        # Button to save price
        def save_price():
            entered_quantity = quantity_entry.get()
            entered_price = price_entry.get()
            entered_period = period_entry.get()
            entered_total_price = total_price_entry.get()

            # Store the entered data in the dictionary
            self.price_data[selected_index[0]] = {
                'Quantity': entered_quantity,
                'Unit price': entered_price,
                'Period': entered_period,
                'Total price': entered_total_price
            }

            print(f"Saved data for item {selected_index[0]}: {self.price_data[selected_index[0]]}")
            price_window.destroy()

        save_button = tk.Button(price_window, text="Save", command=save_price, font=label_font)
        save_button.grid(row=row_index + 5, columnspan=2, pady=20)

    # Create the main window


# Create the main window

root = tkinter.Tk()
app = Project1(root)

    # Start the Tkinter event loop
root.mainloop()