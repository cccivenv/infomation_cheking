import tkinter as tk
from tkinter import filedialog, ttk
import pandas as pd
import numpy as np
import xlsxwriter

db = pd.read_csv('1your_file.csv')
orders = pd.DataFrame()  # Initialize an empty DataFrame


def import_csv_data():
    csv_file_path = filedialog.askopenfilename()
    global orders
    orders = pd.read_csv(csv_file_path)
    update_treeview()


def update_treeview():
    for column in treeview['columns']:
        treeview.delete(column)
    treeview['columns'] = list(orders.columns)
    for column in orders.columns:
        treeview.heading(column, text=column)
        treeview.column(column, width=100)
    for index, row in orders.iterrows():
        treeview.insert('', 'end', values=list(row))


def process_orders(orders, evg_base, plt_base):
    new_data_list = []
    summary_data_list = []
    max_boxes = 0
    for index, order in orders.iterrows():
        item = order['ITEM']
        po = order['PO']
        order_qty = order['ORDER QTY']
        item_info = db[db['ITEM'] == item]
        if len(item_info) == 0:
            print(f'No such ITEM: {item}')
            continue
        item_info = item_info.iloc[0]
        cartons = order_qty // item_info['HSU/Carton']
        leftover_hsu = order_qty % item_info['HSU/Carton']
        hsu_carton = item_info['Carton/Pallet']
        full_pallets = cartons // hsu_carton
        leftover_cartons = cartons % hsu_carton
        max_boxes = max(max_boxes, hsu_carton)

        qty_hsu = item_info['Qty/HSU']
        hsu = order_qty
        carton_pallet = item_info['Carton/Pallet']
        pallet = cartons // carton_pallet
        carton_type = item_info['Carton-Type']
        pallet_weight = 18
        total_pallet_weight = pallet_weight * pallet
        qty = order_qty * qty_hsu

        summary_row = {
            "PO": po,
            "ITEM": item,
            "QTY": qty,
            "QTY/HSU": qty_hsu,
            "HSU": hsu,
            "HSU/Carton": item_info['HSU/Carton'],
            "Carton": cartons,
            "Carton/Pallet": carton_pallet,
            "Pallet": pallet,
            "Carton_type": carton_type,
            "Pallet-weight": pallet_weight,
            "棧板總重量": total_pallet_weight,
            "每箱淨重(KG)": item_info['N.W'],  # assuming N.W is the column name in db
            "每箱毛重(KG)": item_info['G.W'],  # assuming G.W is the column name in db
            "產品總重量": item_info['G.W'] * cartons
        }
        summary_data_list.append(summary_row)

        for i in range(full_pallets):
            new_row = {'PO': po, 'ITEM': item, 'PLT': 'EVG' + str(plt_base).zfill(6)}
            plt_base += 1
            for j in range(hsu_carton):
                new_row[str(j)] = 'EVG' + str(evg_base).zfill(7)
                evg_base += 1
            new_data_list.append(new_row)

        if leftover_cartons > 0:
            new_row = {'PO': po, 'ITEM': item, 'PLT': 'Incomplete Pallet'}
            for j in range(leftover_cartons):
                new_row[str(j)] = 'EVG' + str(evg_base).zfill(7)
                evg_base += 1
            for k in range(leftover_cartons, hsu_carton):
                new_row[str(k)] = np.nan
            new_data_list.append(new_row)

        if leftover_hsu > 0:
            new_row = {'PO': po, 'ITEM': item, 'PLT': 'Incomplete Carton', '0': 'EVG' + str(evg_base).zfill(7)}
            for k in range(1, hsu_carton):
                new_row[str(k)] = np.nan
            evg_base += 1
            new_data_list.append(new_row)

    summary_data = pd.DataFrame(summary_data_list)
    new_data = pd.DataFrame(new_data_list)
    # new_data = new_data[['PO', 'ITEM', 'PLT'] + [str(i) for i in range(max_boxes)]]

    new_data.columns = ['PO', 'ITEM', 'PLT'] + [str(i) for i in range(1, max_boxes + 1)]

    return new_data, summary_data


def is_consecutive(series):
    if series.empty or series.isnull().any():
        return False
    sorted_series = series.sort_values().reset_index(drop=True)
    return (sorted_series == range(sorted_series.iloc[0], sorted_series.iloc[-1] + 1)).all()


def export_data():
    evg_base = int(evg_entry.get())
    plt_base = int(plt_entry.get())
    po = po_entry.get()  # 获取输入的PO
    port = port_entry.get()  # 获取输入的港口名字
    new_data, summary_data = process_orders(orders, evg_base, plt_base)
    file_path = filedialog.asksaveasfilename(defaultextension='.xlsx')

    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
        summary_data.to_excel(writer, sheet_name='Summary', index=False)
        new_data.to_excel(writer, sheet_name='Data', index=False)

        pallets = new_data['PLT'].unique()  # 获取所有不重复的棧板号
        for plt in pallets:
            if plt != 'Incomplete Pallet' and plt != 'Incomplete Carton':
                pallet_data = new_data[new_data['PLT'] == plt]
                pallet_data = pallet_data.dropna(axis=1, how='all')  # Add this line
        for plt in pallets:
            if plt != 'Incomplete Pallet' and plt != 'Incomplete Carton':
                pallet_data = new_data[new_data['PLT'] == plt]

                # 提取箱号并转换为整数进行检查
                box_numbers = pallet_data.loc[:, '3':].values.flatten()
                box_numbers = box_numbers[~pd.isnull(box_numbers)]
                box_numbers = pd.Series([int(bn[3:]) for bn in box_numbers])

                if not is_consecutive(box_numbers):
                    # 如果箱号不连续，将sheet名修改为红色并添加提示信息
                    plt = plt + ' (箱号不连续)'
                    red_format = writer.book.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
                    writer.sheets[plt].set_column('A:E', None, red_format)

                # 获取worksheet
                worksheet = writer.book.add_worksheet(plt)
                worksheet.set_column(0, 0, 60)

                # 将字典更改为列表以改变输出格式
                pallet_info = [
                    [f"PO# {po}", ""],
                    [f"{port}", ""],
                    ["MADE IN TAIWAN", ""],
                    [f"PLT/NO.: {plt}", ""],
                    [f"C/NO.:{pallet_data.iloc[0, 3]}-{pallet_data.iloc[-1, -1]}", ""]
                ]

                # 遍历数据并添加到工作表中
                for row_num, row_data in enumerate(pallet_info):
                    for col_num, col_data in enumerate(row_data):
                        worksheet.write(row_num, col_num, col_data)

                # 获取worksheet
                worksheet = writer.sheets[plt]

                # 更改字体大小
                format1 = writer.book.add_format({'font_size': 40, 'font_name': 'Arial'})
                worksheet.set_column('A:E', None, format1)

                # 设置打印样式
                worksheet.set_print_scale(100)  # 设置为100%
                worksheet.set_landscape()  # 设置为横向

        workbook = writer.book
        worksheet = writer.sheets['Summary']

        # 定義細胞格式
        orange_format = workbook.add_format()
        orange_format.set_bg_color('gray')

        # 設置'O'範圍的細胞顏色
        worksheet.conditional_format('A1:O1', {'type': 'no_errors', 'format': orange_format})


root = tk.Tk()
frame = tk.Frame(root)
frame.pack()

evg_label = tk.Label(frame, text="Starting EVG:")
evg_label.pack(side=tk.LEFT)
evg_entry = tk.Entry(frame)
evg_entry.pack(side=tk.LEFT)

plt_label = tk.Label(frame, text="Starting PLT:")
plt_label.pack(side=tk.LEFT)
plt_entry = tk.Entry(frame)
plt_entry.pack(side=tk.LEFT)
po_label = tk.Label(frame, text="PO:")
po_label.pack(side=tk.LEFT)
po_entry = tk.Entry(frame)
po_entry.pack(side=tk.LEFT)

port_label = tk.Label(frame, text="Port:")
port_label.pack(side=tk.LEFT)
port_entry = tk.Entry(frame)
port_entry.pack(side=tk.LEFT)

import_button = tk.Button(frame, text="Import CSV", command=import_csv_data)
import_button.pack()

export_button = tk.Button(frame, text="Export Data", command=export_data)
export_button.pack()

frame2 = tk.Frame(root)
frame2.pack()

# Set up the treeview
treeview = ttk.Treeview(frame2, show='headings')
treeview.pack()

root.mainloop()