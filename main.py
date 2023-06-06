

from openpyxl.styles import Font
import sys
import pandas as pd
import sqlite3
import numpy as np
from PyQt5.QtWidgets import QHBoxLayout,QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QTableView, QTextEdit, QLabel, QMessageBox
from PyQt5.QtSql import QSqlDatabase, QSqlTableModel, QSqlQuery
from PyQt5.QtCore import Qt
db = pd.read_csv('1your_file.csv')

def process_orders(orders, evg_base, plt_base):
    new_data_list = []
    summary_data_list = []
    max_boxes = 0
    for index, order in orders.iterrows():
        item = order['ITEM']
        po = order['PO']
        order_qty = order['ORDER QTY']
        item_info = db[db['ITEM'] == int(item)]
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
            "每箱淨重(KG)": item_info['N.W'],
            "每箱毛重(KG)": item_info['G.W'],
            "產品總重量": item_info['G.W'] * cartons
        }
        summary_data_list.append(summary_row)
        for i in range(full_pallets):
            new_row = {'PO': po, 'ITEM': item, 'PLT': 'EVG' + str(plt_base).zfill(6),"This Pallet HSU":hsu_carton*item_info['HSU/Carton']}
            plt_base += 1
            for j in range(hsu_carton):
                new_row[str(j)] = 'EVG' + str(evg_base).zfill(7)
                evg_base += 1
            new_data_list.append(new_row)
        if leftover_cartons > 0:
            new_row = {'PO': po, 'ITEM': item, 'PLT': 'Incomplete Pallet' + str(item),"This Pallet HSU":leftover_cartons*item_info['HSU/Carton']}
            for j in range(leftover_cartons):
                new_row[str(j)] = 'EVG' + str(evg_base).zfill(7)
                evg_base += 1
            for k in range(leftover_cartons, hsu_carton):
                new_row[str(k)] = np.nan
            new_data_list.append(new_row)
        if leftover_hsu > 0:
            new_row = {'PO': po, 'ITEM': item, 'PLT': 'Incomplete Carton', '0': 'EVG' + str(evg_base).zfill(7),"This Pallet HSU":leftover_hsu*item_info['HSU/Carton']}
            for k in range(1, hsu_carton):
                new_row[str(k)] = np.nan
            evg_base += 1
            new_data_list.append(new_row)
    summary_data = pd.DataFrame(summary_data_list)
    new_data = pd.DataFrame(new_data_list)
    new_data.columns = ['PO', 'ITEM', 'PLT','This Pallet HSU'] + [str(i) for i in range(1, max_boxes + 1)]

    return new_data, summary_data


class MyApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('HTT——出貨訂單——貨櫃 App')
        self.setGeometry(300, 300, 800, 600)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # Create a QHBoxLayout for buttons
        button_layout = QHBoxLayout()

        self.export_selected_button = QPushButton('Export Selected Rows')
        self.export_selected_button.clicked.connect(self.export_selected_rows)
        button_layout.addWidget(self.export_selected_button)

        self.import_button = QPushButton('Import Data')
        self.import_button.clicked.connect(self.import_data)
        button_layout.addWidget(self.import_button)

        self.export_button = QPushButton('Export Data')
        self.export_button.clicked.connect(self.export_data)
        button_layout.addWidget(self.export_button)

        self.refresh_button = QPushButton('Refresh Database')
        self.refresh_button.clicked.connect(self.refresh_database)
        button_layout.addWidget(self.refresh_button)

        self.delete_button = QPushButton('Delete Database')
        self.delete_button.clicked.connect(self.delete_database)
        button_layout.addWidget(self.delete_button)

        layout.addLayout(button_layout)

        h_layout = QHBoxLayout()

        # Create a sub layout for EVG
        evg_layout = QHBoxLayout()
        self.evg_starting = QTextEdit()
        self.evg_starting.setMaximumWidth(100)
        self.evg_starting.setFixedHeight(25)
        evg_layout.addWidget(QLabel("EVG Starting"))
        evg_layout.addWidget(self.evg_starting)
        evg_layout.addStretch()  # add a stretchable empty space to push your widgets together

        # Create a sub layout for PLT
        plt_layout = QHBoxLayout()
        self.plt_starting = QTextEdit()
        self.plt_starting.setMaximumWidth(100)
        self.plt_starting.setFixedHeight(25)
        plt_layout.addWidget(QLabel("PLT Starting"))
        plt_layout.addWidget(self.plt_starting)
        plt_layout.addStretch()  # add a stretchable empty space to push your widgets together

        h_layout.addLayout(plt_layout)
        h_layout.addLayout(evg_layout)


        layout.addLayout(h_layout)

        self.view_orders = QTableView()
        layout.addWidget(QLabel("db.orders"))
        layout.addWidget(self.view_orders)

        self.view_pallet_data = QTableView()
        layout.addWidget(QLabel("pallet_data"))
        layout.addWidget(self.view_pallet_data)
        self.view_pallet_data.setSelectionMode(QTableView.MultiSelection)

        self.setLayout(layout)

        self.conn = sqlite3.connect('database.db')
        self.cursor = self.conn.cursor()

        self.cursor.execute('CREATE TABLE IF NOT EXISTS db_orders (PO TEXT, ITEM TEXT, ORDER_QTY REAL)')
        column_definitions = ['PO TEXT', 'ITEM TEXT', 'PLT TEXT', 'This Pallet HSU REAL'] + [f'"{i}" TEXT' for i in
                                                                                             range(1, 151)]

        columns_string = ', '.join(column_definitions)

        self.cursor.execute(f"""
            CREATE TABLE IF NOT EXISTS pallet_data ({columns_string})
        """)

        self.db = QSqlDatabase.addDatabase('QSQLITE')
        self.db.setDatabaseName('database.db')
        self.db.open()

        self.refresh_database()

    def export_selected_rows(self):
        indexes = self.view_pallet_data.selectionModel().selectedRows()
        model = self.view_pallet_data.model()
        data = []
        ids_to_delete = []  # 这个列表用于存储选中的行的ID，以便后续删除
        hsu_sums = {}
        column_item = 1
        column_hsu = 3
        column_po = 0

        for index in sorted(indexes):
            row_data = []
            item = model.data(model.index(index.row(), column_item))
            hsu = model.data(model.index(index.row(), column_hsu))
            po = model.data(model.index(index.row(), column_po))
            key = (po, item)
            if key in hsu_sums:
                hsu_sums[key] += hsu
            else:
                hsu_sums[key] = hsu
            for column in range(model.columnCount()):
                cell_index = model.index(index.row(), column)
                cell_data = model.data(cell_index)
                row_data.append(cell_data)
                if model.headerData(column, Qt.Horizontal) == "PLT":
                    ids_to_delete.append(cell_data)
            data.append(row_data)
        print(hsu_sums)
        df = pd.DataFrame(data, columns=[model.headerData(i, Qt.Horizontal) for i in range(model.columnCount())])

        # file_path, _ = QFileDialog.getSaveFileName(self, 'Export Selected Rows', '', "csv Files (*.csv)")
        # if file_path:
        #     print('Export data to:', file_path)
        #     df.to_csv(file_path, index=False)
        #
        for id in ids_to_delete:
            self.cursor.execute('DELETE FROM pallet_data WHERE PLT=?', (id,))
            print("Deleted row with ID:", id)
            self.conn.commit()

        self.refresh_database()

        summary_data_list = []
        db = pd.read_csv('1your_file.csv')
        db.set_index('ITEM', inplace=True)
        for key in hsu_sums:
            po, item = key
            order_qty = hsu_sums[key]
            item_info = db.loc[int(item)]
            cartons = order_qty // item_info['HSU/Carton']
            hsu_carton = item_info['Carton/Pallet']
            full_pallets = cartons // hsu_carton
            qty_hsu = item_info['Qty/HSU']
            hsu = order_qty
            carton_pallet = item_info['Carton/Pallet']
            print(carton_pallet)
            print(cartons)
            if cartons >= carton_pallet:  # If cartons fill at least a full pallet
                pallet = cartons // carton_pallet
            else:  # If cartons do not fill a full pallet
                pallet = 0
            print(pallet)
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
                "每箱淨重(KG)": item_info['N.W'],
                "每箱毛重(KG)": item_info['G.W'],
                "產品總重量": item_info['G.W'] * cartons
            }
            summary_data_list.append(summary_row)
        summary_data = pd.DataFrame(summary_data_list)

        # Use ExcelWriter to save df and summary_data into an Excel file with two sheets
        file_path, _ = QFileDialog.getSaveFileName(self, 'Export Selected Rows', '', "xlsx Files (*.xlsx)")
        if file_path:
            print('Export data to:', file_path)
            with pd.ExcelWriter(file_path) as writer:
                df.to_excel(writer, sheet_name='Pallet Data')
                summary_data.to_excel(writer, sheet_name='Summary Data')
                pallets = df['PLT'].unique()  # 获取所有不重复的棧板号
                for plt in pallets:
                    if plt != 'Incomplete Pallet' and plt != 'Incomplete Carton':
                        pallet_data = df[df['PLT'] == plt]
                        po = pallet_data['PO'].iloc[0]

                        # Here, we create an empty DataFrame for each pallet

                        # We then write the relevant data into the Excel sheet for each pallet
                        writer.book.create_sheet(plt)
                        worksheet = writer.sheets[plt]
                        worksheet['A1'] = 'PO#' + po
                        worksheet['A1'].font = Font(name='Arial',size=40)

                        worksheet['A2'] = 'BELTON'
                        worksheet['A2'].font = Font(name='Arial',size=40)

                        worksheet['A3'] = 'MADE IN TAIWAN'
                        worksheet['A3'].font = Font(name='Arial',size=40)
                        worksheet['A4'] = f"PLT/NO.: {plt}"
                        worksheet['A4'].font = Font(name='Arial',size=40)

                        # 使用numpy找到最后一个非空字符串的索引
                        last_nonempty_index = np.max(np.where(pallet_data.iloc[-1] != ''))
                        last_nonempty_value = pallet_data.iloc[-1][last_nonempty_index]

                        worksheet['A5'] = f"C/NO.:{pallet_data.iloc[0, 4]}-{last_nonempty_value}"
                        worksheet['A5'].font = Font(name='Arial',size=40)
            self.refresh_database()
    # def export_selected_rows(self):
    #     indexes = self.view_pallet_data.selectionModel().selectedRows()
    #     model = self.view_pallet_data.model()
    #     data = []
    #     ids_to_delete = []  # 这个列表用于存储选中的行的ID，以便后续删除
    #     hsu_sums = {}
    #     column_item = 1
    #     column_hsu = 3
    #     column_po = 0
    #     for index in sorted(indexes):
    #         row_data = []
    #         item = model.data(model.index(index.row(), column_item))
    #         hsu = model.data(model.index(index.row(), column_hsu))
    #         po = model.data(model.index(index.row(), column_po))
    #         key = (po, item)
    #         if key in hsu_sums:
    #             hsu_sums[key] += hsu
    #         else:
    #             hsu_sums[key] = hsu
    #         for column in range(model.columnCount()):
    #             cell_index = model.index(index.row(), column)
    #             cell_data = model.data(cell_index)
    #             row_data.append(cell_data)
    #             if model.headerData(column, Qt.Horizontal) == "PLT":
    #                 ids_to_delete.append(cell_data)
    #         data.append(row_data)
    #     print(hsu_sums)
    #     df = pd.DataFrame(data, columns=[model.headerData(i, Qt.Horizontal) for i in range(model.columnCount())])
    #
    #     file_path, _ = QFileDialog.getSaveFileName(self, 'Export Selected Rows', '', "csv Files (*.csv)")
    #     if file_path:
    #         print('Export data to:', file_path)
    #         df.to_csv(file_path, index=False)
    #
    #     for id in ids_to_delete:
    #         self.cursor.execute('DELETE FROM pallet_data WHERE PLT=?', (id,))
    #         print("Deleted row with ID:", id)
    #         self.conn.commit()
    #
    #     self.refresh_database()
    #
    #     summary_data_list = []
    #     db = pd.read_csv('1your_file.csv')
    #     db.set_index('ITEM', inplace=True)
    #     for key in hsu_sums:
    #         po, item = key
    #         order_qty = hsu_sums[key]
    #         item_info = db.loc[int(item)]
    #         cartons = order_qty // item_info['HSU/Carton']
    #         hsu_carton = item_info['Carton/Pallet']
    #         full_pallets = cartons // hsu_carton
    #         qty_hsu = item_info['Qty/HSU']
    #         hsu = order_qty
    #         carton_pallet = item_info['Carton/Pallet']
    #         pallet = cartons // carton_pallet
    #         carton_type = item_info['Carton-Type']
    #         pallet_weight = 18
    #         total_pallet_weight = pallet_weight * pallet
    #         qty = order_qty * qty_hsu
    #         summary_row = {
    #             "PO": po,
    #             "ITEM": item,
    #             "QTY": qty,
    #             "QTY/HSU": qty_hsu,
    #             "HSU": hsu,
    #             "HSU/Carton": item_info['HSU/Carton'],
    #             "Carton": cartons,
    #             "Carton/Pallet": carton_pallet,
    #             "Pallet": pallet,
    #             "Carton_type": carton_type,
    #             "Pallet-weight": pallet_weight,
    #             "棧板總重量": total_pallet_weight,
    #             "每箱淨重(KG)": item_info['N.W'],
    #             "每箱毛重(KG)": item_info['G.W'],
    #             "產品總重量": item_info['G.W'] * cartons
    #         }
    #         summary_data_list.append(summary_row)
    #     summary_data = pd.DataFrame(summary_data_list)
    #
    #     # Use ExcelWriter to save df and summary_data into an Excel file with two sheets
    #     file_path, _ = QFileDialog.getSaveFileName(self, 'Export Selected Rows', '', "xlsx Files (*.xlsx)")
    #     if file_path:
    #         print('Export data to:', file_path)
    #         with pd.ExcelWriter(file_path) as writer:
    #             df.to_excel(writer, sheet_name='Pallet Data')
    #             summary_data.to_excel(writer, sheet_name='Summary Data')

    def update_order_qty(self, df_exported):
        for index, row in df_exported.iterrows():
            po = row['PO']
            item = row['ITEM']

            # 计算出的 QTY 需要根据你的数据进行调整
            qty_to_subtract = row['QTY']

            # 更新 orders 表中的 ORDER QTY
            self.cursor.execute('UPDATE db_orders SET ORDER_QTY = ORDER_QTY - ? WHERE PO = ? AND ITEM = ?',
                                (qty_to_subtract, po, item))
            self.conn.commit()
    def import_data(self):
        evg_starting_text = self.evg_starting.toPlainText().strip()  # strip is used to remove leading/trailing white spaces
        plt_starting_text = self.plt_starting.toPlainText().strip()  # strip is used to remove leading/trailing white spaces
        if evg_starting_text and plt_starting_text:
            file_path, _ = QFileDialog.getOpenFileName(self, 'Import Data', '', "csv Files (*.csv)")
            if file_path:
                print('Import data from:', file_path)

                try:
                    # Read and Import db.orders
                    df_orders = pd.read_csv(file_path)
                    df_orders.to_sql('db_orders', self.conn, if_exists='append', index=False)
                except Exception as e:
                    QMessageBox.critical(self, 'Import Data Error', f'Error importing data: {e}')
                    return

                # Load and process the file
                db = pd.read_csv('1your_file.csv')
                db.set_index('ITEM', inplace=True)

                evg_starting_text = self.evg_starting.toPlainText()
                plt_starting_text = self.plt_starting.toPlainText()

                if not evg_starting_text.isdigit() or not plt_starting_text.isdigit():
                    QMessageBox.critical(self, 'Invalid Input', 'EVG Starting and PLT Starting must be numbers.')
                    return

                evg_starting = int(evg_starting_text)
                plt_starting = int(plt_starting_text)

                new_data, summary_data = process_orders(df_orders, evg_starting, plt_starting)

                # Append new_data and summary_data to their respective tables in the database
                new_data.to_sql('pallet_data', self.conn, if_exists='append', index=False)
                summary_data.to_sql('summary_data', self.conn, if_exists='append', index=False)

                # Refresh views
                self.refresh_database()

                QMessageBox.information(self, 'Import Data', 'Data imported successfully!')

    def export_data(self):
        file_path, _ = QFileDialog.getSaveFileName(self, 'Export Data', '', "csv Files (*.csv)")
        if file_path:
            print('Export data to:', file_path)

            df = pd.read_sql_query("SELECT * from pallet_data", self.conn)
            df.to_csv(file_path, index=False)

    def refresh_database(self):
        self.display_data_in_tableview('db_orders', self.view_orders)
        self.display_data_in_tableview('pallet_data', self.view_pallet_data)

    def delete_database(self):
        self.cursor.execute('DROP TABLE IF EXISTS db_orders')
        self.cursor.execute('DROP TABLE IF EXISTS pallet_data')

        self.refresh_database()

    def display_data_in_tableview(self, table_name, table_view):
        model = QSqlTableModel(db=self.db)
        model.setTable(table_name)
        model.select()
        table_view.setModel(model)

    # ... (the rest of your MyApp methods)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MyApp()
    window.show()
    sys.exit(app.exec_())