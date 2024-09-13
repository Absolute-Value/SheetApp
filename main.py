import os
import sys
from PyQt6.QtWidgets import QApplication, QWidget, QTableWidget, QTableWidgetItem, QMenuBar, QFileDialog
from openpyxl import load_workbook

class Window(QWidget):
    def __init__(self):
        super().__init__()

        self.file_path = "Book1.xlsx"
        self.setWindowTitle(os.path.basename(self.file_path))
        self.setGeometry(100, 100, 800, 600)

        self.table = QTableWidget(self)
        self.table.setGeometry(0, 0, self.width(), self.height())

        self.create_menu_bar()

    def create_menu_bar(self):
        menu_bar = QMenuBar(self)
        file_menu = menu_bar.addMenu("ファイル")

        load_action = file_menu.addAction("開く")
        load_action.triggered.connect(self.load_file)

        save_action = file_menu.addAction("保存")
        save_action.triggered.connect(self.save_file)

        save_as_action = file_menu.addAction("コピーを保存")
        save_as_action.triggered.connect(self.save_file_as)

    def load_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "ファイルを開く", "", "Excelファイル (*.xlsx)")
        if file_path:
            self.load_excel_data(file_path)
            self.file_path = file_path
            self.setWindowTitle(os.path.basename(self.file_path))

    def save_file(self):
        if self.file_path == "Book1.xlsx":
            self.save_file_as()
        else:
            self.save_excel_data(self.file_path)

    def save_file_as(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "ファイルを保存", "", "Excelファイル (*.xlsx)")
        if file_path:
            self.save_excel_data(file_path)
        
    def load_excel_data(self, file_path):
        workbook = load_workbook(file_path)
        sheet = workbook.active

        self.table.setColumnCount(sheet.max_column)
        self.table.setRowCount(sheet.max_row)

        rd = sheet.row_dimensions
        for row_index in rd.keys():
            self.table.setRowHeight(row_index, int(rd[row_index].height))
        sc = sheet.column_dimensions
        for col_index in sc.keys():
            self.table.setColumnWidth(ord(col_index) - 65, int(sc[col_index].width*7))
            for col_index2 in range(sc[col_index].min+1, sc[col_index].max+1):
                self.table.setColumnWidth(col_index2, int(sc[col_index].width*7))

        for row_index, row in enumerate(sheet.iter_rows()):
            for col_index, cell in enumerate(row):
                cell_value = str(cell.value) if cell.value is not None else ""
                item = QTableWidgetItem(cell_value)
                self.table.setItem(row_index, col_index, item)
                if cell.font.bold:
                    font = item.font()
                    font.setBold(True)
                if cell.data_type == "n": # 数字を右寄せにする
                    item.setTextAlignment(0x0082)

    def resizeEvent(self, event):
        self.table.setGeometry(0, 0, self.width(), self.height())
        super().resizeEvent(event)

def main():
    app = QApplication(sys.argv)
    window = Window()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()