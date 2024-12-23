import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QTableWidget, QTableWidgetItem, QVBoxLayout, QPushButton,
    QHBoxLayout, QWidget, QFileDialog, QMessageBox, QStyle
)
from PyQt5.QtCore import Qt

class ExcelViewer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Tabulator9000 by KENNEDY")
        self.resize(800, 600)
        self.setWindowIcon(self.style().standardIcon(QStyle.SP_VistaShield))

        # Center the window on the screen
        self.center_window()

        self.data = None
        self.undo_stack = []  # To store undo states

        # Enable drag and drop
        self.setAcceptDrops(True)

        # Main layout
        main_layout = QVBoxLayout()
        button_layout = QHBoxLayout()
        button_layout_2 = QHBoxLayout()

        self.table = QTableWidget()
        self.table.setSelectionMode(QTableWidget.MultiSelection)
        self.table.setSelectionBehavior(QTableWidget.SelectItems)

        # Buttons
        self.load_button = QPushButton("Load Excel")
        self.load_button.setIcon(self.style().standardIcon(QStyle.SP_FileIcon))
        self.load_button.clicked.connect(self.load_excel)
        button_layout.addWidget(self.load_button)

        self.undo_button = QPushButton("Undo")
        self.undo_button.setIcon(self.style().standardIcon(QStyle.SP_ArrowBack))
        self.undo_button.clicked.connect(self.undo)
        self.undo_button.setEnabled(False)
        button_layout.addWidget(self.undo_button)

        self.copy_button = QPushButton("Copy Cells")
        self.copy_button.setIcon(self.style().standardIcon(QStyle.SP_TitleBarNormalButton))
        self.copy_button.clicked.connect(self.copy_cells)
        button_layout.addWidget(self.copy_button)

        self.filter_columns_button = QPushButton("Filter Columns (GFK)")
        self.filter_columns_button.setIcon(self.style().standardIcon(QStyle.SP_TitleBarUnshadeButton))
        self.filter_columns_button.clicked.connect(self.filter_columns_gfk)
        button_layout_2.addWidget(self.filter_columns_button)

        self.remove_duplicates_button = QPushButton("Merge Duplicates")
        self.remove_duplicates_button.clicked.connect(self.remove_duplicates)
        button_layout_2.addWidget(self.remove_duplicates_button)

        self.clear_button = QPushButton("Clear Data")
        self.clear_button.setIcon(self.style().standardIcon(QStyle.SP_DialogCancelButton))
        self.clear_button.clicked.connect(self.clear_data)
        button_layout.addWidget(self.clear_button)

        main_layout.addLayout(button_layout)
        main_layout.addLayout(button_layout_2)
        main_layout.addWidget(self.table)

        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)

        # Enable keypress events
        self.table.keyPressEvent = self.handle_keypress

        # Ensure window is always on top
        # self.setWindowFlag(Qt.WindowStaysOnTopHint)

    def center_window(self):
        frame_geometry = self.frameGeometry()
        screen = QApplication.primaryScreen()
        center_point = screen.availableGeometry().center()
        frame_geometry.moveCenter(center_point)
        self.move(frame_geometry.topLeft())

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            self.handle_file(file_path)
            break

    def load_excel(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Open Excel File", "", "Excel Files (*.xlsx *.xls *.csv)")
        if file_name:
            self.handle_file(file_name)

    def handle_file(self, file_path):
        try:
            if file_path.endswith('.xls') or file_path.endswith('.xlsx'):
                self.data = pd.read_excel(file_path, header=None)
        except: pass
        try:
            self.data = self.recover_corrupt_excel(file_path)
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))
        self.populate_table()

    def recover_corrupt_excel(self, filename):
        encoding_list = [
            'ascii', 'big5', 'big5hkscs', 'cp037', 'cp273', 'cp424', 'cp437', 'cp500', 'cp720', 'cp737',
            'cp775', 'cp850', 'cp852', 'cp855', 'cp856', 'cp857', 'cp858', 'cp860', 'cp861', 'cp862',
            'cp863', 'cp864', 'cp865', 'cp866', 'cp869', 'cp874', 'cp875', 'cp932', 'cp949', 'cp950',
            'cp1006', 'cp1026', 'cp1125', 'cp1140', 'cp1250', 'cp1251', 'cp1252', 'cp1253', 'cp1254',
            'cp1255', 'cp1256', 'cp1257', 'cp1258', 'euc_jp', 'euc_jis_2004', 'euc_jisx0213', 'euc_kr',
            'gb2312', 'gbk', 'gb18030', 'hz', 'iso2022_jp', 'iso2022_jp_1', 'iso2022_jp_2',
            'iso2022_jp_2004', 'iso2022_jp_3', 'iso2022_jp_ext', 'iso2022_kr', 'latin_1', 'iso8859_2',
            'iso8859_3', 'iso8859_4', 'iso8859_5', 'iso8859_6', 'iso8859_7', 'iso8859_8', 'iso8859_9',
            'iso8859_10', 'iso8859_11', 'iso8859_13', 'iso8859_14', 'iso8859_15', 'iso8859_16', 'johab',
            'koi8_r', 'koi8_t', 'koi8_u', 'kz1048', 'mac_cyrillic', 'mac_greek', 'mac_iceland', 'mac_latin2',
            'mac_roman', 'mac_turkish', 'ptcp154', 'shift_jis', 'shift_jis_2004', 'shift_jisx0213', 'utf_32',
            'utf_32_be', 'utf_32_le', 'utf_16', 'utf_16_be', 'utf_16_le', 'utf_7', 'utf_8', 'utf_8_sig'
        ]
        for encoding in encoding_list:
            try:
                df = pd.read_csv(filename, sep='\t', header=None, encoding=encoding)
                print(encoding)
                return df
            except: pass
        raise ValueError(f"Recovery failed")
    
    def filter_columns_gfk(self):
        if self.data is None:
            QMessageBox.warning(self, "No Data", "No data available to filter.")
            return

        self.save_undo_state()

        # Specify the columns to keep
        columns_to_keep = ["so_number", "item_number", "item_desc", "so_qty", "orderamt"]

        try:
            # Filter columns based on their names
            self.data.columns = self.data.iloc[0] # Treat first row as header for filtering
            self.data = self.data.loc[:, columns_to_keep]
            self.data.reset_index(drop=True, inplace=True)

            # Update the table
            self.populate_table()
        except KeyError:
            QMessageBox.critical(self, "Error", "Some required columns are missing.")

    def remove_duplicates(self):
        if self.data is None:
            QMessageBox.warning(self, "No Data", "No data available to process.")
            return

        self.save_undo_state()

        try:
            # Preserve headers
            headers = self.data.iloc[0]

            # Exclude the headers row from processing
            data_body = self.data[1:].copy()

            # Ensure the relevant columns exist
            required_columns = ["so_number", "item_desc", "so_qty"]
            for col in required_columns:
                if col not in headers.values:
                    QMessageBox.critical(self, "Error", f"Column '{col}' is missing.")
                    return

            # Assign column names for processing
            data_body.columns = headers

            # Convert `so_qty` to numeric for aggregation
            data_body["so_qty"] = pd.to_numeric(data_body["so_qty"], errors="coerce").fillna(0)

            # Group by `so_number` and `item_desc`, summing up `so_qty`
            grouped_data = (
                data_body.groupby(["so_number", "item_desc"], as_index=False)
                .agg({"so_qty": "sum", **{col: "first" for col in data_body.columns if col not in required_columns}})
            )

            # Reinsert headers
            self.data = pd.concat([headers.to_frame().T, grouped_data], ignore_index=True)

            # Update the table
            self.populate_table()

        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred while processing duplicates: {str(e)}")

    def populate_table(self):
        if self.data is not None:
            self.table.clear()
            rows, cols = self.data.shape
            self.table.setRowCount(rows)
            self.table.setColumnCount(cols)

            for row in range(rows):
                for col in range(cols):
                    cell_value = self.data.iat[row, col]
                    item = QTableWidgetItem(str(cell_value) if pd.notna(cell_value) else "")
                    # item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled) # shit is now editable
                    self.table.setItem(row, col, item)

    def copy_cells(self):
        selected_ranges = self.table.selectedRanges()
        if not selected_ranges:
            QMessageBox.warning(self, "No Selection", "Please select cells to copy.")
            return

        copied_data = []
        for selection in selected_ranges:
            for row in range(selection.topRow(), selection.bottomRow() + 1):
                copied_row = []
                for col in range(selection.leftColumn(), selection.rightColumn() + 1):
                    copied_row.append(self.table.item(row, col).text())
                copied_data.append(copied_row)

        # Save to clipboard in Excel-compatible format
        clipboard = QApplication.clipboard()
        clipboard.setText('\n'.join(['\t'.join(row) for row in copied_data]))
        # QMessageBox.information(self, "Copied", "Selected cells have been copied to the clipboard.")

    def clear_data(self):
        self.save_undo_state()
        self.data = None
        self.table.clear()
        self.table.setRowCount(0)
        self.table.setColumnCount(0)

    def handle_keypress(self, event):
        if event.key() == Qt.Key_Space:
            if event.modifiers() & Qt.ShiftModifier:
                self.table.selectRow(self.table.currentRow())
            if event.modifiers() & Qt.ControlModifier:
                self.table.selectColumn(self.table.currentColumn())

        elif event.key() == Qt.Key_Delete:
            if event.modifiers() & Qt.ShiftModifier:
                self.delete_entire_selection()
            else:
                self.clear_selection()

        elif event.key() == Qt.Key.Key_Escape:
            self.table.clearSelection()

        elif event.key() == Qt.Key.Key_A and event.modifiers() & Qt.ControlModifier:
            self.table.selectAll()

        elif event.key() == Qt.Key.Key_Z and event.modifiers() & Qt.ControlModifier:
            self.undo()

        elif event.key() == Qt.Key.Key_C and event.modifiers() & Qt.ControlModifier:
            self.copy_cells()

        elif event.key() == Qt.Key_Plus and event.modifiers() & Qt.ControlModifier and event.modifiers() & Qt.ShiftModifier:
            self.insert_blank_row_or_column()

        else:
            super(QTableWidget, self.table).keyPressEvent(event)

    def insert_blank_row_or_column(self):
        if self.data is not None:
            self.save_undo_state()

            selected_ranges = self.table.selectedRanges()
            if not selected_ranges:
                return

            for selection in selected_ranges:
                if selection.leftColumn() == 0 and selection.rightColumn() == self.data.shape[1] - 1:
                    # Insert blank rows
                    for row in range(selection.topRow(), selection.bottomRow() + 1):
                        self.data = pd.concat([
                            self.data.iloc[:row],
                            pd.DataFrame([[None] * self.data.shape[1]], columns=self.data.columns),
                            self.data.iloc[row:]
                        ]).reset_index(drop=True)
                elif selection.topRow() == 0 and selection.bottomRow() == self.data.shape[0] - 1:
                    # Insert blank columns
                    for col in range(selection.leftColumn(), selection.rightColumn() + 1):
                        self.data.insert(col, None, [None] * self.data.shape[0])

            self.populate_table()

    def clear_selection(self):
        if self.data is not None:
            self.save_undo_state()

            selected_ranges = self.table.selectedRanges()
            for selection in selected_ranges:
                for row in range(selection.topRow(), selection.bottomRow() + 1):
                    for col in range(selection.leftColumn(), selection.rightColumn() + 1):
                        self.data.iat[row, col] = None

            self.populate_table()

    def delete_entire_selection(self):
        if self.data is not None:
            self.save_undo_state()

            selected_ranges = self.table.selectedRanges()

            rows_to_delete = set()
            cols_to_delete = set()

            for selection in selected_ranges:
                if selection.leftColumn() == 0 and selection.rightColumn() == self.data.shape[1] - 1:
                    rows_to_delete.update(range(selection.topRow(), selection.bottomRow() + 1))
                elif selection.topRow() == 0 and selection.bottomRow() == self.data.shape[0] - 1:
                    cols_to_delete.update(range(selection.leftColumn(), selection.rightColumn() + 1))

            # Delete rows
            for row in sorted(rows_to_delete, reverse=True):
                self.data = self.data.drop(index=row).reset_index(drop=True)

            # Delete columns
            for col in sorted(cols_to_delete, reverse=True):
                self.data = self.data.drop(columns=self.data.columns[col])

            self.populate_table()

    def save_undo_state(self):
        if self.data is not None:
            self.undo_stack.append(self.data.copy())
            self.undo_button.setEnabled(True)

    def undo(self):
        if self.undo_stack:
            self.data = self.undo_stack.pop()
            self.populate_table()

        if not self.undo_stack:
            self.undo_button.setEnabled(False)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    viewer = ExcelViewer()
    viewer.show()
    sys.exit(app.exec_())
