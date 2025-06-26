import os.path
import pandas as pd
import re
import sys
from PyQt5.QtWidgets import QFileDialog, QTableWidgetItem, QApplication, QMainWindow, QCheckBox, QMessageBox
from PyQt5.QtGui import QIcon, QDesktopServices,QMovie
from gui import Ui_MainWindow
from PyQt5.QtCore import QUrl
from pathlib import Path


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # Load the UI from the .ui file
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        icon = QIcon('images/PROCESS.bmp')

        # Get the toolButton3 from the UI
        self.toolButton3 = self.ui.toolButton_3
        self.tableWidget = self.ui.tableWidget
        self.pushButton = self.ui.pushButton
        self.pushButton_2 = self.ui.pushButton_2
        self.pushButton_3 = self.ui.pushButton_3
        self.progressBar = self.ui.progressBar
        self.label = self.ui.label
        self.label_6 = self.ui.label_6
        self.verticalLayout_5 = self.ui.verticalLayout_5
        self.toolButton3.setIcon(icon)
        self.list_widget = self.ui.listWidget
        self.open = self.ui.pushButton_4
        self.gif = self.ui.label_4
        self.movie = QMovie('images/text.gif')
        self.movie.setScaledSize(self.gif.size())

        self.ui.label_4.setMovie(self.movie)
        self.movie.start()

        # Check if the button is found, and connect signal if it is
        if self.toolButton3 is None:
            print("Error: toolButton3 not found in the UI")
        else:
            self.toolButton3.clicked.connect(self.handle_button_click)
            self.pushButton.clicked.connect(self.file_picker)
            self.pushButton_2.clicked.connect(self.refresh)
            self.pushButton_3.clicked.connect(self.process)
            self.open.clicked.connect(self.open_files_folder)

    def load(self):

        df = self.merge_chunk_df2_
        self.tableWidget.setRowCount(df.shape[0])
        self.tableWidget.setColumnCount(df.shape[1])
        self.tableWidget.setHorizontalHeaderLabels(df.columns)


        for rows in range(df.shape[0]):
            for col in range(df.shape[1]):
                value = str(df.iat[rows,col])
                item = QTableWidgetItem(value)
                self.tableWidget.setItem(rows,col,item)
        self.tableWidget.setStyleSheet("font : 4pt")

        self.label.setText('Done... Check files in Documents/Extractor')

    def refresh(self):
        try:
            self.label_6.setText('Available Units')
            file_path, _ = QFileDialog.getOpenFileNames(self, "Open Excel File", "", "Excel Files (*.xls)")
            self.file_path_new = file_path


            if file_path:
                self.file_path = file_path
                df = pd.read_excel(self.file_path[0], engine='xlrd')  # Use 'xlrd' for .xls files
                # column names
                df.columns = ['Date', 'Item Code', 'Description', 'Department', 'Qty', 'Unit', 'Packaging']
                checkboxvalues = [val for val in df['Packaging'].unique() if not pd.isna(val) and val != 'Unit']
                self.checkboxes = []
                while self.verticalLayout_5.count():
                    item = self.verticalLayout_5.takeAt(0)
                    widget = item.widget()
                    if widget:
                        widget.deleteLater()

                # Clear your checkbox references as well
                self.checkboxes.clear()

                for val in checkboxvalues:
                    self.checkbox = QCheckBox(str(val))
                    self.checkbox.setStyleSheet("font : 4pt")
                    self.verticalLayout_5.addWidget(self.checkbox)
                    self.checkboxes.append(self.checkbox)
                self.setLayout(self.verticalLayout_5)
        except Exception as e:
            QMessageBox.warning(self, "Excel File error", "Received Bad Structure... Please modify Excel file and try again")
            print(f"Error details: {e}")
            return

    def process(self):
        try:

            self.checked_values = [cb.text() for cb in self.checkboxes if cb.isChecked()]
            self.split_excel_into_chunks(self.file_path[0])
            self.load()
            self.packing()
        except Exception as e:
            QMessageBox.warning(self, "Excel File error", "Please click Refresh to Load data for Processing")
            print(f"Error details: {e}")
            return



    def packing(self):
        # Get unique non-null 'Description' values
        self.unique_values = self.merge_chunk_df_['DESCRIPTION'].dropna().unique()
        self.list_widget.clear()
        # Add unique values to the QListWidget
        for val in self.unique_values:
            self.list_widget.addItem(str(val))
        self.list_widget.setStyleSheet("font : 4pt")


    def file_picker(self):
        try:

            file_path, _ = QFileDialog.getOpenFileNames(self, "Open Excel File", "", "Excel Files (*.xls)")
            self.file_path_new = file_path
            if file_path:
                self.file_path = file_path
                self.split_excel_into_chunks(self.file_path[0])
                self.load()
                self.packing()
        except Exception as e:
            QMessageBox.warning(self, "Excel File error", "Please modify Excel file and try again")
            print(f"Error details: {e}")
            return

    def o_file(self, path):
        file_name = os.path.basename(path)
        new = os.path.splitext(file_name)[0]
        self.new = new
        return self.new

    def handle_button_click(self):
        try:
            path = ''
            text = self.ui.plainTextEdit.toPlainText().strip()
            t = text
            path = t.replace('"', "")
            if path:
                self.split_excel_into_chunks(path)
                self.load()
                self.path_ = path
            else:
                QMessageBox.warning(self, "Excel import error", "Please paste a valid path above for Processing")
        except Exception as e:
            QMessageBox.warning(self, "Excel File error", "Please modify Excel file and try again")
            print(f"Error details: {e}")
            return

    def file(self,df,path):
        try:
            base, ext = os.path.splitext(path)
            i = 1
            while os.path.exists(path):
                path = f"{base}_{i}{ext}"
                i += 1
            df.to_excel(path,index =False)
        except Exception as e:
            QMessageBox.warning(self, "Excel File error", "Please modify Excel file and try again")
            print(f"Error details: {e}")
            return

    def open_files_folder(self):
        try:
            # Get path to "extractor" folder inside user's Documents
            documents_folder = Path.home() / "Documents"
            extractor_folder = documents_folder / "extractor"
            folder_path = str(extractor_folder)  # convert Path to str for Qt

            if os.path.exists(folder_path):
                QDesktopServices.openUrl(QUrl.fromLocalFile(folder_path))
            else:
                print(f"Folder does not exist: {folder_path}")
        except Exception as e:
            QMessageBox.warning(self, "Excel File error", "Please modify Excel file and try again")
            print(f"Error details: {e}")
            return

    def split_excel_into_chunks(self, file_path):
        self.label.setText('Extracting and Processing... ')
        # Load the Excel file
        df = pd.read_excel(file_path, engine='xlrd')  # Use 'xlrd' for .xls files
        # column names
        df.columns = ['Date', 'Item Code', 'Description', 'Department', 'Qty', 'Unit', 'Packaging']
        # Drop the first two rows
        df = df.drop(index=[0, 1])  # Drop the first two rows
        # Initialize an empty list to store chunks
        chunks = []
        chunk = []
        chunk_name = 'New-sales'

        # Define the regex pattern for matching the header pattern
        header_pattern = r'^[A-Za-z]'
        merge_chunk_df = pd.DataFrame()
        # Iterate through the DataFrame rows


        for i, row in df.iterrows():
            # Skip rows that are headers or non-relevant
            if pd.isna(row.iloc[0]) or 'Salesmen Transactions Listing Report' in str(row.iloc[0]) or 'TOTALS' in str(
                    row.iloc[0]):
                continue
            if pd.isna(row.iloc[0]) or 'TOTALS' in str(row.iloc[0]):
                row.iloc[0] = ''
                continue
            # Check for section headers
            if isinstance(row.iloc[0], str) and re.match(header_pattern, row.iloc[0]):
                    # Save the previous chunk if it's not empty
                if chunk:
                    # Save chunk to a new DataFrame with the chunk name
                    chunk_df = pd.DataFrame(chunk)
                    chunk_df.columns = ['DATE', 'ITEM CODE', 'DESCRIPTION', 'DEPARTMENT', 'QTY', 'UNIT', 'PACKAGING']
                    # Ensure the chunk name is safe for file paths
                    chunk_df = chunk_df.drop(columns=['DATE', 'DEPARTMENT', 'QTY'])
                    chunk_df = chunk_df.groupby(['ITEM CODE', 'DESCRIPTION', 'PACKAGING'])['UNIT'].sum().reset_index()
                    chunk_df['text_p'] = chunk_df['DESCRIPTION'].str.extract(r'(\d+)\s*[A-Z]*\s*\*',
                                                                                 flags=re.IGNORECASE)
                    chunk_df['text_p'] = chunk_df['text_p'].fillna('1')
                    chunk_df['text_p'] = chunk_df['text_p'].astype(float)
                    chunk_df['ITEM CODE'] = chunk_df['ITEM CODE'].astype(int)

                    try:
                        chunk_df['CTN'] = chunk_df.apply(
                            lambda row: row['UNIT'] / row['text_p'] if row['PACKAGING'].upper() not in self.checked_values else
                            row['UNIT'], axis=1)
                    except:
                        chunk_df['CTN'] = chunk_df.apply(
                            lambda row : row['UNIT'] / row['text_p'] if row['PACKAGING'].upper() not in ['BOX', 'PCS', 'BALE', 'PARCEL'] else
                            row['UNIT'], axis=1)

                    chunk_df['CTN'] = chunk_df['CTN'].apply(lambda x: int(x) if pd.notna(x) and int(x) != 0 else 0)
                    chunk_name = chunk_name.replace('/', '_')
                    chunk_df['NAME'] = chunk_name.split('-')[1]
                    path = os.path.join("Files", f"{chunk_name}.xlsx")
                    merge_chunk_df = pd.concat([merge_chunk_df, chunk_df], ignore_index=True)
                    # chunk_df.to_excel(path, index=False)
                    chunks.append(chunk_df)  # Add the chunk DataFrame to the list
                # Start a new chunk with the current row as the header
                chunk_name = row.iloc[0]  # Set the name for the chunk
                chunk = []  # Reset the current chunk for new data
                self.progressBar.setValue(int((i + 1) / df.shape[0]) * 100)
            else:
                # Otherwise, add this row to the current chunk
                chunk.append(row.tolist())
            self.progressBar.setValue(int((i + 1) / df.shape[0]) * 100)

        # Save the last chunk if there's any remaining data

        if chunk:
            chunk_df = pd.DataFrame(chunk)
            chunk_df.columns = ['DATE', 'ITEM CODE', 'DESCRIPTION', 'DEPARTMENT', 'QTY', 'UNIT', 'PACKAGING']
            # Ensure the chunk name is safe for file paths
            chunk_name.replace('/', '_')
            chunk_df = chunk_df.drop(columns=['DATE', 'DEPARTMENT', 'QTY'])
            chunk_df = chunk_df.groupby(['ITEM CODE', 'DESCRIPTION', 'PACKAGING'])['UNIT'].sum().reset_index()
            chunk_df['text_p'] = chunk_df['DESCRIPTION'].str.extract(r'(\d+)\s*[A-Z]*\s*\*', flags=re.IGNORECASE)
            chunk_df['text_p'] = chunk_df['text_p'].fillna('1')
            chunk_df['text_p'] = chunk_df['text_p'].astype(float)
            chunk_df['ITEM CODE'] = chunk_df['ITEM CODE'].astype(int)
            if self.checked_values:
                chunk_df['CTN'] = chunk_df.apply(lambda row: row['UNIT'] / row['text_p'] if row['PACKAGING'].upper() not in self.checked_values else row['UNIT'], axis=1)
            else:
                chunk_df['CTN'] = chunk_df.apply(lambda row: row['UNIT'] / row['text_p'] if row['PACKAGING'].upper() not in ['BOX', 'PCS', 'BALE', 'PARCEL'] else row['UNIT'], axis=1)
            chunk_df['CTN'] = chunk_df['CTN'].apply(lambda x: int(x) if pd.notna(x) and int(x) != 0 else 0)
            chunk_df['NAME'] = chunk_name.split('-')[1]
            path = os.path.join("Files", f"{chunk_name}.xlsx")
            merge_chunk_df = pd.concat([merge_chunk_df, chunk_df], ignore_index=True)
            chunks.append(chunk_df)  # Add the chunk DataFrame to the list


        try:
            documents_folder = Path.home() / "Documents"
            extractor_folder = documents_folder / "extractor"
            extractor_folder.mkdir(parents=True, exist_ok=True)

            path2 = extractor_folder / f"{self.o_file(self.file_path[0])}_SALENAMES_AND_PRODUCTS.xlsx"
            path3 = extractor_folder / f"{self.o_file(self.file_path[0])}_FINAL_REPORT.xlsx"
        except Exception:
            path2 = extractor_folder / f"{self.o_file(self.path_)}_SALENAMES_AND_PRODUCTS.xlsx"
            path3 = extractor_folder / f"{self.o_file(self.path_)}_FINAL_REPORT.xlsx"
        merge_chunk_df2 = merge_chunk_df
        merge_chunk_df = merge_chunk_df[merge_chunk_df['NAME'] != 'sales']

        merge_chunk_df2 = merge_chunk_df2[merge_chunk_df2['NAME'] != 'sales']
        merge_chunk_df2 = merge_chunk_df2.groupby(['NAME'])['CTN'].sum().reset_index()
        merge_chunk_df2['NO'] = range(1, len(merge_chunk_df2) + 1)
        merge_chunk_df2 = merge_chunk_df2[['NO', 'NAME', 'CTN']]
        merge_chunk_df2.columns = ['NO', 'NAME', 'QTY']

        merge_chunk_df['NAME'] = merge_chunk_df['NAME'].mask(merge_chunk_df['NAME'].duplicated(), '')
        merge_chunk_df['NO'] = (merge_chunk_df['NAME'] != '').cumsum()
        merge_chunk_df['NO'] = merge_chunk_df['NO'].where(merge_chunk_df['NAME'] != '', '')
        self.merge_chunk_df_ = merge_chunk_df[['NO', 'NAME', 'ITEM CODE', 'DESCRIPTION', 'UNIT', 'PACKAGING', 'CTN']]
        self.file(merge_chunk_df[['NO', 'NAME', 'ITEM CODE', 'DESCRIPTION', 'UNIT', 'PACKAGING', 'CTN']], path2)
        self.file(merge_chunk_df2, path3)
        self.merge_chunk_df2_ = merge_chunk_df2


app = QApplication(sys.argv)
window = MyWindow()
window.show()
sys.exit(app.exec_())
