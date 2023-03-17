import sys
from PyQt5 import QtWidgets, QtGui
from PyQt5.QtWidgets import QFileDialog , QProgressBar , QMessageBox
import win32com.client as win32
from PyQt5.QtCore import Qt


class FileSelectUI(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()

        # Create buttons
        self.xlsx_button = QtWidgets.QPushButton("Select MDL file", self)
        self.txt_button = QtWidgets.QPushButton("Select Variable file", self)
        self.check_button = QtWidgets.QPushButton("Check", self)

        # Create labels
        self.xlsx_label = QtWidgets.QLabel("MDL file:")
        self.txt_label = QtWidgets.QLabel("Variable file:")

        # Create text fields
        self.xlsx_text = QtWidgets.QLineEdit()
        self.txt_text = QtWidgets.QLineEdit()

        # Connect buttons to functions
        self.xlsx_button.clicked.connect(self.select_xlsx)
        self.txt_button.clicked.connect(self.select_txt)
        self.check_button.clicked.connect(self.check_files)

        # Create progress bar
        self.progress_bar = QProgressBar()

        # Set style sheet for buttons
        self.xlsx_button.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold;")
        self.txt_button.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold;")
        self.check_button.setStyleSheet("background-color: #008CBA; color: white; font-weight: bold;")

        # Set font and alignment for labels
        self.xlsx_label.setFont(QtGui.QFont("Arial", 14, QtGui.QFont.Bold))
        self.txt_label.setFont(QtGui.QFont("Arial", 14, QtGui.QFont.Bold))
        self.xlsx_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.txt_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)

        # Set window title and icon
        self.setWindowTitle("MDL Checker")
        self.setWindowIcon(QtGui.QIcon("icon.png"))

        # Set tool tips for buttons
        self.xlsx_button.setToolTip("Select the MDL file to check")
        self.txt_button.setToolTip("Select the Variable file to check against")
        self.check_button.setToolTip("Check the files")

        # Set maximum width for buttons and minimum width for text fields
        self.xlsx_button.setMaximumWidth(150)
        self.txt_button.setMaximumWidth(150)
        self.xlsx_text.setMinimumWidth(400)
        self.txt_text.setMinimumWidth(400)

        # Create layout
        grid_layout = QtWidgets.QGridLayout()
        grid_layout.addWidget(self.xlsx_label, 0, 0)
        grid_layout.addWidget(self.xlsx_text, 0, 1)
        grid_layout.addWidget(self.xlsx_button, 0, 2)
        grid_layout.addWidget(self.txt_label, 1, 0)
        grid_layout.addWidget(self.txt_text, 1, 1)
        grid_layout.addWidget(self.txt_button, 1, 2)
        grid_layout.addWidget(self.check_button, 2, 0, 1, 3)
        grid_layout.addWidget(self.progress_bar, 3, 0, 1, 3)

        # Set spacing and margins for layout
        grid_layout.setHorizontalSpacing(20)
        grid_layout.setVerticalSpacing(10)
        grid_layout.setContentsMargins(20, 20, 20, 20)

        # Set layout for widget
        self.setLayout(grid_layout)

        # Create instance variable for progress bar
        self.progress = 0


    def select_xlsx(self):
        xlsx_file, _ = QFileDialog.getOpenFileName(self, "Select MDL file", "", "Excel files (*.xlsx)")
        self.xlsx_text.setText(xlsx_file)

    def select_txt(self):
        txt_file, _ = QFileDialog.getOpenFileName(self, "Select Variable file", "", "Text files (*.txt)")
        self.txt_text.setText(txt_file)

    def check_files(self):
        MDLFile = self.xlsx_text.text()
        VarFile = self.txt_text.text()

        f = open(VarFile, "r")
        File = f.read()

        # Opening Workbook and Worksheet
        xl = win32.Dispatch('Excel.Application')
        wb = xl.Workbooks.Open(MDLFile)
        ws = wb.Worksheets('Mapping_File_MDL_LINK')
        self.progress_bar.setMaximum(ws.UsedRange.Rows.Count)

        def rgbToInt(rgb):
            colorInt = rgb[0] + (rgb[1] * 256) + (rgb[2] * 256 * 256)
            return colorInt

        # Update Values
        for i in range(1, ws.UsedRange.Rows.Count + 1):
            self.progress_bar.setValue(i)
            try:
                y = str((ws.Cells(i, 3).Value).strip())
                if y in File:
                    ws.Cells(i, 5).Value = "OK"
                    ws.Cells(i, 5).Interior.Color = rgbToInt((0, 255, 0))
                elif y not in File and y.startswith('Platform'):
                    ws.Cells(i, 5).Value = "NOK"
                    ws.Cells(i, 5).Interior.Color = rgbToInt((255, 0, 0))
            except:
                pass

        ws.Columns("E:E").AutoFilter(1)

        # Close and save the workbook
        wb.Close(True)
        # xl.Quit()


        # Show finished message box
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setWindowTitle("MDL Checker")
        msg.setText("File checking is finished.")
        msg.exec_()


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    file_select_ui = FileSelectUI()
    file_select_ui.show()
    sys.exit(app.exec_())
