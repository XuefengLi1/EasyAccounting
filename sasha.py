import pandas as pd
import datetime
import shutil
import os, sys
import pdb
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtWidgets import QMainWindow, QWidget, QLabel, QLineEdit
from PyQt5.QtWidgets import QPushButton
from PyQt5.QtCore import QSize    


# need permission to move files

# /Users/smountain/Desktop/Retail Batch/PAYLIST FOLDER/paylistnew.xls
# /Users/smountain/Desktop/Retail Batch

def move_files(folder_path, sheet_path):
	# path = "Retail Batch"
	dest = "Retail Batch/PAYLIST FOLDER"

	path = folder_path
	# xl = pd.ExcelFile("Retail Batch/PAYLIST FOLDER/paylistnew.xls")
	xl = pd.ExcelFile(sheet_path)

	df = xl.parse("Sheet1")


	index = df['Unnamed: 1'].notnull()

	df = df[index]

	# df = df[['Printed by:  SASHAY','Unnamed: 1']]

	names = []
	for index, row in df.iterrows():

		prefix = row['Printed by:  SASHAY']
		date = row['Unnamed: 1']
		code = row['Unnamed: 6']
		if isinstance(date, datetime.datetime):
			name = prefix + '_' + code
			names.append(name)


	for filename in os.listdir(path):
		for name in names:
			if name in filename:

				shutil.move(path + '/' + filename, dest + '/' + filename)
				break

class MainWindow(QMainWindow):
    def __init__(self):
        QMainWindow.__init__(self)

        self.setMinimumSize(QSize(320, 140))    
        self.setWindowTitle("Sasha Accounting Tool") 

        self.nameLabel = QLabel(self)
        self.nameLabel.setText('Folder Path:')
        self.line = QLineEdit(self)

        self.line.move(100, 20)
        self.line.resize(200, 32)
        self.nameLabel.move(20, 20)


        self.nameLabel2 = QLabel(self)
        self.nameLabel2.setText('Sheet Path:')
        self.line2 = QLineEdit(self)

        self.line2.move(100, 60)
        self.line2.resize(200, 32)
        self.nameLabel2.move(20, 60)


        pybutton = QPushButton('Start', self)
        pybutton.clicked.connect(self.clickMethod)
        pybutton.resize(200,32)
        pybutton.move(100, 100)        

    def clickMethod(self):
        move_files(self.line.text(), self.line2.text())



if __name__ == "__main__":

    app = QtWidgets.QApplication(sys.argv)
    mainWin = MainWindow()
    mainWin.show()
    sys.exit( app.exec_() )

