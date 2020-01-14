import os
import sys
from spreadsheet_handler import SpreadsheetHandler
from PyQt5.QtWidgets import QApplication, QPushButton, QMessageBox,\
                            QDoubleSpinBox, QWidget, QVBoxLayout,\
                            QHBoxLayout, QDateEdit, QDialog,\
                            QStyleFactory, QComboBox, QLineEdit,\
                            QCheckBox
from PyQt5.QtCore import QDate
from PyQt5.QtGui import QIcon


class WidgetGallery(QDialog):

    def __init__(self, parent=None):
        super(WidgetGallery, self).__init__(parent)

        mainLayout = QVBoxLayout()

        self.createHorizontalLayout()

        submitbutton = QPushButton("Submit Expense")
        submitbutton.clicked.connect(self.submitButtonClicked)

        self.vrCheckBox = QCheckBox("Gasto com VR")

        mainLayout.addLayout(self.horizontalLayout)
        mainLayout.addWidget(self.vrCheckBox)
        mainLayout.addWidget(submitbutton)

        self.setLayout(mainLayout)
        self.setWindowTitle("Expenses Tracker")

        QApplication.setStyle(QStyleFactory.create("Fusion"))

    def addCategories(self, new_items):
        self.categoriesComboBox.insertItems(0, new_items)

    def submitButtonClicked(self):
        spreadsheet_hdl = SpreadsheetHandler()
        alert = QMessageBox()
        alert.setWindowTitle("Expense Submitted")
        alert.setWindowIcon(QIcon(os.path.join(os.path.dirname(os.getcwd()),
                                               "icons\\submit.ico")))
        data = [
            ["=MONTH(\""+self.dateEdit.date().toString("MM/dd/yyyy")+"\")",
             self.dateEdit.date().toString("MM/dd/yyyy"),
             str(self.doubleSpinBox.value()),
             self.categoriesComboBox.currentText(),
             self.specificationLine.text(),
             self.observationLine.text()
             ]
        ]

        if self.vrCheckBox.checkState() == 2:
            data[0].append("VR")

        spreadsheet_hdl.append_data(data, range='B4')
        alert.setText("The expense was submitted!")
        alert.exec_()

    def createHorizontalLayout(self):
        self.horizontalLayout = QHBoxLayout()

        self.dateEdit = QDateEdit(calendarPopup=True, displayFormat="dd/MM/yy",
                                  date=QDate.currentDate())

        self.doubleSpinBox = QDoubleSpinBox(maximum=1000, decimals=2,
                                            minimum=0)

        # TODO: Find out how to get insertPolicy working
        self.categoriesComboBox = QComboBox(insertPolicy=QComboBox.InsertAlphabetically)

        # TODO: Remove hardcoded categories. Read categories from file
        self.addCategories(["Bandejão", "Supermercado", "Contas", "Lanche",
                            "Almoço", "Ônibus", "Outros"])

        self.specificationLine = QLineEdit()

        self.observationLine = QLineEdit()

        self.horizontalLayout.addWidget(self.doubleSpinBox)
        self.horizontalLayout.addWidget(self.dateEdit)
        self.horizontalLayout.addWidget(self.categoriesComboBox)
        self.horizontalLayout.addWidget(self.specificationLine)
        self.horizontalLayout.addWidget(self.observationLine)
