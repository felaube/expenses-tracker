import os
import sys
from spreadsheet_handler import SpreadsheetHandler
from PyQt5.QtWidgets import QApplication, QPushButton, QMessageBox,\
                            QDoubleSpinBox, QWidget, QVBoxLayout,\
                            QHBoxLayout, QDateEdit, QDialog,\
                            QStyleFactory, QComboBox, QLineEdit,\
                            QCheckBox, QTabWidget, QGridLayout
from PyQt5.QtCore import QDate
from PyQt5.QtGui import QIcon
from fbs_runtime.application_context.PyQt5 import ApplicationContext


class WidgetGallery(QDialog):

    def __init__(self, parent=None, submit_icon=None):
        super(WidgetGallery, self).__init__(parent)

        tabsWidget = QTabWidget()

        self.createExpensesLayout()
        expensesWidget = QWidget()
        expensesWidget.setLayout(self.expensesLayout)

        tabsWidget.addTab(expensesWidget, "Expenses")

        self.layout = QVBoxLayout()
        self.layout.addWidget(tabsWidget)
        self.setLayout(self.layout)
        self.setWindowTitle("Expenses Tracker")

        QApplication.setStyle(QStyleFactory.create("Fusion"))

    def createExpensesLayout(self):
        self.expensesLayout = QGridLayout()

        self.doubleSpinBox = QDoubleSpinBox(maximum=1000, decimals=2,
                                            minimum=0)

        self.dateEdit = QDateEdit(calendarPopup=True, displayFormat="dd/MM/yy",
                                  date=QDate.currentDate())

        # TODO: Find out how to get insertPolicy working
        self.categoriesComboBox = QComboBox(insertPolicy=QComboBox.InsertAlphabetically)

        # TODO: Remove hardcoded categories. Read categories from file
        self.addCategories(["Bandejão", "Supermercado", "Contas", "Lanche",
                            "Almoço", "Ônibus", "Outros"])

        self.specificationLine = QLineEdit()

        self.observationLine = QLineEdit()

        self.expensesLayout.addWidget(self.doubleSpinBox, 0, 0)
        self.expensesLayout.addWidget(self.dateEdit, 0, 1)
        self.expensesLayout.addWidget(self.categoriesComboBox, 0, 2)
        self.expensesLayout.addWidget(self.specificationLine, 0, 3)
        self.expensesLayout.addWidget(self.observationLine, 0, 4)

        submitbutton = QPushButton("Submit Expense")
        submitbutton.clicked.connect(self.submitButtonClicked)

        self.vrCheckBox = QCheckBox("Gasto com VR")

        self.expensesLayout.addWidget(self.vrCheckBox, 1, 0)
        self.expensesLayout.addWidget(submitbutton, 2, 0, -1, -1)

    def addCategories(self, new_items):
        self.categoriesComboBox.insertItems(0, new_items)

    def submitButtonClicked(self):
        spreadsheet_hdl = SpreadsheetHandler()
        appctx = ApplicationContext()

        alert = QMessageBox()
        alert.setWindowTitle("Expense Submitted")
        alert.setWindowIcon(QIcon(appctx.get_resource("submit.ico")))
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
