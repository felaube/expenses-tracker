import os
import sys
import webbrowser
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

        self.createIncomesLayout()
        incomeWidget = QWidget()
        incomeWidget.setLayout(self.incomesLayout)

        tabsWidget.addTab(incomeWidget, "Income")
        
        self.createSpreadsheetActionsLayout()
        spreadsheetActions = QWidget()
        spreadsheetActions.setLayout(self.spreadsheetActionsLayout)

        tabsWidget.addTab(spreadsheetActions, "Spreadsheet Actions")

        self.layout = QVBoxLayout()
        self.layout.addWidget(tabsWidget)
        self.setLayout(self.layout)
        self.setWindowTitle("Expenses Tracker")

        QApplication.setStyle(QStyleFactory.create("Fusion"))

    def createExpensesLayout(self):
        self.expensesLayout = QGridLayout()

        self.expenseDoubleSpinBox = QDoubleSpinBox(maximum=1000, decimals=2,
                                            minimum=0)

        self.expenseDateEdit = QDateEdit(calendarPopup=True, displayFormat="dd/MM/yy",
                                         date=QDate.currentDate())

        # TODO: Find out how to get insertPolicy working
        self.expenseCategoriesComboBox = QComboBox(insertPolicy=QComboBox.InsertAlphabetically)

        # TODO: Remove hardcoded categories. Read categories from file
        self.addCategories(["Bandejão", "Supermercado", "Contas", "Lanche",
                            "Almoço", "Ônibus", "Outros"])

        self.expenseSpecificationLine = QLineEdit()

        self.expenseObservationLine = QLineEdit()

        submitbutton = QPushButton("Submit Expense")
        submitbutton.clicked.connect(self.submitExpenseButtonClicked)

        self.vrCheckBox = QCheckBox("Gasto com VR")

        self.expensesLayout.addWidget(self.expenseDoubleSpinBox, 0, 0)
        self.expensesLayout.addWidget(self.expenseDateEdit, 0, 1)
        self.expensesLayout.addWidget(self.expenseCategoriesComboBox, 0, 2)
        self.expensesLayout.addWidget(self.expenseSpecificationLine, 0, 3)
        self.expensesLayout.addWidget(self.expenseObservationLine, 0, 4)
        self.expensesLayout.addWidget(self.vrCheckBox, 1, 0)
        self.expensesLayout.addWidget(submitbutton, 2, 0, -1, -1)

    def createSpreadsheetActionsLayout(self):
        self.spreadsheetActionsLayout = QVBoxLayout()

        access_spreadsheet_button = QPushButton("Access Spreadsheet")
        access_spreadsheet_button.clicked.connect(self.accessSpreadsheetButtonClicked)

        create_and_maintain_button = QPushButton("Create New Spreadsheet Maintaining the Old One")
        create_and_maintain_button.clicked.connect(self.createAndMaintainButtonClicked)

        create_and_delete_button = QPushButton("Create New Spreadsheet Deleting the Old One")
        create_and_delete_button.clicked.connect(self.createAndDeleteButtonClicked)

        self.spreadsheetActionsLayout.addWidget(access_spreadsheet_button)
        self.spreadsheetActionsLayout.addWidget(create_and_maintain_button)
        self.spreadsheetActionsLayout.addWidget(create_and_delete_button)


    def createIncomesLayout(self):
        self.incomesLayout = QGridLayout()

        self.incomesDoubleSpinBox = QDoubleSpinBox(maximum=1000, decimals=2,
                                                   minimum=0)

        self.incomesDateEdit = QDateEdit(calendarPopup=True, displayFormat="dd/MM/yy",
                                         date=QDate.currentDate())

        self.incomesSpecificationLine = QLineEdit()

        self.incomesObservationLine = QLineEdit()

        submitbutton = QPushButton("Submit Income")
        submitbutton.clicked.connect(self.submitIncomeButtonClicked)

        self.incomesLayout.addWidget(self.incomesDoubleSpinBox, 0, 0)
        self.incomesLayout.addWidget(self.incomesDateEdit, 0, 1)
        self.incomesLayout.addWidget(self.incomesSpecificationLine, 0, 2)
        self.incomesLayout.addWidget(self.incomesObservationLine, 0, 3)
        self.incomesLayout.addWidget(submitbutton, 1, 0, -1, -1)

    def addCategories(self, new_items):
        self.expenseCategoriesComboBox.insertItems(0, new_items)

    def submitExpenseButtonClicked(self):
        spreadsheet_hdl = SpreadsheetHandler()
        appctx = ApplicationContext()

        alert = QMessageBox()
        alert.setWindowTitle("Expense Submitted")
        alert.setWindowIcon(QIcon(appctx.get_resource("submit.ico")))
        data = [
            ["=MONTH(\""+self.expenseDateEdit.date().toString("MM/dd/yyyy")+"\")",
             self.expenseDateEdit.date().toString("MM/dd/yyyy"),
             str(self.expenseDoubleSpinBox.value()),
             self.expenseCategoriesComboBox.currentText(),
             self.expenseSpecificationLine.text(),
             self.expenseObservationLine.text()
             ]
        ]

        if self.vrCheckBox.checkState() == 2:
            data[0].append("VR")

        spreadsheet_hdl.append_data(data, range='B4')
        alert.setText("The expense was submitted!")
        alert.exec_()

    def submitIncomeButtonClicked(self):
        spreadsheet_hdl = SpreadsheetHandler()
        appctx = ApplicationContext()

        alert = QMessageBox()
        alert.setWindowTitle("Income Submitted")
        alert.setWindowIcon(QIcon(appctx.get_resource("submit.ico")))
        data = [
            ["=MONTH(\""+self.incomesDateEdit.date().toString("MM/dd/yyyy")+"\")",
             self.incomesDateEdit.date().toString("MM/dd/yyyy"),
             str(self.incomesDoubleSpinBox.value()),
             self.incomesSpecificationLine.text(),
             self.incomesObservationLine.text()
             ]
        ]

        # spreadsheet_hdl.append_data(data, range='B4')
        alert.setText("The income was submitted!")
        alert.exec_()

    def accessSpreadsheetButtonClicked(self):
        spreadsheet_hdl = SpreadsheetHandler()
        webbrowser.open("https://docs.google.com/spreadsheets/d/" + 
                        spreadsheet_hdl.spreadsheet_id)

    def createAndMaintainButtonClicked(self):
        spreadsheet_hdl = SpreadsheetHandler()
        spreadsheet_hdl.rename_spreadsheet(spreadsheet_hdl.file_name + '_OLD')
        spreadsheet_hdl.create_spreadsheet()
        
        alert = QMessageBox()
        alert.setWindowTitle("Spreadsheet Reset")
        alert.setText("A new spreadsheet was created! To access the old "
                      "spreadsheet, look for Expenses Tracker_OLD in your Drive.")
        alert.exec_()

    def createAndDeleteButtonClicked(self):
        alert = QMessageBox.question(self, "Spreadsheet Reset",
                                     "All information present "
                                     "on the current spreadsheet "
                                     "will be lost. "
                                     "Are you sure you wish to continue?",
                                     QMessageBox.Yes | QMessageBox.No,
                                     QMessageBox.No)

        if alert == QMessageBox.Yes:
            spreadsheet_hdl = SpreadsheetHandler()
            spreadsheet_hdl.delete_spreadsheet()
            spreadsheet_hdl.create_spreadsheet()

            alert = QMessageBox()
            alert.setWindowTitle("Spreadsheet Reset")
            alert.setText("A new spreadsheet was created!")
            alert.exec_()
