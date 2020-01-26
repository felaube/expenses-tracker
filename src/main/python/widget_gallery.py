import os
import sys
import webbrowser
from spreadsheet_handler import SpreadsheetHandler
from PyQt5.QtWidgets import QApplication, QPushButton, QMessageBox,\
                            QDoubleSpinBox, QWidget, QVBoxLayout,\
                            QHBoxLayout, QDateEdit, QDialog,\
                            QStyleFactory, QComboBox, QLineEdit,\
                            QCheckBox, QTabWidget, QGridLayout,\
                            QLabel, QGroupBox, QBoxLayout
from PyQt5.QtCore import QDate, Qt
from PyQt5.QtGui import QIcon
from fbs_runtime.application_context.PyQt5 import ApplicationContext


class WidgetGallery(QDialog):

    def __init__(self, parent=None, submit_icon=None):
        super(WidgetGallery, self).__init__(parent)

        appctx = ApplicationContext()

        spreadsheet_hdl = SpreadsheetHandler()
        categories = spreadsheet_hdl.read_categories()

        tabsWidget = QTabWidget()

        self.createExpensesLayout()
        expensesWidget = QWidget()
        expensesWidget.setLayout(self.expensesLayout)

        tabsWidget.addTab(expensesWidget,
                          QIcon(appctx.get_resource("submit.ico")),
                          "Expenses")
        tabsWidget.setTabIcon

        self.createIncomesLayout()
        incomeWidget = QWidget()
        incomeWidget.setLayout(self.incomesLayout)

        tabsWidget.addTab(incomeWidget,
                          QIcon(appctx.get_resource("submit.ico")),
                          "Income")

        self.createSpreadsheetActionsLayout()
        spreadsheetActions = QWidget()
        spreadsheetActions.setLayout(self.spreadsheetActionsLayout)

        tabsWidget.addTab(spreadsheetActions,
                          QIcon(appctx.get_resource("sheets.ico")),
                          "Spreadsheet Actions")

        self.addCategories(categories)

        self.layout = QVBoxLayout()
        self.layout.addWidget(tabsWidget)
        self.setLayout(self.layout)
        self.setWindowTitle("Expenses Tracker")

        QApplication.setStyle(QStyleFactory.create("Fusion"))

    def createExpensesLayout(self):
        self.expensesLayout = QGridLayout()

        expenseDoubleSpinBox_label = QLabel("Value")
        expenseDoubleSpinBox_label.setAlignment(Qt.AlignBottom | Qt.AlignHCenter)
        self.expenseDoubleSpinBox = QDoubleSpinBox(maximum=1000, decimals=2,
                                                   minimum=0)

        expenseDateEdit_label = QLabel("Date")
        expenseDateEdit_label.setAlignment(Qt.AlignBottom | Qt.AlignHCenter)
        self.expenseDateEdit = QDateEdit(calendarPopup=True, displayFormat="dd/MM/yy",
                                         date=QDate.currentDate())

        expenseCategoriesComboBox_label = QLabel("Category")
        expenseCategoriesComboBox_label.setAlignment(Qt.AlignBottom | Qt.AlignHCenter)
        self.expenseCategoriesComboBox = QComboBox()

        expenseSpecificationLine_label = QLabel("Specification")
        expenseSpecificationLine_label.setAlignment(Qt.AlignBottom | Qt.AlignHCenter)
        self.expenseSpecificationLine = QLineEdit()

        expenseObservationLine_label = QLabel("Observation")
        expenseObservationLine_label.setAlignment(Qt.AlignBottom | Qt.AlignHCenter)
        self.expenseObservationLine = QLineEdit()

        submitbutton = QPushButton("Submit Expense")
        submitbutton.clicked.connect(self.submitExpenseButtonClicked)

        self.vrCheckBox = QCheckBox("Gasto com VR")

        self.expensesLayout.addWidget(expenseDoubleSpinBox_label, 0, 0)
        self.expensesLayout.addWidget(expenseDateEdit_label, 0, 1)
        self.expensesLayout.addWidget(expenseCategoriesComboBox_label, 0, 2)
        self.expensesLayout.addWidget(expenseSpecificationLine_label, 0, 3)
        self.expensesLayout.addWidget(expenseObservationLine_label, 0, 4)

        self.expensesLayout.addWidget(self.expenseDoubleSpinBox, 1, 0)
        self.expensesLayout.addWidget(self.expenseDateEdit, 1, 1)
        self.expensesLayout.addWidget(self.expenseCategoriesComboBox, 1, 2)
        self.expensesLayout.addWidget(self.expenseSpecificationLine, 1, 3)
        self.expensesLayout.addWidget(self.expenseObservationLine, 1, 4)
        self.expensesLayout.addWidget(self.vrCheckBox, 2, 0)
        self.expensesLayout.addWidget(submitbutton, 3, 0, -1, -1)

    def createSpreadsheetActionsLayout(self):
        self.spreadsheetActionsLayout = QVBoxLayout()

        access_spreadsheet_button = QPushButton("Access Spreadsheet")
        access_spreadsheet_button.clicked.connect(self.accessSpreadsheetButtonClicked)

        create_and_maintain_button = QPushButton("Create New Spreadsheet Maintaining the Old One")
        create_and_maintain_button.clicked.connect(self.createAndMaintainButtonClicked)

        create_and_delete_button = QPushButton("Create New Spreadsheet Deleting the Old One")
        create_and_delete_button.clicked.connect(self.createAndDeleteButtonClicked)

        self.addCategoryLine = QLineEdit()

        add_category_button = QPushButton("Add New Category")
        add_category_button.clicked.connect(self.addCategoryButtonClicked)

        self.delCategoryComboBox = QComboBox()

        del_category_button = QPushButton("Delete Category")
        del_category_button.clicked.connect(self.delCategoryButtonClicked)

        categories_group = QGroupBox(title="Expenses Categories")
        categories_layout = QHBoxLayout()

        add_categories_layout = QVBoxLayout()
        del_categories_layout = QVBoxLayout()

        add_categories_layout.addWidget(self.addCategoryLine)
        add_categories_layout.addWidget(add_category_button)

        del_categories_layout.addWidget(self.delCategoryComboBox)
        del_categories_layout.addWidget(del_category_button)

        add_categories_widget = QWidget()
        del_categories_widget = QWidget()

        add_categories_widget.setLayout(add_categories_layout)
        del_categories_widget.setLayout(del_categories_layout)

        categories_layout.addWidget(add_categories_widget)
        categories_layout.addWidget(del_categories_widget)

        categories_group.setLayout(categories_layout)

        self.spreadsheetActionsLayout.addWidget(access_spreadsheet_button)
        self.spreadsheetActionsLayout.addWidget(create_and_maintain_button)
        self.spreadsheetActionsLayout.addWidget(create_and_delete_button)
        self.spreadsheetActionsLayout.addWidget(categories_group)

    def createIncomesLayout(self):
        self.incomesLayout = QGridLayout()

        incomesDoubleSpinBox_label = QLabel("Value")
        incomesDoubleSpinBox_label.setAlignment(Qt.AlignBottom | Qt.AlignHCenter)
        self.incomesDoubleSpinBox = QDoubleSpinBox(maximum=1000, decimals=2,
                                                   minimum=0)

        incomesDateEdit_label = QLabel("Date")
        incomesDateEdit_label.setAlignment(Qt.AlignBottom | Qt.AlignHCenter)
        self.incomesDateEdit = QDateEdit(calendarPopup=True, displayFormat="dd/MM/yy",
                                         date=QDate.currentDate())

        incomesSpecificationLine_label = QLabel("Specification")
        incomesSpecificationLine_label.setAlignment(Qt.AlignBottom | Qt.AlignHCenter)
        self.incomesSpecificationLine = QLineEdit()

        incomesObservationLine_label = QLabel("Observation")
        incomesObservationLine_label.setAlignment(Qt.AlignBottom | Qt.AlignHCenter)
        self.incomesObservationLine = QLineEdit()

        submitbutton = QPushButton("Submit Income")
        submitbutton.clicked.connect(self.submitIncomeButtonClicked)

        self.incomesLayout.addWidget(incomesDoubleSpinBox_label, 0, 0)
        self.incomesLayout.addWidget(incomesDateEdit_label, 0, 1)
        self.incomesLayout.addWidget(incomesSpecificationLine_label, 0, 2)
        self.incomesLayout.addWidget(incomesObservationLine_label, 0, 3)

        self.incomesLayout.addWidget(self.incomesDoubleSpinBox, 1, 0)
        self.incomesLayout.addWidget(self.incomesDateEdit, 1, 1)
        self.incomesLayout.addWidget(self.incomesSpecificationLine, 1, 2)
        self.incomesLayout.addWidget(self.incomesObservationLine, 1, 3)
        self.incomesLayout.addWidget(submitbutton, 2, 0, -1, -1)

    def addCategories(self, new_items):
        self.expenseCategoriesComboBox.insertItems(0, new_items)
        self.delCategoryComboBox.insertItems(0, new_items)

    def resetCategoriesComboBox(self):
        spreadsheet_hdl = SpreadsheetHandler()

        categories = spreadsheet_hdl.read_categories()

        self.expenseCategoriesComboBox.clear()
        self.delCategoryComboBox.clear()

        self.addCategories(categories)

    def submitExpenseButtonClicked(self):
        spreadsheet_hdl = SpreadsheetHandler()
        appctx = ApplicationContext()

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

        spreadsheet_hdl.append_data(data, range="Expenses")
        spreadsheet_hdl.expenses_sort_by_date()

        alert = QMessageBox()
        alert.setWindowTitle("Expense Submitted")
        alert.setWindowIcon(QIcon(appctx.get_resource("submit.ico")))
        alert.setText("The expense was submitted!")
        alert.exec_()

    def submitIncomeButtonClicked(self):
        spreadsheet_hdl = SpreadsheetHandler()
        appctx = ApplicationContext()

        data = [
            ["=MONTH(\""+self.incomesDateEdit.date().toString("MM/dd/yyyy")+"\")",
             self.incomesDateEdit.date().toString("MM/dd/yyyy"),
             str(self.incomesDoubleSpinBox.value()),
             self.incomesSpecificationLine.text(),
             self.incomesObservationLine.text()
             ]
        ]

        spreadsheet_hdl.append_data(data, range="Income")
        spreadsheet_hdl.income_sort_by_date()

        alert = QMessageBox()
        alert.setWindowTitle("Income Submitted")
        alert.setWindowIcon(QIcon(appctx.get_resource("submit.ico")))
        alert.setText("The income was submitted!")
        alert.exec_()

    def accessSpreadsheetButtonClicked(self):
        spreadsheet_hdl = SpreadsheetHandler()
        webbrowser.open("https://docs.google.com/spreadsheets/d/" +
                        spreadsheet_hdl.spreadsheet_id)

    def createAndMaintainButtonClicked(self):
        appctx = ApplicationContext()

        spreadsheet_hdl = SpreadsheetHandler()
        spreadsheet_hdl.rename_spreadsheet(spreadsheet_hdl.file_name + '_OLD')
        spreadsheet_hdl.create_spreadsheet()

        self.resetCategoriesComboBox()

        alert = QMessageBox()
        alert.setWindowTitle("Spreadsheet Reset")
        alert.setWindowIcon(QIcon(appctx.get_resource("sheets.ico")))
        alert.setText("A new spreadsheet was created! To access the old "
                      "spreadsheet, look for Expenses Tracker_OLD in your Drive.")
        alert.exec_()

    def createAndDeleteButtonClicked(self):
        appctx = ApplicationContext()

        alert = QMessageBox()
        alert.setIcon(QMessageBox.Question)
        alert.setText("All information present "
                      "on the current spreadsheet "
                      "will be lost. "
                      "Are you sure you wish to continue?")
        alert.setWindowTitle("Spreadsheet Reset Confirmation")
        alert.setWindowIcon(QIcon(appctx.get_resource("sheets.ico")))
        alert.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        alert.setDefaultButton(QMessageBox.No)
        reply = alert.exec_()

        if reply == alert.Yes:
            spreadsheet_hdl = SpreadsheetHandler()
            spreadsheet_hdl.delete_spreadsheet()
            spreadsheet_hdl.create_spreadsheet()

            self.resetCategoriesComboBox()

            alert = QMessageBox()
            alert.setWindowTitle("Spreadsheet Reset")
            alert.setWindowIcon(QIcon(appctx.get_resource("sheets.ico")))
            alert.setText("A new spreadsheet was created!")
            alert.exec_()

    def addCategoryButtonClicked(self):
        spreadsheet_hdl = SpreadsheetHandler()
        appctx = ApplicationContext()

        new_category = self.addCategoryLine.text()

        categories = spreadsheet_hdl.read_categories()

        if new_category in categories:
            alert = QMessageBox()
            alert.setWindowTitle("Category Adding")
            alert.setWindowIcon(QIcon(appctx.get_resource("sheets.ico")))
            alert.setText("The category " + new_category + " already exists.")
            alert.exec_()
            return

        spreadsheet_hdl.add_category(new_category)

        self.resetCategoriesComboBox()

        alert = QMessageBox()
        alert.setWindowTitle("Category Adding")
        alert.setWindowIcon(QIcon(appctx.get_resource("sheets.ico")))
        alert.setText("The category " + new_category + " was succesfully added!")
        alert.exec_()

    def delCategoryButtonClicked(self):
        spreadsheet_hdl = SpreadsheetHandler()
        appctx = ApplicationContext()

        category_to_be_del = self.delCategoryComboBox.currentText()

        spreadsheet_hdl.delete_category(category_to_be_del)

        self.resetCategoriesComboBox()

        alert = QMessageBox()
        alert.setWindowTitle("Category Deleting")
        alert.setWindowIcon(QIcon(appctx.get_resource("sheets.ico")))
        alert.setText("The category " + category_to_be_del + " was succesfully deleted!")
        alert.exec_()
