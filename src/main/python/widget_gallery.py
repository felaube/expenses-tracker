import webbrowser
from spreadsheet_handler import SpreadsheetHandler
from PyQt5.QtWidgets import QApplication, QPushButton, QMessageBox,\
                            QDoubleSpinBox, QWidget, QVBoxLayout,\
                            QHBoxLayout, QDateEdit, QDialog,\
                            QStyleFactory, QComboBox, QLineEdit,\
                            QCheckBox, QTabWidget, QGridLayout,\
                            QLabel, QGroupBox, QBoxLayout,\
                            QDesktopWidget, QSizePolicy, QTableWidget,\
                            QTableWidgetItem, QAbstractItemView
from PyQt5.QtCore import QDate, Qt, QRect
from PyQt5.QtGui import QIcon, QFont
from fbs_runtime.application_context.PyQt5 import ApplicationContext


class WidgetGallery(QDialog):

    def __init__(self, parent=None, submit_icon=None):
        super(WidgetGallery, self).__init__(parent)

        appctx = ApplicationContext()

        spreadsheet_hdl = SpreadsheetHandler()
        categories = spreadsheet_hdl.read_categories()

        self.tabsWidget = QTabWidget()

        # Add 'Expenses' tab to main widget
        self.createExpensesLayout()
        self.tabsWidget.addTab(self.expensesWidget,
                               QIcon(appctx.get_resource("submit.ico")),
                               "Expenses")

        # Add 'Incomes' tab to main widget
        self.createIncomesLayout()
        self.tabsWidget.addTab(self.incomesWidget,
                               QIcon(appctx.get_resource("submit.ico")),
                               "Incomes")

        # Add 'Latest Uploads' tab to main widget
        self.createLatestUploads()
        self.tabsWidget.addTab(self.latestUploadsWidget,
                               QIcon(appctx.get_resource("sheets.ico")),
                               "Latest Uploads")

        # Add 'Spreadsheet Actions' tab to main widget
        self.createSpreadsheetActionsLayout()
        self.tabsWidget.addTab(self.spreadsheetActionsWidget,
                               QIcon(appctx.get_resource("sheets.ico")),
                               "Spreadsheet Actions")

        # Set the current available expenses categories
        self.addCategories(categories)

        # Set main window size
        self.resize(570, 320)

        self.tabsWidget.currentChanged.connect(self.adjustTabWidgetSize)

        self.layout = QVBoxLayout()
        self.layout.addWidget(self.tabsWidget)
        self.setLayout(self.layout)
        self.setWindowTitle("Expenses Tracker")

        QApplication.setStyle(QStyleFactory.create("Fusion"))

    def createExpensesLayout(self):
        self.expensesWidget = QWidget()
        self.expensesWidget.setGeometry(QRect(10, 10, 550, 300))
        self.expensesWidget.size = (570, 320)

        expenseDoubleSpinBox_label = QLabel("Value", self.expensesWidget)
        expenseDoubleSpinBox_label.setGeometry(QRect(30, 80, 40, 20))
        expenseDoubleSpinBox_label.setAlignment(Qt.AlignBottom | Qt.AlignHCenter)
        self.expenseDoubleSpinBox = QDoubleSpinBox(self.expensesWidget)
        self.expenseDoubleSpinBox.setMaximum(10000)
        self.expenseDoubleSpinBox.setDecimals(2)
        self.expenseDoubleSpinBox.setMinimum(0)
        self.expenseDoubleSpinBox.setGeometry(QRect(10, 100, 80, 20))

        expenseDateEdit_label = QLabel("Date", self.expensesWidget)
        expenseDateEdit_label.setGeometry(QRect(120, 80, 40, 20))
        expenseDateEdit_label.setAlignment(Qt.AlignBottom | Qt.AlignHCenter)
        self.expenseDateEdit = QDateEdit(self.expensesWidget)
        self.expenseDateEdit.setCalendarPopup(True)
        self.expenseDateEdit.setDisplayFormat("dd/MM/yy")
        self.expenseDateEdit.setDate(QDate.currentDate())
        self.expenseDateEdit.setGeometry(QRect(100, 100, 80, 20))

        expenseCategoriesComboBox_label = QLabel("Category", self.expensesWidget)
        expenseCategoriesComboBox_label.setGeometry(QRect(210, 80, 60, 20))
        expenseCategoriesComboBox_label.setAlignment(Qt.AlignBottom | Qt.AlignHCenter)
        self.expenseCategoriesComboBox = QComboBox(self.expensesWidget)
        self.expenseCategoriesComboBox.setGeometry(QRect(190, 100, 100, 20))

        expenseSpecificationLine_label = QLabel("Specification", self.expensesWidget)
        expenseSpecificationLine_label.setGeometry(QRect(320, 80, 70, 20))
        expenseSpecificationLine_label.setAlignment(Qt.AlignBottom | Qt.AlignHCenter)
        self.expenseSpecificationLine = QLineEdit(self.expensesWidget)
        self.expenseSpecificationLine.setGeometry(QRect(300, 100, 115, 20))

        expenseObservationLine_label = QLabel("Observation", self.expensesWidget)
        expenseObservationLine_label.setGeometry(QRect(440, 80, 80, 20))
        expenseObservationLine_label.setAlignment(Qt.AlignBottom | Qt.AlignHCenter)
        self.expenseObservationLine = QLineEdit(self.expensesWidget)
        self.expenseObservationLine.setGeometry(QRect(420, 100, 115, 20))

        self.vrCheckBox = QCheckBox("Gasto com VR", self.expensesWidget)
        self.vrCheckBox.setGeometry(QRect(10, 140, 110, 20))

        submitbutton = QPushButton("Submit Expense", self.expensesWidget)
        submitbutton.clicked.connect(self.submitExpenseButtonClicked)
        submitbutton.setGeometry(QRect(10, 170, 520, 25))

    def createLatestUploads(self):
        self.latestUploadsWidget = QWidget()
        self.latestUploadsWidget.setGeometry(QRect(10, 10, 550, 300))
        self.latestUploadsWidget.size = (824, 403)

        # Construct Latest Expenses Group
        expenses_group = QGroupBox("Latest Expenses", self.latestUploadsWidget)

        self.expensesTable = QTableWidget(expenses_group)
        self.expensesTable.setColumnCount(4)
        self.expensesTable.verticalHeader().setVisible(False)
        self.expensesTable.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.expensesTable.setSelectionMode(QAbstractItemView.NoSelection)
        self.expensesTable.setFixedSize(415, 265)
        self.expensesTable.setHorizontalHeaderLabels(["Date", "Value",
                                                      "Category", "Specification"])

        self.fillExpensesTableData()

        update_expenses_table_button = QPushButton("Update Expenses Table")
        update_expenses_table_button.clicked.connect(self.updateExpensesTableButtonClicked)

        layout = QVBoxLayout()
        layout.addWidget(self.expensesTable)
        layout.addWidget(update_expenses_table_button)

        expenses_group.setLayout(layout)
        expenses_group.move(10, 5)
        expenses_group.setStyleSheet("QGroupBox { font-weight: bold; } ")

        # Construct Latest Incomes Group
        incomes_group = QGroupBox("Latest Incomes", self.latestUploadsWidget)

        self.incomesTable = QTableWidget(incomes_group)
        self.incomesTable.setColumnCount(3)
        self.incomesTable.verticalHeader().setVisible(False)
        self.incomesTable.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.incomesTable.setSelectionMode(QAbstractItemView.NoSelection)
        self.incomesTable.setFixedSize(301, 265)
        self.incomesTable.setHorizontalHeaderLabels(["Date", "Value",
                                                     "Category"])

        self.fillIncomesTableData()

        update_incomes_table_button = QPushButton("Update Incomes Table")
        update_incomes_table_button.clicked.connect(self.updateIncomesTableButtonClicked)

        layout = QVBoxLayout()
        layout.addWidget(self.incomesTable)
        layout.addWidget(update_incomes_table_button)

        incomes_group.setLayout(layout)
        incomes_group.move(460, 5)
        incomes_group.setStyleSheet("QGroupBox { font-weight: bold; } ")

    def createIncomesLayout(self):
        self.incomesWidget = QWidget()
        self.incomesWidget.setGeometry(QRect(10, 10, 460, 300))
        self.incomesWidget.size = (480, 320)

        incomesDoubleSpinBox_label = QLabel("Value", self.incomesWidget)
        incomesDoubleSpinBox_label.setAlignment(Qt.AlignBottom | Qt.AlignHCenter)
        incomesDoubleSpinBox_label.setGeometry(QRect(40, 80, 40, 20))
        self.incomesDoubleSpinBox = QDoubleSpinBox(self.incomesWidget)
        self.incomesDoubleSpinBox.setMaximum(10000)
        self.incomesDoubleSpinBox.setDecimals(2)
        self.incomesDoubleSpinBox.setMinimum(0)
        self.incomesDoubleSpinBox.setGeometry(QRect(20, 100, 80, 20))

        incomesDateEdit_label = QLabel("Date", self.incomesWidget)
        incomesDateEdit_label.setAlignment(Qt.AlignBottom | Qt.AlignHCenter)
        incomesDateEdit_label.setGeometry(QRect(130, 80, 40, 20))
        self.incomesDateEdit = QDateEdit(self.incomesWidget)
        self.incomesDateEdit.setCalendarPopup(True)
        self.incomesDateEdit.setDisplayFormat("dd/MM/yy")
        self.incomesDateEdit.setDate(QDate.currentDate())
        self.incomesDateEdit.setGeometry(QRect(110, 100, 80, 20))

        incomesSpecificationLine_label = QLabel("Specification", self.incomesWidget)
        incomesSpecificationLine_label.setAlignment(Qt.AlignBottom | Qt.AlignHCenter)
        incomesSpecificationLine_label.setGeometry(QRect(220, 80, 70, 20))
        self.incomesSpecificationLine = QLineEdit(self.incomesWidget)
        self.incomesSpecificationLine.setGeometry(QRect(200, 100, 115, 20))

        incomesObservationLine_label = QLabel("Observation", self.incomesWidget)
        incomesObservationLine_label.setAlignment(Qt.AlignBottom | Qt.AlignHCenter)
        incomesObservationLine_label.setGeometry(QRect(340, 80, 80, 20))
        self.incomesObservationLine = QLineEdit(self.incomesWidget)
        self.incomesObservationLine.setGeometry(QRect(320, 100, 115, 20))

        submitbutton = QPushButton("Submit Income", self.incomesWidget)
        submitbutton.setGeometry(QRect(10, 170, 430, 25))
        submitbutton.clicked.connect(self.submitIncomeButtonClicked)

    def createSpreadsheetActionsLayout(self):
        self.spreadsheetActionsWidget = QWidget()
        self.spreadsheetActionsWidget.setGeometry(QRect(10, 10, 550, 300))
        self.spreadsheetActionsWidget.size = (570, 320)

        access_spreadsheet_button = QPushButton("Access Spreadsheet",
                                                self.spreadsheetActionsWidget)
        access_spreadsheet_button.clicked.connect(self.accessSpreadsheetButtonClicked)
        access_spreadsheet_button.setGeometry(QRect(10, 10, 525, 25))

        create_and_maintain_button = QPushButton("Create New Spreadsheet Maintaining the Old One",
                                                 self.spreadsheetActionsWidget)
        create_and_maintain_button.clicked.connect(self.createAndMaintainButtonClicked)
        create_and_maintain_button.setGeometry(QRect(10, 40, 525, 25))

        create_and_delete_button = QPushButton("Create New Spreadsheet Deleting the Old One",
                                               self.spreadsheetActionsWidget)
        create_and_delete_button.clicked.connect(self.createAndDeleteButtonClicked)
        create_and_delete_button.setGeometry(QRect(10, 70, 525, 25))

        categories_group = QGroupBox("Expenses Categories", self.spreadsheetActionsWidget)
        categories_group.setGeometry(QRect(10, 110, 525, 140))
        categories_group.setStyleSheet("QGroupBox { font-weight: bold; } ")
        categories_layout = QHBoxLayout()

        self.addCategoryLine = QLineEdit()

        add_category_button = QPushButton("Add New Category")
        add_category_button.clicked.connect(self.addCategoryButtonClicked)

        self.delCategoryComboBox = QComboBox()

        del_category_button = QPushButton("Delete Category")
        del_category_button.clicked.connect(self.delCategoryButtonClicked)

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

    def fillExpensesTableData(self):
        latest_expenses = SpreadsheetHandler().get_latest_upload("expenses")

        self.expensesTable.setRowCount(len(latest_expenses))

        for row_index, row in enumerate(latest_expenses):
            for column_index, cell_content in enumerate(row):
                self.expensesTable.setItem(row_index, column_index,
                                           QTableWidgetItem(cell_content))

    def fillIncomesTableData(self):
        latest_incomes = SpreadsheetHandler().get_latest_upload("incomes")

        self.incomesTable.setRowCount(len(latest_incomes))

        for row_index, row in enumerate(latest_incomes):
            for column_index, cell_content in enumerate(row):
                self.incomesTable.setItem(row_index, column_index,
                                          QTableWidgetItem(cell_content))

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

        for index in range(len(data[0])):
            if data[0][index] == "":
                data[0][index] = "-"

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

        for index in range(len(data[0])):
            if data[0][index] == "":
                data[0][index] = "-"

        spreadsheet_hdl.append_data(data, range="Incomes")
        spreadsheet_hdl.income_sort_by_date()

        alert = QMessageBox()
        alert.setWindowTitle("Income Submitted")
        alert.setWindowIcon(QIcon(appctx.get_resource("submit.ico")))
        alert.setText("The income was submitted!")
        alert.exec_()

    def updateExpensesTableButtonClicked(self):
        self.expensesTable.clearContents()
        self.fillExpensesTableData()

    def updateIncomesTableButtonClicked(self):
        self.incomesTable.clearContents()
        self.fillIncomesTableData()

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

    def adjustTabWidgetSize(self):
        current_tab = self.tabsWidget.currentWidget()
        self.resize(current_tab.size[0],
                    current_tab.size[1])
