# from PyQt5.QtWidgets import *
# from PyQt5.QtGui import *
# from PyQt5.QtCore import *
# from PyQt5.uic import loadUi
import sys
from libs.PyQt5 import QtCore
from libs.PyQt5 import QtWidgets
from libs.PyQt5 import QtGui
from libs.PyQt5 import uic
from createDocuments import CreateDocuments


class Four(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        super(Four, self).__init__(parent)
        uic.loadUi('window4.ui', self)
        self.back.clicked.connect(self.on_back)
        self.save.clicked.connect(self.file_save1)
        self.save_2.clicked.connect(self.file_save2)
        self.save_3.clicked.connect(self.file_save3)
        self.save_4.clicked.connect(self.file_save4)
        self.save_5.clicked.connect(self.file_save5)
        self.save_6.clicked.connect(self.file_save6)

    def on_back(self):
        settings = QtCore.QSettings('myorg', 'myapp')

        settings.remove('wbs')
        theclass = Third(self)
        theclass.restoreSettings()
        Four.hide(self)
        theclass.show()

    def file_save1(self):
        settings = QtCore.QSettings('myorg', 'myapp')

        wb = settings.value('wb1', )

        create = CreateDocuments()
        name = QtWidgets.QFileDialog.getSaveFileName(self, 'Save File')
        if name[0]:
            print(name[0])
            create.save(wb, name[0])
        else:
            pass

    def file_save2(self):
        settings = QtCore.QSettings('myorg', 'myapp')
        wb = settings.value('wb2', )
        create = CreateDocuments()
        name = QtWidgets.QFileDialog.getSaveFileName(self, 'Save File')
        if name[0]:
            print(name[0])
            create.save(wb, name[0])
        else:
            pass

    def file_save3(self):
        settings = QtCore.QSettings('myorg', 'myapp')
        wb = settings.value('wb3', )
        create = CreateDocuments()
        name = QtWidgets.QFileDialog.getSaveFileName(self, 'Save File')
        if name[0]:
            print(name[0])
            create.save(wb, name[0])
        else:
            pass

    def file_save4(self):
        settings = QtCore.QSettings('myorg', 'myapp')
        wb = settings.value('wb4', )
        create = CreateDocuments()
        name = QtWidgets.QFileDialog.getSaveFileName(self, 'Save File')
        if name[0]:
            print(name[0])
            create.save(wb, name[0])
        else:
            pass

    def file_save5(self):
        settings = QtCore.QSettings('myorg', 'myapp')
        wb = settings.value('wb5', )
        create = CreateDocuments()
        name = QtWidgets.QFileDialog.getSaveFileName(self, 'Save File')
        if name[0]:
            print(name[0])
            create.save(wb, name[0])
        else:
            pass

    def file_save6(self):
        settings = QtCore.QSettings('myorg', 'myapp')
        wb = settings.value('wb6', )
        create = CreateDocuments()
        name = QtWidgets.QFileDialog.getSaveFileName(self, 'Save File')
        if name[0]:
            create.save(wb, name[0])
        else:
            pass


class Third(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        super(Third, self).__init__(parent)
        uic.loadUi('window3.ui', self)
        reg_ex = QtCore.QRegExp('^[0-9]{1,2}$')
        validator = QtGui.QRegExpValidator(reg_ex)
        self.supplieNum.setValidator(validator)
        self.confirm.clicked.connect(self.on_supplieChange)
        self.clear.clicked.connect(self.clear_table)
        self.back.clicked.connect(self.on_back)
        self.next.clicked.connect(self.on_next)
        self.actionPlan.triggered.connect(self.on_settingPlan)
        self.actionProduce.triggered.connect(self.on_settingProduce)

    def on_settingPlan(self):
        dialog = SettingPlan(self)
        dialog.show()

    def on_settingProduce(self):
        dialog = SettingProduce(self)
        dialog.show()

    def on_supplieLine(self, part):
        rows = self.supplieTable.rowCount()
        data = []
        for i in range(rows):
            it = self.supplieTable.cellWidget(i, 0)
            if it and it.text():
                data.append(it.text())

        for i in range(len(part)):
            supplieItem = str(part[i])
            if supplieItem in data:
                pass
            else:
                reg_ex = QtCore.QRegExp('^[0-9]{1,100}$')
                validator = QtGui.QRegExpValidator(reg_ex)
                lineEdit = QtWidgets.QLineEdit()
                lineEdit.setValidator(validator)
                rowPosition = self.supplieTable.rowCount()
                self.supplieTable.insertRow(rowPosition)
                self.supplieTable.setCellWidget(rowPosition, 0, QtWidgets.QLineEdit(supplieItem))
                self.supplieTable.setCellWidget(rowPosition, 1, QtWidgets.QLineEdit())
                self.supplieTable.setCellWidget(rowPosition, 2, lineEdit)
                self.supplieTable.setCellWidget(rowPosition, 3, QtWidgets.QLineEdit())
                self.supplieTable.setCellWidget(rowPosition, 4, QtWidgets.QDateEdit(QtCore.QDate(2020, 1, 1)))
                self.supplieTable.setCellWidget(rowPosition, 5, QtWidgets.QDateEdit(QtCore.QDate(2020, 1, 1)))
                self.supplieTable.setCellWidget(rowPosition, 6, QtWidgets.QDateEdit(QtCore.QDate(2020, 1, 1)))
                self.supplieTable.setCellWidget(rowPosition, 7, QtWidgets.QDateEdit(QtCore.QDate(2020, 1, 1)))

    def on_supplieChange(self):

        supplieNum = self.supplieNum.text()
        if supplieNum:
            supplieNum = int(supplieNum)

            for i in range(0, supplieNum):
                reg_ex = QtCore.QRegExp('^[0-9]{1,100}$')
                validator = QtGui.QRegExpValidator(reg_ex)
                lineEdit = QtWidgets.QLineEdit()
                lineEdit.setValidator(validator)
                rowPosition = self.supplieTable.rowCount()
                self.supplieTable.insertRow(rowPosition)
                self.supplieTable.setCellWidget(rowPosition, 0, QtWidgets.QLineEdit())
                self.supplieTable.setCellWidget(rowPosition, 1, QtWidgets.QLineEdit())
                self.supplieTable.setCellWidget(rowPosition, 2, lineEdit)
                self.supplieTable.setCellWidget(rowPosition, 3, QtWidgets.QLineEdit())
                self.supplieTable.setCellWidget(rowPosition, 4, QtWidgets.QDateEdit(QtCore.QDate(2020, 1, 1)))
                self.supplieTable.setCellWidget(rowPosition, 5, QtWidgets.QDateEdit(QtCore.QDate(2020, 1, 1)))
                self.supplieTable.setCellWidget(rowPosition, 6, QtWidgets.QDateEdit(QtCore.QDate(2020, 1, 1)))
                self.supplieTable.setCellWidget(rowPosition, 7, QtWidgets.QDateEdit(QtCore.QDate(2020, 1, 1)))
        else:
            pass

    def clear_table(self):
        self.supplieTable.setRowCount(0)

    def saveInstance(self):
        settings = QtCore.QSettings('myorg', 'myapp')

        rows = self.supplieTable.rowCount()

        if rows == 0:
            pass
        else:

            data = []
            for row in range(rows):
                data.append([])
                for column in range(0, 4):
                    it = self.supplieTable.cellWidget(row, column)
                    if it and it.text():
                        pass
                    else:
                        item = QtWidgets.QLineEdit()
                        item.setText("")
                        self.supplieTable.setCellWidget(row, column, item)
                    data[row].append(self.supplieTable.cellWidget(row, column).text())
                for column in range(4, 8):
                    data[row].append(self.supplieTable.cellWidget(row, column).date())

            settings.setValue('bang3', data)

    def restoreSettings(self):
        settings = QtCore.QSettings('myorg', 'myapp')

        data = []
        data = settings.value('bang3', data)
        numrows = len(data)
        if numrows == 0:
            pass
        else:
            for row in range(numrows):
                reg_ex = QtCore.QRegExp('^[0-9]{1,100}$')
                validator = QtGui.QRegExpValidator(reg_ex)
                lineEdit = QtWidgets.QLineEdit(str(data[row][2]))
                lineEdit.setValidator(validator)
                rowPosition = self.supplieTable.rowCount()
                self.supplieTable.insertRow(rowPosition)
                self.supplieTable.setCellWidget(rowPosition, 0, QtWidgets.QLineEdit(str(data[row][0])))
                self.supplieTable.setCellWidget(rowPosition, 1, QtWidgets.QLineEdit(str(data[row][1])))
                self.supplieTable.setCellWidget(rowPosition, 2, lineEdit)
                self.supplieTable.setCellWidget(rowPosition, 3, QtWidgets.QLineEdit(str(data[row][3])))
                self.supplieTable.setCellWidget(rowPosition, 4, QtWidgets.QDateEdit(QtCore.QDate(data[row][4])))
                self.supplieTable.setCellWidget(rowPosition, 5, QtWidgets.QDateEdit(QtCore.QDate(data[row][5])))
                self.supplieTable.setCellWidget(rowPosition, 6, QtWidgets.QDateEdit(QtCore.QDate(data[row][6])))
                self.supplieTable.setCellWidget(rowPosition, 7, QtWidgets.QDateEdit(QtCore.QDate(data[row][7])))

    def on_back(self):
        Third.saveInstance(self)
        theclass = Second(self)
        theclass.restoreSettings()
        Third.hide(self)
        theclass.show()

    def on_next(self):
        Third.saveInstance(self)
        theclass = Four(self)

        settings = QtCore.QSettings('myorg', 'myapp')
        orderDate = settings.value('dateOrder', '')
        deliDate = settings.value('deliveryDate', '')
        name = settings.value('ten', '')
        data = []
        data = settings.value('bang', data)

        settings1 = QtCore.QSettings('myorg', 'mysetting')
        hour = settings1.value('minWorkh', '')
        hour = int(hour)
        closingDay = settings1.value('closingDay', '')
        closingDay = 0 - int(closingDay)

        create = CreateDocuments(name, orderDate)

        closingDay1 = deliDate.addDays(closingDay)
        data1 = []
        data1 = settings.value('bang2', data1)

        dateOfCheck = settings1.value('dateOfCheck', '')
        dateOfCheck = 0 - int(dateOfCheck)
        dateOfCheck1 = deliDate.addDays(dateOfCheck)
        closingDay_PXXK = settings1.value('closingDay_PXXK', '')
        closingDay_PXXK = 0 - int(closingDay_PXXK)
        closingDay_PXXK1 = deliDate.addDays(closingDay_PXXK)

        data2 = []
        data2 = settings.value('bang3', data2)

        wb1 = create.closePlan(data, closingDay1.toString('MMMM d, yyyy'))
        wb2 = create.embryosPlan(data1)
        wb3 = create.inventPlan(data, dateOfCheck1.toString('MMMM d, yyyy'), closingDay1.toString('MMMM d, yyyy'),
                                closingDay_PXXK1.toString('MMMM d, yyyy'))
        wb5 = create.productPlan(data1)
        wb4 = create.embryosPlanPart(data1, data2)
        wb6 = create.productPlanPart(data1, hour, orderDate)

        settings.setValue('wb1', wb1)
        settings.setValue('wb2', wb2)
        settings.setValue('wb3', wb3)
        settings.setValue('wb4', wb4)
        settings.setValue('wb5', wb5)
        settings.setValue('wb6', wb6)
        Third.hide(self)
        theclass.show()


class Second(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        super(Second, self).__init__(parent)
        uic.loadUi('window2.ui', self)
        reg_ex = QtCore.QRegExp('^[0-9]{1,100}$')
        validator = QtGui.QRegExpValidator(reg_ex)
        self.lineNum.setValidator(validator)
        self.confirm.clicked.connect(self.on_productChange)
        self.clear.clicked.connect(self.clear_table)
        self.back.clicked.connect(self.on_back)
        self.next.clicked.connect(self.on_next)
        self.actionPlan.triggered.connect(self.on_settingPlan)
        self.actionProduce.triggered.connect(self.on_settingProduce)

    def on_settingPlan(self):
        dialog = SettingPlan(self)
        dialog.show()

    def on_settingProduce(self):
        dialog = SettingProduce(self)
        dialog.show()

    def on_productLine(self, part):
        rows = self.lineTable.rowCount()

        data = []
        for i in range(rows):
            it = self.lineTable.cellWidget(i, 0)
            if it and it.text():
                data.append(it.text())

        for n, i in enumerate(part):
            for j in range(int(i)):
                productCode = str('{}.{}'.format(n + 1, j + 1))
                if productCode in data:
                    pass
                else:

                    reg_ex1 = QtCore.QRegExp('^[0-9]{1,1}$')
                    validator = QtGui.QRegExpValidator(reg_ex1)
                    lineEdit1 = QtWidgets.QLineEdit()
                    lineEdit1.setValidator(validator)
                    reg_ex = QtCore.QRegExp('^[0-9]{1,10}$')
                    validator = QtGui.QRegExpValidator(reg_ex)
                    lineEdit = QtWidgets.QLineEdit()
                    lineEdit.setValidator(validator)
                    lineEdit2 = QtWidgets.QLineEdit()
                    lineEdit2.setValidator(validator)
                    lineEdit3 = QtWidgets.QLineEdit()
                    lineEdit3.setValidator(validator)

                    rowPosition = self.lineTable.rowCount()
                    self.lineTable.insertRow(rowPosition)

                    self.lineTable.setCellWidget(rowPosition, 0, QtWidgets.QLineEdit(productCode))
                    self.lineTable.setCellWidget(rowPosition, 1, QtWidgets.QLineEdit())
                    self.lineTable.setCellWidget(rowPosition, 2, QtWidgets.QLineEdit())
                    self.lineTable.setCellWidget(rowPosition, 3, lineEdit2)
                    self.lineTable.setCellWidget(rowPosition, 4, lineEdit3)
                    self.lineTable.setCellWidget(rowPosition, 5, QtWidgets.QLineEdit())
                    self.lineTable.setCellWidget(rowPosition, 6, QtWidgets.QLineEdit())
                    self.lineTable.setCellWidget(rowPosition, 7, QtWidgets.QLineEdit())
                    self.lineTable.setCellWidget(rowPosition, 8, lineEdit)
                    self.lineTable.setCellWidget(rowPosition, 9, lineEdit1)

    def on_productChange(self):
        lineNum = self.lineNum.text()
        if lineNum:
            lineNum = int(lineNum)

            for i in range(0, lineNum):
                reg_ex = QtCore.QRegExp('^[0-9]{1,10}$')
                validator = QtGui.QRegExpValidator(reg_ex)
                lineEdit = QtWidgets.QLineEdit()
                lineEdit.setValidator(validator)
                lineEdit2 = QtWidgets.QLineEdit()
                lineEdit2.setValidator(validator)
                lineEdit3 = QtWidgets.QLineEdit()
                lineEdit3.setValidator(validator)
                reg_ex1 = QtCore.QRegExp('^[0-9]{1,1}$')
                validator = QtGui.QRegExpValidator(reg_ex1)

                lineEdit1 = QtWidgets.QLineEdit()
                lineEdit1.setValidator(validator)

                rowPosition = self.lineTable.rowCount()
                self.lineTable.insertRow(rowPosition)

                self.lineTable.setCellWidget(rowPosition, 0, QtWidgets.QLineEdit())
                self.lineTable.setCellWidget(rowPosition, 1, QtWidgets.QLineEdit())
                self.lineTable.setCellWidget(rowPosition, 2, QtWidgets.QLineEdit())
                self.lineTable.setCellWidget(rowPosition, 3, lineEdit2)
                self.lineTable.setCellWidget(rowPosition, 4, lineEdit3)
                self.lineTable.setCellWidget(rowPosition, 5, QtWidgets.QLineEdit())
                self.lineTable.setCellWidget(rowPosition, 6, QtWidgets.QLineEdit())
                self.lineTable.setCellWidget(rowPosition, 7, QtWidgets.QLineEdit())
                self.lineTable.setCellWidget(rowPosition, 8, lineEdit)
                self.lineTable.setCellWidget(rowPosition, 9, lineEdit1)
        else:
            pass

    def clear_table(self):
        self.lineTable.setRowCount(0)

    def saveInstance(self):
        settings = QtCore.QSettings('myorg', 'myapp')

        rows = self.lineTable.rowCount()
        columns = self.lineTable.columnCount()

        data = []
        for row in range(rows):
            data.append([])
            for column in range(columns):
                it = self.lineTable.cellWidget(row, column)
                if it and it.text():
                    pass
                else:
                    item = QtWidgets.QLineEdit()
                    item.setText("")
                    self.lineTable.setCellWidget(row, column, item)
                data[row].append(self.lineTable.cellWidget(row, column).text())

        settings.setValue('bang2', data)

    def restoreSettings(self):
        settings = QtCore.QSettings('myorg', 'myapp')

        data = []
        data = settings.value('bang2', data)
        numrows = len(data)
        if numrows == 0:
            pass
        else:
            for row in range(numrows):
                rowPosition = self.lineTable.rowCount()
                self.lineTable.insertRow(rowPosition)
                reg_ex = QtCore.QRegExp('^[0-9]{1,10}$')
                validator = QtGui.QRegExpValidator(reg_ex)
                lineEdit = QtWidgets.QLineEdit(str(data[row][8]))
                lineEdit.setValidator(validator)
                lineEdit2 = QtWidgets.QLineEdit(str(data[row][3]))
                lineEdit2.setValidator(validator)
                lineEdit3 = QtWidgets.QLineEdit(str(data[row][4]))
                lineEdit3.setValidator(validator)
                reg_ex1 = QtCore.QRegExp('^[0-9]{1,1}$')
                validator = QtGui.QRegExpValidator(reg_ex1)
                lineEdit1 = QtWidgets.QLineEdit(str(data[row][9]))
                lineEdit1.setValidator(validator)

                self.lineTable.setCellWidget(rowPosition, 0, QtWidgets.QLineEdit(str(data[row][0])))
                self.lineTable.setCellWidget(rowPosition, 1, QtWidgets.QLineEdit(str(data[row][1])))
                self.lineTable.setCellWidget(rowPosition, 2, QtWidgets.QLineEdit(str(data[row][2])))
                self.lineTable.setCellWidget(rowPosition, 3, lineEdit2)
                self.lineTable.setCellWidget(rowPosition, 4, lineEdit3)
                self.lineTable.setCellWidget(rowPosition, 5, QtWidgets.QLineEdit(str(data[row][5])))
                self.lineTable.setCellWidget(rowPosition, 6, QtWidgets.QLineEdit(str(data[row][6])))
                self.lineTable.setCellWidget(rowPosition, 7, QtWidgets.QLineEdit(str(data[row][7])))
                self.lineTable.setCellWidget(rowPosition, 8, lineEdit)
                self.lineTable.setCellWidget(rowPosition, 9, lineEdit1)

    def on_back(self):
        Second.saveInstance(self)
        theclass = First(self)
        theclass.restoreSettings()
        Second.hide(self)
        theclass.show()

    def on_next(self):
        Second.saveInstance(self)
        theclass = Third(self)
        settings = QtCore.QSettings('myorg', 'myapp')
        if settings.contains('bang3'):
            theclass.restoreSettings()

        data = []
        data = settings.value('bang2', data)
        part = []
        for each in data:
            part.append(each[1])
        # part = list(filter(None, part))

        theclass.on_supplieLine(part)
        Second.hide(self)
        theclass.show()


class First(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        super(First, self).__init__(parent)
        uic.loadUi('window1.ui', self)
        reg_ex = QtCore.QRegExp('^[0-9]{1,5}$')
        validator = QtGui.QRegExpValidator(reg_ex)
        self.productNum.setValidator(validator)
        self.confirm.clicked.connect(self.on_productNum)
        self.clear.clicked.connect(self.clear_table)
        self.next.clicked.connect(self.on_next)
        self.actionNew.triggered.connect(self.on_new)
        self.actionQuit_2.triggered.connect(self.close)
        self.actionPlan.triggered.connect(self.on_settingPlan)
        self.actionProduce.triggered.connect(self.on_settingProduce)

    def on_productNum(self):
        productNum = self.productNum.text()

        if productNum:
            productNum = int(productNum)

            for i in range(0, productNum):
                rowPosition = self.tableWidget.rowCount()
                self.tableWidget.insertRow(rowPosition)
                lineEdit = QtWidgets.QLineEdit()
                reg_ex = QtCore.QRegExp('^[0-9]{1,10}$')
                validator = QtGui.QRegExpValidator(reg_ex)
                lineEdit.setValidator(validator)
                lineEdit1 = QtWidgets.QLineEdit()
                reg_ex = QtCore.QRegExp('^[0-9]{1,2}$')
                validator = QtGui.QRegExpValidator(reg_ex)
                lineEdit1.setValidator(validator)
                lineEdit2 = QtWidgets.QLineEdit()
                reg_ex = QtCore.QRegExp('^[0-9]{1,10}$')
                validator = QtGui.QRegExpValidator(reg_ex)
                lineEdit2.setValidator(validator)
                self.tableWidget.setCellWidget(rowPosition, 0, QtWidgets.QLineEdit())
                self.tableWidget.setCellWidget(rowPosition, 1, QtWidgets.QLineEdit())
                self.tableWidget.setCellWidget(rowPosition, 2, QtWidgets.QLineEdit())
                self.tableWidget.setCellWidget(rowPosition, 5, lineEdit2)
                self.tableWidget.setCellWidget(rowPosition, 6, QtWidgets.QLineEdit())
                self.tableWidget.setCellWidget(rowPosition, 7, QtWidgets.QLineEdit())
                self.tableWidget.setCellWidget(rowPosition, 3, lineEdit)
                self.tableWidget.setCellWidget(rowPosition, 4, lineEdit1)

        else:
            pass

    def clear_table(self):
        self.tableWidget.setRowCount(0)

    def on_settingPlan(self):
        dialog = SettingPlan(self)
        dialog.show()

    def on_settingProduce(self):
        dialog = SettingProduce(self)
        dialog.show()

    def on_next(self):
        theclass = Second(self)
        First.saveInstance(self)

        settings = QtCore.QSettings('myorg', 'myapp')
        if settings.contains('bang2'):
            theclass.restoreSettings()

        data = []
        data = settings.value('bang', data)
        part = []
        for each in data:
            part.append(each[4])
        for n, i in enumerate(part):
            if not i.isdigit():
                part[n] = 0
            else:
                part[n] = int(i)

        theclass.on_productLine(part)
        First.hide(self)
        theclass.show()

    def on_new(self):
        settings = QtCore.QSettings('myorg', 'myapp')
        settings.clear()

        self.hide()
        self.show()

    def saveInstance(self):
        settings = QtCore.QSettings('myorg', 'myapp')
        settings.setValue('ten', self.name.text())
        settings.setValue('loai', self.productNum.text())
        settings.setValue('dateOrder', self.dateOfOrder.date())
        settings.setValue('deliveryDate', self.deliveryDate.date())

        rows = self.tableWidget.rowCount()
        columns = self.tableWidget.columnCount()
        data = []
        for row in range(rows):
            data.append([])
            for column in range(columns):
                it = self.tableWidget.cellWidget(row, column)
                if it and it.text():
                    pass
                else:
                    item = QtWidgets.QLineEdit()
                    item.setText("")
                    self.tableWidget.setCellWidget(row, column, item)
                data[row].append(self.tableWidget.cellWidget(row, column).text())

        settings.setValue('bang', data)

    def restoreSettings(self):
        settings = QtCore.QSettings('myorg', 'myapp')
        name = settings.value('ten', '')
        productNum = settings.value('loai', '')
        orderDate = settings.value('dateOrder', '')
        deliDate = settings.value('deliveryDate', '')

        self.name.setText(name)
        self.productNum.setText(productNum)
        self.dateOfOrder.setDate(QtCore.QDate(orderDate))
        self.deliveryDate.setDate(QtCore.QDate(deliDate))
        data = []
        data = settings.value('bang', data)
        numrows = len(data)
        if numrows == 0:
            pass
        else:
            for row in range(numrows):
                rowPosition = self.tableWidget.rowCount()
                self.tableWidget.insertRow(rowPosition)
                lineEdit = QtWidgets.QLineEdit(str(data[row][3]))
                reg_ex = QtCore.QRegExp('^[0-9]{1,1000}$')
                validator = QtGui.QRegExpValidator(reg_ex)
                lineEdit.setValidator(validator)
                lineEdit1 = QtWidgets.QLineEdit(str(data[row][4]))
                reg_ex = QtCore.QRegExp('^[0-9]{1,2}$')
                validator = QtGui.QRegExpValidator(reg_ex)
                lineEdit1.setValidator(validator)
                lineEdit2 = QtWidgets.QLineEdit(str(data[row][5]))
                reg_ex = QtCore.QRegExp('^[0-9]{1,2}$')
                validator = QtGui.QRegExpValidator(reg_ex)
                lineEdit2.setValidator(validator)
                self.tableWidget.setCellWidget(rowPosition, 0, QtWidgets.QLineEdit(str(data[row][0])))
                self.tableWidget.setCellWidget(rowPosition, 1, QtWidgets.QLineEdit(str(data[row][1])))
                self.tableWidget.setCellWidget(rowPosition, 2, QtWidgets.QLineEdit(str(data[row][2])))
                self.tableWidget.setCellWidget(rowPosition, 3, lineEdit)
                self.tableWidget.setCellWidget(rowPosition, 4, lineEdit1)
                self.tableWidget.setCellWidget(rowPosition, 5, lineEdit2)
                self.tableWidget.setCellWidget(rowPosition, 6, QtWidgets.QLineEdit(str(data[row][6])))
                self.tableWidget.setCellWidget(rowPosition, 7, QtWidgets.QLineEdit(str(data[row][7])))


class SettingPlan(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super(SettingPlan, self).__init__(parent)
        uic.loadUi('settingPlan.ui', self)
        self.save.clicked.connect(self.on_save)
        self.restoreSettings()

    def on_save(self):
        settings = QtCore.QSettings('myorg', 'mysetting')
        settings.setValue('dateOfCheck', self.dateOfCheck.text())
        settings.setValue('closingDay', self.closingDay.text())
        settings.setValue('closingDay_PXXK', self.closingDay_PXXK.text())
        settings.setValue('minWorkh', self.minWorkh.text())
        settings.setValue('maxWorkh', self.maxWorkh.text())

    def restoreSettings(self):
        settings = QtCore.QSettings('myorg', 'mysetting')
        dateOfCheck = settings.value('dateOfCheck', '')
        closingDay = settings.value('closingDay', '')
        closingDay_PXXK = settings.value('closingDay_PXXK', '')
        minWorkh = settings.value('minWorkh', '')
        maxWorkh = settings.value('maxWorkh', '')

        self.dateOfCheck.setText(dateOfCheck)
        self.closingDay.setText(closingDay)
        self.closingDay_PXXK.setText(closingDay_PXXK)
        self.minWorkh.setText(minWorkh)
        self.maxWorkh.setText(maxWorkh)


class SettingProduce(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super(SettingProduce, self).__init__(parent)
        uic.loadUi('settingProduce.ui', self)
        self.save.clicked.connect(self.on_save)
        self.restoreSettings()

    def on_save(self):
        settings = QtCore.QSettings('myorg', 'mysetting')
        settings.setValue('inProgress', self.inProgress.text())
        settings.setValue('endProgress', self.endProgress.text())

    def restoreSettings(self):
        settings = QtCore.QSettings('myorg', 'mysetting')
        inProgress = settings.value('inProgress', '')
        endProgress = settings.value('endProgress', '')

        self.inProgress.setText(inProgress)
        self.endProgress.setText(endProgress)


def main():
    app = QtWidgets.QApplication(sys.argv)
    settings = QtCore.QSettings('myorg', 'myapp')
    settings.clear()
    main = First()
    main.show()

    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
