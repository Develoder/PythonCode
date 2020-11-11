# pyuic5 untitled.ui -o untitled.py
# coding: utf8

import datetime, os, openpyxl, Table, Otchet

from PyQt5.QtWidgets import *
from PyQt5.QtGui import QPixmap
from PyQt5 import QtGui

from mainUI import *
from Conection import *


class MyWin(QtWidgets.QMainWindow):
    indexTO = {"a": ["В рейсе", [85, 170, 0]], "b": ["Обеспечение рейся", [85, 170, 0]], "c": ["Задержки по метеоусловиям и в связи с запретами", [85, 170, 0]], "d": ["В резерве", [85, 170, 0]],
               "e": ["Исправный - неиспользуемый", [85, 170, 0]], "f": ["Устранение неисправностей при оперативном ТО", [0, 0, 255]], "g": ["Ожидание технического обслуживания по Ф-Б", [255, 0, 0]], "h": ["Техобслуживание по Ф-Б", [0, 0, 255]],
               "i": ["Ожидание периодического техобслуживания", [255, 0, 0]], "j": ["Техобслуживание по периодическим формам", [0, 0, 255]], "k": ["Межсменные или ночные перерывы", [255, 0, 0]], "l": ["Ожидание ремонта", [150, 75, 0]],
               "m": ["В ремонте", [150, 75, 0]], "n": ["Отсуствие запчастей", [255, 0, 0]], "o": ["Отсутствие двигателей", [255, 0, 0]], "p": ["Доработки по бюллетеням", [0, 0, 255]],
               "q": ["Рекламация промышленности", [255, 0, 0]], "r": ["Рекламация ремзаводам", [255, 0, 0]], "s": ["Расследование летных происшедствий, поломок, повреждений", [255, 0, 0]], "t": ["Восстановление самолета после повреждения", [255, 0, 0]],
               "u": ["Ожидание списания", [255, 255, 0]]}

    currentSymbolTO = ''
    now = datetime.datetime
    headerIndex = dict()

    def __init__(self, parent=None):
        QtWidgets.QWidget.__init__(self, parent)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.connect = Connection()

        self.now = datetime.datetime.now()
        self.ui.dateEdit.setDate(QtCore.QDate(self.now.year, self.now.month, self.now.day))
        self.ui.dateEdit.setMaximumDate(QtCore.QDate(self.now.year, self.now.month, self.now.day))

        self.GenerateIndexTO()
        self.GenerateTable("%04i-%02i-%02i" % (self.now.year, self.now.month, self.now.day))

        self.ui.tableWidget.cellClicked.connect(self.Clicked)
        self.ui.listWidget.itemClicked.connect(self.ChangeIndexTO)
        self.ui.dateEdit.userDateChanged.connect(self.ChangeDate)
        self.ui.action_5.triggered.connect(self.close)
        self.ui.action.triggered.connect(self.ObAvtore)
        self.ui.action_2.triggered.connect(self.Oprogram)
        self.ui.action_4.triggered.connect(self.ShowOtchet)
        self.ui.action_3.triggered.connect(self.ShowTable)
        self.ui.action_6.triggered.connect(self.Printing)

        self.ui.tableWidget.resizeColumnsToContents()

    def Printing(self):
        fileName, _ = QFileDialog.getSaveFileName(None, 'Open File', 'file.xlsx', "Exele (*.xlsx)")
        if fileName:
            self.CreateExale(fileName)

    def CreateExale(self, fileName):
        print(fileName)
        wb = openpyxl.Workbook()

        sheets = wb.sheetnames
        wb.create_sheet(title='Лист 1', index=0)
        sheet = wb['Лист 1']

        for i in range(1, self.ui.tableWidget.columnCount() + 1):
            cell = sheet.cell(row=2, column=i+1)
            cell.value = self.ui.tableWidget.horizontalHeaderItem(i - 1).text()

        for i in range(self.ui.tableWidget.rowCount()):
            cell = sheet.cell(row=i+3, column=1)
            cell.value = self.ui.tableWidget.verticalHeaderItem(i).text()

        for row in range(3, self.ui.tableWidget.rowCount() + 3):
            for col in range(1, self.ui.tableWidget.columnCount() + 1):
                cell = sheet.cell(row=row, column=col+1)
                try:
                    if type(self.ui.tableWidget.cellWidget(row - 3, col - 1)) == QtWidgets.QComboBox:
                        cell.value = self.ui.tableWidget.cellWidget(row - 3, col - 1).currentText()
                    else:
                        cell.value = self.ui.tableWidget.item(row - 3, col - 1).text()
                except:
                    pass

        lenTitle = 0
        for i in range(self.ui.tableWidget.columnCount()):
            if self.ui.tableWidget.horizontalHeaderItem(i).text() != "":
                lenTitle += 1
        letters = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]
        exele = dict()
        for i in range(150):
            exele[i] = letters[i // (len(letters))] + letters[i % (len(letters))]
        print(exele)
        #sheet.merge_cells('A1:%s1' % exele[lenTitle])
        sheet['A1'] = self.ui.label.text()

        try:
            wb.save(fileName)
        except Exception as e:
            self.ErrorMesage("Фаил не может сохраниться", str(e))

        os.startfile(fileName, "print")

    def ShowOtchet(self):
        self.otchet = Otchet.Otchet()
        self.otchet.show()

    def ShowTable(self):
        self.table = Table.Table()
        self.table.show()


    def ObAvtore(self):
        msg = QMessageBox()
        msg.setWindowTitle("Об авторе")
        msg.setText('Данный программный продукт разработан курсантам 431 группы Кайгородцевым Юрием Витальевичем по специальности: 09.02.03 «Программирование в компьютерных системах»')
        msg.exec()

    def Oprogram(self):
        msg = QMessageBox()
        msg.setWindowTitle("О программе")
        msg.setText('Данный программный продукт разработан для автоматизация учета исправности и использования самолетов в эксплуатационном авиапредприятии')
        msg.exec()

    def ChangeDate(self, date):
        self.GenerateTable(date.toString('yyyy-MM-dd'))

    def GenerateTable(self, date):
        self.ui.tableWidget.clear()

        tables = []
        num = 0
        for i in range(24):
            for j in range(6):
                tables.append("  " + ('%02i' % i) + ":" + ('%02i' % (j * 10)) + "  ")
                self.headerIndex[num] = ('%02i' % i) + ('%02i' % (j * 10))
                num += 1

        collumn = [i[1] + ": " + i[0] for i in self.connect.SelectTable("Plane", ['name_plane', "bort_number"])]
        print(collumn)

        self.ui.tableWidget.setRowCount(len(collumn))
        self.ui.tableWidget.setColumnCount(len(tables))
        self.ui.tableWidget.setHorizontalHeaderLabels(tables)
        self.ui.tableWidget.setVerticalHeaderLabels(collumn)

        font = QtGui.QFont()
        font.setFamily("Tabellsymboler")
        font.setPointSize(36)
        brush = QtGui.QBrush()

        elements = list(self.connect.Selection("SELECT [id_plane] ,[time] ,[simbol] FROM [HealthAccounting] WHERE [date] = '%s';"
                                               % date))
        print(elements)
        for i in elements:
            item = QtWidgets.QTableWidgetItem(i[2])
            item.setFont(font)
            brush = QtGui.QBrush(QtGui.QColor(self.indexTO[i[2]][1][0], self.indexTO[i[2]][1][1], self.indexTO[i[2]][1][2]))
            item.setForeground(brush)

            self.ui.tableWidget.setItem(i[0]-1, self.get_key(self.headerIndex, i[1]), item)

        self.ui.tableWidget.resizeColumnsToContents()


    def GenerateIndexTO(self):
        self.ui.listWidget.addItem(QListWidgetItem(QtGui.QIcon('eraser.png'), ""))
        for i in self.indexTO:
            item = QListWidgetItem(i)
            item.setForeground(QtGui.QBrush(QtGui.QColor(self.indexTO[i][1][0], self.indexTO[i][1][1], self.indexTO[i][1][2])))
            self.ui.listWidget.addItem(item)

    def ChangeIndexTO(self, item: QListWidgetItem):
        if item.text() == "":
            self.currentSymbolTO = "delete"
            self.ui.label_3.setText("Выберете ячейку для удаления")
            self.ui.label_4.setText("")
            return
        self.currentSymbolTO = item.text()
        self.ui.label_3.setText(self.indexTO[item.text()][0])
        self.ui.label_4.setText(item.text())
        self.ui.label_4.setStyleSheet("color: rgb(%i, %i, %i);" % (self.indexTO[item.text()][1][0], self.indexTO[item.text()][1][1], self.indexTO[item.text()][1][2]))


    def Clicked(self, i, j):
        if self.currentSymbolTO == "":
            return
        selectItem = self.ui.tableWidget.item(i, j)
        if self.currentSymbolTO == "delete":
            self.ui.tableWidget.setItem(i, j, QtWidgets.QTableWidgetItem(""))
        else:
            item = QtWidgets.QTableWidgetItem(self.currentSymbolTO)
            brush = QtGui.QBrush(QtGui.QColor(self.indexTO[self.currentSymbolTO][1][0], self.indexTO[self.currentSymbolTO][1][1], self.indexTO[self.currentSymbolTO][1][2]))

            font = QtGui.QFont()
            font.setFamily("Tabellsymboler")
            font.setPointSize(36)
            item.setFont(font)
            item.setForeground(brush)

            self.ui.tableWidget.setItem(i, j, item)

        if selectItem == None:
            print(self.currentSymbolTO + " " + str(i+1) + " " + self.headerIndex[j] + " " + str("%04i-%02i-%02i" % (self.ui.dateEdit.date().year(), self.ui.dateEdit.date().month(), self.ui.dateEdit.date().day())))
            self.connect.Create("HealthAccounting", ["id_plane", "time", "simbol", "date"], [i+1,
                                                                                             self.headerIndex[j],
                                                                                             self.currentSymbolTO,
                "%04i-%02i-%02i" % (self.ui.dateEdit.date().year(), self.ui.dateEdit.date().month(), self.ui.dateEdit.date().day())])
        else:
            element = list(self.connect.Selection(
                "SELECT [id] FROM [HealthAccounting] WHERE [date] = '%s' AND [id_plane] = %s AND [time] = '%s';"
                % ("%04i-%02i-%02i" % (
                self.ui.dateEdit.date().year(), self.ui.dateEdit.date().month(), self.ui.dateEdit.date().day()),
                   i + 1, self.headerIndex[j])))
            print(element[0][0])
            if self.currentSymbolTO == "delete":
                self.connect.Delete("HealthAccounting", "id", element[0][0])
            else:
                self.connect.Update("HealthAccounting", "simbol", self.currentSymbolTO, "id", element[0][0])

    def get_key(self, d, value):
        for k, v in d.items():
            if str(v) == str(value):
                return k
        return None


if __name__ == "__main__":
    import sys

    app = QtWidgets.QApplication(sys.argv)
    myapp = MyWin()
    myapp.show()
    sys.exit(app.exec_())
