# import pymysql
import sys
from PyQt5.QtWidgets import *
import pymysql
import csv
import decimal
import json
import xml.etree.ElementTree as ET

class DB_Utils:

    def queryExecutor(self, db, sql, params):
        conn = pymysql.connect(host='localhost', user='guest', password='bemyguest', db=db, charset='utf8')

        try:
            with conn.cursor(pymysql.cursors.DictCursor) as cursor:  # dictionary based cursor
                cursor.execute(sql, params)
                rows = cursor.fetchall()
                return rows
        except Exception as e:
            print(e)
            print(type(e))
        finally:
            conn.close()

class DB_Queries:

    def selectCustomerName(self):
        sql = "SELECT DISTINCT name FROM customers ORDER BY name"
        params = ()

        util = DB_Utils()
        rows = util.queryExecutor(db="classicmodels", sql=sql, params=params)
        return rows

    def selectCustomerCountry(self):
        sql = "SELECT DISTINCT country FROM customers ORDER BY country"
        params = ()

        util = DB_Utils()
        rows = util.queryExecutor(db="classicmodels", sql=sql, params=params)
        return rows

    def selectCustomerCity(self):
        sql = "SELECT DISTINCT city FROM customers ORDER BY city"
        params = ()

        util = DB_Utils()
        rows = util.queryExecutor(db="classicmodels", sql=sql, params=params)
        return rows

    def selectCountryCity(self, value):
        if value == "ALL":
            sql = "SELECT DISTINCT city FROM customers ORDER BY city"
            params = ()
        else:
            sql = "SELECT DISTINCT city FROM customers WHERE country = %s ORDER BY city"
            params = (value)

        util = DB_Utils()
        rows = util.queryExecutor(db="classicmodels", sql=sql, params=params)
        return rows

    def selectOrder(self, value, isCustomer, isCountry, isCity):

        if value == "ALL":
            sql = """SELECT orderNo, orderDate, requiredDate, shippedDate, status, name, comments 
            FROM orders o JOIN customers c ON c.customerId = o.customerId ORDER BY orderNo"""
            params = ()
        else:
            if isCustomer == 1:
                sql = """SELECT orderNo, orderDate, requiredDate, shippedDate, status, name, comments 
                FROM orders o JOIN customers c ON c.customerId = o.customerId 
                WHERE name = %s ORDER BY orderNo"""
                params = (value)

            elif isCountry == 1:
                sql = """SELECT orderNo, orderDate, requiredDate, shippedDate, status, name, comments 
                FROM orders o JOIN customers c ON c.customerId = o.customerId 
                WHERE country = %s ORDER BY orderNo"""
                params = (value)

            elif isCity == 1:
                sql = """SELECT orderNo, orderDate, requiredDate, shippedDate, status, name, comments 
                FROM orders o JOIN customers c ON c.customerId = o.customerId 
                WHERE city = %s ORDER BY orderNo"""
                params = (value)

        util = DB_Utils()
        rows = util.queryExecutor(db="classicmodels", sql=sql, params=params)
        return rows

    def selectSaleDetails(self, value):
        sql = """SELECT orderLineNo, o.productCode, name, quantity, priceEach
            FROM orderdetails o JOIN products p ON p.productCode = o.productCode
            WHERE orderNo = %s ORDER BY o.orderLineNo"""
        params = (value)

        util = DB_Utils()
        rows = util.queryExecutor(db="classicmodels", sql=sql, params=params)
        return rows


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setupUI()

    def setupUI(self):
        self.setWindowTitle("주문 검색")
        self.setGeometry(0, 0, 1200, 800)

        query = DB_Queries()
        customers = query.selectCustomerName()
        countries = query.selectCustomerCountry()
        cities = query.selectCustomerCity()

        self.label1 = QLabel("주문 검색")
        font1 = self.label1.font()
        font1.setPointSize(15)
        self.label1.setFont(font1)

        self.comboValue = "ALL"
        self.isCustomer = 0
        self.isCountry = 0
        self.isCity = 0

        self.label2 = QLabel("고객:")
        self.name = QComboBox(self)
        self.name.addItem('ALL')
        for customer in customers:
            self.name.addItem(customer.get('name'))
        self.name.activated.connect(self.customerBox_Activated)

        self.label3 = QLabel("국가:")
        self.country = QComboBox(self)
        self.country.addItem('ALL')
        for country in countries:
            self.country.addItem(country.get('country'))
        self.country.activated.connect(self.countryBox_Activated)

        self.label4 = QLabel("도시:")
        self.city = QComboBox(self)
        self.city.addItem('ALL')
        for city in cities:
            self.city.addItem(city.get('city'))
        self.city.activated.connect(self.cityBox_Activated)

        self.search = QPushButton("검색")
        self.search.clicked.connect(self.search_Clicked)
        self.reset = QPushButton("초기화")
        self.reset.clicked.connect(self.reset_Clicked)

        self.order_list = QTableWidget()
        self.order_list.cellClicked.connect(self.order_Clicked)
        query = DB_Queries()
        orders = query.selectOrder("ALL", self.isCustomer, self.isCountry, self.isCity)

        self.order_list.clearContents()
        self.column = list(orders[0].keys())
        self.column[5] = 'customer'
        self.order_list.setRowCount(len(orders))
        self.order_list.setColumnCount(len(self.column))
        self.order_list.setHorizontalHeaderLabels(self.column)
        self.order_list.setEditTriggers(QAbstractItemView.NoEditTriggers)

        for rowIDX, order in enumerate(orders):
            for columnIDX, (k, v) in enumerate(order.items()):
                if v == None:
                    continue
                else:
                    item = QTableWidgetItem(str(v))

                self.order_list.setItem(rowIDX, columnIDX, item)

        self.order_list.resizeColumnsToContents()
        self.order_list.resizeRowsToContents()


        self.search_cnt = len(orders)
        self.search_no = QLabel("검색된 주문의 개수:")
        self.search_no_result = QLabel(str(self.search_cnt))


        button_box = QVBoxLayout()
        button_box.addWidget(self.search)
        button_box.addWidget(self.reset)

        combo_box = QHBoxLayout()
        combo_box.addWidget(self.label2)
        combo_box.addWidget(self.name)
        combo_box.addStretch(1)
        combo_box.addWidget(self.label3)
        combo_box.addWidget(self.country)
        combo_box.addStretch(1)
        combo_box.addWidget(self.label4)
        combo_box.addWidget(self.city)
        combo_box.addStretch(1)

        search_box = QHBoxLayout()
        search_box.addWidget(self.search_no)
        search_box.addWidget(self.search_no_result)
        search_box.addStretch(1)

        box1 = QVBoxLayout()
        box1.addLayout(combo_box)
        box1.addLayout(search_box)

        box2 = QHBoxLayout()
        box2.addLayout(box1)
        box2.addLayout(button_box)

        layout = QVBoxLayout()
        layout.addWidget(self.label1)
        layout.addLayout(box2)
        layout.addWidget(self.order_list)

        self.setLayout(layout)

    def customerBox_Activated(self):

        self.isCustomer = 1
        self.isCountry = 0
        self.isCity = 0
        self.comboValue = self.name.currentText()

    def countryBox_Activated(self):

        self.isCustomer = 0
        self.isCountry = 1
        self.isCity = 0
        query = DB_Queries()
        self.comboValue = self.country.currentText()
        cities = query.selectCountryCity(self.comboValue)
        self.city.clear()
        self.city.addItem('ALL')
        for city in cities:
            self.city.addItem(city.get('city'))

    def cityBox_Activated(self):

        self.isCustomer = 0
        self.isCountry = 0
        self.isCity = 1
        self.comboValue = self.city.currentText()

    def search_Clicked(self):

        query = DB_Queries()
        orders = query.selectOrder(self.comboValue, self.isCustomer, self.isCountry, self.isCity)

        self.order_list.clearContents()
        self.order_list.setRowCount(len(orders))
        self.order_list.setColumnCount(len(self.column))
        self.order_list.setHorizontalHeaderLabels(self.column)
        self.order_list.setEditTriggers(QAbstractItemView.NoEditTriggers)

        for rowIDX, order in enumerate(orders):
            for columnIDX, (k, v) in enumerate(order.items()):
                if v == None:
                    continue
                else:
                    item = QTableWidgetItem(str(v))

                self.order_list.setItem(rowIDX, columnIDX, item)

        self.order_list.resizeColumnsToContents()
        self.order_list.resizeRowsToContents()
        self.search_cnt = len(orders)
        self.search_no_result.setText(str(self.search_cnt))

    def reset_Clicked(self):

        query = DB_Queries()
        print(self.comboValue)
        orders = query.selectOrder("ALL", self.isCustomer, self.isCountry, self.isCity)

        customers = query.selectCustomerName()
        countries = query.selectCustomerCountry()
        cities = query.selectCustomerCity()

        self.name.clear()
        self.name.addItem('ALL')
        for customer in customers:
            self.name.addItem(customer.get('name'))

        self.country.clear()
        self.country.addItem('ALL')
        for country in countries:
            self.country.addItem(country.get('country'))

        self.city.clear()
        self.city.addItem('ALL')
        for city in cities:
            self.city.addItem(city.get('city'))

        self.order_list.clearContents()
        self.order_list.setRowCount(len(orders))
        self.order_list.setColumnCount(len(self.column))
        self.order_list.setHorizontalHeaderLabels(self.column)
        self.order_list.setEditTriggers(QAbstractItemView.NoEditTriggers)

        for rowIDX, order in enumerate(orders):
            for columnIDX, (k, v) in enumerate(order.items()):
                if v == None:
                    continue
                else:
                    item = QTableWidgetItem(str(v))

                self.order_list.setItem(rowIDX, columnIDX, item)

        self.order_list.resizeColumnsToContents()
        self.order_list.resizeRowsToContents()
        self.search_cnt = len(orders)
        self.search_no_result.setText(str(self.search_cnt))

    def order_Clicked(self):
        x = self.order_list.selectedIndexes()
        r = x[0].row()
        c = 0
        orderNo = self.order_list.item(r, c).text()

        dialogue = saleDetails(orderNo)
        dialogue.exec_()



class saleDetails(QDialog):
    def __init__(self, orderNo):
        super().__init__()
        self.setupUI(orderNo)

    def setupUI(self, orderNo):

        self.orderNo = orderNo
        self.setWindowTitle("주문 상세 내역")
        self.setGeometry(0, 0, 900, 800)

        self.label1 = QLabel("주문 상세 내역")
        font1 = self.label1.font()
        font1.setPointSize(15)
        self.label1.setFont(font1)

        query = DB_Queries()
        self.saleDetailsTable = QTableWidget()
        sales = query.selectSaleDetails(self.orderNo)

        self.column = list(sales[0].keys())
        self.column.append('상품주문액')
        self.column[2] = ' productName'
        self.saleDetailsTable.setRowCount(len(sales))
        self.saleDetailsTable.setColumnCount(len(self.column))
        self.saleDetailsTable.setHorizontalHeaderLabels(self.column)
        self.saleDetailsTable.setEditTriggers(QAbstractItemView.NoEditTriggers)

        for rowIDX, sale in enumerate(sales):
            for columnIDX, (k, v) in enumerate(sale.items()):
                if v == None:
                    continue
                else:
                    item = QTableWidgetItem(str(v))

                self.saleDetailsTable.setItem(rowIDX, columnIDX, item)

        self.sum = 0
        for rowIDX, sale in enumerate(sales):
            arr = list(sale.values())
            total = arr[3] * arr[4]
            self.sum += total
            item = QTableWidgetItem(str(total))
            self.saleDetailsTable.setItem(rowIDX, 5, item)

        self.saleDetailsTable.resizeColumnsToContents()
        self.saleDetailsTable.resizeRowsToContents()

        self.label2 = QLabel("주문번호: ")
        self.label2_1 = QLabel(self.orderNo)

        self.label3 = QLabel("상품개수: ")
        self.label3_1 = QLabel(str(len(sales)))

        self.label4 = QLabel("주문액: ")
        self.label4_1 = QLabel(str(self.sum))

        self.label5 = QLabel("파일 출력")
        font2 = self.label5.font()
        font2.setPointSize(15)
        self.label5.setFont(font2)

        self.radioBtn1 = QRadioButton("CSV", self)
        self.radioBtn1.setChecked(True)

        self.radioBtn2 = QRadioButton("JSON", self)

        self.radioBtn3 = QRadioButton("XML", self)

        self.save = QPushButton("저장")
        self.save.clicked.connect(self.saveBtn_Clicked)

        details_box = QHBoxLayout()
        details_box.addWidget(self.label2)
        details_box.addWidget(self.label2_1)
        details_box.addStretch(1)
        details_box.addWidget(self.label3)
        details_box.addWidget(self.label3_1)
        details_box.addStretch(1)
        details_box.addWidget(self.label4)
        details_box.addWidget(self.label4_1)
        details_box.addStretch(1)

        radio_box = QHBoxLayout()
        radio_box.addWidget(self.radioBtn1)
        radio_box.addStretch(1)
        radio_box.addWidget(self.radioBtn2)
        radio_box.addStretch(1)
        radio_box.addWidget(self.radioBtn3)
        radio_box.addStretch(1)

        button_box = QHBoxLayout()
        button_box.addStretch(3)
        button_box.addWidget(self.save)

        layout = QVBoxLayout()
        layout.addWidget(self.label1)
        layout.addLayout(details_box)
        layout.addWidget(self.saleDetailsTable)
        layout.addWidget(self.label5)
        layout.addLayout(radio_box)
        layout.addLayout(button_box)

        self.setLayout(layout)

    def writeCSV(self, orderNo):
        fileName = orderNo + '.csv'
        query = DB_Queries()
        rows = query.selectSaleDetails(orderNo)
        with open(fileName, 'w', encoding='utf-8', newline='') as f:
            wr = csv.writer(f)

            for row in rows:
                arr = list(row.values())
                total = arr[3] * arr[4]
                row['상품주문액'] = total

            self.column = list(rows[0].keys())
            wr.writerow(self.column)

            for row in rows:
                arr = list(row.values())
                wr.writerow(arr)

    def writeJSON(self, orderNo):
        fileName = orderNo + '.json'
        query = DB_Queries()
        rows = query.selectSaleDetails(orderNo)

        for row in rows:
            arr = list(row.values())
            total = arr[3] * arr[4]
            row['상품주문액'] = total

            for k, v in row.items():
                if type(v) == decimal.Decimal:
                    row[k] = str(v)

        newDict = dict(orderDetails=rows)
        with open(fileName, 'w', encoding='utf-8') as f:
            json.dump(newDict, f, indent=4, ensure_ascii=False)

    def writeXML(self, orderNo):
        fileName = orderNo + '.xml'
        query = DB_Queries()
        rows = query.selectSaleDetails(orderNo)

        for row in rows:
            arr = list(row.values())
            total = arr[3] * arr[4]
            row['상품주문액'] = total

            for k, v in row.items():
                if type(v) == decimal.Decimal:
                    row[k] = str(v)

        newDict = dict(orderDetails=rows)

        tableName = list(newDict.keys())[0]
        tableRows = list(newDict.values())[0]

        rootElement = ET.Element('TABLE')
        rootElement.attrib['name'] = tableName

        for row in tableRows:
            rowElement = ET.Element('ROW')
            rootElement.append(rowElement)

            for columnName in list(row.keys()):
                if row[columnName] == None:
                    rowElement.attrib[columnName] = ''
                elif type(row[columnName]) == int:
                    rowElement.attrib[columnName] = str(row[columnName])
                else:
                    rowElement.attrib[columnName] = row[columnName]

        ET.ElementTree(rootElement).write(fileName, encoding='utf-8', xml_declaration=True)


    def saveBtn_Clicked(self):

        if self.radioBtn1.isChecked():
            self.writeCSV(self.orderNo)
        elif self.radioBtn2.isChecked():
            self.writeJSON(self.orderNo)
        else:
            self.writeXML(self.orderNo)


#########################################

if __name__ == "__main__":
    app = QApplication(sys.argv)
    mainWindow = MainWindow()
    mainWindow.show()
    app.exec_()