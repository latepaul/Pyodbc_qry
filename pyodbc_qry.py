#!/usr/bin/env python

import sys

from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
import pypyodbc

ERROR = 1
INFO = 2
NO_DSN = 1
CUST_DSN = 2

class Ui_custDSNForm(QDialog):

    def __init__(self, parent=None):
        super(Ui_custDSNForm, self).__init__(parent)
        self.setupUi(self)
        self.conn_string = ''

    def setupUi(self, custDSNForm):
        custDSNForm.setObjectName("custDSNForm")
        custDSNForm.resize(321, 206)
        custDSNForm.setMinimumSize(QSize(321, 206))
        custDSNForm.setMaximumSize(QSize(321, 206))
        self.horizontalLayout_3 = QHBoxLayout(custDSNForm)
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.verticalLayout = QVBoxLayout()
        self.verticalLayout.setObjectName("verticalLayout")
        self.horizontalLayout_2 = QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.label_2 = QLabel(custDSNForm)
        self.label_2.setObjectName("label_2")
        self.horizontalLayout_2.addWidget(self.label_2)
        spacerItem = QSpacerItem(40, 20, QSizePolicy.Minimum, QSizePolicy.Minimum)
        self.horizontalLayout_2.addItem(spacerItem)
        self.vnode = QLineEdit(custDSNForm)
        self.vnode.setObjectName("vnode")
        self.horizontalLayout_2.addWidget(self.vnode)
        self.verticalLayout.addLayout(self.horizontalLayout_2)
        self.horizontalLayout = QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.label = QLabel(custDSNForm)
        self.label.setObjectName("label")
        self.horizontalLayout.addWidget(self.label)
        spacerItem1 = QSpacerItem(25, 20, QSizePolicy.Minimum, QSizePolicy.Minimum)
        self.horizontalLayout.addItem(spacerItem1)
        self.dbName = QLineEdit(custDSNForm)
        self.dbName.setObjectName("dbName")
        self.horizontalLayout.addWidget(self.dbName)
        self.verticalLayout.addLayout(self.horizontalLayout)
        self.horizontalLayout_4 = QHBoxLayout()
        self.horizontalLayout_4.setObjectName("horizontalLayout_4")
        self.label_3 = QLabel(custDSNForm)
        self.label_3.setObjectName("label_3")
        self.horizontalLayout_4.addWidget(self.label_3)
        spacerItem2 = QSpacerItem(5, 5, QSizePolicy.Minimum, QSizePolicy.Minimum)
        self.horizontalLayout_4.addItem(spacerItem2)
        self.params = QLineEdit(custDSNForm)
        self.params.setObjectName("params")
        self.horizontalLayout_4.addWidget(self.params)
        self.verticalLayout.addLayout(self.horizontalLayout_4)
        self.buttonBox = QDialogButtonBox(custDSNForm)
        self.buttonBox.setStandardButtons(QDialogButtonBox.Cancel|QDialogButtonBox.Ok)
        self.buttonBox.setObjectName("buttonBox")
        self.verticalLayout.addWidget(self.buttonBox, Qt.AlignHCenter)
        self.horizontalLayout_3.addLayout(self.verticalLayout)

        self.retranslateUi(custDSNForm)
        QMetaObject.connectSlotsByName(custDSNForm)

    def retranslateUi(self, custDSNForm):
        _translate = QCoreApplication.translate
        custDSNForm.setWindowTitle(_translate("custDSNForm", "Custom Datasource"))
        self.label_2.setText(_translate("custDSNForm", "vnode"))
        self.vnode.setText(_translate("custDSNForm", "(local)"))
        self.label.setText(_translate("custDSNForm", "database"))
        self.label_3.setText(_translate("custDSNForm", "Other params"))
        self.buttonBox.rejected.connect(self.cancel_close)
        self.buttonBox.accepted.connect(self.save_close)

    def cancel_close(self):
        self.conn_string = ''
        self.close()

    def save_close(self):
        self.conn_string = 'Driver={Ingres};SERVER='+self.vnode.text()+';DB='+self.dbName.text()
        if self.params.text() != '':
            self.conn_string += ';'+self.params.text()
        self.close()

    def getConnStr(self):
        return self.conn_string

class Ui_MainWindow(QMainWindow):
    global conn, curr, ingres_dsns, noni_dsns

    def __init__(self, parent=None):
        super(Ui_MainWindow, self).__init__(parent)
        self.setupUi(self)

    def setupUi(self, MainWindow):

        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(523, 300)
        MainWindow.setStatusTip("")
        self.centralwidget = QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.horizontalLayout = QHBoxLayout(self.centralwidget)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.verticalLayout = QVBoxLayout()
        self.verticalLayout.setObjectName("verticalLayout")
        self.label = QLabel(self.centralwidget)
        self.label.setObjectName("label")
        self.verticalLayout.addWidget(self.label)
        self.qryText = QTextEdit(self.centralwidget)
        self.qryText.setMinimumSize(QSize(420, 70))
        self.qryText.setMaximumSize(QSize(16777215, 130))
        font = QFont()
        font.setPointSize(12)
        self.qryText.setFont(font)
        self.qryText.setObjectName("qryText")
        self.verticalLayout.addWidget(self.qryText)
        self.label_2 = QLabel(self.centralwidget)
        self.label_2.setObjectName("label_2")
        self.verticalLayout.addWidget(self.label_2)
        self.tblResults = QTableWidget(self.centralwidget)
        font = QFont()
        font.setPointSize(10)
        self.tblResults.setFont(font)
        self.tblResults.setObjectName("tblResults")
        self.tblResults.setColumnCount(0)
        self.tblResults.setRowCount(0)
        self.verticalLayout.addWidget(self.tblResults)
        self.horizontalLayout.addLayout(self.verticalLayout)
        self.verticalLayout_2 = QVBoxLayout()
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        spacerItem = QSpacerItem(20, 20, QSizePolicy.Minimum, QSizePolicy.Maximum)
        self.verticalLayout_2.addItem(spacerItem)
        self.goButton = QPushButton(self.centralwidget)
        self.goButton.setObjectName("goButton")
        self.verticalLayout_2.addWidget(self.goButton)
        spacerItem1 = QSpacerItem(20, 15, QSizePolicy.Minimum, QSizePolicy.Maximum)
        self.verticalLayout_2.addItem(spacerItem1)
        self.closeButton = QPushButton(self.centralwidget)
        self.closeButton.setObjectName("closeButton")
        self.verticalLayout_2.addWidget(self.closeButton)
        spacerItem2 = QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding)
        self.verticalLayout_2.addItem(spacerItem2)
        self.horizontalLayout.addLayout(self.verticalLayout_2)
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QMenuBar(MainWindow)
        self.menubar.setGeometry(QRect(0, 0, 523, 21))
        self.menubar.setObjectName("menubar")
        self.menuConnect = QMenu(self.menubar)
        self.menuConnect.setObjectName("menuConnect")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QStatusBar(MainWindow)
        self.statusbar.setToolTip("")
        self.statusbar.setStatusTip("")
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.actionCustom_connection = QAction(MainWindow)
        self.actionCustom_connection.setObjectName("actionCustom_connection")
        self.actionCustom_connection.triggered.connect(lambda tval: self.dostuff("custom",tval))
        self.menuConnect.addAction(self.actionCustom_connection)
        self.menuConnect.addSeparator()
        self.menubar.addAction(self.menuConnect.menuAction())
        for dsn in ingres_dsns:
            entry = QAction(dsn,MainWindow)
            self.menuConnect.addAction(entry)
            entry.triggered.connect(lambda tval,menuitem=dsn: self.dostuff(menuitem,tval))
            entry.setText(dsn)
        self.menuConnect.addSeparator()
        for dsn in noni_dsns:
            entry = QAction(dsn,MainWindow)
            self.menuConnect.addAction(entry)
            entry.triggered.connect(lambda tval,menuitem=dsn: self.dostuff(menuitem,tval))
            entry.setText(dsn)

        self.retranslateUi(MainWindow)
        QMetaObject.connectSlotsByName(MainWindow)


    def dostuff(self,dsn,tval):
        global conn, curr

        print ("doing stuff to ", dsn)
        print("meanwhile tval=",tval)
        if dsn == "custom":
            cform = Ui_custDSNForm()
            cform.exec_()
            print ("Returned:",cform.getConnStr())
            connection_string = cform.getConnStr()
            if connection_string == '':
                self.tblMessage('{no DB connection}')
                return
        else:
            connection_string = "DSN="+dsn

        print ("Attemping connection to %s" % connection_string)
        try:
            conn = pypyodbc.connect(connection_string)
            curr = conn.cursor()
            self.tblMessage("Connected to ["+dsn+"]")
        except:
            print(sys.exc_info())

            etype, evalue, etrace = sys.exc_info()
            if etype == pypyodbc.DatabaseError:
                msg=evalue.value[1]
            else:
                msg=str(evalue)
            QMessageBox.information(self,"Error connecting",msg,QMessageBox.Ok)
            conn = None

    def retranslateUi(self, MainWindow):
        _translate = QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "ODBC Query Runner"))
        self.label.setText(_translate("MainWindow", "Query:"))
        self.label_2.setText(_translate("MainWindow", "Output:"))
        self.goButton.setText(_translate("MainWindow", "Go"))
        self.closeButton.setText(_translate("MainWindow", "Close"))
        self.menuConnect.setTitle(_translate("MainWindow", "Connect"))
        self.actionCustom_connection.setText(_translate("MainWindow", "Custom connection"))

        self.goButton.clicked.connect(self.button_pushed)
        self.closeButton.clicked.connect(self.close)
        self.tblMessage("{no database connection}")

    def button_pushed(self):
        global conn, curr

        self.tblResults.setRowCount(0)
        self.tblResults.setColumnCount(0)

        if not conn:
            self.errMessage(INFO,"No DB connection","Unable to run queries as no database connection exists")
            return

        qtext = self.qryText.toPlainText()
        if qtext == "":
            return

        try:
            c=curr.execute(qtext)
        except pypyodbc.DatabaseError:
            self.errMessage(ERROR,"Database Error",sys.exc_info()[1].value[1])
            return
        except:
            print (sys.exc_info())
            self.errMessage(ERROR,"Sys Error",str(sys.exc_info()[1]))
            return

        if c.rowcount < 0:

            rs = curr.fetchall()
            colnames =[]
            for row in curr.description:
                colnames.append(row[0])

            l = len(rs)
            if l == 0:
                self.tblMessage("0 rows returned")
            else:
                self.tblMessage("%d rows returned" % l)
                r = len(rs[0])
                self.tblResults.setRowCount(l)
                self.tblResults.setColumnCount(r)

                self.tblResults.setHorizontalHeaderLabels(colnames)
                for i,row in enumerate(rs):
                    for j, col in enumerate(row):
                        item = QTableWidgetItem(str(col))
                        item.setFlags(item.flags() & ~ Qt.ItemIsEditable)
                        self.tblResults.setItem(i,j,item)
        else:
            self.tblMessage("%d rows affected." % c.rowcount)

        curr.commit()

    def tblMessage(self,msg):
        self.statusbar.showMessage(msg)

    def errMessage(self,type,title,msg):
        if type == ERROR:
            QMessageBox.critical(self,title,msg,QMessageBox.Ok)
        else:
            QMessageBox.information(self,title,msg,QMessageBox.Ok)
        pass


def main():
    global curr, conn, ingres_dsns, noni_dsns

    conn = None
    app = QApplication(sys.argv)

    sources = pypyodbc.dataSources()

    ingres_dsns = []
    for item in sources:
        if sources[item].decode() == "Ingres":
            ingres_dsns.append(item.decode())

    noni_dsns = []
    for item in sources:
        if sources[item].decode() != "Ingres":
            noni_dsns.append(item.decode())

    form = Ui_MainWindow()

    form.show()
    app.exec_()
    del form
    if conn:
        curr.close()
        conn.close()


if __name__ == '__main__':
    main()
