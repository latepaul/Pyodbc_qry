#!/usr/bin/env python

import sys

from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
import pypyodbc

class Ui_MainWindow(QMainWindow):
    global conn, curr

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
        self.menuConnect.addAction(self.actionCustom_connection)
        self.menuConnect.addSeparator()
        self.menubar.addAction(self.menuConnect.menuAction())

        self.retranslateUi(MainWindow)
        QMetaObject.connectSlotsByName(MainWindow)

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

    def contextMenuEvent(self, event):
        menu = QMenu(self)
        menu.addAction(self.cutAct)
        menu.addAction(self.copyAct)
        menu.addAction(self.pasteAct)
        menu.exec_(event.globalPos())

    def button_pushed(self):

        self.tblResults.setRowCount(0)
        self.tblResults.setColumnCount(0)

        if not conn:
            QMessageBox.information(self,"No DB connection","Unable to run queries as no database connection exists",QMessageBox.Ok)
            return

        qtext = self.qryText.toPlainText()
        if qtext == "":
            return

        try:
            c=curr.execute(qtext)
        except pypyodbc.ProgrammingError:
            self.tblMessage(sys.exc_info()[1].value[1])
            return
        except:
            print (sys.exc_info())
            self.tblMessage("SYS:"+sys.exc_info()[1].value[1])
            return

        if c.rowcount < 0:

            rs = curr.fetchall()

            l = len(rs)
            if l == 0:
                self.tblMessage("0 rows returned")
            else:
                self.tblMessage("%d rows returned" % l)
                r = len(rs[0])
                self.tblResults.setRowCount(l)
                self.tblResults.setColumnCount(r)

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



def main():
    global curr, conn

    app = QApplication(sys.argv)

    # Would normally be invoked as modal dialog.
    # But for simplicity we use it as the main form here.
    form = Ui_MainWindow()
    print(type(form))

    sources = pypyodbc.dataSources()
    print (sources)
    dsns = list(sources.keys())
    print (dsns)
    dsns.sort()
    sl = []
    for dsn in dsns:
        sl.append('%s [%s]' % (dsn, sources[dsn]))
    print('\n'.join(sl))

    connection_string = 'Driver={Ingres};SERVER=(local);DB=pmdb2'

    connection_string = "DSN="+dsn.decode("utf-8")
    print (connection_string)
    try:
        conn = pypyodbc.connect(connection_string)
        curr = conn.cursor()
    except:
        msg=sys.exc_info()[1].value[1]
        QMessageBox.information(form,"Error connecting",msg,QMessageBox.Ok)
        conn = None

    form.show()
    app.exec_()
    del form
    if conn:
        curr.close()
        conn.close()


if __name__ == '__main__':
    main()
