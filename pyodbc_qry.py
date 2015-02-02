#!/usr/bin/env python

import sys

from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
import pypyodbc

class Ui_Form(QWidget):
    global curr, conn

    def __init__(self, parent=None):
        super(Ui_Form, self).__init__(parent)
        self.setupUi(self)

    def setupUi(self,Dialog):
        self.resize(650, 302)
        self.horizontalLayout_2 = QHBoxLayout(Dialog)
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.verticalLayout = QVBoxLayout()
        self.verticalLayout.setObjectName("verticalLayout")
        self.label = QLabel(Dialog)
        self.label.setObjectName("label")
        self.verticalLayout.addWidget(self.label)
        self.qryText = QTextEdit(Dialog)
        self.qryText.setMinimumSize(QSize(420, 71))
        self.qryText.setMaximumSize(QSize(16777215, 130))
        self.qryText.setObjectName("qryText")
        self.verticalLayout.addWidget(self.qryText)
        self.label_2 = QLabel(Dialog)
        self.label_2.setObjectName("label_2")
        self.verticalLayout.addWidget(self.label_2)
        self.tblResults = QTableWidget(Dialog)
        self.tblResults.setObjectName("tblResults")
        self.tblResults.setColumnCount(0)
        self.tblResults.setRowCount(0)
        self.verticalLayout.addWidget(self.tblResults)
        self.label_3 = QLabel(Dialog)
        self.label_3.setObjectName("label_3")
        self.verticalLayout.addWidget(self.label_3)
        self.lineEdit = QLineEdit(Dialog)
        self.lineEdit.setAlignment(Qt.AlignCenter)
        self.lineEdit.setObjectName("lineEdit")
        self.lineEdit.setReadOnly(True)
        self.verticalLayout.addWidget(self.lineEdit)
        self.horizontalLayout_2.addLayout(self.verticalLayout)
        self.verticalLayout_3 = QVBoxLayout()
        self.verticalLayout_3.setObjectName("verticalLayout_3")
        spacerItem = QSpacerItem(20, 5, QSizePolicy.Minimum, QSizePolicy.Maximum)
        self.verticalLayout_3.addItem(spacerItem)
        self.goButton = QPushButton(Dialog)
        self.goButton.setObjectName("goButton")
        self.verticalLayout_3.addWidget(self.goButton)
        spacerItem1 = QSpacerItem(20, 15, QSizePolicy.Minimum, QSizePolicy.Maximum)
        self.verticalLayout_3.addItem(spacerItem1)
        self.closeButton = QPushButton(Dialog)
        self.closeButton.setObjectName("closeButton")
        self.verticalLayout_3.addWidget(self.closeButton)
        spacerItem2 = QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding)
        self.verticalLayout_3.addItem(spacerItem2)
        self.horizontalLayout_2.addLayout(self.verticalLayout_3)

        self.retranslateUi(Dialog)
        QMetaObject.connectSlotsByName(Dialog)

    def retranslateUi(self, Dialog):
        _translate = QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Simple Query Runner"))
        self.label.setText(_translate("Dialog", "Query"))
        self.label_2.setText(_translate("Dialog", "Output"))
        self.label_3.setText(_translate("Dialog", "Status"))
        self.lineEdit.setText(_translate("Dialog", "0 rows returned."))
        self.goButton.setText(_translate("Dialog", "Go"))
        self.closeButton.setText(_translate("Dialog", "Close"))

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
        self.lineEdit.setText(msg)



def main():
    global curr, conn

    app = QApplication(sys.argv)

    # Would normally be invoked as modal dialog.
    # But for simplicity we use it as the main form here.
    form = Ui_Form()
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
