import sys
import numpy as np
from ERD import *
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QMessageBox,QComboBox,QLabel,QSizePolicy
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import pyqtSlot
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import matplotlib.pyplot as plt
import random

class Finance_Passion_Interface:
   def __init__(self,file):
      self.window(file)
   def window(self,file):
      self.app = QApplication(sys.argv)
      self.win = QWidget()
      self.e1 = erd(file)
      self.e2=self.e1.get_category()
      self.combo1 = QComboBox(self.win)
      self.combo1.addItem("Select Category")
      for i in self.e2:
         self.combo1.addItem(i)
      self.combo1.move(80,25)
      self.label1 = QLabel(self.win)
      self.label1.move(25, 25)
      self.label2 = QLabel(self.win)
      self.label2.move(25, 50)
      self.label2.setText("Total Current Value:")
      self.label21 = QLabel(self.win)
      self.label21.move(230, 50)
      self.label21.setText("Total Monthly Installment Amount:")
      self.label22 = QLabel(self.win)
      self.label22.move(425,50)
      self.label22.setText(str(self.e1.get_total_installment_value()))
      self.label3=QLabel(self.win)
      self.label3.move(180, 50)
      self.label3.setText(str(self.e1.get_total_current_value()))
      self.label3.adjustSize()
      self.label4 = QLabel(self.win)
      self.label4.move(200,25)
      self.label5 = QLabel(self.win)
      self.label5.move(2, 75)
      self.label5.setText("---------------------------------------------------------------------------------------------------------------------------")
      self.label6 = QLabel(self.win)
      self.label6.move(25, 85)
      self.label6.setText("Account Details")
      self.label1.setText("Category")
      self.label7 = QLabel(self.win)
      self.label7.move(2, 95)
      self.label7.setText("---------------------------------------------------------------------------------------------------------------------------")
      self.label5.hide()
      self.label6.hide()
      self.label7.hide()
      self.label8 = QLabel(self.win)
      self.label8.move(25,110)
      self.label8.setText("Account #:")
      self.label8.hide()
      self.label9 = QLabel(self.win)
      self.label9.move(80, 110)
      self.label9.setText("")
      self.label9.hide()
      self.label10 = QLabel(self.win)
      self.label10.setText("Installement Amount:")
      self.label10.move(175,110)
      self.label10.hide()
      self.label11 = QLabel(self.win)
      self.label11.move(280, 110)
      self.label11.setText("")
      self.label11.hide()
      self.label12 = QLabel(self.win)
      self.label12.move(25,130)
      self.label12.setText("Remaining Installement:")
      self.label12.hide()
      self.label13 = QLabel(self.win)
      self.label13.move(145, 130)
      self.label13.setText("")
      self.label13.hide()
      self.label14 = QLabel(self.win)
      self.label14.move(325, 110)
      self.label14.setText("Maturity Amount:")
      self.label14.hide()
      self.label15 = QLabel(self.win)
      self.label15.move(410, 110)
      self.label15.setText("")
      self.label15.hide()
      self.label16 = QLabel(self.win)
      self.label16.move(175, 130)
      self.label16.setText("Start Date:")
      self.label16.hide()
      self.label17 = QLabel(self.win)
      self.label17.move(235, 130)
      self.label17.setText("")
      self.label17.hide()
      self.label18 = QLabel(self.win)
      self.label18.move(325, 130)
      self.label18.setText("Maturity Date:")
      self.label18.hide()
      self.label19 = QLabel(self.win)
      self.label19.move(400, 130)
      self.label19.setText("")
      self.label19.hide()
      self.combo1.activated[str].connect(self.catChanged)
      self.combo2 = QComboBox(self.win)
      self.combo2.move(350,25)
      self.combo2.hide()
      self.combo2.activated[str].connect(self.acctChanged)

      self.m=[]
      self.m.append()
      self.E=[]
      self.expl=[]
      for i in self.e2:
         e3=self.e1.get_current_value_by_category(i)
         self.E.append(e3)
         self.expl.append(0)
      self.m=MultiCanvas(self.win, width=5, height=4,list1=self.e2,list2=self.E,list3=self.expl)
      self.m.move(2,175)

      self.win.setGeometry(50, 50, 500, 600)
      self.win.setWindowTitle("Finance Passion")
      self.win.show()
      sys.exit(self.app.exec_())

   def catChanged(self, text):
      self.label5.hide()
      self.label6.hide()
      self.label7.hide()
      self.label8.hide()
      self.label9.hide()
      self.label10.hide()
      self.label11.hide()
      self.label12.hide()
      self.label13.hide()
      self.label14.hide()
      self.label15.hide()
      self.label16.hide()
      self.label17.hide()
      self.label18.hide()
      self.label19.hide()
      if(text=="Select Category"):
         self.label2.setText("Total Current Value:")
         self.label2.adjustSize()
         self.label21.setText("Total Monthly Installment:")
         self.label4.setText("")
         self.combo2.hide()
         self.combo2.clear()
         expl=[0 for i in range(len(self.m.lab))]
         self.m.graph_clear()
         self.m.pie_plot(expl)
         self.m.move(2, 175)
      else:
         self.label2.setText("Total "+text+" current value")
         self.label21.setText("Total "+text+" Monthly Installment :")
         self.label4.setText(text+" Account List:")
         self.combo2.clear()
         self.combo2.show()
         self.m.graph_clear()
         expl=[]
         for i in self.e1.get_category():
            if(i==text):
               expl.append(0.1)
            else:
               expl.append(0)
         del self.m
         self.m = MultiCanvas(self.win, width=5, height=4, list1=self.e2, list2=self.E, list3=self.expl)
         self.m.move(2, 175)
         e3=self.e1.get_account_list_by_category(text)
         self.combo2.addItem("Account List")
         for i in e3:
            self.combo2.addItem(i)
         self.combo2.adjustSize()
      self.label2.adjustSize()
      self.label4.adjustSize()
      self.label21.adjustSize()
      if (text == "Select Category"):
         self.label3.setText(str(self.e1.get_total_current_value()))
         self.label22.setText(str(self.e1.get_total_installment_value()))
      else:
         self.label3.setText(str(self.e1.get_current_value_by_category(text)))
         self.label22.setText(str(self.e1.get_installement_value_by_category(text)))
      self.label3.adjustSize()
      self.label22.adjustSize()

   def acctChanged(self,acct):
      self.label5.show()
      self.label6.show()
      self.label7.show()
      self.label8.show()
      self.label9.show()
      self.label10.show()
      self.label11.show()
      self.label12.show()
      self.label13.show()
      self.label14.show()
      self.label15.show()
      self.label16.show()
      self.label17.show()
      self.label18.show()
      self.label19.show()
      if (acct == "Account List"):
         self.label5.hide()
         self.label6.hide()
         self.label7.hide()
         self.label8.hide()
         self.label9.hide()
         self.label10.hide()
         self.label11.hide()
         self.label12.hide()
         self.label13.hide()
         self.label14.hide()
         self.label15.hide()
         self.label16.hide()
         self.label17.hide()
         self.label18.hide()
         self.label19.hide()
      else:
         self.label9.setText(acct)
         self.label9.adjustSize()
         self.label11.setText(str(self.e1.get_amount_by_account(acct)))
         self.label11.adjustSize()
         self.label13.setText(str(self.e1.get_remaining_by_account(acct)))
         self.label13.adjustSize()
         self.label15.setText(str(self.e1.get_maturity_amount_by_account(acct)))
         self.label15.adjustSize()
         self.label17.setText(str(self.e1.get_start_date_by_account(acct)))
         self.label17.adjustSize()
         self.label19.setText(str(self.e1.get_maturity_date_by_account(acct)))
         self.label19.adjustSize()

class MultiCanvas(FigureCanvas):
   def absolute_value(self,val):
      a = np.round(val / 100. * self.nsizes.sum(), 0)
      return a
   def __init__(self,parent=None, width=5, height=4, dpi=100,list1=[],list2=[],list3=[]):
      self.fig = Figure(figsize=(width, height), dpi=dpi)
      self.axes = self.fig.add_subplot(111)
      FigureCanvas.__init__(self, self.fig)
      self.setParent(parent)
      self.lab=list1[:]
      self.dat=list2[:]
      FigureCanvas.setSizePolicy(self,
                                 QSizePolicy.Expanding,
                                 QSizePolicy.Expanding)
      FigureCanvas.updateGeometry(self)
      self.pie_plot(list3)
   def pie_plot(self,explode):
      self.nsizes=np.array(self.dat)
      self.axes.pie(self.nsizes, explode=explode, labels=self.lab, autopct=self.absolute_value, shadow=False)
      self.axes.axis('equal')
      self.draw()
   def graph_clear(self):
      self.fig.clear()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    main_file="E:\Finance Passion\Deposits.xls"
    ex = Finance_Passion_Interface(main_file)
    sys.exit(app.exec_())
