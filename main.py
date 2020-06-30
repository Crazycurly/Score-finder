import sys
import requests
import xlwings as xw
from bs4 import BeautifulSoup

from Ui_main import *
from PyQt5.QtWidgets import QWidget, QLabel, QApplication,QFileDialog
from PyQt5.QtCore import QThread, Qt, pyqtSignal, pyqtSlot
from PyQt5.Qt import Qt

base_url = 'https://www.com.tw/'

class MyWindow(Ui_Dialog, QtWidgets.QLabel):

    def __init__(self, parent=None):
        super(MyWindow, self).__init__(parent)
        self.setupUi(self)
        self.pushButton.clicked.connect(self.start)
        
    def start(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        f = QFileDialog.getSaveFileName(self,"QFileDialog.getSaveFileName()","","Excel 活頁簿 (*.xlsx)", options=options)

        if f[0] != '':
            app = xw.App(visible=True,add_book=False)
            wb = app.books.add()
            group = self.groupid_get(base_url+'vtech/groupid_list'+self.lineEdit.text()+'.html')

            for g in reversed(group):
                v = self.vtech_get(g['href'])
                t = self.techreg_get(g['href'])
                self.creat_exl(wb,g.text,v,t)
                print(g.text,'\t',g['href'])
            wb.save(f[0]+'.xlsx')



    def groupid_get(self,url):
        response = requests.get(url)
        soup = BeautifulSoup(response.text, "html.parser")
        table = soup.find('table', {'id': 'table1'})
        trs = table.find_all('a')
        return(trs)


    def vtech_get(self,url):
        response = requests.get(base_url+'vtech/'+ url)
        soup = BeautifulSoup(response.text, "html.parser")

        table = soup.find('table', {'id': 'table1'})
        trs = table.find_all('tr')[3:]
        rows = list()
        for tr in trs:
            rows.append([td.text.replace('\xa0', '').replace(' ', '')
                        for td in tr.find_all('td') if td != ['']])
        while [''] in rows:
            rows.remove([''])
        for i in rows:
            i[0] = str(i[0]).split('\n')[0].replace('(', '').replace(')', '')

        return rows


    def techreg_get(self,url):
        response = requests.get(base_url+'techreg/'+ url)
        soup = BeautifulSoup(response.text, "html.parser")

        table = soup.find('table', {'id': 'table1'})
        trs = table.find_all('tr')[3:]
        rows = list()

        for tr in trs:
            rows.append([td.text.replace('\xa0', '').replace(' ', '')
                        for td in tr.find_all('td') if td != ['']])
        while [''] in rows:
            rows.remove([''])
        for i in rows:
            i[0] = str(i[0]).split('\n')[0].replace('(', '').replace(')', '')
            del i[2]
            tmp = i[2].split('/')
            i[2] = tmp[0]
            i.append(tmp[1])
        return rows


    def creat_exl(self,book, n, vtech, techreg):
        sht = book.sheets.add(name=n)
        sht.range('a1').value = ['甄選入學']
        sht.range('a1:c1').api.merge()
        sht.range('a1:c2').api.HorizontalAlignment = -4108

        sht.range('a2').value = [['校系代碼', '科系名稱', '通過篩選標準']]
        sht.range('e1').value = ['登記分發']
        sht.range('e1:h1').api.merge()
        sht.range('e1:h2').api.HorizontalAlignment = -4108

        sht.range('e2').value = [['校系代碼', '科系名稱', '加權總分', '平均']]
        sht.range('a3').value = vtech
        sht.range('e3').value = techreg
        sht.autofit()


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    myWin = MyWindow()
    myWin.show()
    sys.exit(app.exec_())