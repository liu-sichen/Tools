from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog
import sys
import Ui_main_window
from pdf2docx import parse
import os
import pandas as pd
import regex as re

class MainWindow(QMainWindow, Ui_main_window.Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.initUI()
        pattern = re.compile(r"[a-zA-Z0-9_-]+@[a-zA-Z0-9_-]+(?:\.[a-zA-Z0-9_-]+)")
        info = '私人邮箱xxxx@xxxx.com，工作邮箱是yyyy@yyyy.com。'
        result = pattern.findall(info)
        self.infoContentLabel.setText(result[0])

    def initUI(self):
        self.PDFToWordBtn.clicked.connect(self.PDFToWordBtnRun)
        self.searchExcelBtn.clicked.connect(self.SearchExcelBtnRun)
        self.StatisticsBtn.clicked.connect(self.StatisticsExcelBtnRun)
        return

    def PDFToWordBtnRun(self):
        filename = QFileDialog.getOpenFileName(self, "选取文件", './', "PDF Files (*.PDF);;All Files (*)")
        self.runOutPutTextBrowser.setText("开始转换 %s，请稍等。" %filename[0])
        baseName=os.path.splitext(filename[0])[0]
        docxFile = os.path.join(baseName, '.docx')
        parse(filename[0], docxFile)
        self.runOutPutTextBrowser.append("转换完成。")
        return

    def SearchExcelBtnRun(self):
        filename = QFileDialog.getOpenFileName(self, "选取文件", './', "Excel Files (*.xlsx);;All Files (*)")
        sheet = pd.read_excel(filename[0], header = None)
        rs = int(self.rsLineEdit.text()) - 1
        cs = int(self.csLineEdit.text()) - 1
        re = int(self.reLineEdit.text())
        ce = int(self.ceLineEdit.text())
        target = self.targetLineEdit.text()
        self.runOutPutTextBrowser.setText("开始搜索 %s，请稍等。" %filename[0])
        for i in range(rs, re):
            for j in range(cs, ce):
                    print(sheet.iloc[i, j], i ,j)
                    if sheet.iloc[i, j] == target:
                        self.runOutPutTextBrowser.append("搜索到%s, 在%d行 %d列。" %(target, (i + 1), (j + 1)))
        self.runOutPutTextBrowser.append("搜索完成。")
        return

    def StatisticsExcelBtnRun(self):
        filename = QFileDialog.getOpenFileName(self, "选取文件", './', "Excel Files (*.xlsx);;All Files (*)")
        originSheet = pd.read_excel(filename[0], sheet_name="origin", header = None)
        StatisticsSheet = pd.read_excel(filename[0], sheet_name="Statistics", header = None)
        rsOrigin = int(self.rsStaLineEdit.text()) - 1
        csOrigin = int(self.csStaLineEdit.text()) - 1
        reOrigin = int(self.reStaLineEdit.text())
        ceOrigin = int(self.ceStaLineEdit.text())
        rsStatistics = int(self.rsStaTarLineEdit.text()) - 1
        reStatistics = int(self.reStaTarLineEdit.text())
        split1Col = int(self.colStaSplit1LineEdit.text()) - 1
        split2Col = int(self.colStaSplit2LineEdit.text()) - 1
        split1Count = 0
        split2Count = 0
        staPos1 = int(self.staPos1LineEdit.text()) - 1
        staPos2 = int(self.staPos2LineEdit.text()) - 1
        self.runOutPutTextBrowser.setText("开始统计 %s，请稍等。" %filename[0])
        for si in range(rsStatistics, reStatistics):
            target = StatisticsSheet.iloc[si, 1]
            split1Count = 0
            split2Count = 0
            for oi in range(rsOrigin, reOrigin):
                for oj in range(csOrigin, ceOrigin):
                    if originSheet.iloc[oi, oj] == target:
                        if oj >= 0 and oj < split1Col:
                            split1Count = split1Count + 1
                        elif oj >= split1Col and oj < split2Col:
                            split2Count = split2Count + 1
            StatisticsSheet.iloc[si, staPos1] = split1Count
            StatisticsSheet.iloc[si, staPos2] = split2Count
        dfResult = pd.DataFrame(StatisticsSheet)
        dfResult.to_excel('result.xlsx', index=False)
        self.runOutPutTextBrowser.append("统计完成。")
        return

if __name__ == '__main__':
    app = QApplication(sys.argv)
    main = MainWindow()
    main.show()
    sys.exit(app.exec_())
