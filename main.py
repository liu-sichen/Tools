import sys
from PyQt5.QtWidgets import QApplication, QMainWindow
import Ui_main_window

str = 'Hello World'
print(str)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    MainWindow = QMainWindow()
    ui = Ui_main_window.Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())

# import pandas as pd
#
# fileExamPlan = "data/监考安排.xlsx"
# fileExamTimes = "data/监考次数.xlsx"
#
# sheetExamPlan = pd.read_excel(fileExamPlan)
# sheetExamTimes = pd.read_excel(fileExamTimes, sheet_name=0, index_col=0)
#
# print(sheetExamPlan.iloc[35, 8])
# print(sheetExamPlan.iloc[60, 9])
# for i in range(0, 94):
#     for j in range(35, 61):
#         for k in range(8, 10):
#             if sheetExamTimes.iloc[i, 0] == sheetExamPlan.iloc[j, k]:
#                 print(sheetExamTimes.iloc[i, 0] )
#                 sheetExamTimes.iloc[i, 5] = 1
#
# sheetExamTimes.to_excel('data/test2.xlsx')

# import sys
# from PyQt5.QtWidgets import QApplication , QMainWindow
# import test
# if __name__ == '__main__':
#     app = QApplication(sys.argv)
#     main= QMainWindow()
#     ui = test.Ui_Form()
#     ui.setupUi(main)
#     main.show()
#     sys.exit(app.exec_())