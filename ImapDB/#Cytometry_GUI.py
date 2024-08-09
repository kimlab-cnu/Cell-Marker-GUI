import sys
from PyQt5.QtWidgets import *
from PyQt5 import uic
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt

import pandas as pd
Lym_file = pd.read_excel("C:/Users/DELL/Desktop/Python/Lym_input.xlsx")
marker_data = Lym_file['Cell Type_1']
marker_data1 = marker_data.values.tolist()
marker_data2 = pd.Series(marker_data1).dropna().tolist()
marker_data2.insert(0, 'Cell Type')
new_list = []
for v in marker_data2:
    if v not in new_list:
        new_list.append(v)
Lym_word = new_list

My_file = pd.read_excel("C:/Users/DELL/Desktop/Python/Myeloid_input.xlsx")
my_data = My_file['Cell Type_1']
my_data1 = my_data.values.tolist()
my_data2 = pd.Series(my_data1).dropna().tolist()
my_data2.insert(0, 'Cell Type')
my_list = []
for v in my_data2:
    if v not in my_list:
        my_list.append(v)
My_word = my_list
All_file = pd.read_excel("C:/Users/DELL/Desktop/Python/All_input.xlsx")

form_home = uic.loadUiType("Intro_Real_fin.ui")[0]
form_immune = uic.loadUiType("Immune_Dialog.ui")[0]
form_CyTOF = uic.loadUiType("CyTOF_Dialog.ui")[0]
form_FACS = uic.loadUiType("FACS_Dialog.ui")[0]

class Main(QMainWindow, QWidget, form_home):
    def __init__(self):
        super().__init__()
        self.initUI()


    def initUI(self):
        KimLab = QLabel(self)
        KimLab.setPixmap(QPixmap("C:/Users/DELL/Desktop/Python/KimLab_Logo.jpg"))
        KimLab.setGeometry(345, 450, 220, 80)

        KimLab.setScaledContents(True)
        KimLab.setAlignment(Qt.AlignCenter)

        self.setupUi(self)
        self.Immune_btn.clicked.connect(self.btn_second)
        self.CyTOF_btn.clicked.connect(self.btn_CyTOF)
        self.FACS_btn.clicked.connect(self.btn_FACS)
        self.show()

    def btn_second(self): #Immune cell
        self.close()
        self.second = HomeWindow()
        self.second.exec()
        self.show()


    def btn_CyTOF(self): #CyTOF panel
        self.close()
        self.CyTOF = CyTOFWindow()
        self.CyTOF.exec()
        self.show()

    def btn_FACS(self): #FACS panel
        self.close()
        self.FACS = FACSWindow()
        self.FACS.exec()
        self.show()


class HomeWindow(QDialog, QWidget, form_immune):
    def __init__(self):
        super(HomeWindow, self).__init__()
        self.initUI()
        self.show()
        self.setWindowTitle("Immune Cell Marker")

        self.C1.currentIndexChanged.connect(self.c1_changed)
        self.C2.currentIndexChanged.connect(self.c2_changed)
        self.Done_btn.clicked.connect(self.accept)

    def c1_changed(self, index):
        selected_option = self.C1.currentText()
        self.C2.clear()
        if selected_option == "Lymphocyte":
            self.C2.addItems(Lym_word)
        elif selected_option == "Myeloid cell":
            self.C2.addItems(My_word)
        elif selected_option == "Cell type":
            self.C2.addItem("Cell type")

    def c2_changed(self, index):
        selected_option1 = self.C1.currentText()
        selected_option = self.C2.currentText()

        plan = Lym_file[Lym_file['Cell Type_1'] == selected_option]
        plan1 = plan.values.tolist()
        plan2 = pd.Series(plan1).dropna().tolist()
        plan2.insert(0, 'Cell Type')

        want = plan['Cell Type_2']
        want1 = want.values.tolist()
        want2 = pd.Series(want1).dropna().tolist()
        want2.insert(0, 'Cell Type')

        new_list1 = []
        for v in want2:
            if v not in new_list1:
                new_list1.append(v)
        List = new_list1

        planB = My_file[My_file['Cell Type_1'] == selected_option]
        planB1 = planB.values.tolist()
        planB2 = pd.Series(planB1).dropna().tolist()
        planB2.insert(0, 'Cell Type')

        wantB = planB['Cell Type_2']
        wantB1 = wantB.values.tolist()
        wantB2 = pd.Series(wantB1).dropna().tolist()
        wantB2.insert(0, 'Cell Type')

        my_list1 = []
        for v in wantB2:
            if v not in my_list1:
                my_list1.append(v)
        ListB = my_list1

        self.C3.clear()
        if selected_option1 == "Lymphocyte":
            self.C3.addItems(List)
        elif selected_option1 == "Myeloid cell":
            self.C3.addItems(ListB)

    def accept(self):
        selected_option1 = self.C2.currentText()
        selected_option2 = self.C3.currentText()
        marker = All_file[(All_file['Cell Type_1'] == selected_option1) & (All_file['Cell Type_2'] == selected_option2)]
        marker1 = marker.values.tolist()
        marker2 = pd.Series(marker1).dropna().tolist()
        print(marker2)

        self.Show_Marker.setRowCount(marker.shape[0])
        self.Show_Marker.setColumnCount(marker.shape[1])
        for row in range(marker.shape[0]):
           for col in range(marker.shape[1]):
                item = QTableWidgetItem(str(marker.iloc[row, col]))
                self.Show_Marker.setItem(row, col, item)


    def initUI(self):
        self.setupUi(self)
        self.Home_btn.clicked.connect(self.close_btn)
        self.show()

    def close_btn(self):
        self.close()


class CyTOFWindow(QDialog, QWidget, form_CyTOF):
    def __init__(self):
        super(CyTOFWindow, self).__init__()
        self.initUI()
        self.show()
        self.resize(1200, 800)
        self.df_list = []

        horizon = QHBoxLayout()
        horizon.addWidget(self.sheet, 4)
        horizon.addWidget(self.Com_btn,2)
        horizon.addWidget(self.Ex_btn,2)
        horizon.addWidget(self.Home_btn,1)

        vertical = QVBoxLayout()
        vertical.addWidget(self.label)
        vertical.addLayout(horizon)
        vertical.addWidget(self.panel)
        self.setLayout(vertical)
        self.Com_btn.clicked.connect(self.Com_open)
        self.Ex_btn.clicked.connect(self.Ex_open)
        self.sheet.currentIndexChanged[int].connect(self.cmbChanged)

    def Com_open(self):
        self.sheet.clear()
        file_path = 'C:/Users/DELL/Desktop/Python/Com_CyTOF_panel.xlsx'
        if file_path:
            self.df_list = self.loadData(file_path)
            for i in self.df_list:
                self.sheet.addItem(i.name)

            self.initTableWidget(0)

    def Ex_open(self):
        self.sheet.clear()
        file_path = 'C:/Users/DELL/Desktop/Python/Ex_CyTOF_panel.xlsx'
        if file_path:
            self.df_list = self.loadData(file_path)
            for i in self.df_list:
                self.sheet.addItem(i.name)

            self.initTableWidget(0)

    def cmbChanged(self, id):
        self.initTableWidget(id)

    def loadData(self, file_name):
        df_list = []
        with pd.ExcelFile(file_name) as wb:
            for i, sn in enumerate(wb.sheet_names):
                try:
                    df = pd.read_excel(wb, sheet_name=sn)
                except Exception as e:
                    print('File read error:', e)
                else:
                    df = df.fillna(0)
                    df.name = sn
                    df_list.append(df)
        return df_list

    def initTableWidget(self, id):
        self.panel.clear()
        df = self.df_list[id]
        col = len(df.keys())
        self.panel.setColumnCount(col)
        self.panel.setHorizontalHeaderLabels(df.keys())

        row = len(df.index)
        self.panel.setRowCount(row)
        self.writeTableWidget(id, df, row, col)

    def writeTableWidget(self, id, df, row, col):
        for r in range(row):
            for c in range(col):
                item = QTableWidgetItem(str(df.iloc[r][c]))
                self.panel.setItem(r, c, item)
        self.panel.resizeColumnsToContents()

    def initUI(self):
        self.setupUi(self)
        self.Home_btn.clicked.connect(self.close_btn)
        self.show()

    def close_btn(self):
        self.close()


class FACSWindow(QDialog, QWidget, form_FACS):
    def __init__(self):
        super(FACSWindow, self).__init__()
        self.initUI()
        self.show()
        self.resize(1200, 800)  # 위젯 사이즈
        self.df_list = []

        horizon = QHBoxLayout()
        horizon.addWidget(self.sheet1, 4)
        horizon.addWidget(self.Com_btn,2)
        horizon.addWidget(self.Ex_btn,2)
        horizon.addWidget(self.Home_btn,1)

        vertical = QVBoxLayout()

        vertical.addWidget(self.label1)
        vertical.addLayout(horizon)
        vertical.addWidget(self.panel1)

        self.setLayout(vertical)

        self.Com_btn.clicked.connect(self.Com_open)
        self.Ex_btn.clicked.connect(self.Ex_open)

        self.sheet1.currentIndexChanged[int].connect(self.cmbChanged)

    def Com_open(self):
        self.sheet1.clear()
        file_path = 'C:/Users/DELL/Desktop/Python/Com_FACS_Panel.xlsx'
        if file_path:
            self.df_list = self.loadData(file_path)
            for i in self.df_list:
                self.sheet1.addItem(i.name)

            self.initTableWidget(0)

    def Ex_open(self):
        self.sheet1.clear()
        file_path = 'C:/Users/DELL/Desktop/Python/Ex_FACS_panel.xlsx'
        if file_path:
            self.df_list = self.loadData(file_path)
            for i in self.df_list:
                self.sheet1.addItem(i.name)

            self.initTableWidget(0)


    def cmbChanged(self, id):
        self.initTableWidget(id)

    def loadData(self, file_name):
        df_list = []
        with pd.ExcelFile(file_name) as wb:
            for i, sn in enumerate(wb.sheet_names):
                try:
                    df = pd.read_excel(wb, sheet_name=sn)
                except Exception as e:
                    print('File read error:', e)
                else:
                    df = df.fillna(0)
                    df.name = sn
                    df_list.append(df)
        return df_list

    def initTableWidget(self, id):
        self.panel1.clear()
        df = self.df_list[id];
        col = len(df.keys())
        self.panel1.setColumnCount(col)
        self.panel1.setHorizontalHeaderLabels(df.keys())

        row = len(df.index)
        self.panel1.setRowCount(row)
        self.writeTableWidget(id, df, row, col)

    def writeTableWidget(self, id, df, row, col):
        for r in range(row):
            for c in range(col):
                item = QTableWidgetItem(str(df.iloc[r][c]))
                self.panel1.setItem(r, c, item)
        self.panel1.resizeColumnsToContents()

    def initUI(self):
        self.setupUi(self)
        self.Home_btn.clicked.connect(self.close_btn)
        self.show()

    def close_btn(self):
        self.close()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    m = Main()
    m.show()
    sys.exit(app.exec_())
