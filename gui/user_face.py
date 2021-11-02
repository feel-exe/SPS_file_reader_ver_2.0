import sys
from PyQt5 import *
from PyQt5.QtWidgets import *
# from scripts.converter_vlt import *
# from scripts.converter_txt import *
from scripts.converter_vlt_cl import *
from scripts.converter_txt_cl import *


class DlgMain(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("SPS reader")
        # self.ledText = QLineEdit("Default text", self)
        self.resize(300, 70)
        # self.ledText.move(100, 50)

        self.btn_vlt = QPushButton("vlt in TXT", self)
        self.btn_txt = QPushButton("TXT in excel", self)
        # setGeometry(200, 150, 100, 40)
        self.btn_vlt.setGeometry(200, 150, 100, 40)
        self.btn_txt.setGeometry(100, 150, 100, 40)

        self.btn_vlt.move(25, 20)
        self.btn_txt.move(175, 20)

        self.btn_vlt.clicked.connect(self.evt_vlt_clicked)
        self.btn_txt.clicked.connect(self.evt_txt_clicked)

    def evt_vlt_clicked(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fname, _ = QFileDialog.getOpenFileName(self, "QFileDialog.getOpenFileName()", "",
                                               "(*vlt)", options=options)

        if not fname:
            return

        file_name = fname.split("/")[6].split(".")[0]
        g = fname.split("/")[:-1]
        # print("g", g)
        file_path = "/".join(g) + "/"
        # print("file_name", file_name)
        # print("file_path", file_path)
        s = ConvertVltToTXT(file_path, file_name, columns)
        s.load_vlt()
        s.conversion_vlt_to_txt()

    def evt_txt_clicked(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fname, _ = QFileDialog.getOpenFileName(self, "QFileDialog.getOpenFileName()", "",
                                               "(*TXT)", options=options)
        print("fname",fname)
        if not fname:
            return
        file_extension = fname.split(".")[-1:]
        print("file_extension", file_extension)
        file_name = fname.split("/")[6].split(".")[0]
        g = fname.split("/")[:-1]
        # print("g", g)
        file_path = "/".join(g) + "/"
        # print("file_name", file_name)
        # print("file_path", file_path)
        gg = ConvertTXTToExcel(file_path, file_name, columns_min, file_extension)
        gg.load_txt()
        gg.create_exls()



