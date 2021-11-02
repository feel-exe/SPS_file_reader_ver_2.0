from scripts.converter_txt_cl import *
from scripts.converter_vlt_cl import *
from gui.user_face import *


if __name__ == "__main__":
    app = QApplication(sys.argv)
    dlgMain = DlgMain()
    dlgMain.show()
    sys.exit(app.exec_())
