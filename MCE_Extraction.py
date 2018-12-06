import sys
from PySide2.QtWidgets import QWidget,QFileDialog,QMessageBox,QApplication
from extraction_ui import Ui_Form
import ColorFidelity_extraction
import Flash_extraction
import Texture_extraction
import Greychart_extraction

class test(QWidget):
    def __init__(self,parent=None):#此处的parent是干嘛的？
        QWidget.__init__(self,parent)
        self.ui=Ui_Form()
        self.ui.setupUi(self)

        self.path=''
        self.ui.dir_line.setText(self.path)

        self.ui.Choose_btn.clicked.connect(self.opendirectory)
        self.ui.Texture_btn.clicked.connect(self.Texture_output)
        self.ui.GreyChart_btn.clicked.connect(self.Grey_output)
        self.ui.Flash_btn.clicked.connect(self.Flash_output)
        self.ui.ColorChecker_btn.clicked.connect(self.Color_output)

        self.ui.Choose_btn.setToolTip('待处理的文件夹，选就完了')
        self.ui.ColorChecker_btn.setToolTip('输出按照Color Fidelity排序的表格')
        self.ui.Flash_btn.setToolTip('同时读取flsh和vc2然后输出表格')
        self.ui.GreyChart_btn.setToolTip('肯定是输出提取GreyChart的光源表格了...')
        self.ui.Texture_btn.setToolTip('输出提取Texture的光源表格')



    def opendirectory(self):
        self.path = QFileDialog.getExistingDirectory(self)+'/'
        self.ui.dir_line.setText(self.path)

    def Texture_output(self):
        self.ui.Texture_btn.setEnabled(False)
        self.ui.ColorChecker_btn.setEnabled(False)
        self.ui.Flash_btn.setEnabled(False)
        self.ui.GreyChart_btn.setEnabled(False)
        self.ui.Choose_btn.setEnabled(False)

        a,b=Texture_extraction.batch_process(self.path)
        Texture_extraction.write_excel(a,b)

        self.ui.Choose_btn.setEnabled(True)
        self.ui.Texture_btn.setEnabled(True)
        self.ui.ColorChecker_btn.setEnabled(True)
        self.ui.Flash_btn.setEnabled(True)
        self.ui.GreyChart_btn.setEnabled(True)
        QMessageBox.question(self, 'Message', "Finished", QMessageBox.Ok, QMessageBox.Ok)


    def Color_output(self):
        self.ui.Texture_btn.setEnabled(False)
        self.ui.ColorChecker_btn.setEnabled(False)
        self.ui.Flash_btn.setEnabled(False)
        self.ui.GreyChart_btn.setEnabled(False)
        self.ui.Choose_btn.setEnabled(False)
        a, b = ColorFidelity_extraction.batch_process(self.path)
        Texture_extraction.write_excel(a, b)
        self.ui.Choose_btn.setEnabled(True)
        self.ui.Texture_btn.setEnabled(True)
        self.ui.ColorChecker_btn.setEnabled(True)
        self.ui.Flash_btn.setEnabled(True)
        self.ui.GreyChart_btn.setEnabled(True)
        QMessageBox.question(self, 'Message', "Finished", QMessageBox.Ok, QMessageBox.Ok)
    def Grey_output(self):
        self.ui.Texture_btn.setEnabled(False)
        self.ui.ColorChecker_btn.setEnabled(False)
        self.ui.Flash_btn.setEnabled(False)
        self.ui.GreyChart_btn.setEnabled(False)
        self.ui.Choose_btn.setEnabled(False)
        a, b = Greychart_extraction.batch_process(self.path)
        Greychart_extraction.write_excel(a,b)
        self.ui.Choose_btn.setEnabled(True)
        self.ui.Texture_btn.setEnabled(True)
        self.ui.ColorChecker_btn.setEnabled(True)
        self.ui.Flash_btn.setEnabled(True)
        self.ui.GreyChart_btn.setEnabled(True)
        QMessageBox.question(self, 'Message', "Finished", QMessageBox.Ok, QMessageBox.Ok)
    def Flash_output(self):
        self.ui.Texture_btn.setEnabled(False)
        self.ui.ColorChecker_btn.setEnabled(False)
        self.ui.Flash_btn.setEnabled(False)
        self.ui.GreyChart_btn.setEnabled(False)
        self.ui.Choose_btn.setEnabled(False)
        a, b, c = Flash_extraction.stimul_read(self.path)
        Flash_extraction.write_excel(a, b, c)
        self.ui.Choose_btn.setEnabled(True)
        self.ui.Texture_btn.setEnabled(True)
        self.ui.ColorChecker_btn.setEnabled(True)
        self.ui.Flash_btn.setEnabled(True)
        self.ui.GreyChart_btn.setEnabled(True)
        QMessageBox.question(self, 'Message', "Finished", QMessageBox.Ok, QMessageBox.Ok)

if __name__ == "__main__":
    # 所有应用必须创建一个应用（Application）对象
    app = QApplication(sys.argv)
    form = test()
    form.show()
    sys.exit(app.exec_())