import os
import sys  # pyoiwin32
from PyQt5.QtCore import QCoreApplication
from win32com.client import constants, DispatchEx
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QLabel, QLineEdit, \
    QGridLayout, QMessageBox, QFileDialog, QProgressBar, QTextEdit


class UI(QWidget):
    def __init__(self):
        super().__init__()
        self.file_path = QLabel('文件夹路径')
        self.pdf_path = QLabel('PDF输出路径')
        self.progress_label = QLabel('生成进度')
        self.progress_tips = QLabel('(0/0)')
        self.log = QLabel('运行日志')
        self.file_pathEdit = QLineEdit(self, placeholderText="请选择或输入包含word文件所在文件夹")
        self.pdf_pathEdit = QLineEdit(self, placeholderText="请选择或输入输出pdf文件的文件夹")
        self.log_Text = QTextEdit(self)
        self.btn = QPushButton('生成', self, clicked=self.inputFilePath)
        self.qbtn = QPushButton('退出', self)
        self.selectFileBtn = QPushButton('选择文件夹', self, clicked=self.selectFileMenu)
        self.selectPdfBtn = QPushButton('选择文件夹', self, clicked=self.selectPdfMenu)
        self.timerProgress = QProgressBar(self, minimum=0, maximum=100, objectName="RedProgressBar")
        self.timerProgress.setValue(0)

        self.initUI()  # 界面绘制交给InitUi方法

    def initUI(self):
        # 设置窗口的位置和大小
        self.setGeometry(400, 350, 550, 300)
        # 设置窗口的标题
        self.setWindowTitle('PDF生成器')

        # 设置窗口的图标，引用当前目录下的web.png图片
        # self.setWindowIcon(QIcon('web.png'))

        grid = QGridLayout()
        grid.setSpacing(20)

        grid.addWidget(self.file_path, 1, 0)
        grid.addWidget(self.file_pathEdit, 1, 1, 1, 3)
        grid.addWidget(self.selectFileBtn, 1, 4)

        grid.addWidget(self.pdf_path, 2, 0)
        grid.addWidget(self.pdf_pathEdit, 2, 1, 1, 3)
        grid.addWidget(self.selectPdfBtn, 2, 4)

        self.btn.setToolTip('点击生成')
        self.btn.resize(self.btn.sizeHint())

        self.qbtn.clicked.connect(QCoreApplication.instance().quit)
        self.qbtn.resize(self.qbtn.sizeHint())

        # 进度条
        grid.addWidget(self.progress_label, 3, 0)
        grid.addWidget(self.timerProgress, 3, 1, 1, 3)
        grid.addWidget(self.progress_tips, 3, 4)

        # 日志
        grid.addWidget(self.log, 4, 0)
        grid.addWidget(self.log_Text, 4, 1, 1, 4)

        # 生成或退出按钮
        grid.addWidget(self.btn, 5, 1)
        grid.addWidget(self.qbtn, 5, 2)

        self.setLayout(grid)

        # 显示窗口
        self.show()

    #  生成按钮点击触发
    def inputFilePath(self):
        self.path = self.file_pathEdit.text().strip()
        self.pdfpath = self.pdf_pathEdit.text().strip()
        try:
            assert self.path != ''
            assert self.pdfpath != ''
            self.btn.setEnabled(False)
            self.qbtn.setEnabled(False)
            self.convert_word_to_pdf()
            s = QMessageBox()
            s.about(self, "提示", '生成成功！')
        except AssertionError:
            s = QMessageBox()
            s.warning(self, "提示", "输入和输出目录不可为空")
        except Exception as e:
            s = QMessageBox()
            s.warning(self, "提示", "生成失败！失败原因：" + str(e))
        self.btn.setEnabled(True)
        self.qbtn.setEnabled(True)

    def selectFileMenu(self):
        _dir = QFileDialog.getExistingDirectory(self, "选择文件夹", "/")
        self.file_pathEdit.setText(_dir)
        self.pdf_pathEdit.setText(_dir + r'/out')

    def selectPdfMenu(self):
        self.pdf_pathEdit.setText(QFileDialog.getExistingDirectory(self, "选择文件夹", "/"))


class WordToPDF(UI):
    def __init__(self):
        super().__init__()
        self.path = ''
        self.pdfpath = ''
        self.error_str = ''
        self.total_num = 0
        self.finish_num = 0

    def get_path(self):
        filename_list = os.listdir(self.path)
        # 生成迭word文件和输出pdf文件迭代器
        wordname_list = self.getwordlist(filename_list)
        for wordname in wordname_list:
            pdfname = os.path.splitext(wordname)[0] + '.pdf'
            if pdfname in filename_list:
                continue
            wordpath = os.path.join(self.path, wordname)
            pdfpathx = os.path.join(self.pdfpath, pdfname)
            yield wordpath, pdfpathx

    def getwordlist(self, filename_list):
        if not os.path.exists(self.pdfpath):
            os.makedirs(self.pdfpath)
        # 屏蔽已读占用
        wordname_list = [filename for filename in filename_list if
                         (filename.endswith((".doc", ".docx")) and not filename.startswith('~$'))]

        self.total_num = wordname_list.__len__()
        self.progress_tips.setText('(0/{num})'.format(num=self.total_num))
        QApplication.processEvents()
        return wordname_list

    def convert_word_to_pdf(self):
        for wordpath, pdfpathss in self.get_path():
            try:
                self.createPdf(wordpath, pdfpathss)
                self.finish_num += 1
                self.timerProgress.setValue(int((self.finish_num / self.total_num) * 100))
                self.progress_tips.setText('({now}/{num})'.format(num=self.total_num, now=self.finish_num))
                QApplication.processEvents()
                self.logger('Success:' + pdfpathss)
            except Exception as e:
                self.logger('Error:' + wordpath + '生成失败，' + str(e))

    @classmethod
    def createPdf(cls, wordPath, pdfPath):
        """
        创建主方法 word转pdf
        :param wordPath: word文件路径
        :param pdfPath:  生成pdf文件路径
        """
        word = DispatchEx('Word.Application')
        doc = word.Documents.Open(wordPath, ReadOnly=1)
        # 开启工作空间
        doc.ExportAsFixedFormat(pdfPath,
                                constants.wdExportFormatPDF,
                                Item=constants.wdExportDocumentWithMarkup,
                                CreateBookmarks=constants.wdExportCreateHeadingBookmarks)
        word.Quit(constants.wdDoNotSaveChanges)

    def logger(self, msg):
        self.error_str += str(msg) + '\n'
        self.log_Text.setText(self.error_str)
        QApplication.processEvents()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = WordToPDF()
    sys.exit(app.exec_())
