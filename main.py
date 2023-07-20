from PyQt5.QtWidgets import (
    QDialog,
    QLabel,
    QVBoxLayout,
    QMainWindow,
    QWidget,
    QScrollBar,
    QAction,
    QActionGroup,
    QToolBar,
    QTextEdit,
    QLineEdit,
    QPushButton,
    QGridLayout,
    QStatusBar,
    QMessageBox,
    QFileDialog,
    QApplication,
)
from PyQt5.QtCore import QThread, pyqtSignal, QCoreApplication, QSize, Qt
from PyQt5.QtGui import QMovie
from sys import exit, argv
import requests
from bs4 import BeautifulSoup
from os.path import splitext as sp
from os.path import abspath, dirname, sep
from pandas import DataFrame
import openpyxl


# QThread爬取中文释义，防止主界面卡死
class LongTask(QThread):
    # 定义一个信号，用于通知主界面任务完成
    finished = pyqtSignal()
    # 定义回传变量
    output = pyqtSignal(list)

    def __init__(self, data, plat):
        super().__init__()
        self.data = data
        self.plat = plat

    def run(self):
        # 爬取bing词典
        def bing(inpu):
            result = []
            r = requests.get(
                f"https://cn.bing.com/dict/search?q={inpu}"
            ).content.decode()
            soup = BeautifulSoup(r, "html.parser")
            tags = soup.find(
                name=None, attrs={"class": "qdef"}
            )  # 查找文本当中class为title的<p>元素
            text = tags.find("ul")
            speeches = text.find_all(name=None, attrs={"class": "pos"})
            chins = text.find_all(name=None, attrs={"class": "b_regtxt"})
            for i in range(len(speeches) - 1):
                speech = speeches[i].text
                chin = chins[i].text
                result.append(speech + chin)
            result = "---".join(result)
            return result

        # 爬取dict.cn
        def dictcn(inpu):
            result = []
            r = requests.get(f"http://dict.cn/{inpu}").content.decode()
            soup = BeautifulSoup(r, "html.parser")
            tags = soup.find(
                name=None, attrs={"class": "layout dual"}
            )  # 查找文本当中class为title的<p>元素
            text = tags.find_all("li")
            for i in text:
                result.append(i.find("strong").text)
            result = "---".join(result)
            return result

        # 爬取有道词典
        def youdao(inpu):
            result = []

            r = requests.get(
                f"https://dict.youdao.com/result?word={inpu}&lang=en"
            ).content.decode()
            soup = BeautifulSoup(r, "html.parser")
            tags = soup.find(
                name=None, attrs={"class": "dict-book"}
            )  # 查找文本当中class为title的<p>元素
            chins = tags.find_all(name=None, attrs={"class": "trans"})
            speeches = tags.find_all(name=None, attrs={"class": "pos"})
            for i in range(len(speeches)):
                speech = speeches[i].text
                chin = chins[i].text
                result.append(speech + chin)
            result = "---".join(result)
            return result

        def test(a, i):
            if a == 0:
                result = bing(i)
            if a == 1:
                result = youdao(i)
            if a == 2:
                result = dictcn(i)
            return result

        res_list = []
        for i in self.data:
            lists = [0, 1, 2]
            if i != "":
                try:
                    dig = test(self.plat, i)
                    if dig == "":
                        raise
                    else:
                        res_list.append(dig)
                except:
                    lists.remove(self.plat)
                    try:
                        dig = test(lists[0], i)
                        if dig == "":
                            raise
                        else:
                            res_list.append(dig)
                    except:
                        try:
                            dig = test(lists[1], i)
                            if dig == "":
                                raise
                            else:
                                res_list.append(dig)
                        except:
                            res_list.append("")
        self.output.emit(res_list)
        self.finished.emit()


class waiting(QDialog):  # 显示等待窗口
    def __init__(self):
        super(waiting, self).__init__()
        self.setWindowTitle("Waiting...")
        self.resize(100, 150)
        self.label = QLabel()
        # 创建一个QMovie对象，加载动画文件
        self.movie = QMovie()
        current_path = abspath(__file__)
        father_name = abspath(dirname(current_path) + sep + ".")
        ft_name = father_name.replace("\\", r"\\")
        self.movie.setFileName(f"{ft_name}\\image\\loading.gif")
        # 设置标签的动画为QMovie对象
        self.label.setMovie(self.movie)
        # 创建一个垂直布局，添加按钮和标签
        self.layout = QVBoxLayout()
        self.layout.addWidget(self.label)
        # 设置窗口的布局为垂直布局
        self.setLayout(self.layout)
        self.movie.start()


class MainWindow(QMainWindow):  # 主界面
    def __init__(self, parent=None):
        super(QWidget, self).__init__(parent)

        self.initUI()

    def initUI(self):  # 界面初始化
        self.setFixedSize(800, 600)
        self.setWindowTitle("单词本")

        # 滑动条qss样式与初始化
        qss_H = """
            QScrollBar:horizontal {
                border-width: 0px;
                border: none;
                background:rgba(0,0,0, 0);
                height:12px;
                margin: 0px 0px 0px 0px;
            }
            QScrollBar::handle:horizontal {
                background: rgba(139,139,139,0.8);
                min-height: 20px;
                max-height: 20px;
                margin: 0 0px 0 0px;
                border-radius: 6px;
            }
            QScrollBar::handle:horizontal:pressed {
                background: rgb(96,96,96);
            }
  
            """
        qss_V = """
            QScrollBar:vertical {
                border-width: 0px;
                border: none;
                background:rgba(0,0,0, 0);
                width:12px;
                margin: 0px 0px 0px 0px;
            }
            QScrollBar::handle:vertical  {
                background: rgba(139,139,139,0.8);
                min-height: 20px;
                max-height: 20px;
                margin: 0 0px 0 0px;
                border-radius: 6px;
            }
            QScrollBar::handle:vertical:pressed {
                background: rgb(96,96,96);
            }
            """

        self.textEdit_V = QScrollBar()
        self.textEdit_V.setStyleSheet(qss_V)
        self.textEdit_H = QScrollBar()
        self.textEdit_H.setStyleSheet(qss_H)

        # 菜单栏设置
        # --->文件
        bar = self.menuBar()
        file = bar.addMenu("文件")

        select = QAction("导入", self)
        select.setShortcut("Ctrl+O")
        file.addAction(select)
        select.triggered.connect(self.select)

        out = QAction("导出", self)
        out.setShortcut("Ctrl+M")
        file.addAction(out)
        out.triggered.connect(self.out)

        broken = QAction("退出", self)
        file.addAction(broken)
        broken.triggered.connect(QCoreApplication.instance().quit)

        # --->平台
        way = bar.addMenu("平台")
        self.action_group = QActionGroup(self)
        self.action_group.setExclusive(True)

        dictcn = QAction("dict.cn(更快速)", self, checkable=True)
        way.addAction(dictcn)
        dictcn.setData(2)
        dictcn.setChecked(True)
        self.action_group.addAction(dictcn)

        youdao = QAction("有道(更准确)", self, checkable=True)
        way.addAction(youdao)
        youdao.setData(1)
        self.action_group.addAction(youdao)

        bing = QAction("必应", self, checkable=True)
        way.addAction(bing)
        bing.setData(0)
        self.action_group.addAction(bing)
        self.action_group.triggered.connect(self.setstatusBar)

        self.main_widget = QWidget()
        # 创建主部件的网格布局

        # 设置窗口主部件布局为网格布局
        # 侧边栏设置
        self.toolBar = QToolBar()
        self.toolBar.setFixedWidth(250)

        self.space3 = QWidget()
        self.space3.setFixedSize(QSize(10, 15))
        self.toolBar.addWidget(self.space3)

        self.english = QLabel("英文单词")
        self.english.setAlignment(Qt.AlignCenter | Qt.AlignBottom)
        self.toolBar.addWidget(self.english)
        self.toolBar.addSeparator()

        self.spacebar2 = QToolBar()
        self.toolBar.addWidget(self.spacebar2)
        self.space2 = QWidget()
        self.space2.setFixedSize(QSize(10, 40))
        self.spacebar2.addWidget(self.space2)

        self.lolabel = QTextEdit()
        self.lolabel.setFixedSize(QSize(200, 330))

        self.spacebar2.addWidget(self.lolabel)

        self.toolBar.addSeparator()

        self.spacebar = QToolBar()
        self.toolBar.addWidget(self.spacebar)
        self.space = QWidget()
        self.space.setFixedSize(QSize(10, 40))
        self.spacebar.addWidget(self.space)

        self.label = QLineEdit()
        self.spacebar.addWidget(self.label)
        self.label.returnPressed.connect(self.changed)

        self.change = QPushButton("输入")
        self.change.setFixedSize(QSize(60, 40))
        self.change.clicked.connect(self.changed)
        self.spacebar.addWidget(self.change)

        self.addToolBar(Qt.LeftToolBarArea, self.toolBar)
        self.toolBar.setStyleSheet("QToolBar {spacing: 8px;alignment:center;}")

        # 创建窗口主部件

        self.centralWidget = QWidget()
        self.setCentralWidget(self.centralWidget)
        self.gridLayout = QGridLayout()
        self.centralWidget.setLayout(self.gridLayout)  # 设置格栅布局为中央部件的布局

        self.chin = QLabel("中文释义")
        self.chin.setAlignment(Qt.AlignCenter | Qt.AlignBottom)
        self.gridLayout.addWidget(self.chin, 0, 0, 1, 7)

        self.prlabel = QTextEdit()
        self.prlabel.setFixedSize(QSize(520, 330))
        self.prlabel.setLineWrapMode(QTextEdit.NoWrap)
        self.gridLayout.addWidget(self.prlabel, 2, 0, 6, 6)  # 竖2横0

        self.single_query = QPushButton("清空")
        self.single_query.clicked.connect(self.delete)
        self.gridLayout.addWidget(self.single_query, 7, 1, 3, 1)

        self.single_query = QPushButton("转换")
        self.single_query.clicked.connect(self.translate)
        self.gridLayout.addWidget(self.single_query, 7, 3, 3, 2)

        # 状态栏初始化
        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)
        self.statusBar.setStyleSheet("background-color:rgba(210,210,210,1)")
        self.statusBar.showMessage("目前依赖平台：dictcn", 0)

        self.prlabel.setVerticalScrollBar(self.textEdit_V)
        self.prlabel.setHorizontalScrollBar(self.textEdit_H)
        # self.lolabel.setVerticalScrollBar(self.textEdit_V)

    def setstatusBar(self, action):  # 状态栏响应平台改变事件
        state = action.data()
        if state == 0:
            self.statusBar.showMessage("目前依赖平台：bing", 0)
        if state == 1:
            self.statusBar.showMessage("目前依赖平台：youdao", 0)
        if state == 2:
            self.statusBar.showMessage("目前依赖平台：dict.cn", 0)

    def delete(self):  # 清空
        reply = QMessageBox.warning(
            self, "提示", "确定删除全部？", QMessageBox.No | QMessageBox.Yes, QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            self.lolabel.clear()
            self.prlabel.clear()

    def changed(self):  # 输入
        a = self.label.text()
        self.label.clear()
        if a != "":
            self.lolabel.insertPlainText(a + "\n")
            self.lolabel.moveCursor(self.lolabel.textCursor().End)

    def translate(self):  # 接收LongTask传入
        action = self.action_group.checkedAction()
        plat = action.data()
        reply = QMessageBox.question(
            self, "提示", "确定开始翻译？", QMessageBox.No | QMessageBox.Yes, QMessageBox.Yes
        )
        if reply == QMessageBox.Yes:
            a = self.lolabel.toPlainText()
            list2 = a.split("\n")
            list2 = [x for x in list2 if x != ""]
            self.lolabel.setPlainText("\n".join(list2) + "\n")
            self.prlabel.clear()

            self.LongTask = LongTask(list2, plat)
            self.waiting = waiting()
            self.waiting.show()
            self.LongTask.start()
            self.single_query.setEnabled(False)

            self.LongTask.output.connect(self.outPUT)
            self.LongTask.finished.connect(self.task_finished)

    def outPUT(self, output):  # 显示翻译
        self.prlabel.insertPlainText("\n".join(output))
        QApplication.processEvents()

    def task_finished(self):  # 停止播放动画
        self.waiting.setVisible(False)
        self.single_query.setEnabled(True)

    def select(self):  # 选取导入文件
        directory = QFileDialog.getOpenFileName(self, "选取文件", "./", "文本文件 (*.txt)")
        types = sp(directory[0])[-1]
        if types == ".txt":
            f = open(directory[0], encoding="utf8")
            txt = []
            for line in f:
                txt.append(line.strip())
            self.lolabel.setPlainText("\n".join(txt) + "\n")
        elif types == "":
            pass
        else:
            reply = QMessageBox.about(self, "提示", "文件路径无效！")

    def out(self):  # 选择导出文件目录
        directory = QFileDialog.getExistingDirectory(self, "选择导出目录(保存文件名为output.xlsx)")
        lol = self.lolabel.toPlainText()
        prout = self.prlabel.toPlainText()
        if directory == "":
            pass
        elif lol == "" or prout == "":
            reply = QMessageBox.warning(self, "提示", "请检查输入输出")

        else:
            try:
                list2 = lol.split("\n")
                list3 = prout.split("\n")
                list4 = []
                for i in list3:
                    k = i.replace("---", "\n")
                    list4.append(k)
                list4.append("")
                data = DataFrame({"英文单词": list2, "中文释义": list4})
                data.to_excel(f"{directory}/output.xlsx", index=False)
                fil = openpyxl.load_workbook(f"{directory}/output.xlsx")
                sheet = fil.active
                for a in sheet:
                    for c in a:
                        c.alignment = openpyxl.styles.Alignment(wrapText=True)
                sheet.column_dimensions["B"].width = 100
                sheet.column_dimensions["A"].width = 30

                fil.save(f"{directory}/output.xlsx")
                reply = QMessageBox.information(self, "提示", "导出成功")
            except PermissionError:
                reply = QMessageBox.warning(self, "提示", "操作无法完成，因为文件在其他程序中打开")


if __name__ == "__main__":
    app = QApplication(argv)
    win = MainWindow()
    win.show()
    exit(app.exec_())
