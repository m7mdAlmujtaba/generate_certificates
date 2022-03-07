from PyQt5.QtGui import*
from PyQt5.QtWidgets import QDateEdit,QHBoxLayout, QListView, QTableWidgetItem, QMainWindow, QApplication, QWidget,QTableWidget, QVBoxLayout, QAction, QFileDialog, QLabel, QPushButton, QComboBox, QAbstractItemView
from PyQt5.QtCore import*
import sys
import xlrd
from image2pdf import*


class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        self.setGeometry(30, 30, 1320, 720)
        self.setWindowTitle("GCM")
        self.setWindowIcon(QIcon('img/icon.png'))

        self.courses_file = xlrd.open_workbook('excel/courses.xlsx')
        sheet = self.courses_file.sheet_by_index(0)
        self.courses = sheet.col_values(1)

        self.instructors_file = xlrd.open_workbook('excel/instructors.xlsx')
        isheet = self.instructors_file.sheet_by_index(0)
        self.instructors = isheet.col_values(0)

        self.index = []
        self.serials = []

        style = """QMainWindow{background-color:gray}
        QMainMenu{
        background-color:gary;
        }
        """

        wid_style = """QWidget{background-color:#efefef;
                        margin:0;    }

                        QPushButton{
                        background-color:#efefef;
                        color:#2a2e3b;
                        font-size:20px;
                        min-width:180px;
                        min-height:55px;
                        max-width:280px;
                        max-height:60px;
                        border-width:1px;
                        border-color:#39455b;
                        border-style:solid;}

                        QPushButton:hover{
                        background-color:#39455b;
                        color:white;
                        }
                        """

        wid3_style = """QWidget{background-color:#2a2e3b;}
                                QComboBox{
                                background-color:#494f62;
                                color:white;
                                font-size:13px;
                                min-width:120px;
                                min-height:30px;
                                max-width:150px;
                                border-radius:4px;
                                border-width:1px;
                                border-color:white;
                                border-style:solid;
                                }
                                QDateEdit{
                                background-color:#494f62;
                                color:white;
                                font-size:13px;
                                min-width:120px;
                                min-height:30px;
                                border-radius:4px;
                                border-width:1px;
                                border-color:white;
                                border-style:solid;
                                }
                                        """

        wid4_style = """QWidget{background-color:#2a2e3b;}
                        QTableWidget{
                        color:white;
                        font-size:13px;
                        }
                                                """

        wid5_style = """QWidget{background-color:#2a2e3b;}
                                QListView{
                                color:white;
                                font-size:15px;
                                min-width:120px;
                                min-height:30px;
                                }
                        QPushButton{
                        background-color:#494f62;
                        color:white;
                        font-size:14px;
                        border-width:1px;
                        border-color:white;
                        border-style:solid;}

                        QPushButton:hover{
                        background-color:#39455b;
                        color:white;
                        }
                                        """

        self.setStyleSheet(style)

        self.wid = QWidget(self)
        self.wid.setGeometry(0, 0, 300, 720)
        self.wid.setStyleSheet(wid_style)

        self.wid_2 = QWidget(self)
        self.wid_2.setGeometry(300, 0, 1120, 720)
        self.wid_2.setStyleSheet('background-color:#1d212d')

        self.wid_0 = QWidget(self)
        self.wid_0.setGeometry(300, 0, 1320, 100)
        self.wid_0.setStyleSheet('background-color:#2a2e3b;color:white;')

        self.wid_3 = QWidget(self)
        self.wid_3.setGeometry(380, 180, 850, 120)
        self.wid_3.setStyleSheet(wid3_style)

        self.wid_4 = QWidget(self)
        self.wid_4.setGeometry(380, 310, 420, 350)
        self.wid_4.setStyleSheet(wid4_style)

        self.wid_5 = QWidget(self)
        self.wid_5.setGeometry(810, 310, 420, 350)
        self.wid_5.setStyleSheet(wid5_style)

        self.widget1_ui()
        self.widget3_ui()
        self.widget4_ui()
        self.widget5_ui()

        self.st_btn.clicked.connect(self.ns_upload)
        self.co_btn.clicked.connect(self.nc_upload)
        self.in_btn.clicked.connect(self.ni_upload)

        self.genserial_btn.clicked.connect(self.serial_number)
        self.gen_btn.clicked.connect(self.gen_certificate)

        self.logo_label = QLabel(self.wid_0)
        self.logo_label.setGeometry(400,40,280,40)
        self.myPixmap = QPixmap('img/logo.png')
        self.myScaledPixmap = self.myPixmap.scaled(self.logo_label.size(), Qt.KeepAspectRatio)
        self.logo_label.setPixmap(self.myScaledPixmap)

        self.ui()
        self.home()

    def widget1_ui(self):
        # LAYOUT and BUTTONS
        self.v_lay = QVBoxLayout(self.wid)
        self.v_lay.addStretch()
        self.v_lay.addStretch()
        self.st_btn = QPushButton('STUDENTS', self)
        self.v_lay.addWidget(self.st_btn)
        self.co_btn = QPushButton('COURSES', self)
        self.v_lay.addWidget(self.co_btn)
        self.in_btn = QPushButton('INSTRUCTORS', self)
        self.v_lay.addWidget(self.in_btn)
        self.v_lay.addStretch()
        self.v_lay.setSpacing(18)
        self.v_lay.addStretch()
        self.v_lay.addStretch()
        self.v_lay.addStretch()
        self.v_lay.addStretch()

    def widget3_ui(self):
        # Create combobox and add items.
        self.s_comboBox = QComboBox()
        # self.comboBox.setGeometry(QRect(40, 40, 491, 31))
        self.s_comboBox.setObjectName("Students comboBox")
        self.s_comboBox.addItem("All")
        # self.s_comboBox.addItem("Qt")
        self.s_comboBox.setEnabled(False)
        self.s_comboBox.currentIndexChanged.connect(self.all)

        self.c_comboBox = QComboBox()
        self.c_comboBox.setObjectName(("comboBox"))
        for course in self.courses:
            self.c_comboBox.addItem(course)

        self.i_comboBox = QComboBox()
        self.i_comboBox.setObjectName(("comboBox"))
        for instructor in self.instructors:
            self.i_comboBox.addItem(instructor)

        # nice widget for editing the date
        self.dateEdit = QDateEdit()
        self.dateEdit.setCalendarPopup(True)
        self.dateEdit.setDate(QDate.currentDate())

        self.h_lay_3 = QHBoxLayout(self.wid_3)
        self.h_lay_3.addStretch()
        self.h_lay_3.addWidget(self.s_comboBox)
        self.h_lay_3.addStretch()
        self.h_lay_3.addWidget(self.c_comboBox)
        self.h_lay_3.addStretch()
        self.h_lay_3.addWidget(self.i_comboBox)
        self.h_lay_3.addStretch()
        self.h_lay_3.addWidget(self.dateEdit)
        self.h_lay_3.addStretch()

    def widget4_ui(self):
        self.tableWidget = QTableWidget(self.wid_4)
        self.tableWidget.setGeometry(10, 10, 400, 340)
        self.tableWidget.setColumnCount(2)
        self.tableWidget.setRowCount(11)
        self.tableWidget.verticalHeader().setVisible(False)
        self.tableWidget.horizontalHeader().setVisible(False)
        self.tableWidget.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tableWidget.setColumnWidth(0, 295)

    def widget5_ui(self):
        self.listView = QListView(self.wid_5)
        self.listView.setGeometry(10, 10, 195, 340)

        self.genserial_btn = QPushButton('Serial Numbers', self.wid_5)
        self.genserial_btn.setGeometry(255, 40, 120, 30)

        self.gen_btn = QPushButton('Certificate', self.wid_5)
        self.gen_btn.setGeometry(255, 80, 120, 30)

    def ui(self):
        openFile = QAction('&Open', self)
        openFile.setShortcut("Ctrl+O")
        openFile.setStatusTip('Open CSV File')
        # openFile.triggered.connect(self.open)

        # Save
        saveAction = QAction('&Save', self)
        saveAction.setShortcut("Ctrl+S")
        saveAction.setStatusTip('Save The New File')
        # saveAction.triggered.connect(self.close)

        # quit
        eAction = QAction('&Quit', self)
        eAction.setShortcut("Ctrl+Q")
        eAction.setStatusTip('Leave The App')
        # eAction.triggered.connect(self.close)

        self.statusBar()

        mainMenu = self.menuBar()
        fileMenu = mainMenu.addMenu('&File')
        fileMenu.addAction(openFile)
        fileMenu.addAction(saveAction)
        fileMenu.addAction(eAction)

        viewMenu = mainMenu.addMenu('&View')
        toolsMenu = mainMenu.addMenu('&Tools')
        helpMenu = mainMenu.addMenu('&Help')

    def home(self):
        self.showMaximized()
        # self.show()

    # Upload Buttons
    def ns_upload(self):
        name = QFileDialog.getOpenFileName(self, 'Open File')
        print(str(name[0]))
        #wb = xlrd.open_workbook('excel/stds.xlsx')
        wb = xlrd.open_workbook(str(name[0]))
        print(2)
        sheet = wb.sheet_by_index(0)
        print(sheet.col_values(0))
        print(sheet.col_values(1))
        self.students = sheet.col_values(1)
        self.students.sort()
        for i in range(sheet.nrows):
            self.tableWidget.setItem(i, 0, QTableWidgetItem(self.students[i]))
            self.index.append('%02d' % (i+1))
        # a = 5
        # print '%02d' % a
        # output: 05

    def nc_upload(self):
        cource = QFileDialog.getOpenFileName(self, 'Open File')
        wb = xlrd.open_workbook(str(cource))
        sheet = wb.sheet_by_index(0)
        print(sheet.col_values(0))
        print(sheet.col_values(1))
        self.c_comboBox.clear()
        self.courses = sheet.col_values(1)
        for course in self.courses:
            self.c_comboBox.addItem(course)
        self.c_comboBox.update()

    def ni_upload(self):
        instructor = QFileDialog.getOpenFileName(self, 'Open File')
        wb = xlrd.open_workbook(str(instructor))
        sheet = wb.sheet_by_index(0)
        print(sheet.col_values(0))
        print(sheet.col_values(1))
        self.i_comboBox.clear()
        self.instructors = sheet.col_values(1)
        for instructor in self.instructors:
            self.i_comboBox.addItem(instructor)
        self.i_comboBox.update()

    #
    def all(self):
        pass

    # Generate
    def serial_number(self):
        cn = self.c_comboBox.findText(self.c_comboBox.currentText())
        print ('%02d' % (cn+1))
        temp_var = self.dateEdit.date()
        var_name = temp_var.toPyDate()
        print(str(var_name), str(var_name)[5:7], str(var_name)[2:4])
        # print len(self.students)
        for i in range(len(self.students)):
            serial = ('%02d' % (cn+1)) + str(var_name)[5:7] + str(var_name)[2:4] + ('%02d' % (i+1))
            self.serials.append(serial)
            self.tableWidget.setItem(i, 1, QTableWidgetItem(serial))
        print(self.serials)
        print(self.students)

    def gen_certificate(self):
        icon = QIcon('icon/pdf_5.png')
        directory = QFileDialog.getExistingDirectory(self, 'Select directory')
        print(directory)
        model = QStandardItemModel(self.listView)
        for i in range(len(self.students)):
            im2pdf((self.students[i]).upper(), str(self.c_comboBox.currentText()), 'For 21 hours from 22 Feb - 27 Feb / 2020', ('SN:'+self.serials[i]),
                   str(self.i_comboBox.currentText()))

            # create an item with a caption
            item = QStandardItem(self.serials[i]+'.pdf')
            # Add the item to the model
            item.setIcon(icon)
            model.appendRow(item)

        # Apply the model to the list view
        self.listView.setModel(model)


def run():
    app = QApplication(sys.argv)
    gui = MainWindow()
    sys.exit(app.exec_())

run()