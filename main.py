from PyQt5.QtGui import*
from PyQt5.QtWidgets import QDesktopWidget, QDateEdit,QHBoxLayout, QListView, QTableWidgetItem, QMainWindow, QApplication, QWidget,QTableWidget, QVBoxLayout, QAction, QFileDialog, QLabel, QPushButton, QComboBox, QAbstractItemView
from PyQt5.QtCore import*
import sys
import xlrd
from image2pdf import*


class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()

        geometry = QDesktopWidget().screenGeometry()
        self.setGeometry(geometry)
        screen_width = geometry.width()
        screen_height = geometry.height()
        self.setWindowTitle("GCM")
        self.setWindowIcon(QIcon('img/icon.png'))

        # Open options file and get courses, instructors and directors sheets
        self.options_file = xlrd.open_workbook('files/options.xlsx')

        courses_sheet = self.options_file.sheet_by_name('courses')
        self.courses = courses_sheet.col_values(1)

        instructors_sheet = self.options_file.sheet_by_name('instructors')
        self.instructors = instructors_sheet.col_values(0)

        director_sheet = self.options_file.sheet_by_name('directors')
        self.directors = director_sheet.col_values(0)

        #self.index = []
        self.serials = []

        style = """QMainWindow{background-color:#1d212d}
        QMainMenu{
        background-color:gary;
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
                                QLabel{
                                color:white;
                                font-size:13px;
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

        self.header_widget = QWidget(self)
        self.header_widget.setGeometry(0, 0, screen_width,  int(screen_height*.14))
        self.header_widget.setStyleSheet('background-color:#2a2e3b;color:white;')

        self.options_widget = QWidget(self)
        self.options_widget.setGeometry(int(screen_width*.1),  int(screen_height*.18), int(screen_width*.8), int(screen_height*.16))
        self.options_widget.setStyleSheet(wid3_style)

        self.students_widget = QWidget(self)
        self.students_widget.setGeometry(int(screen_width*.1), int(screen_height*.35), int(screen_width*.39), int(screen_height*.5))
        self.students_widget.setStyleSheet(wid4_style)

        self.control_widget = QWidget(self)
        self.control_widget.setGeometry(int(screen_width*.51), int(screen_height*.35), int(screen_width*.39), int(screen_height*.5))
        self.control_widget.setStyleSheet(wid5_style)

        self.options_items()
        self.students_items()
        self.control_items()

        self.st_btn.clicked.connect(self.ns_upload)
        
        self.genserial_btn.clicked.connect(self.serial_number)
        self.gen_btn.clicked.connect(self.gen_certificate)

        self.logo_label = QLabel(self.header_widget)
        self.logo_label.setGeometry(400,40,280,40)
        self.myPixmap = QPixmap('img/logo.png')
        self.myScaledPixmap = self.myPixmap.scaled(self.logo_label.size(), Qt.KeepAspectRatio)
        self.logo_label.setPixmap(self.myScaledPixmap)

        self.ui()
        self.home()

    def options_items(self):
        # Create combobox and add items.
        self.d_lbl = QLabel('Director')
        self.d_comboBox = QComboBox()
        self.d_comboBox.setObjectName("director_comboBox")
        for director in self.directors:
            self.d_comboBox.addItem(director)

        self.c_lbl = QLabel('Course')
        self.c_comboBox = QComboBox()
        self.c_comboBox.setObjectName(("course_comboBox"))
        for course in self.courses:
            self.c_comboBox.addItem(course)

        self.i_lbl = QLabel('Instructor')
        self.i_comboBox = QComboBox()
        self.i_comboBox.setObjectName(("instructor_comboBox"))
        for instructor in self.instructors:
            self.i_comboBox.addItem(instructor)

        # From date
        self.from_lbl = QLabel('From')
        self.from_date = QDateEdit()
        self.from_date.setCalendarPopup(True)
        self.from_date.setDate(QDate.currentDate())

        # From date
        self.to_lbl = QLabel('To')
        self.to_date = QDateEdit()
        self.to_date.setCalendarPopup(True)
        self.to_date.setDate(QDate.currentDate())

        self.h_lay_3 = QHBoxLayout(self.options_widget)
        self.h_lay_3.addStretch()
        self.h_lay_3.addWidget(self.d_lbl)
        self.h_lay_3.addWidget(self.d_comboBox)
        self.h_lay_3.addStretch()
        self.h_lay_3.addWidget(self.c_lbl)
        self.h_lay_3.addWidget(self.c_comboBox)
        self.h_lay_3.addStretch()
        self.h_lay_3.addWidget(self.i_lbl)
        self.h_lay_3.addWidget(self.i_comboBox)
        self.h_lay_3.addStretch()
        self.h_lay_3.addWidget(self.from_lbl)
        self.h_lay_3.addWidget(self.from_date)
        self.h_lay_3.addStretch()
        self.h_lay_3.addWidget(self.to_lbl)
        self.h_lay_3.addWidget(self.to_date)

    def students_items(self):
        self.tableWidget = QTableWidget(self.students_widget)
        self.tableWidget.setGeometry(10, 10, 400, 340)
        self.tableWidget.setColumnCount(2)
        self.tableWidget.setRowCount(11)
        self.tableWidget.verticalHeader().setVisible(False)
        self.tableWidget.horizontalHeader().setVisible(False)
        self.tableWidget.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tableWidget.setColumnWidth(0, 295)

    def control_items(self):
        self.listView = QListView(self.control_widget)
        self.listView.setGeometry(10, 10, 195, 340)
        
        self.st_btn = QPushButton('Upload Students', self.control_widget)
        self.st_btn.setGeometry(255, 40, 120, 30)

        self.genserial_btn = QPushButton('Serial Numbers', self.control_widget)
        self.genserial_btn.setGeometry(255, 80, 120, 30)

        self.gen_btn = QPushButton('Certificate', self.control_widget)
        self.gen_btn.setGeometry(255, 120, 120, 30)

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
        #print(str(name[0]))
        #wb = xlrd.open_workbook('excel/stds.xlsx')
        wb = xlrd.open_workbook(str(name[0]))
        sheet = wb.sheet_by_index(0)
        #print(sheet.col_values(0))
        #print(sheet.col_values(1))
        self.students = sheet.col_values(1)
        self.students.sort()
        for i in range(sheet.nrows):
            self.tableWidget.setItem(i, 0, QTableWidgetItem(self.students[i]))
            #self.index.append('%02d' % (i+1))
        # a = 5
        # print '%02d' % a
        # output: 05
    #

    # Generate
    def serial_number(self):
        cn = self.c_comboBox.findText(self.c_comboBox.currentText())
        #print ('%02d' % (cn+1))
        temp_var = self.from_date.date()
        var_name = temp_var.toPyDate()
        print(str(var_name), str(var_name)[5:7], str(var_name)[2:4])
        # print len(self.students)
        for i in range(len(self.students)):
            serial = ('%02d' % (cn+1)) + str(var_name)[5:7] + str(var_name)[2:4] + ('%02d' % (i+1))
            self.serials.append(serial)
            self.tableWidget.setItem(i, 1, QTableWidgetItem(serial))
        #print(self.serials)
        #print(self.students)

    def gen_certificate(self):
        icon = QIcon('files/icon/pdf_5.png')
        self.directory = QFileDialog.getExistingDirectory(self, 'Select directory')

        self.from_date_value = self.from_date.date().toPyDate().strftime('%m/%d/%Y')

        self.to_date_value = self.to_date.date().toPyDate().strftime('%m/%d/%Y')

        #print(directory)
        model = QStandardItemModel(self.listView)
        for i in range(len(self.students)):
            im2pdf((self.students[i]).upper(), str(self.c_comboBox.currentText()),self.from_date_value , self.to_date_value , ('SN:'+self.serials[i]),
                   str(self.i_comboBox.currentText()), str(self.d_comboBox.currentText()),  self.directory)

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