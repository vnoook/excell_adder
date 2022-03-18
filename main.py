# ...
# INSTALL
# pip install openpyxl
# pip install PyQt5
# COMPILE
# pyinstaller -F -w main.py
# ...

import sys
import time
import PyQt5
import PyQt5.QtWidgets
import PyQt5.QtCore
import PyQt5.QtGui
import openpyxl
import openpyxl.utils


# класс главного окна
class Window(PyQt5.QtWidgets.QMainWindow):
    # описание главного окна
    def __init__(self):
        super(Window, self).__init__()

        # переменные, атрибуты
        self.info_for_open_file = ''
        self.info_path_open_file = ''
        self.info_extention_open_file = 'Файлы Excel xlsx (*.xlsx)'
        self.text_empty_path_file = 'файл пока не выбран'
        self.file_full = ''
        self.file_half = ''
        self.max_string = 1000
        self.spec_list = ('111', '222', '333', '444', '555', '666', '777', '888', '999', '000')
        self.range_full_file = 'A2:J11501'
        self.range_half_file = 'A2:J256'

        # главное окно, надпись на нём и размеры
        self.setWindowTitle('Добор в эксель')
        self.setGeometry(600, 300, 900, 340)

        # объекты на главном окне
        # label_full_file
        self.label_full_file = PyQt5.QtWidgets.QLabel(self)
        self.label_full_file.setObjectName('label_full_file')
        self.label_full_file.setText('1. Выберите полный файл')
        self.label_full_file.setGeometry(PyQt5.QtCore.QRect(10, 10, 150, 40))
        font = PyQt5.QtGui.QFont()
        font.setPointSize(12)
        self.label_full_file.setFont(font)
        self.label_full_file.adjustSize()
        self.label_full_file.setToolTip(self.label_full_file.objectName())

        # toolButton_select_full_file
        self.toolButton_select_full_file = PyQt5.QtWidgets.QPushButton(self)
        self.toolButton_select_full_file.setObjectName('toolButton_select_full_file')
        self.toolButton_select_full_file.setText('...')
        self.toolButton_select_full_file.setGeometry(PyQt5.QtCore.QRect(10, 40, 50, 20))
        self.toolButton_select_full_file.setFixedWidth(50)
        self.toolButton_select_full_file.clicked.connect(self.select_file)
        self.toolButton_select_full_file.setToolTip(self.toolButton_select_full_file.objectName())

        # label_path_full_file
        self.label_path_full_file = PyQt5.QtWidgets.QLabel(self)
        self.label_path_full_file.setObjectName('label_path_full_file')
        self.label_path_full_file.setText(self.text_empty_path_file)
        self.label_path_full_file.setGeometry(PyQt5.QtCore.QRect(70, 40, 820, 16))
        font = PyQt5.QtGui.QFont()
        font.setPointSize(10)
        self.label_path_full_file.setFont(font)
        self.label_path_full_file.adjustSize()
        self.label_path_full_file.setToolTip(self.label_path_full_file.objectName())

        # label_half_file
        self.label_half_file = PyQt5.QtWidgets.QLabel(self)
        self.label_half_file.setObjectName('label_half_file')
        self.label_half_file.setText('2. Выберите неполный файл')
        self.label_half_file.setGeometry(PyQt5.QtCore.QRect(10, 70, 150, 40))
        font = PyQt5.QtGui.QFont()
        font.setPointSize(12)
        self.label_half_file.setFont(font)
        self.label_half_file.adjustSize()
        self.label_half_file.setToolTip(self.label_half_file.objectName())

        # toolButton_select_half_file
        self.toolButton_select_half_file = PyQt5.QtWidgets.QPushButton(self)
        self.toolButton_select_half_file.setObjectName('toolButton_select_half_file')
        self.toolButton_select_half_file.setText('...')
        self.toolButton_select_half_file.setGeometry(PyQt5.QtCore.QRect(10, 100, 50, 20))
        self.toolButton_select_half_file.setFixedWidth(50)
        self.toolButton_select_half_file.clicked.connect(self.select_file)
        self.toolButton_select_half_file.setToolTip(self.toolButton_select_half_file.objectName())

        # label_path_half_file
        self.label_path_half_file = PyQt5.QtWidgets.QLabel(self)
        self.label_path_half_file.setObjectName('label_path_half_file')
        self.label_path_half_file.setText(self.text_empty_path_file)
        self.label_path_half_file.setGeometry(PyQt5.QtCore.QRect(70, 100, 820, 20))
        font = PyQt5.QtGui.QFont()
        font.setPointSize(10)
        self.label_path_half_file.setFont(font)
        self.label_path_half_file.adjustSize()
        self.label_path_half_file.setToolTip(self.label_path_half_file.objectName())

        # label_max_string
        self.label_max_string = PyQt5.QtWidgets.QLabel(self)
        self.label_max_string.setObjectName('label_max_string')
        self.label_max_string.setText('3. Сколько добавить строк в файл')
        self.label_max_string.setGeometry(PyQt5.QtCore.QRect(10, 130, 150, 40))
        font = PyQt5.QtGui.QFont()
        font.setPointSize(12)
        self.label_max_string.setFont(font)
        self.label_max_string.adjustSize()
        self.label_max_string.setToolTip(self.label_max_string.objectName())

        # lineEdit_max_string
        self.lineEdit_max_string = PyQt5.QtWidgets.QLineEdit(self)
        self.lineEdit_max_string.setObjectName('lineEdit_max_string')
        self.lineEdit_max_string.setText('')
        self.lineEdit_max_string.setGeometry(PyQt5.QtCore.QRect(10, 160, 90, 20))
        self.lineEdit_max_string.setClearButtonEnabled(True)
        self.lineEdit_max_string.setToolTip(self.lineEdit_max_string.objectName())

        # label_spec_string
        self.label_spec_string = PyQt5.QtWidgets.QLabel(self)
        self.label_spec_string.setObjectName('label_spec_string')
        self.label_spec_string.setText('4. Введите специализации через запятую')
        self.label_spec_string.setGeometry(PyQt5.QtCore.QRect(10, 190, 150, 40))
        font = PyQt5.QtGui.QFont()
        font.setPointSize(12)
        self.label_spec_string.setFont(font)
        self.label_spec_string.adjustSize()
        self.label_spec_string.setToolTip(self.label_spec_string.objectName())

        # lineEdit_spec_string
        self.lineEdit_spec_string = PyQt5.QtWidgets.QLineEdit(self)
        self.lineEdit_spec_string.setObjectName('lineEdit_spec_string')
        self.lineEdit_spec_string.setText('')
        self.lineEdit_spec_string.setGeometry(PyQt5.QtCore.QRect(10, 220, 300, 20))
        self.lineEdit_spec_string.setClearButtonEnabled(True)
        self.lineEdit_spec_string.setToolTip(self.lineEdit_spec_string.objectName())



        # # comboBox_specialization
        # self.comboBox_specialization = PyQt5.QtWidgets.QComboBox(self)
        # self.comboBox_specialization.setObjectName('comboBox_specialization')
        # self.comboBox_specialization.setGeometry(PyQt5.QtCore.QRect(10, 350, 70, 20))
        # self.comboBox_specialization.addItem('пусто')
        # self.comboBox_specialization.setEnabled(True)
        # self.comboBox_specialization.setVisible(True)
        #
        # item = QListWidgetItem(cfg.get_description())
        # item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
        #
        # self.comboBox_specialization.adjustSize()
        # self.comboBox_specialization.setToolTip(self.comboBox_specialization.objectName())
        # self.comboBox_specialization = PyQt5.QtWidgets.QLabel(self)
        # font = PyQt5.QtGui.QFont()
        # font.setPointSize(12)
        # self.comboBox_specialization.setFont(font)




        # pushButton_do_fill_data
        self.pushButton_do_fill_data = PyQt5.QtWidgets.QPushButton(self)
        self.pushButton_do_fill_data.setObjectName('pushButton_do_fill_data')
        self.pushButton_do_fill_data.setEnabled(False)
        self.pushButton_do_fill_data.setText('Произвести заполнение')
        self.pushButton_do_fill_data.setGeometry(PyQt5.QtCore.QRect(10, 260, 180, 25))
        self.pushButton_do_fill_data.setFixedWidth(130)
        self.pushButton_do_fill_data.clicked.connect(self.do_fill_data)
        self.pushButton_do_fill_data.setToolTip(self.pushButton_do_fill_data.objectName())

        # button_exit
        self.button_exit = PyQt5.QtWidgets.QPushButton(self)
        self.button_exit.setObjectName('button_exit')
        self.button_exit.setText('Выход')
        self.button_exit.setGeometry(PyQt5.QtCore.QRect(10, 300, 180, 25))
        self.button_exit.setFixedWidth(50)
        self.button_exit.clicked.connect(self.click_on_btn_exit)
        self.button_exit.setToolTip(self.button_exit.objectName())

    # событие - нажатие на кнопку выбора файла
    def select_file(self):
        # запоминание старого значения пути выбора файлов
        old_path_of_selected_full_file = self.label_path_full_file.text()
        old_path_of_selected_half_file = self.label_path_half_file.text()

        # определение какая кнопка выбора файла нажата
        if self.sender().objectName() == self.toolButton_select_full_file.objectName():
            self.info_for_open_file = 'Выберите полный файл формата Excel, версии старше 2007 года (.XLSX)'
        elif self.sender().objectName() == self.toolButton_select_half_file.objectName():
            self.info_for_open_file = 'Выберите неполный файл формата Excel, версии старше 2007 года (.XLSX)'

        # непосредственное окно выбора файла и переменная для хранения пути файла
        data_of_open_file_name = PyQt5.QtWidgets.QFileDialog.getOpenFileName(self,
                                                                             self.info_for_open_file,
                                                                             self.info_path_open_file,
                                                                             self.info_extention_open_file)
        # вычленение пути файла из data_of_open_file_name
        file_name = data_of_open_file_name[0]

        # выбор где и что менять исходя из выбора пользователя
        # нажата кнопка выбора полного файла
        if self.sender().objectName() == self.toolButton_select_full_file.objectName():
            if file_name == '':
                self.label_path_full_file.setText(old_path_of_selected_full_file)
                self.label_path_full_file.adjustSize()
            else:
                old_path_of_selected_full_file = self.label_path_full_file.text()
                self.label_path_full_file.setText(file_name)
                self.label_path_full_file.adjustSize()

        # нажата кнопка выбора неполного файла
        if self.sender().objectName() == self.toolButton_select_half_file.objectName():
            if file_name == '':
                self.label_path_half_file.setText(old_path_of_selected_half_file)
                self.label_path_half_file.adjustSize()
            else:
                old_path_of_selected_half_file = self.label_path_half_file.text()
                self.label_path_half_file.setText(file_name)
                self.label_path_half_file.adjustSize()

        # активация и деактивация объектов на форме зависящее от выбраны ли все файлы и они разные
        if self.label_path_full_file.text() != self.label_path_half_file.text():
            if self.text_empty_path_file not in (self.label_path_full_file.text(), self.label_path_half_file.text()):
                self.pushButton_do_fill_data.setEnabled(True)
        else:
            self.pushButton_do_fill_data.setEnabled(False)

    # событие - нажатие на кнопку заполнения файла
    def do_fill_data(self):
        # считаю время заполнения
        time_start = time.time()

        # определение множеств
        set_data_full_file = set()
        set_data_half_file = set()

        print(f'специальности {self.spec_list = }')

        # открыть файлы Полный и НЕПолный, и выбрать листы
        wb_full = openpyxl.load_workbook(self.label_path_full_file.text())
        wb_full_s = wb_full.active
        wb_half = openpyxl.load_workbook(self.label_path_half_file.text())
        wb_half_s = wb_half.active

        # сформированные диапазоны
        wb_full_range = wb_full_s[self.range_full_file]
        wb_half_range = wb_half_s[self.range_half_file]

        # TODO
        # 4.1) взять строку из Полного
        # 5.1) проверить, есть ли она в НЕПолном (проверять по ФИО+почта)
        # 6.1) вставить в НЕПолный или взять новую
        #
        # 4.2) взять все строки в Полном
        # 5.2) взять все строки в НЕПолном,
        # 6.2) вычесть из Полных все строки из НЕПолного файла
        # 7.2) из полученного множества случайным образом брать строки для добавления в НЕПолный
        for row_in_range_full in wb_full_range:
            print()
            for cell_in_row_full in row_in_range_full:
                print(cell_in_row_full.value, end=' ')
                # TODO

                # if wb_GASPS_cells_range[indexR_GASPS][indexC_GASPS].value == None:
                #     wb_GASPS_cell_value = 'None'
                # else:
                #     wb_GASPS_cell_value = str(wb_GASPS_cells_range[indexR_GASPS][indexC_GASPS].value)
                #
                # for ikud in wb_GASPS_cell_value.split(";"):
                #     set_data_GASPS.add(ikud.strip().replace('.', ''))
                #
                # tuple_data_GASPS = tuple(set_data_GASPS)

        # # обработка файла ИЦ
        # for row_in_range_IC in wb_IC_cells_range:
        #     for cell_in_row_IC in row_in_range_IC:
        #         # определение адреса ячейки из области данных
        #         indexR_IC = wb_IC_cells_range.index(row_in_range_IC)
        #         indexC_IC = row_in_range_IC.index(cell_in_row_IC)
        #
        #         # получение координаты и значения ячейки IC
        #         if wb_IC_cells_range[indexR_IC][indexC_IC].value == None:
        #             wb_IC_cell_value = 'None'
        #         else:
        #             wb_IC_cell_value = str(wb_IC_cells_range[indexR_IC][indexC_IC].value)
        #
        #         # очистка множества для номеров дел из колонки и
        #         # разбивка строки на несколько номеров дел если есть ";"
        #         set_data_IC.clear()
        #         for ikud in wb_IC_cell_value.split(";"):
        #             set_data_IC.add(ikud.strip().replace('.', ''))
        #
        #         tuple_data_IC = tuple(set_data_IC)
        #
        #         # раскраска колонок УД в ИЦ файле
        #         for ikud in wb_IC_cell_value.split(";"):
        #             ikud_split = ikud.strip().replace('.', '').replace(' ', '')
        #
        #             if (ikud_split in tuple_data_GASPS) and (ikud_split in tuple_data_IC):
        #                 wb_IC_cells_range[indexR_IC][indexC_IC].fill =\
        #                     openpyxl.styles.PatternFill(start_color='FF0000', end_color='FF0000',
        #                                                 fill_type='solid')
        #             elif ikud_split not in tuple_data_GASPS:
        #                 wb_IC_cells_range[indexR_IC][indexC_IC].fill =\
        #                     openpyxl.styles.PatternFill(start_color='878787', end_color='878787',
        #                                                 fill_type='solid')
        #
        #             # обработка колонки преступности - добавляется номер УД к номеру преступления
        #             if self.flag_edit_prest:
        #                 wb_IC_cells_range_prest[indexR_IC][indexC_IC].value =\
        #                     ikud_split + wb_IC_cells_range_prest[indexR_IC][indexC_IC].value
        #
        # # сохраняю файл и закрываю оба
        # self.wb_file_IC.save(self.file_IC)
        # self.wb_file_IC.close()
        # self.wb_file_GASPS.close()
        #
        # # считаю время заполнения
        # time_finish = time.time()
        # '\n' + '.' * 30 + 'закончено за', round(time_finish - time_start, 1), 'секунд'
        #
        # # информационное окно о сохранении файлов
        # self.window_info = PyQt5.QtWidgets.QMessageBox()
        # self.window_info.setWindowTitle('Файлы')
        # self.window_info.setText(f'Файлы сохранены и закрыты.\n{self.file_IC}\n'
        #                          f'Заполнение сделано за {round(time_finish - time_start, 1)} секунд')
        # self.window_info.exec_()
        #
        # # очистка переменных от повторного использования
        # del set_data_IC
        # del set_data_GASPS
        # self.flag_edit_prest = None

    # событие - нажатие на кнопку Выход
    def click_on_btn_exit(self):
        exit()


# создание основного окна
def main_app():
    app = PyQt5.QtWidgets.QApplication(sys.argv)
    app.setStyle('Fusion')
    app_window_main = Window()
    app_window_main.show()
    sys.exit(app.exec_())


# запуск основного окна
if __name__ == '__main__':
    main_app()
