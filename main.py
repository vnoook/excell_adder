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
# import openpyxl
# import openpyxl.utils
# import openpyxl.styles


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
        # self.text_empty_combobox = 'не выбрано'
        # self.file_IC = ''
        # self.file_GASPS = ''
        # self.wb_file_IC = ''
        # self.wb_file_IC_s = ''
        # self.wb_file_GASPS = ''
        # self.wb_file_GASPS_s = ''

        # главное окно, надпись на нём и размеры
        self.setWindowTitle('Добор в эксель')
        self.setGeometry(600, 400, 900, 300)

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
        self.label_half_file.setGeometry(PyQt5.QtCore.QRect(10, 120, 150, 40))
        font = PyQt5.QtGui.QFont()
        font.setPointSize(12)
        self.label_half_file.setFont(font)
        self.label_half_file.adjustSize()
        self.label_half_file.setToolTip(self.label_half_file.objectName())

        # label_path_half_file
        self.label_path_half_file = PyQt5.QtWidgets.QLabel(self)
        self.label_path_half_file.setObjectName('label_path_half_file')
        self.label_path_half_file.setText(self.text_empty_path_file)
        self.label_path_half_file.setGeometry(PyQt5.QtCore.QRect(70, 150, 820, 20))
        font = PyQt5.QtGui.QFont()
        font.setPointSize(10)
        self.label_path_half_file.setFont(font)
        self.label_path_half_file.adjustSize()
        self.label_path_half_file.setToolTip(self.label_path_half_file.objectName())

        # toolButton_select_half_file
        self.toolButton_select_half_file = PyQt5.QtWidgets.QPushButton(self)
        self.toolButton_select_half_file.setObjectName('toolButton_select_half_file')
        self.toolButton_select_half_file.setText('...')
        self.toolButton_select_half_file.setGeometry(PyQt5.QtCore.QRect(10, 150, 50, 20))
        self.toolButton_select_half_file.setFixedWidth(50)
        self.toolButton_select_half_file.clicked.connect(self.select_file)
        self.toolButton_select_half_file.setToolTip(self.toolButton_select_half_file.objectName())

        # pushButton_do_fill_data
        self.pushButton_do_fill_data = PyQt5.QtWidgets.QPushButton(self)
        self.pushButton_do_fill_data.setObjectName('pushButton_do_fill_data')
        self.pushButton_do_fill_data.setEnabled(False)
        self.pushButton_do_fill_data.setText('Произвести заполнение')
        self.pushButton_do_fill_data.setGeometry(PyQt5.QtCore.QRect(10, 225, 180, 25))
        self.pushButton_do_fill_data.setFixedWidth(130)
        self.pushButton_do_fill_data.clicked.connect(self.do_fill_data)
        self.pushButton_do_fill_data.setToolTip(self.pushButton_do_fill_data.objectName())

        # button_exit
        self.button_exit = PyQt5.QtWidgets.QPushButton(self)
        self.button_exit.setObjectName('button_exit')
        self.button_exit.setText('Выход')
        self.button_exit.setGeometry(PyQt5.QtCore.QRect(10, 260, 180, 25))
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

        # # проверяю чекбокс "с преступлениями" и другими условиями
        # if self.checkBox_prest_IC.checkState() == 2 and\
        #         (self.comboBox_liter_prest_IC.itemText(self.comboBox_liter_prest_IC.currentIndex()) not in
        #          (self.text_empty_combobox, self.comboBox_liter_IC.itemText(
        #              self.comboBox_liter_IC.currentIndex()))):
        #     self.flag_edit_prest = True
        #
        # elif (self.checkBox_prest_IC.checkState() == 2) and\
        #         (self.comboBox_liter_prest_IC.itemText(
        #             self.comboBox_liter_prest_IC.currentIndex()) == self.text_empty_combobox):
        #     # информационное окно о предупреждении выбора полей
        #     self.window_select = PyQt5.QtWidgets.QMessageBox()
        #     self.window_select.setWindowTitle('Поля')
        #     self.window_select.setText(f'Выберите пустые поля или уберите галочку "с преступлениями"')
        #     self.window_select.exec_()
        #     self.flag_edit_prest = False
        #
        # elif (self.checkBox_prest_IC.checkState() == 2) and\
        #         (self.comboBox_liter_prest_IC.itemText(self.comboBox_liter_prest_IC.currentIndex()) ==
        #          self.comboBox_liter_IC.itemText(self.comboBox_liter_IC.currentIndex())):
        #     # информационное окно о сравнении значений двух комбо
        #     self.window_select = PyQt5.QtWidgets.QMessageBox()
        #     self.window_select.setWindowTitle('Сравнение')
        #     self.window_select.setText(f'Поля в строке ИЦ не должны совпадать')
        #     self.window_select.exec_()
        #     self.flag_edit_prest = False
        # else:
        #     self.flag_edit_prest = False
        #
        # # формируются диапазоны для обработки данных в файлах из комбобоксов
        # range_file_IC = self.comboBox_liter_IC.itemText(self.comboBox_liter_IC.currentIndex()) +\
        #                 self.comboBox_digit_IC.itemText(self.comboBox_digit_IC.currentIndex()) +\
        #                 ':' +\
        #                 self.comboBox_liter_IC.itemText(self.comboBox_liter_IC.currentIndex()) +\
        #                 self.comboBox_digit_IC.itemText(self.comboBox_digit_IC.count()-1)
        #
        # range_file_GASPS = self.comboBox_liter_GASPS.itemText(self.comboBox_liter_GASPS.currentIndex()) +\
        #                    self.comboBox_digit_GASPS.itemText(self.comboBox_digit_GASPS.currentIndex()) +\
        #                    ':' +\
        #                    self.comboBox_liter_GASPS.itemText(self.comboBox_liter_GASPS.currentIndex()) +\
        #                    self.comboBox_digit_GASPS.itemText(self.comboBox_digit_GASPS.count()-1)
        #
        # if self.flag_edit_prest:
        #     # формируется диапазон для обработки колонки преступлений
        #     range_file_IC_prest = self.comboBox_liter_prest_IC.itemText(
        #         self.comboBox_liter_prest_IC.currentIndex()) + \
        #                           self.comboBox_digit_IC.itemText(self.comboBox_digit_IC.currentIndex()) + \
        #                           ':' + \
        #                           self.comboBox_liter_prest_IC.itemText(
        #                               self.comboBox_liter_prest_IC.currentIndex()) + \
        #                           self.comboBox_digit_IC.itemText(self.comboBox_digit_IC.count() - 1)
        #     wb_IC_cells_range_prest = self.wb_file_IC_s[range_file_IC_prest]
        #
        # # сформированные диапазоны из выбранных комбобоксов
        # wb_IC_cells_range = self.wb_file_IC_s[range_file_IC]
        # wb_GASPS_cells_range = self.wb_file_GASPS_s[range_file_GASPS]
        #
        # if (self.checkBox_prest_IC.checkState() == 0) or\
        #         (self.checkBox_prest_IC.checkState() == 2 and self.flag_edit_prest):
        #     # формирование множества из обработанных значений ячеек GASPS
        #     for row_in_range_GASPS in wb_GASPS_cells_range:
        #         for cell_in_row_GASPS in row_in_range_GASPS:
        #             indexR_GASPS = wb_GASPS_cells_range.index(row_in_range_GASPS)
        #             indexC_GASPS = row_in_range_GASPS.index(cell_in_row_GASPS)
        #
        #             if wb_GASPS_cells_range[indexR_GASPS][indexC_GASPS].value == None:
        #                 wb_GASPS_cell_value = 'None'
        #             else:
        #                 wb_GASPS_cell_value = str(wb_GASPS_cells_range[indexR_GASPS][indexC_GASPS].value)
        #
        #             for ikud in wb_GASPS_cell_value.split(";"):
        #                 set_data_GASPS.add(ikud.strip().replace('.', ''))
        #
        #             tuple_data_GASPS = tuple(set_data_GASPS)
        #
        #     # обработка файла ИЦ
        #     for row_in_range_IC in wb_IC_cells_range:
        #         for cell_in_row_IC in row_in_range_IC:
        #             # определение адреса ячейки из области данных
        #             indexR_IC = wb_IC_cells_range.index(row_in_range_IC)
        #             indexC_IC = row_in_range_IC.index(cell_in_row_IC)
        #
        #             # получение координаты и значения ячейки IC
        #             if wb_IC_cells_range[indexR_IC][indexC_IC].value == None:
        #                 wb_IC_cell_value = 'None'
        #             else:
        #                 wb_IC_cell_value = str(wb_IC_cells_range[indexR_IC][indexC_IC].value)
        #
        #             # очистка множества для номеров дел из колонки и
        #             # разбивка строки на несколько номеров дел если есть ";"
        #             set_data_IC.clear()
        #             for ikud in wb_IC_cell_value.split(";"):
        #                 set_data_IC.add(ikud.strip().replace('.', ''))
        #
        #             tuple_data_IC = tuple(set_data_IC)
        #
        #             # раскраска колонок УД в ИЦ файле
        #             for ikud in wb_IC_cell_value.split(";"):
        #                 ikud_split = ikud.strip().replace('.', '').replace(' ', '')
        #
        #                 if (ikud_split in tuple_data_GASPS) and (ikud_split in tuple_data_IC):
        #                     wb_IC_cells_range[indexR_IC][indexC_IC].fill =\
        #                         openpyxl.styles.PatternFill(start_color='FF0000', end_color='FF0000',
        #                                                     fill_type='solid')
        #                 elif ikud_split not in tuple_data_GASPS:
        #                     wb_IC_cells_range[indexR_IC][indexC_IC].fill =\
        #                         openpyxl.styles.PatternFill(start_color='878787', end_color='878787',
        #                                                     fill_type='solid')
        #
        #                 # обработка колонки преступности - добавляется номер УД к номеру преступления
        #                 if self.flag_edit_prest:
        #                     wb_IC_cells_range_prest[indexR_IC][indexC_IC].value =\
        #                         ikud_split + wb_IC_cells_range_prest[indexR_IC][indexC_IC].value
        #
        #     # сохраняю файл и закрываю оба
        #     self.wb_file_IC.save(self.file_IC)
        #     self.wb_file_IC.close()
        #     self.wb_file_GASPS.close()
        #
        #     # считаю время заполнения
        #     time_finish = time.time()
        #     '\n' + '.' * 30 + 'закончено за', round(time_finish - time_start, 1), 'секунд'
        #
        #     # информационное окно о сохранении файлов
        #     self.window_info = PyQt5.QtWidgets.QMessageBox()
        #     self.window_info.setWindowTitle('Файлы')
        #     self.window_info.setText(f'Файлы сохранены и закрыты.\n{self.file_IC}\n'
        #                              f'Заполнение сделано за {round(time_finish - time_start, 1)} секунд')
        #     self.window_info.exec_()
        #
        #     # очистка переменных от повторного использования
        #     del set_data_IC
        #     del set_data_GASPS
        #     self.flag_edit_prest = None

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
