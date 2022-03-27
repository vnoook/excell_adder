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
import random

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
        self.max_string = '260'
        self.header_list = ('Фамилия', 'Имя', 'Отчество', 'Email', 'Дата рождения(дд.мм.гггг)', 'Телефон', 'Город',
                            'Основное место работы(сокращения допускаются)', 'Должность', 'Специальность')
        self.spec_set = set()
        self.range_full_file = 'A2:J11501'
        self.range_half_file = 'A2:J215'

        # главное окно, надпись на нём и размеры
        self.setWindowTitle('Добор в эксель')
        self.setGeometry(600, 200, 700, 610)

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
        self.label_half_file.setText('2. Выберите Неполный файл')
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
        self.label_max_string.setText('3. Сколько должно быть строк в файле')
        self.label_max_string.setGeometry(PyQt5.QtCore.QRect(10, 130, 150, 40))
        font = PyQt5.QtGui.QFont()
        font.setPointSize(12)
        self.label_max_string.setFont(font)
        self.label_max_string.adjustSize()
        self.label_max_string.setToolTip(self.label_max_string.objectName())

        # lineEdit_max_string
        self.lineEdit_max_string = PyQt5.QtWidgets.QLineEdit(self)
        self.lineEdit_max_string.setObjectName('lineEdit_max_string')
        self.lineEdit_max_string.setText(self.max_string)
        self.lineEdit_max_string.setGeometry(PyQt5.QtCore.QRect(10, 160, 90, 20))
        self.lineEdit_max_string.setClearButtonEnabled(True)
        self.lineEdit_max_string.setToolTip(self.lineEdit_max_string.objectName())

        # label_spec_string
        self.label_spec_string = PyQt5.QtWidgets.QLabel(self)
        self.label_spec_string.setObjectName('label_spec_string')
        self.label_spec_string.setText('4. Выберите специальности, одну или несколько')
        self.label_spec_string.setGeometry(PyQt5.QtCore.QRect(10, 190, 150, 40))
        font = PyQt5.QtGui.QFont()
        font.setPointSize(12)
        self.label_spec_string.setFont(font)
        self.label_spec_string.adjustSize()
        self.label_spec_string.setToolTip(self.label_spec_string.objectName())

        # listWidget_specialization
        self.listWidget_specialization = PyQt5.QtWidgets.QListWidget(self)
        self.listWidget_specialization.setObjectName('listWidget_specialization')
        self.listWidget_specialization.setGeometry(PyQt5.QtCore.QRect(10, 220, 400, 300))
        self.listWidget_specialization.setSelectionMode(PyQt5.QtWidgets.QListView.MultiSelection)
        font = PyQt5.QtGui.QFont()
        font.setPointSize(10)
        self.listWidget_specialization.setFont(font)
        self.listWidget_specialization.setResizeMode(PyQt5.QtWidgets.QListView.Adjust)
        self.listWidget_specialization.sortItems(True)
        self.listWidget_specialization.setEnabled(False)

        # pushButton_do_fill_data
        self.pushButton_do_fill_data = PyQt5.QtWidgets.QPushButton(self)
        self.pushButton_do_fill_data.setObjectName('pushButton_do_fill_data')
        self.pushButton_do_fill_data.setEnabled(False)
        self.pushButton_do_fill_data.setText('Произвести заполнение')
        self.pushButton_do_fill_data.setGeometry(PyQt5.QtCore.QRect(10, 535, 180, 25))
        self.pushButton_do_fill_data.setFixedWidth(130)
        self.pushButton_do_fill_data.clicked.connect(self.do_fill_data)
        self.pushButton_do_fill_data.setToolTip(self.pushButton_do_fill_data.objectName())

        # button_exit
        self.button_exit = PyQt5.QtWidgets.QPushButton(self)
        self.button_exit.setObjectName('button_exit')
        self.button_exit.setText('Выход')
        self.button_exit.setGeometry(PyQt5.QtCore.QRect(10, 570, 180, 25))
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
            self.info_for_open_file = 'Выберите Неполный файл формата Excel, версии старше 2007 года (.XLSX)'

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

        # нажата кнопка выбора Неполного файла
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
            # если выбранные файлы разные
            if self.text_empty_path_file not in (self.label_path_full_file.text(), self.label_path_half_file.text()):
                self.pushButton_do_fill_data.setEnabled(True)
                self.listWidget_specialization.setEnabled(True)
                self.listWidget_specialization.clear()
        # если выбранные файлы одинаковые
        else:
            self.pushButton_do_fill_data.setEnabled(False)
            self.listWidget_specialization.setEnabled(False)
            self.listWidget_specialization.clear()

        if self.listWidget_specialization.isEnabled():
            # открыть файлы Полный и Неполный, и выбрать листы
            wb_full = openpyxl.load_workbook(self.label_path_full_file.text())
            wb_full_s = wb_full.active
            wb_half = openpyxl.load_workbook(self.label_path_half_file.text())
            wb_half_s = wb_half.active

            # посчитать количество строк и вывести на форме
            self.label_full_file.setText(self.label_full_file.text() + f' (строк в файле {str(wb_full_s.max_row -1)})')
            self.label_full_file.adjustSize()
            self.label_half_file.setText(self.label_half_file.text() + f' (строк в файле {str(wb_half_s.max_row -1)})')
            self.label_half_file.adjustSize()

            # сформированные диапазоны обработки
            wb_full_range = wb_full_s[self.range_full_file]

            # множество для хранения специальностей
            self.spec_set = set()

            # цикл прохода по полному файлу
            for row_in_range_full in wb_full_range:
                if row_in_range_full[-1].value:
                    self.spec_set.add(row_in_range_full[-1].value)

            self.listWidget_specialization.addItems(sorted(self.spec_set, reverse=False))
            wb_full.close()
            wb_half.close()

    # событие - нажатие на кнопку заполнения файла
    def do_fill_data(self):
        # выбор выбранных строк в списке специальностей
        spec_selected = [item.text() for item in self.listWidget_specialization.selectedItems()]

        # проверка на количество выбранных строк в listWidget_specialization
        if len(spec_selected) == 0:
            # информационное окно о сохранении файлов
            self.window_info = PyQt5.QtWidgets.QMessageBox()
            self.window_info.setWindowTitle('Число')
            self.window_info.setText(f'В списке специальностей ничего не выбрано,\n'
                                     f'выберите хотя бы одну строку')
            self.window_info.exec_()
        else:
            # считаю время заполнения
            time_start = time.time()

            # открыть файл Полный и Неполный, и выбрать листы
            wb_full = openpyxl.load_workbook(self.label_path_full_file.text())
            wb_full_s = wb_full.active
            wb_half = openpyxl.load_workbook(self.label_path_half_file.text())
            wb_half_s = wb_half.active

            # сформированные диапазоны обработки
            wb_full_range = wb_full_s[self.range_full_file]
            wb_half_range = wb_half_s[self.range_half_file]

            # списки всех строк, одной строки прохода, выбранных строк по специальностям
            list_one_string = []  # временная переменная для значения ячейки
            list_filtered_string = []  # фильтрованные строки из Полного файла которые устраивают выбранным специальностям
            list_half_file = []  # весь Неполный файл

            # счётчик удачных добавлений из выбранных строк
            count_add_succes = 0

            # цикл прохода по полному файлу, для заполнения list_filtered_string фильтрованных из spec_selected
            for row_in_range_full in wb_full_range:
                # чищу список для временной строки
                list_one_string = []

                # прохожу строку
                for cell_in_row_full in row_in_range_full:
                    list_one_string.append(cell_in_row_full.value)

                # если последнее значение в списке специальностей, то добавляю его в список выбранных из полного файла
                if list_one_string[-1] in spec_selected:
                    list_filtered_string.append(list_one_string)

            # цикл прохода по Неполному файлу
            for row_in_range_half in wb_half_range:
                # чищу список для временной строки
                list_one_string = []

                # прохожу строку
                for cell_in_row_half in row_in_range_half:
                    list_one_string.append(cell_in_row_half.value)

                # все записи из Неполного файла
                list_half_file.append(list_one_string)

            # количество строк "сколько хочу строк" (перевод значения в поле шага 3)
            count_string_want = int(self.lineEdit_max_string.text())

            # количество строк в Неполном файле (-1 потому что верхняя строка это шапка)
            count_string_half = wb_half_s.max_row -1

            # сколько нужно добавить строк в Неполный файл
            # разница количества строк между тем, что "сколько хочу строк" и строк уже имеется в файле
            count_add_string = count_string_want - count_string_half

            # количество строк в отфильтрованном списке
            count_filter_string = len(list_filtered_string)

            # количество строк которых будет реально добавлены в Неполный файл
            count_real_data_add = count_filter_string - count_add_string

            # если количество строк в Неполном меньше, чем хочется, то добавить разницу строк
            if count_add_string <= 0:
                # информационное окно
                self.window_info = PyQt5.QtWidgets.QMessageBox()
                self.window_info.setWindowTitle('Строки')
                self.window_info.setText(f'Количество строк в Неполном файле больше или одинаково,\n'
                                         f'чем в ПУНКТЕ 3, их разница равна {count_add_string}')
                self.window_info.exec_()
            else:
                # если добавляемых больше, чем отфильтрованых, то добавлять всё из list_filtered_string
                if count_add_string > count_filter_string:
                    # флаг добавления "всё что есть в list_filtered_string"
                    flag_add_all = True

                    # информационное окно
                    self.window_info = PyQt5.QtWidgets.QMessageBox()
                    self.window_info.setWindowTitle('Строки')
                    self.window_info.setText(f'Количество строк в Полном файле по этим специальностям\n'
                                             f'меньше, чем в ПУНКТЕ 3, их разница равна {count_real_data_add}')
                    self.window_info.exec_()
                else:
                    # флаг добавления "всё что есть в list_filtered_string"
                    flag_add_all = False

                    print(11111)

                    # кортеж Неполного файла с первыми тремя ячейками ФИО для проверки
                    # вхождения случайно выбранного из list_filtered_string
                    list_dif = []
                    # print(f'{list_half_file = }')
                    # print(f'{len(list_half_file) = }')
                    for str_half in list_half_file:
                        print(f'{str_half = }')
                        str_temp = ''
                        for cell_half in range(0, 3):  # беру первые три значения где ФИО
                            # преобразую их в безпробельную строку в нижнем регистре
                            str_temp = str_temp + ''.join(str(str_half[cell_half]).lower().split())
                        print(f'{str_temp = }')
                        list_dif.append(str_temp)
                    tuple_half_file = tuple(list_dif)

                    # print(f'{list_dif = }')
                    # print(f'{tuple_half_file = }')

                    print(22222)

                    # выбрать count_add_string штук из list_filtered_string и добавить только их
                    flag_add_succes = False  # условие выхода - достижение количества нужный выбраных случайных строк
                    while flag_add_succes == False:
                        # выбираю случайную строку из подготовленных по специальностям
                        random_string = random.choice(list_filtered_string)

                        # преобразую её в безпробельную строку в нижнем регистре
                        compare_string = ''.join(random_string[0:3]).lower()

                        # проверяю есть ли рандомная строка в выбраном списке
                        if compare_string in tuple_half_file:
                            print(compare_string)
                            # TODO
                        else:
                            count_add_succes += 1
                            if count_add_succes == count_add_string:
                                flag_add_succes = True
                            print(count_add_succes)
                    else:
                        # информационное окно о сохранении файлов
                        self.window_info = PyQt5.QtWidgets.QMessageBox()
                        self.window_info.setWindowTitle('Строки')
                        self.window_info.setText(f'Не хватило данных для добавления,\n'
                                                 f'добавлено в файл сколько было.')
                        self.window_info.exec_()

                    print(333333)


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

            pass

            # сохраняю файл и закрываю оба
            # self.wb_file_IC.save(self.file_IC)
            wb_full.close()
            wb_half.close()

            # считаю время заполнения
            time_finish = time.time()

            # информационное окно о сохранении файлов
            self.window_info = PyQt5.QtWidgets.QMessageBox()
            self.window_info.setWindowTitle('Файл')
            self.window_info.setText(f'Файлы сохранены и закрыты.\n'
                                     f'Заполнение сделано за {round(time_finish - time_start, 1)} секунд')
            self.window_info.exec_()

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
