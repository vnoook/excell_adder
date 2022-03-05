import time
import openpyxl

# считаю время скрипта
time_start = time.time()
print('начинается' + '.'*20)

# файлы для работы
xl_file = 'res/111222333.xlsx'

# переменные для работы
min_row_xl_my_with_id = 2
max_row_xl_my_with_id = 1249
min_col_xl_my_with_id = 1
max_col_xl_my_with_id = 23

# открываю книги
wb_file = openpyxl.load_workbook(xl_file)
wb_file_s = wb_file.active

# алгоритмы добавления данных
for i in range(min_row_xl_my_with_id, max_row_xl_my_with_id+1):
    user_id = wb_file_s.cell(i, 1).value
    user_grade = 'через поиск совпадений в xl_all_users'
    user_title = 'через поиск совпадений в xl_all_users'
    user_date_reg = 'нет такой информации, надо выдумать алгоритм подсчёта от даты мероприятия в xl_real_data'
    user_platform = 'xl_real_data'
    user_country = 'xl_real_data'

    print(i, ' == ', user_id)
    print(i, ' == ', user_platform)

# сохраняю файл и закрываю его
wb_file.save(xl_file)
wb_file.close()

# считаю время скрипта
time_finish = time.time()
print('\n' + '.'*30 + 'закончено за', round(time_finish-time_start, 3), 'секунд')

# закрываю программу
# input('\nНажмите ENTER')
