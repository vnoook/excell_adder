import time
import openpyxl

# считаю время скрипта
time_start = time.time()
print('начинается' + '.'*20)

# файл для работы
xl_file = 'res/spisok_main1.xlsx'

# открываю книги
wb_file = openpyxl.load_workbook(xl_file)
# wb_file_s = wb_file.active
wb_file_s1 = wb_file['Лист1']
wb_file_s2 = wb_file['Лист2']

# переменные для работы
max_row_s1 = wb_file_s1.max_row


# алгоритмы добавления данных
for i in range(1, max_row_s1 + 1):
    user_fio = wb_file_s1.cell(i, 1).value
    user_email = wb_file_s1.cell(i, 6).value

    print(i, ' == ', user_fio)
    print(i, ' == ', user_email)

exit()

# сохраняю файл и закрываю его
# wb_file.save(xl_file)
# wb_file.close()

# считаю время скрипта
time_finish = time.time()
print('\n' + '.'*30 + 'закончено за', round(time_finish-time_start, 3), 'секунд')

# закрываю программу
# input('\nНажмите ENTER')
