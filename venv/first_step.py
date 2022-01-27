# -*- coding: utf-8 -*-
import os, time, itertools,re
from datetime import datetime, date
from MyOfficeSDKDocumentAPI import DocumentAPI as sdk
from string import ascii_uppercase


def log_file(filename, cell, is_first_error, log):
    if is_first_error:
        log.append(datetime.now().strftime("%Y-%m-%d %H:%M:%S") + " Ошибка в файле: " + " " + filename+ " в ячейке " + cell)
    else:
        log[len(log) - 1] = log[len(log) - 1] + ''.join(", в ячейке " + cell)
    return log


def log_file_rec(log, folderName):
    f = open(folderName + "\\errors.log", "a")
    for i in log:
        f.write(i + "\n")
    f.close()


def extract_txt_doc(path, folderName, mydirs_, lena,log,workstatuspanel):
    try:
        document_xlsinput = application.loadDocument(path)
    except Exception as e:
        raise Exception("Невозможно открыть документ: " + os.path.basename(path))
    table_input = document_xlsinput.getBlocks().getTable(0)

    last_name = table_input.getCell('C6').getRawValue()
    first_name = table_input.getCell('C8').getRawValue()
    middle_name = table_input.getCell('C10').getRawValue()

    date_birth = table_input.getCell('C12').getFormattedValue()
    country = table_input.getCell('C14').getRawValue()
    district = table_input.getCell('C15').getRawValue()
    post_index = table_input.getCell('C16').getRawValue()
    region = table_input.getCell('C18').getRawValue()
    city = table_input.getCell('C20').getRawValue()
    # address = table_input.getCell('C22').getRawValue()
    school = table_input.getCell('C22').getRawValue()
    school_address = table_input.getCell('C24').getRawValue()
    exp = table_input.getCell('C26').getRawValue()

    cert_1 = table_input.getCell('C28').getRawValue()
    cert_2 = table_input.getCell('C29').getRawValue()
    cert_3 = table_input.getCell('C30').getRawValue()
    cert_4 = table_input.getCell('C31').getRawValue()
    cert_5 = table_input.getCell('C32').getRawValue()
    cert = [cert_1, cert_2, cert_3, cert_4, cert_5]
    ball_list = [3, 5, 10, 5, 7]
    cert = [re.sub('[Дд][Аа]', 'Да', i) for i in cert]
    cert = [re.sub('[Нн][Ее][Тт]', 'Нет', i) for i in cert]
    diplom_list = ["Диплом 1 степени", "Диплом 2 степени", "Диплом 3 степени", "Диплом Почтового комиссара",
                   'Диплом "Наставника"']
    cert_list = [False if a=='' else a if a == "Нет" else b for a, b in zip(cert, diplom_list)] #Для того чтобы написать в сводном
                                                                                # файле название диплома или нет
    certflag = True
    da_count = [i for i in cert if re.search("Да", i)]
    if len(da_count) > 3:
        certflag = False
    cert_ball = [(a == 'Да') * b for a, b in zip(cert, ball_list)]

    phone = table_input.getCell('C33').getRawValue()
    email = table_input.getCell('C35').getRawValue()

    parent_fio = table_input.getCell('C37').getRawValue()
    parent_phone = table_input.getCell('C41').getRawValue()
    parent_email = table_input.getCell('C43').getRawValue()
    parent_work = table_input.getCell('C45').getRawValue()

    regex_error = '№_\d+.*_Ошибка\.xlsx'
    regex_date = '^(0[1-9]|[12][0-9]|3[01])[- /.](0[1-9]|1[012])[- /.](19|20)\d\d$'  # dd-mm-yyyy
    filename = str("№_" + str(lena) + "_" + last_name + "_" + first_name)
    # Раскрашиваем ячейки в анкетах с ошибками
    if not re.search(regex_date, str(date_birth)):
        date_birth=False
    dict_cells = {
        'C6': last_name,
        'C8': first_name,
        'C10': middle_name,

        'C12': date_birth,
        'C14': country,
        'C15': district,
        'C16': post_index,
        'C18': region,
        'C20': city,
        'C22': school,
        'C24': school_address,
        'C26': exp,

        'C27': certflag,
        'C28': cert_1,
        'C29': cert_2,
        'C30': cert_3,
        'C31': cert_4,
        'C32': cert_5,

        'C33': phone,
        'C35': email,

        'C37': parent_fio,
        'C41': parent_phone,
        'C43': parent_email,
        'C45': parent_work,
    }
    cell_properties = table_input.getCell('C6').getCellProperties()
    cell_properties.backgroundColor = sdk.ColorRGBA(255, 0, 0, 1)
    is_changed = False
    is_first_error = True

    error_count=0
    filename_err=str(filename + "_Ошибка.xlsx")
    for key, val in dict_cells.items():
        if not val:
            # Помечаем ошибки красным цветом в входном файле
            table_input.getCell(key).setCellProperties(cell_properties)
            log = log_file(filename_err, key, is_first_error, log)
            is_changed = True
            is_first_error = False
            error_count+=1
    if is_changed:

        # Ошибка
        # document_xlsinput.saveAs((mydirs_[0]+"\\"+filename))
        if not (os.path.exists(mydirs_[0])):
            try:
                os.mkdir(mydirs_[0], 0o777)  # Папка Ошибки
                workstatuspanel.Text = u'Создана папка: ' + os.path.abspath(mydirs_[0])
            except Exception:
                workstatuspanel.Text = u'Неудалось создать папку:' + os.path.basename(mydirs_[0])
                raise Exception("Неудалось создать папку:" + os.path.basename(mydirs_[0]))
        try:
            document_xlsinput.saveAs((mydirs_[0] + "\\" + filename_err))
            workstatuspanel.Text = u'Сохраняю файл с ошибкой:' + filename_err
        except Exception as e:
            raise Exception("Невозможно сохранить документ: " + (mydirs_[0] + "\\" + filename_err))
        os.remove(path)
    else:
        filename_new = str(filename + "_Обработан.xlsx")
        if re.search(regex_error, str(os.path.basename(path))):
            lena = int(os.path.basename(path).split("_")[1])
            filename_new = str(os.path.basename(path)).replace("Ошибка", "Обработан")
        os.rename(path, os.path.dirname(path) + "\\" + filename_new)
    full_row_lst = [
        last_name,
        first_name,
        middle_name,
        date_birth,
        country,
        district,
        post_index,
        region,
        city,
        school,
        school_address,
        exp,
        str(cert_list[0]) + ", " + str(cert_list[1]) + ", " + str(cert_list[2]) + ", " + str(cert_list[3]) + ", " + str(cert_list[4]),
        phone,
        email,
        parent_fio,
        parent_phone,
        parent_email,
        parent_work,
        cert_ball,
        certflag,
        lena,
        error_count,
    ]
    return full_row_lst


def iter_all_strings():
    for size in itertools.count(1):
        for s in itertools.product(ascii_uppercase, repeat=size):
            yield "".join(s)


def list_xls(rang):
    lst_addr = []
    for s in iter_all_strings():
        lst_addr.append(s)
        if s == rang:
            break
    return lst_addr


def write_table(all_str_lst, worker, datetime1_end, n_rows):
    print ("Запись данных.")
    current_row = n_rows
    column = list_xls("AB")
    for str_ in all_str_lst:
        index = current_row - 3  # 1
        #print "Сохраняю в таблицу номер записи: " + str(index)
        worker.ReportProgress(93, "Сохраняю в таблицу номер записи")
        row_str = str(current_row)  # 4
        table_output_xlsx.getCell("G" + row_str).setFormula(
            "=TRUNC(DAYS($C$1" + ",F" + row_str + ")/365.242199, 0" + ")")
        table_output_xlsx.getCell("A" + row_str).setNumber(index)
        table_output_xlsx.getCell("E" + row_str).setText(str_[0] + " " + str_[1] + " " + str_[2])
        #table_output_xlsx.getCell("AL" + row_str).setNumber(str_[21])

        # A4 set text 1
        k = 1  # Начинаем с B
        # print ("str", str_[0]
        for s in str_[:-4]:
            # print ("Столбец ", column, " Строка ", row_str
            # двигаемся построчно
            if k == 4 or k == 6:
                k += 1
            table_output_xlsx.getCell(column[k] + row_str).setText(s)  # column[1]-B+4 settext из лист str_
            k += 1
        k = 23 # Начинаем с X
        for b in str_[19]:
            table_output_xlsx.getCell(column[k] + row_str).setNumber(b)
            k += 1

        # column = chr(k+66) #Получаем из ASCII, 66 = "B"
        # print ("Столбец ", column, " Строка ", row_str
        # table_output_xlsx.getCell(column + row_str).setText(s) #B+4 settext из лист str_
        current_row += 1


def error_data(data_error, worker):  # Раскрашивает в выходном файле строки с ошибками
    worker.ReportProgress(98, "Помечаем ошибки")
    for i in data_error:
        cell_properties = sdk.CellProperties()
        cell_properties.backgroundColor = sdk.ColorRGBA(255, 0, 0, 1)
        cell_properties.verticalAlignment = sdk.VerticalAlignment_Center
        cell_range = table_output_xlsx.getCellRange("B" + str(i + 3) + ":V" + str(i + 3))
        cell_range.setCellProperties(cell_properties)


def set_cells_format(number_rows, worker,n_rows):
    current_row = n_rows
    row_str = str(current_row)
    print ("Применение форматирования.")
    worker.ReportProgress(94, "Применение форматирования.")
    """cell_range_aligment = table_output_xlsx.getCellRange("A"+row_str+":A" + number_rows)
    for c in cell_range_aligment:
        c.setCellProperties(cell_properties_aligment)"""
    # cell_range_aligment.setCellProperties(cell_properties_aligment) Баг-репорт

    # Формат Date для столбца E4 потому что SDK тупит
    worker.ReportProgress(97, "Формат Date для столбца F")
    cell_range_date = table_output_xlsx.getCellRange("F"+row_str+":F" + number_rows)
    for c in cell_range_date:
        c.setFormat(sdk.CellFormat_Date)





application = None
document_xls = None


def main_(worker, folderName, mydirs_, date_end,workstatuspanel):
    log = []
    global application, table_output_xlsx
    application = sdk.Application()
    folder_url = mydirs_[6]  # Анкеты
    try:
        document_xls = application.loadDocument(mydirs_[11])  # load Сводный
    except Exception as e:
        raise Exception("Невозможно открыть документ: " + os.path.basename(mydirs_[11]))
    table_output_xlsx = document_xls.getBlocks().getTable(0)  # table Сводный
    all_str_lst = []
    error_index = []

    def split(s):
        for x, y in re.findall('(\d*)(\D*)', s):
            yield '', int(x or '0')
            yield y, 0

    def s(c):
        return list(split(c))

    regex_not_er = '(?<=_Ошибка).xlsx'
    regex_done = '№_\d+.*_Обработан\.xlsx'
    regex_not_prep = '(?<!_Обработан).xlsx'  # Попадание в список не обработанных, далее работаем только с ними.
    filtered_erorr = []
    for root, dirs, files in os.walk(mydirs_[0]):
        del dirs[:]  # go only one level deep
        filtered_erorr = [i for i in files if re.search(regex_not_er, str(i))]
    for root, dirs, files in os.walk(folder_url):
        del dirs[:]  # go only one level deep
        filtered_done = [i for i in files if re.search(regex_done, str(i))]  # Обработанные
        filtered_not_prep = [i for i in files if re.search(regex_not_prep, str(i))]  # Не обработанные
    filtered_erorr = sorted(filtered_erorr, key=s, reverse=True)
    filtered_done = sorted(filtered_done, key=s)
    print (f"Количество файлов не обработанных: {len(filtered_not_prep)}")
    if len(filtered_done + filtered_not_prep) == 0:
        raise Exception("Нет необходимых xlsx файлов в папке анкеты")
    lena = 1
    if len(filtered_not_prep) == 0:
        print ("Добавление записей не требуется!")
        worker.ReportProgress(100, "Выполнено.")
        return
    else:
        if len(filtered_done) > 0:
            lena = int((filtered_done[len(filtered_done) - 1]).split("_")[1]) + 1  # № последнего
            if filtered_erorr:
                lena_ = int((filtered_erorr[0]).split("_")[1]) + 1
                if lena < lena_:
                    lena = lena_
        for filename in filtered_not_prep:
            percentage = int((filtered_not_prep.index(filename) * 91) / len(filtered_not_prep))
            # print worker.WorkerReportsProgress
            # try:
            worker.ReportProgress(percentage, "Экспорт анкет.")
            if worker.CancellationPending == True:
                worker.ReportProgress(percentage, "Отмена задания")
                time.sleep(1)
                return
            # except Exception as e:
            #     print e.message
            path = mydirs_[6] + "\\" + filename
            # path = path.encode('utf8')
            dict_str = extract_txt_doc(path, folderName, mydirs_, lena,log,workstatuspanel)  # folderName - textboxBrowse.Text
            all_str_lst.append(dict_str)
            lena += 1
    error_files_count = 0
    error_count = 0
    for k, i in enumerate(all_str_lst):
        error_count += i[-1]
        if i[-1] > 0:
            error_index.append(k)
            error_files_count += 1
    print (f"Количество файлов с ошибками: {error_files_count}\nКоличество записей с ошибками: {error_count}")
    if error_count>0:
        log_file_rec(log, folderName)
    for index in sorted(error_index, reverse=True):
        del all_str_lst[index] #удаление строк с ошибками
    # Выделение строк в таблице по количеству строк из документов
    n_rows = table_output_xlsx.getRowsCount()  # количество строк в док перед вставкой
    rows_c = len(all_str_lst)  # количество записей
    # print ("n_rows,rows_c= ", n_rows ,rows_c
    A4_empty = table_output_xlsx.getCell("A4").getRawValue() == ''
    if rows_c < 1:
        pass
    elif rows_c > 0:
        if A4_empty and rows_c == 1:  # без вставки если пусто и 1 запись
            pass
        elif A4_empty and rows_c > 1:
            table_output_xlsx.insertRowAfter(n_rows - 1, copyRowStyle=True, rowsCount=rows_c - 1)  # вставка если пусто для 2 и более
        elif table_output_xlsx.getCell("A4").getRawValue() != '':
            table_output_xlsx.insertRowAfter(n_rows - 1, copyRowStyle=True, rowsCount=rows_c)  # вставка если не пусто
            n_rows = n_rows + 1
    number_rows = str(table_output_xlsx.getRowsCount())
    print (f"Количество строк в документе: {number_rows}")
    # Записываем результат в таблицу
    table_output_xlsx.getCell("C1").setFormattedValue(str(date_end)) #Дата проведения
    write_table(all_str_lst, worker, date_end, n_rows)  # Массив передастся в упорядоченном виде скорее всего
    set_cells_format(number_rows, worker,n_rows)
    # error_data(error_index, worker)
    worker.ReportProgress(99, "Сохранение XLSX.")
    try:
        document_xls.saveAs(mydirs_[11])
    except Exception as e:
        #workstatuspanel.Text = u'Ошибка открыт документ!'
        raise Exception("Открыт документ: "+ mydirs_[11])
    # raise Exception('This is the exception you expect to handle') #Аналог Throw для обработки исключений
    time.sleep(1)
    worker.ReportProgress(100, "Готово.")
