# -*- coding: utf-8 -*-
import os, time,itertools
from datetime import datetime
from MyOfficeSDKDocumentAPI import DocumentAPI as sdk
from string import ascii_uppercase


global application
application = sdk.Application()
cell_properties_win = sdk.CellProperties()
cell_properties_win.backgroundColor = sdk.ColorRGBA(193, 242, 17, 255)
cell_properties_lose = sdk.CellProperties()
cell_properties_lose.backgroundColor = sdk.ColorRGBA(108, 122, 137, 255)


def message(table_input, i,template,mydirs_):
    document = application.loadDocument(template)
    bookmarks = document.getBookmarks()
    last_name = table_input.getCell("B" + str(i)).getFormattedValue()
    first_name = table_input.getCell("C" + str(i)).getFormattedValue()
    bookmarks.getBookmarkRange('name').replaceText(last_name + ' ' + first_name)
    f_path = (mydirs_[17] + "\\" + '№' + str(i - 3) + ' ' + last_name + ' ' + first_name + ' ' + os.path.basename(
            template))
    document.saveAs((f_path))


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

def load_doc(mydirs_):
    input_file_url_xls = mydirs_  # Output с размещением системы рейтингования
    document_xls = application.loadDocument(input_file_url_xls)  # load Сводный
    table_output_xlsx = document_xls.getBlocks().getTable(0)  # table Сводный
    return table_output_xlsx,document_xls



def write_color_win(worker,mydirs_):
    template_win = mydirs_[16]
    template_lose = mydirs_[15]
    worker.ReportProgress(91, "Выделение победителей первого этапа")
    for i in range(1, 5):
        globals()['table_output_xlsx_%d' % i],globals()['document_xls_%d' % i] =load_doc(mydirs_[i+10])
    n_rows = table_output_xlsx_1.getRowsCount()
    for i in range(4, n_rows + 1):
        n_rows = str(i)
        cell_range = table_output_xlsx_1.getCellRange("A" + n_rows + ":AK" + n_rows)
        if float(table_output_xlsx_1.getCell("W" + n_rows).getFormattedValue()) >= 10:
            message(table_output_xlsx_1, i,template_win,mydirs_)
            for i in range(1, 5):
                globals()['cell_range_%s' % i] = globals()['table_output_xlsx_%s' % i].getCellRange(
                    "A" + n_rows + ":F" + n_rows)
                globals()['cell_range_%s' % i].setCellProperties(cell_properties_win)
            cell_range.setCellProperties(cell_properties_win)
        else:
            message(table_output_xlsx_1, i,template_lose,mydirs_)
            cell_range.setCellProperties(cell_properties_lose)
    for i in range(1, 5):
        try:
            globals()['document_xls_%s' % i].saveAs(mydirs_[i+10])
        except Exception as e:
            raise Exception("Открыт документ: " + os.path.basename(mydirs_[i+10]))
    worker.ReportProgress(100, "Завершено")

def write_scores(col,scores_,table_output_xlsx_main):
    n_rows = table_output_xlsx_main.getRowsCount()
    for i in range(4, n_rows + 1):
        table_output_xlsx_main.getCell(col + str(i)).setNumber(scores_[i - 4])
    #document_xls.saveAs(mydirs_)


def get_scores(table_output_xlsx,adr):
    scores_ = []
    n_rows = table_output_xlsx.getRowsCount()
    for i in range(4,n_rows+1):
        scores_.append(int(table_output_xlsx.getCell(str(adr)+str(i)).getFormattedValue()))
    return scores_




def main_score(worker,mydirs_,adr,k,adr_last,proc):
    table_output_xlsx_main, document_xls_main = load_doc(mydirs_[11])
    worker.ReportProgress(0, "Экспорт балов.")
    column = list_xls(str(adr_last))
    j = 0  # Смещение вправо
    for i in range(12, 15):
        table_output_xlsx,document_xls=load_doc(mydirs_[i])
        scores_ = get_scores(table_output_xlsx,adr)                  # Экспорт баллов
        write_scores(column[k + j], scores_, table_output_xlsx_main) #Сохранение в сводный по трем файлам, в цикле
        j += 1
        worker.ReportProgress(30*j, "Экспорт балов.")
    try:
        document_xls_main.saveAs(mydirs_[11])
    except Exception as e:
        raise Exception("Открыт документ: " + os.path.basename(mydirs_[11]))
    worker.ReportProgress(proc, "Экспорт завершен.")


