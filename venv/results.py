# -*- coding: utf-8 -*-
import os, sys,re,inspect
from MyOfficeSDKDocumentAPI import DocumentAPI as sdk



def main_results(worker, mydirs_):
    def load_doc(mydirs_):
        input_file_url_xls = mydirs_  # Output с размещением системы рейтингования
        document_xls = application.loadDocument(input_file_url_xls)  # load Сводный
        table_output_xlsx = document_xls.getBlocks().getTable(0)  # table Сводный
        return table_output_xlsx, document_xls

    def message(table_input, i, mydirs_, int_status):
        last_name = table_input.getCell("B" + str(i)).getRawValue()
        first_name = table_input.getCell("C" + str(i)).getRawValue()
        if int_status==1: #Победитель
            cast=table_input.getCell("AL" + str(i)).getRawValue()
            date_begin = table_input.getCell("AM" + str(i)).getRawValue()
            date_last = table_input.getCell("AN" + str(i)).getRawValue()
            date = table_input.getCell("AO" + str(i)).getRawValue()
            bookmarks_win.getBookmarkRange('name').replaceText(last_name + ' ' + first_name)
            bookmarks_win.getBookmarkRange('cast').replaceText(cast)
            bookmarks_win.getBookmarkRange('date_begin').replaceText(date_begin)
            bookmarks_win.getBookmarkRange('date_last').replaceText(date_last)
            bookmarks_win.getBookmarkRange('date').replaceText(date)
            f_path = (mydirs_[1] + "\\" + '№' + str(i - 3) + ' ' + last_name + ' ' + first_name + ' ' + os.path.basename(
                mydirs_[20]))
            document_win.saveAs((f_path))
        elif int_status == 2: #Не прошел
            scores = table_input.getCell("AJ" + str(i)).getFormattedValue()
            bookmarks_lose.getBookmarkRange('name').replaceText(last_name + ' ' + first_name)
            bookmarks_lose.getBookmarkRange('scores').replaceText(scores)
            f_path = (mydirs_[1] + "\\" + '№' + str(i - 3) + ' ' + last_name + ' ' + first_name + ' ' + os.path.basename(
                mydirs_[21]))
            document_lose.saveAs((f_path))
        elif int_status == 3: #Резерв
            scores = table_input.getCell("AJ" + str(i)).getFormattedValue()
            number_pos = table_input.getCell("AP" + str(i)).getRawValue()
            date_ = table_input.getCell("AQ" + str(i)).getRawValue()
            bookmarks_reserve.getBookmarkRange('name').replaceText(last_name + ' ' + first_name)
            bookmarks_reserve.getBookmarkRange('scores').replaceText(scores)
            bookmarks_reserve.getBookmarkRange('number').replaceText(number_pos)
            bookmarks_reserve.getBookmarkRange('date').replaceText(date_)
            f_path = (mydirs_[1] + "\\" + '№' + str(i - 3) + ' ' + last_name + ' ' + first_name + ' ' + os.path.basename(
                mydirs_[19]))
            document_reserve.saveAs((f_path))


    def paper_win(table_input, i,mydirs_):
        last_name = table_input.getCell("B" + str(i)).getRawValue()
        first_name = table_input.getCell("C" + str(i)).getRawValue()
        middle_name = table_input.getCell("D" + str(i)).getRawValue()
        #Округление на всякий случай
        scores = str(round(float((table_input.getCell("AJ" + str(i)).getFormattedValue()))))
        bookmarks_paper_win.getBookmarkRange('Last_name').replaceText(last_name)
        bookmarks_paper_win.getBookmarkRange('First_middle_name').replaceText(first_name + ' ' + middle_name)
        bookmarks_paper_win.getBookmarkRange('scores').replaceText(scores)
        output_file = mydirs_[2] + "\\" + '№' + str(i - 3) + ' ' + last_name + ' ' + first_name + '.pdf'
        document_paper_win.exportAs(str(output_file), sdk.ExportFormat_PDFA1)

    application = sdk.Application()
    template_win_doc = mydirs_[4]
    template_lose_doc = mydirs_[5]
    template_paper_win_doc = mydirs_[3]
    template_reserve_doc = mydirs_[18]
    document_win = application.loadDocument(template_win_doc)
    bookmarks_win = document_win.getBookmarks()
    document_lose = application.loadDocument(template_lose_doc)
    bookmarks_lose = document_lose.getBookmarks()
    document_paper_win = application.loadDocument(template_paper_win_doc)
    bookmarks_paper_win = document_paper_win.getBookmarks()
    document_reserve=application.loadDocument(template_reserve_doc)
    bookmarks_reserve=document_reserve.getBookmarks()

    table_output_xlsx_main, document_xls_main = load_doc(mydirs_[11])
    n_rows_count = table_output_xlsx_main.getRowsCount()

    k = 0
    for i in range(4, n_rows_count + 1):
        k += 1
        percentage = int((k * 100) / (n_rows_count - 3))
        n_rows = str(i)
        status=table_output_xlsx_main.getCell("AK" + n_rows).getRawValue()
        all_scores=float(table_output_xlsx_main.getCell("W" + n_rows).getFormattedValue())
        worker.ReportProgress(percentage, u"Формирование грамот и писем.")
        if worker.CancellationPending == True:
            worker.ReportProgress(percentage, u"Отмена задания")
            time.sleep(1)
            return
        if all_scores >= 10 and re.search('[Пп]обедитель',status):
            int_status=1
            message(table_output_xlsx_main, i, mydirs_, int_status)
            paper_win(table_output_xlsx_main, i,mydirs_)
        elif all_scores >= 10 and re.search('[Нн]е прошел',status):
            int_status = 2
            message(table_output_xlsx_main, i, mydirs_, int_status)
        elif re.search('[Рр]езерв',status):
            int_status = 3
            message(table_output_xlsx_main, i, mydirs_, int_status)



