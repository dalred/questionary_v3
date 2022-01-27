
from MyOfficeSDKDocumentAPI import DocumentAPI as sdk
import inspect, os,time

application = sdk.Application()



def load_doc(mydirs_):
    input_file_url_xls = mydirs_  # Output с размещением системы рейтингования
    document_xls = application.loadDocument(input_file_url_xls)  # load Сводный
    table_output_xlsx = document_xls.getBlocks().getTable(0)  # table Сводный
    return table_output_xlsx,document_xls


def write_formules(row_str,fio_list,table_output2_xlsx,table_output_xlsx,mydirs_,worker):
    worker.ReportProgress(0, "Формирование файлов Жюри")
    current_row = 4 # c index 4
    for str_ in fio_list:
        if worker.CancellationPending == True:
            worker.ReportProgress(percentage, "Отмена задания")
            time.sleep(1)
            return
        index = current_row - 3
        percentage = int((index * 95) / len(fio_list))
        worker.ReportProgress(percentage, "Формирование файлов Жюри")
        row_str = str(current_row)
        table_output2_xlsx.getCell("B" + row_str).setText(str_)
        table_output2_xlsx.getCell("A" + row_str).setNumber(index)
        table_output2_xlsx.getCell("F" + row_str).setFormula("=SUM(C" + row_str + ":D" + row_str + ":E" + row_str + ")")
        table_output2_xlsx.getCell("K" + row_str).setFormula("=SUM(G" + row_str + ":H" + row_str + ":I" + row_str + ":J" + row_str + ")")
        table_output_xlsx.getCell("W" + row_str).setFormula("=AVERAGE(AC" + row_str + ":AE" + row_str + ")+SUM(" + "X" + row_str + ":AB" + row_str + ")")
        table_output_xlsx.getCell("AF" + row_str).setFormula("=AVERAGE(AG" + row_str + ":AI" + row_str + ")")
        table_output_xlsx.getCell("AJ" + row_str).setFormula("=SUM(AF" + row_str + ",W" + row_str + ")")
        current_row += 1


def main_judges(worker,mydirs_,workstatuspanel):
    worker.ReportProgress(0, "Формирование файлов Жюри")
    try:
        table_output_xlsx,document_xls =load_doc(mydirs_[11]) #Сводный
        table_output2_xlsx,document2_xls=load_doc(mydirs_[8]) #Жюри 1
    except Exception as e:
        raise Exception("Невозможно открыть документ: " + e)
    row_str = table_output_xlsx.getRowsCount()  # количество строк Сводный
    E4_empty = table_output_xlsx.getCell("E4").getRawValue() == ''

    if not E4_empty:
        cell_range = table_output_xlsx.getCellRange("E4:E" + str(row_str))
        fio_list = [i.getRawValue() for i in cell_range]

    # Выделение строк в таблице по количеству строк из документов
    rows_count_2 = len(fio_list)
    if rows_count_2 > 1:
        try:
            table_output2_xlsx.insertRowAfter(3, copyRowStyle=True, rowsCount=rows_count_2 - 1)
            document2_xls.saveAs(mydirs_[12])
        except Exception as e:
            raise Exception("Невозможно сохранить документ: " + mydirs_[12])
    elif rows_count_2 == 1:
        try:
            document2_xls.saveAs(mydirs_[12])
        except Exception as e:
            raise Exception("Невозможно сохранить документ: " + mydirs_[12])

    write_formules(row_str,fio_list,table_output2_xlsx,table_output_xlsx,mydirs_,worker)
    try:
        document_xls.saveAs(mydirs_[11])
        document2_xls.saveAs(mydirs_[12])
        table_output2_xlsx.getCell("B1").setText("Жюри 2")
        document2_xls.saveAs(mydirs_[13])
        table_output2_xlsx.getCell("B1").setText("Жюри 3")
        document2_xls.saveAs(mydirs_[14])
    except Exception as e:
        raise Exception("Невозможно сохранить документ: Жюри ")
    worker.ReportProgress(100, "Завершено формирование файлов Жюри")



