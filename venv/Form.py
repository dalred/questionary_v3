import multiprocessing
import os, sys, time, clr, datetime, shutil, threading
import tkinter.filedialog
from first_step import main_
from results import main_results
from formation_judges import main_judges
from Scores import main_score,  write_color_win
clr.AddReference('System')
from System import DateTime as NetDateTime
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Windows.Forms import *
from System.Drawing import *
from System.ComponentModel import BackgroundWorker
from System.Diagnostics import Process

formConvert = Form()

workstatuspanel = StatusBarPanel()
splashForm = Form()

progressbar1 = ProgressBar()
combobox1 = ComboBox()
textboxBrowse = TextBox()
start = Button()
canceling = Button()
worker = BackgroundWorker()
datetimepicker1=DateTimePicker()
open = Button()

worker.WorkerReportsProgress = True
worker.WorkerSupportsCancellation = True

path_dir_root = os.getcwd()
path_dirname_=path_dir_root+'\\'+'templates'
textboxBrowse.Text = path_dir_root
mydir=textboxBrowse.Text+'\\'+u'Обработка Анкет'

state_dir = True

mydir0 = mydir + "\\" + u"Ошибки"
mydir1 = mydir + "\\" + u"Результаты Финал"
mydir2 = mydir + "\\" + u"Грамоты"
mydir3 = path_dirname_ + "\\" + u"Пример Диплома_27072020.docx"
mydir4 = path_dirname_ + "\\" + u"3.1. информац. сообщение - победитель.docx"
mydir5 = path_dirname_ + "\\" + u"3.3. информац. сообщение - не прошел.docx"
mydir6 = os.path.dirname(os.path.abspath(mydir)) + "\\" + u"Анкеты"
mydir7 = path_dirname_ + "\\" + u"Файл с размещением системы рейтингования _версия 02.09.2020.xlsx"
mydir8= path_dirname_ + "\\" + u"Артек_Программа _Дверь синего цвета_ (жюри 1).xlsx"
mydir9 = path_dirname_ + "\\" + u"Артек_Программа _Дверь синего цвета_ (жюри 2).xlsx"
mydir10 = path_dirname_ + "\\" + u"Артек_Программа _Дверь синего цвета_ (жюри 3).xlsx"
mydir11 = mydir + "\\"+ u"Файл с размещением системы рейтингования.xlsx"

mydir12_out = mydir + "\\" + u"Артек_Программа _Дверь синего цвета_ (жюри 1)_результаты.xlsx"
mydir13_out = mydir + "\\" + u"Артек_Программа _Дверь синего цвета_ (жюри 2)_результаты.xlsx"
mydir14_out = mydir + "\\" + u"Артек_Программа _Дверь синего цвета_ (жюри 3)_результаты.xlsx"
mydir15_out = path_dirname_ + "\\" + u"По итогам 1 этапа - не прошел.docx"
mydir16_out = path_dirname_ + "\\" + u"По итогам 1 этапа - прошел.docx"
mydir17_out = mydir + "\\" + u"Результаты первого этапа"
mydir18 = path_dirname_ + "\\" + u"3.2. информац. сообщение - резерв.docx"
mydir19_out = mydir1 + "\\"+os.path.basename(mydir18)
mydir20_out = mydir1+'\\'+os.path.basename(mydir4)
mydir21_out = mydir1+'\\'+os.path.basename(mydir5)
mydirs_ = [mydir0, mydir1, mydir2, mydir3, mydir4, mydir5, mydir6, mydir7,mydir8,mydir9,mydir10, mydir11,mydir12_out,mydir13_out,mydir14_out,mydir15_out,mydir16_out,mydir17_out,mydir18,mydir19_out,mydir20_out,mydir21_out]

def Cancel_(sender, event):
    worker.CancelAsync()

def message_box(title,Message):
    MessageBox.Show(Message, title, 0, MessageBoxIcon.Error)

def message_warn(title, Message):
    result = MessageBox.Show(Message, title, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)
    return result

def do_work(sender, event):
    time.sleep(1)
    formConvert.Focus()
    """for k,i in enumerate(mydirs_):
        print k,i"""
    date_end=datetimepicker1.Value.ToString("dd.MM.yyyy")
    formConvert.Text = str('0%')
    if not (os.path.exists(mydirs_[6])):
        print(f"Отсутствует папка: {os.path.abspath(mydirs_[6])}")
        workstatuspanel.Text=u'Отсутствует папка:  ' + os.path.abspath(mydirs_[6])
        raise Exception("Отсутствует папка:  " + os.path.abspath(mydirs_[6]))
    for i in range(3, 11):
        if not (os.path.exists(mydirs_[i])):
            workstatuspanel.Text = u'Отсутствует: ' + os.path.abspath(mydirs_[i])
            raise Exception("Отсутствует:  " + os.path.abspath(mydirs_[i])) #Папку Анкеты тоже проверяем.
    if combobox1.SelectedIndex == 0:
        if not (os.path.exists(mydir)):
            try:
                os.mkdir(mydir, 0o777)  # Папка Обработка Анкет
                workstatuspanel.Text = u'Создана папка: ' + os.path.basename(mydir)
            except Exception:
                raise Exception("Неудалось создать папку:" + os.path.basename(mydir))
        if os.path.exists(mydirs_[11]):
            #6 - да - 7 нет
            dialog_result = message_warn('Добавить?','Вы хотите добавить информацию в : ' + os.path.basename(mydirs_[11]))
            if dialog_result==7:
                sender.CancelAsync()
                sender.Dispose()
                return
        else:
            shutil.copyfile(mydirs_[7], mydirs_[11])
        main_(sender, mydir,mydirs_,str(date_end),workstatuspanel)
    elif combobox1.SelectedIndex == 1:
        if not (os.path.exists(mydir)):
            raise Exception("Неудалось найти папку: " + os.path.basename(mydir)) #Папку Обработка Анкет проверяем
        for i in range(12, 15):
            if os.path.exists(mydirs_[i]):
                dialog_result = message_warn('Перзаписать?', 'Вы хотите перезаписать информацию в : ' + os.path.basename(
                    mydirs_[i]))
                if dialog_result==7:
                    sender.CancelAsync()
                    sender.Dispose()
                    return
        for i in range(8, 12):
            if not (os.path.exists(mydirs_[i])): #проверка шаблонов
                workstatuspanel.Text = u'Отсутствует: ' + os.path.abspath(mydirs_[i])
                raise Exception("Отсутствует:  " + os.path.abspath(mydirs_[i]))
        main_judges(sender, mydirs_,workstatuspanel)
    elif combobox1.SelectedIndex == 2:
        if not (os.path.exists(mydirs_[17])):
            try:
                os.mkdir(mydirs_[17], 0o777)  # Папка Результаты I-й этап
                workstatuspanel.Text = u'Создана папка : ' + os.path.basename(mydirs_[17])
            except Exception:
                raise Exception("Неудалось создать папку:" + os.path.basename(mydirs_[17]))
        for i in range(11, 15):
            if not (os.path.exists(mydirs_[i])):
                workstatuspanel.Text = u'Отсутствует:  ' + os.path.abspath(mydirs_[i])
                raise Exception("Отсутствует:  " + os.path.abspath(mydirs_[i]))  # Жюри1,2,3 Сводный.
        main_score(sender,mydirs_,"F",28,"AE",90)  #k = 28 AС до АЕ F из жюри k=32 AG
        time.sleep(0.1)
        write_color_win(sender, mydirs_)
    elif combobox1.SelectedIndex == 3:
        for i in range(11, 15):
            if not (os.path.exists(mydirs_[i])):
                workstatuspanel.Text = u'Отсутствует:  ' + os.path.abspath(mydirs_[i])
                raise Exception("Отсутствует:  " + os.path.abspath(mydirs_[i]))  # Жюри1,2,3 Сводный.
        main_score(sender, mydirs_, "K", 32, "AI",100)
    elif combobox1.SelectedIndex == 4:
        for i in range(1, 3):
            if not (os.path.exists(mydirs_[i])):
                try:
                    os.mkdir(mydirs_[i], 0o777)  # Папка Результаты,Грамоты
                    workstatuspanel.Text = u'Создана папка : ' + os.path.basename(mydirs_[i])
                except Exception:
                    raise Exception("Неудалось создать папку:" + os.path.basename(mydirs_[i]))
        for i in range(11, 15):
            if not (os.path.exists(mydirs_[i])):
                workstatuspanel.Text = u'Отсутствует:  ' + os.path.abspath(mydirs_[i])
                raise Exception("Отсутствует:  " + os.path.abspath(mydirs_[i]))  # Жюри1,2,3 Сводный.
        """t1 = threading.Thread(target=show_form, )
        t1.setDaemon(True)
        t1.start()"""
        main_results(worker,mydirs_)

def begin_dfile(sender, event):
    start.Enabled = False
    #???foldername = textboxBrowse.Text
    if textboxBrowse.Text == 'folder not specified':
        message_box("Предупреждение!", 'Папка не задана!')
    elif combobox1.SelectedIndex == 0 and state_dir is True:
        worker.RunWorkerAsync()
        Application.UseWaitCursor = True
    elif combobox1.SelectedIndex == 1 and state_dir is True:
        worker.RunWorkerAsync()
        Application.UseWaitCursor = True
    elif combobox1.SelectedIndex == 2 and state_dir is True:
        worker.RunWorkerAsync()
        Application.UseWaitCursor = True
    elif combobox1.SelectedIndex == 3 and state_dir is True:
        worker.RunWorkerAsync()
        Application.UseWaitCursor = True
    elif combobox1.SelectedIndex == 4 and state_dir is True:
        worker.RunWorkerAsync()
        Application.UseWaitCursor = True

def bgWorker_ProgressChanged(sender, event):
    start.Enabled = False
    #workstatuspanel.Text = u'Выполняется задача: ' + combobox1.SelectedItem
    formConvert.Text = str(event.ProgressPercentage) + u"%, " + event.UserState
    progressbar1.Value = event.ProgressPercentage
    if progressbar1.Value==93:
        canceling.Enabled=False

def final(sender,event):
    if event.Error != None:
        message_box('Ошибка!', event.Error.Message)
    print("RunWorkerCompleted")
    canceling.Enabled = True
    Application.UseWaitCursor = False
    Cursor.Current = Cursors.Default
    formConvert.Text = 'Задача завершена!'
    sender.Dispose()
    time.sleep(1)
    progressbar1.Value = 0
    workstatuspanel.Text = u'Задача: ' + combobox1.SelectedItem + u' завершена!'
    formConvert.Focus()
    formConvert.Activate()
    start.Enabled = True

# def message_box(title,Message):
#     root = tkinter.Tk()
#     root.withdraw()
#     root.iconbitmap(path_dir_root + "\\images\\" + "folder.ico")
#     root.attributes("-topmost", True)
#     tkinter.messagebox.showerror(title, Message)
#     formConvert.Focus()


def click_open(sender,event):
    if state_dir:
        Process.Start("explorer.exe", textboxBrowse.Text)

def show_dialog(sender, event):
    folderName=FileDialog('Пожалуйста укажите корневой каталог! ')
    if folderName:
        open.Enabled = True
        global state_dir
        state_dir = True
        start.Enabled = True
        global mydirs_,mydir
        textboxBrowse.Text = folderName.replace("/","\\")
        for i in (0, 1, 2, 6, 11, 12, 13, 14, 17,19,20,21):
            mydirs_[i] = mydirs_[i].replace(os.path.dirname(os.path.abspath(mydir)), textboxBrowse.Text)
        #path_dirname_ и Mydir совпасть может поэтому вот такое перечисление
        #mydirs_#Без глобал невозможно сделать присвоение статическому полю, так как ты его до этого не объявил.
                            # Использовать можно, но изменять нет.
        #for k,i in enumerate(mydirs_):
        #    print k,i
        mydir = textboxBrowse.Text+'\\'+u'Обработка Анкет'
    else:
        textboxBrowse.Text = 'folder not specified'
        state_dir = False
        open.Enabled = False
    formConvert.Focus()

worker.DoWork += do_work
worker.ProgressChanged += bgWorker_ProgressChanged
worker.RunWorkerCompleted += final


def FileDialog(title):
    root = tkinter.Tk()
    root.withdraw()
    root.iconbitmap(path_dir_root+"\\images\\" +"folder.ico")
    root.attributes("-topmost", True)
    folderName = tkinter.filedialog.askdirectory(title=title)
    return folderName

def splash_form():
    picturebox1 = PictureBox()
    picturebox1.Padding = Padding(1)
    picturebox1.BorderStyle = 0
    picturebox1.Image = Image.FromFile(path_dir_root+"\\images\\" +'Document-icon.jpg')
    picturebox1.Name = 'picturebox1'
    #picturebox1.ClientSize = Size(335, 187)
    picturebox1.Size = Size(503,280)
    picturebox1.SizeMode = PictureBoxSizeMode.StretchImage
    picturebox1.TabIndex = 0
    picturebox1.TabStop = False
    splashForm.Text = 'Загрузка модулей...'
    splashForm.FormBorderStyle = 0
    splashForm.Size = Size(502, 279)
    splashForm.StartPosition = FormStartPosition.CenterScreen
    splashForm.TopMost = False
    splashForm.MinimizeBox = False
    splashForm.MaximizeBox = False
    splashForm.Controls.Add(picturebox1)
    Application.Run(splashForm)

def show_form():
    formConvert.StartPosition = FormStartPosition.CenterScreen
    formConvert.ClientSize = Size(452, 276)
    formConvert.FormBorderStyle = FormBorderStyle.FixedSingle
    formConvert.Name = 'formConvert'
    formConvert.BackColor = SystemColors.ButtonFace
    formConvert.Text = u'Форма для конвертации'
    formConvert.MaximizeBox = False;
    formConvert.Icon = Icon(path_dir_root+"\\images\\" +"folder.ico")
    #formConvert.Shown += event_show
    #
    # start
    #
    #
    start.Location = Point(12, 210)
    start.Name = 'start'
    start.Size = Size(110, 30)
    start.TabIndex = 0
    start.Text = 'Start'
    start.Click += begin_dfile
    start.UseCompatibleTextRendering = True
    start.UseVisualStyleBackColor = True
    #
    ## Cancel
    canceling.Location = Point(333, 210)
    canceling.Name = 'canceling'
    canceling.Size = Size(110, 30)
    canceling.TabIndex = 0
    canceling.Text = u'Отмена'
    canceling.UseCompatibleTextRendering = True
    canceling.UseVisualStyleBackColor = True
    canceling.Click += Cancel_
    #
    #open
    open.ImageAlign=ContentAlignment.MiddleCenter
    open.Location = Point(95, 79)
    open.Name = 'open'
    open.Size = Size(32, 27)
    open.TabIndex = 12
    open.UseCompatibleTextRendering = True
    open.UseVisualStyleBackColor = True
    open.Image = Image.FromFile(path_dir_root+"\\images\\" +"Open-folder-full.png")
    open.Click+=click_open
    #
    #buttonbrowse
    buttonbrowse = Button()
    buttonbrowse.Location = Point(12, 79)
    buttonbrowse.Name = 'buttonbrowse'
    buttonbrowse.Size = Size(77, 27)
    buttonbrowse.TabIndex = 5
    buttonbrowse.Text = u'Обзор'
    buttonbrowse.Click += show_dialog
    buttonbrowse.UseCompatibleTextRendering = True
    buttonbrowse.UseVisualStyleBackColor = True
    #
    # ProgressBar
    #
    #
    progressbar1.Location = Point(12, 170)
    progressbar1.Name = 'progressbar1'
    progressbar1.Size = Size(431, 34)
    progressbar1.Step = 1
    progressbar1.TabIndex = 1
    progressbar1.Value = 0
    progressbar1.ForeColor = Color.Green
    progressbar1.Style = ProgressBarStyle.Continuous

    #
    # combobox1
    #
    combobox1.FormattingEnabled = True
    combobox1.Items.Add(u'1. Обработка Анкет')
    combobox1.Items.Add(u'2.1 Формирование списков Жюри')
    combobox1.Items.Add(u'2.2 Формирование результатов 1-й этап')
    combobox1.Items.Add(u'2.3 Формирование результатов 2-й этап')
    combobox1.Items.Add(u'3.1 Формирование грамот и писем')
    combobox1.Location = Point(228, 143)
    combobox1.Name = 'combobox1'
    combobox1.Size = Size(215, 21)
    combobox1.TabIndex = 2
    combobox1.SelectedIndex=0
    #combobox1.SelectedIndexChanged+=ComboBox1_SelectedIndexChanged
    combobox1.DropDownStyle = ComboBoxStyle.DropDownList
    #combobox1.SelectedIndexChanged += SelectedIndexChanged
    #
    #label Этап
    label = Label()
    label.Location = Point(12, 142)
    label.Name = 'label'
    label.Size = Size(202, 22)
    label.TabIndex = 3
    label.Text = u'Этап обработки'
    label.TextAlign = ContentAlignment.MiddleLeft
    label.UseCompatibleTextRendering = True
    #
    #
    #Путь к размещению файлов
    label1 = Label()
    label1.Location = Point(12, 45)
    label1.Name = 'label1'
    label1.Size = Size(150, 31)
    label1.TabIndex = 4
    label1.Text = u'Путь к корневому каталогу'
    label1.TextAlign = ContentAlignment.MiddleLeft
    label1.UseCompatibleTextRendering = True
    #
    labeldata = Label()
    labeldata.Location = Point(12, 109)
    labeldata.Name = 'labeldata'
    labeldata.Size = Size(185, 20)
    labeldata.TabIndex = 8
    labeldata.Text = u'Дата проведения конкурса:'
    labeldata.TextAlign = ContentAlignment.MiddleLeft
    labeldata.UseCompatibleTextRendering = True
    # TextBox Путь
    textboxBrowse.Location = Point(133, 79)
    textboxBrowse.Name = 'textboxBrowse'
    textboxBrowse.Size = Size(310, 27)
    textboxBrowse.Multiline = True
    textboxBrowse.TabIndex = 6
    textboxBrowse.ReadOnly = True
    #
    #picturebox
    picturebox1=PictureBox()
    picturebox1.Location = Point(368, 4)
    picturebox1.Name = 'picturebox1'
    picturebox1.Size = Size(75, 72)
    picturebox1.SizeMode = PictureBoxSizeMode.CenterImage
    picturebox1.TabIndex = 9
    picturebox1.TabStop = False
    picturebox1.Image = Image.FromFile(path_dir_root+"\\images\\" +"post_logo.png")
    #
    #
    #datetimepicker1
    datetimepicker1.Format = DateTimePickerFormat.Short
    datetimepicker1.Location = Point(260, 117)
    datetimepicker1.Name = 'datetimepicker1'
    datetimepicker1.Size = Size(183, 20)
    datetimepicker1.TabIndex = 10
    datetimepicker1.Value=NetDateTime(2022, 0o3, 0o1)
    #statusbarpanel
    statusbar1 = StatusBar()
    statuspanel = StatusBarPanel()
    statusbar1.Location = Point(0, 264)
    statusbar1.Name = 'statusbar1'
    statusbar1.Size = Size(452, 22)
    statusbar1.Text = 'statusbar1'
    statusbar1.ShowPanels = True
    statusbar1.Panels.Add(statuspanel)
    statusbar1.Panels.Add(workstatuspanel)
    statuspanel.Text = u'Статус:'
    statuspanel.Alignment = HorizontalAlignment.Center
    workstatuspanel.Text = u'Здесь пишется статус задания'
    workstatuspanel.Width = 402
    statuspanel.Width = 50
    #
    # ControlsAdd

    formConvert.Controls.Add(progressbar1)
    formConvert.Controls.Add(canceling)
    formConvert.Controls.Add(start)
    formConvert.Controls.Add(combobox1)
    formConvert.Controls.Add(label)
    formConvert.Controls.Add(label1)
    formConvert.Controls.Add(buttonbrowse)
    formConvert.Controls.Add(textboxBrowse)
    formConvert.Controls.Add(labeldata)
    formConvert.Controls.Add(datetimepicker1)
    formConvert.Controls.Add(picturebox1)
    formConvert.Controls.Add(open)
    formConvert.Controls.Add(statusbar1)
    Application.Run(formConvert)


if __name__ == '__main__':
    multiprocessing.freeze_support()
    t1 = multiprocessing.Process(target = splash_form)
    t1.start()
    time.sleep(4)
    t1.terminate()
    t2 = threading.Thread(target=show_form,)
    t2.start()