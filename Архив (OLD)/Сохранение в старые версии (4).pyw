#-------------------------------------------------------------------------------
# Name:        Сохранение в старые версии
# version:     v1.1.0.0
#
# Author:      dimak222
#
# Created:     21.01.2022
# Copyright:   (c) dimak222 2022
# Licence:     No
#-------------------------------------------------------------------------------

title = "Сохранение в старые версии"

def Resource_path(relative_path): # для сохранения картинки внутри exe файла

    import os # работа с файовой системой

    try: # попытаться определить путь к папке
        base_path = sys._MEIPASS # путь к временной папки PyInstaller

    except Exception: # если ошибка
        base_path = os.path.abspath(".") # абсолютный путь

    return os.path.join(base_path, relative_path) # объеденяем и возващаем полный путь

def DoubleExe():# проверка на уже запущеное приложение, с отключённым консольным окном "CREATE_NO_WINDOW"

    import subprocess # модуль вывода запущеных процессов
    from sys import exit # для выхода из приложения без ошибки

    CREATE_NO_WINDOW = 0x08000000 # отключённое консольное окно
    processes = subprocess.Popen('tasklist', stdin=subprocess.PIPE, stderr=subprocess.PIPE, stdout=subprocess.PIPE, creationflags=CREATE_NO_WINDOW).communicate()[0] # список всех процессов
    processes = processes.decode('cp866') # декодировка списка

    if str(processes).count(title[0:25]) > 2: # если найдено название программы (два процесса) с ограничением в 25 символов
        Message("Приложение уже запущено!") # сообщение, поверх всех окон и с автоматическим закрытием
        exit() # выходим из програмы

def KompasAPI(): # подключение API компаса

    from win32com.client import Dispatch, gencache # библиотека API Windows
    from sys import exit # для выхода из приложения без ошибки

    try: # попытаться подключиться к КОМПАСу

        global KompasAPI7 # значение делаем глобальным
        global iApplication # значение делаем глобальным
        global iKompasObject # значение делаем глобальным

        KompasConst3D = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0).constants # константа 3D документов
        KompasConst2D = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants # константа 2D документов
        KompasConst = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants # константа для скрытия вопросов перестроения

        KompasAPI5 = gencache.EnsureModule('{0422828C-F174-495E-AC5D-D31014DBBE87}', 0, 1, 0) # API5 КОМПАСа
        iKompasObject = Dispatch('Kompas.Application.5', None, KompasAPI5.KompasObject.CLSID) # интерфейс API КОМПАС

        KompasAPI7 = gencache.EnsureModule('{69AC2981-37C0-4379-84FD-5DD2F3C0A520}', 0, 1, 0) # API7 КОМПАСа
        iApplication = Dispatch('Kompas.Application.7') # интерфейс приложения КОМПАС-3D.

        iKompasDocument = iApplication.ActiveDocument # делаем активный открытый документ

        if iApplication.Visible == False: # если компас невидимый
            iApplication.Visible = True # сделать КОМПАС-3D видемым

    except: # если не получилось выдать сообщение

        message("КОМПАС-3D не найден!\nУстановите или переустановите КОМПАС-3D!") # сообщение, поверх всех окон с автоматическим закрытием
        exit() # выходим из програмы

def Kompas_message(text): # сообщение в окне КОМПАСа если он открыт

    if iApplication.Visible == True: # если компас видимый
        iApplication.MessageBoxEx(text, 'Message:', 64) # сообщение в КОМПАСе

def Message(text = "Ошибка!", counter = 4): # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)

    from threading import Thread # библиотека потоков
    import time # модуль времени

    def Message_Thread(text, counter): # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)

        import tkinter.messagebox as mb # окно с сообщением
        import tkinter as tk # модуль окон

        if counter == 0: # время до закрытия окна (если 0)
            counter = 1 # закрытие через 1 сек
        window_msg = tk.Tk() # создание окна
        window_msg.iconbitmap(default = Resource_path("cat.ico")) # значёк программы
        window_msg.attributes("-topmost",True) # окно поверх всех окон
        window_msg.withdraw() # скрываем окно "невидимое"
        time = counter * 1000 # время в милисекундах
        window_msg.after(time, window_msg.destroy) # закрытие окна через n милисекунд
        if mb.showinfo(title, text, parent = window_msg) == "": # информационное окно закрытое по времени
            pass
        else:
            window_msg.destroy() # окно закрыто по кнопке
        window_msg.mainloop() # отображение окна

    msg_th = Thread(target = Message_Thread, args = (text, counter)) # запуск определения положения мышки в отдельном потоке
    msg_th.start() # запуск потока

    th.join() # ждать завершения процесса, иначе может закрыться следующие окно

##    while msg_th.is_alive(): # ждать запущеный поток
##        time.sleep(0.5) # ждём n сек

def Askyesnocancel(text): # вопросительное сообщение, поверх всех окон

    import tkinter.messagebox as mb # окно с сообщением
    import tkinter as tk # модуль окон
    from sys import exit # для выхода из приложения без ошибки

    ask = tk.Tk() # создание окна
    ask.iconbitmap(default = Resource_path("cat.ico")) # значёк программы
    ask.attributes("-topmost",True) # окно поверх всех окон
    ask.withdraw() # скрываем окно "невидимое"
    window = mb.askyesnocancel(title, text) # задаём вопрос
    ask.destroy() # закрываем окно
    ask = ask.mainloop() # отображение окна

    if window == True or False: # если ответ "Да" или "Нет"
        return window # возвращаем результат вопроса

    elif window == None: # если нажали отмена или крестик
        if iApplication.Visible == False: # если компас невидимый
            iApplication.Quit() # закрываем его
        exit() # выходим из програмы

def Askyesno(text): # вопросительное сообщение, поверх всех окон

    import tkinter.messagebox as mb # окно с сообщением
    import tkinter as tk # модуль окон

    ask = tk.Tk() # создание окна
    ask.iconbitmap(default = Resource_path("cat.ico")) # значёк программы
    ask.attributes("-topmost",True) # окно поверх всех окон
    ask.withdraw() # скрываем окно "невидимое"
    ask_mb = mb.askyesno(title, text) # задаём вопрос
    ask.destroy() # закрываем окно
    ask.mainloop() # отображение окна

    return ask_mb # возвращаем результат вопроса

def Message_count(all_failes_number, msg = "Идёт обработка файлов!"): # выдача сообщений о количестве файлов (количество всех файлов, сообщение) + file_number (номер обрабатываемого файла) + current_file_name (текущее название файла)

    from threading import Thread # библиотека потоков

    global Stop # глобальный параметр остановки сообщения

    def Message_count_Thread(all_failes_number, msg): # сообщений о количестве файлов в потоке

        import tkinter as tk # модуль окон
        import tkinter.ttk as ttk # модуль окон
        import time # модуль времени

        global Stop # глобальный параметр остановки обработки

        class ToolTip(object): # отображает подсказку к виджету

        	def __init__(self, widget, text):
        		self.widget = widget
        		self.text = text
        		self.acid = None
        		self.tipwindow = None
        		self.widget.bind('<Enter>', self.enter)
        		self.widget.bind('<Leave>', self.leave)
        		self.widget.bind('<ButtonRelease>', self.leave)
        		self.widget.bind('<Key>', self.leave)

        	def enter(self, event):
        		self.schedule()

        	def leave(self, event):
        		self.unschedule()
        		self.hidetip()

        	def schedule(self):
        		self.unschedule()
        		self.acid = self.widget.after(300, self.showtip) # через сколько милисунд отображать подсказку

        	def unschedule(self):
        		idac = self.acid
        		if idac:
        			self.widget.after_cancel(idac)
        		self.acid = None

        	def showtip(self):
        		tw = self.tipwindow = tk.Toplevel(self.widget)
        		tw.wm_overrideredirect(1)
        		tw.wm_attributes('-topmost', 1) # поверх всех окон
        		tw.wm_geometry('+%d+%d' % (self.widget.winfo_rootx(), self.widget.winfo_rooty() + self.widget.winfo_height() + 2))
        		tk.Label(tw, text = current_file_name, justify = 'left', bg = '#f2f2f2', relief = 'solid', bd = 1, font = "Verdana 10").pack() # положение, цвет и шрифт текста

        	def hidetip(self):
        		tw = self.tipwindow
        		if tw:
        			tw.destroy()
        		self.tipwindow = None

        def Update_text(): # обновление отчёта цифр

            def Updating_text(): # обновление текста

                if Stop: # если триггер остановки обработки и сообщения включён
                    print("Остановил поток!")
                    window.destroy() # закрываем окно

                else: # триггер выключен
                    print("Обновление текста!")
                    text.config(text = str(file_number) + "/" + str(all_failes_number)) # обновляем текст
                    text.after(300, Updating_text) # через милисекунды запускаем функцию заново

            Updating_text() # обновление текста

        def Update_progress(): # обновление прогресса

            def Updating_progress(): # обновление прогресса

                percent_file_number = percent_all_failes_number * file_number # процент выполнения

                if Stop: # если триггер остановки обработки и сообщения включён
                    print("Остановил поток прогресса!")
                    window.destroy() # закрываем окно

                else: # триггер выключен
                    print("Обновление прогресса!")
                    progress['value'] = percent_file_number # процент выполнения
                    window.update() # (update_idletasks не сбрасывет дпока не дошёл до конца)
                    progress.after(300, Updating_progress) # через милисекунды запускаем функцию заново

            Updating_progress() # обновление прогресса

        def Button_exit(): # кнопка "Отмена"
            window.destroy() # закрываем окно

        window = tk.Tk() # создание окна
        window.iconbitmap(default = Resource_path("cat.ico")) # значёк программы
        window.title(title) # заголовок окна
        window.attributes("-topmost",True) # окно поверх всех окон
        x = (window.winfo_screenwidth() - window.winfo_reqwidth()) / 2 # положение по центру монитора
        y = (window.winfo_screenheight() - window.winfo_reqheight()) / 2 # положение по центру монитора
        window.wm_geometry("+%d+%d" % (x-50, y)) # положение по центру монитора -50 из-за логотипа
##        window.geometry('200x100') # размер окна
        window.resizable(width = False, height = False) # блокировка изменение размера окна

##        logo = tk.PhotoImage(file = Resource_path("cat.png")) # логотип
##        logo = logo.subsample(1, 1) # мастаб картинки
##        tk.Label(window, image=logo).pack(side="right") # расположение картинки в окне

        f_top = tk.Frame(window) # блок окна (вверх)
        f_top.pack(expand = True, fill = "both") # размещение блока (с возможностью расширяться и заполненем окна во всех направлениях)

        text = tk.Label(f_top, justify=tk.LEFT, font = "Verdana 10", text = msg) # текст в окне
        text.pack(padx = 5, pady = 2) # размещение блока

        text = tk.Label(f_top, fg="green", justify=tk.LEFT, padx = 3, pady = 3, font = "Verdana 10") # текст
        ToolTip(text, current_file_name) # имя текущего файла в виде всплывающего окна
        Update_text() # обновление отчёта цифр
        text.pack() # размещение блока

        progress = ttk.Progressbar(f_top, orient = "horizontal", length = 250, mode = 'determinate') # панель прогресса (положение, длина, вид отображения)
        percent_all_failes_number = 100/all_failes_number # перевод в процент от общего числа
        Update_progress() # обновление прогресса
        progress.pack(padx = 4) # размещение блока

        button = tk.Button(f_top, font = "Verdana 11", command = Button_exit, text = "Отмена") # действие кнопки
        button.pack(side = "bottom", pady = 3) # размещение блока

        window.mainloop() # отображение окна

        Stop = True # триггер остановки обработки и сообщения
        print("Stop =", Stop)

    Stop = False # триггер остановки сообщения (для работы сообщений при повторном вызове)

    msg_th = Thread(target = Message_count_Thread, args = (all_failes_number, msg, )) # запуск определения положения мышки в отдельном потоке
    msg_th.start() # запуск потока

def text_processing(lines, Path): # обработка текста (строки текста, путь к файлу)

    import re # модуль регулярных выражений
    from sys import exit # для выхода из приложения без ошибки

    global parameters # делаем глобальным список с параметрами

    if lines == []: # если в файле нет записи, вписываем в файл опции для редактирования и инструкцию по использованию
        to_create_txt_file(Path) # создать txt файл с записью в него значений

    else: # если есть текст обратобать его

        try: # попытаться обработать значения в txt файле

            lines_clean = clearing_the_list(lines) # очистка списка строк от "#"

            for line in lines_clean: # для каждой строки производим обработку
                parameter = line.split("=") # делим по "="
                parameter[0] = parameter[0].strip() # убираем пробелы по бокам
                parameter[1] = parameter[1].strip().strip('"') # убираем пробелы и "..." по бокам

                if parameter[1].find("True") != -1: # если есть параметр со словом True, обрабатываем его
                    parameter[1] = True # присвоем значение "True"

                elif parameter[1].find("False") != -1 or parameter[1].strip() == "": # если есть параметр со словом False или "", обрабатываем его
                    parameter[1] = False # присвоем значение "False"

                elif parameter[1].find(";") != -1: # если есть параметр с ";", обрабатываем его
                    parameter[1] = parameter[1].split(";") # разделяем параметр по ";", создаёться список

                try: # пытаемся добавить параметры в словарь
                    parameters[parameter[0]] = parameter[1] # добавляем в словарь параметры

                except NameError: # если нет словаря создаём его и добавляем параметры
                    parameters = {} # создаём словарь с параметрами
                    parameters[parameter[0]] = parameter[1] # добавляем в словарь параметры

        except:
            message("Проверте правильность записи файла: \n\"" + Path + "\"\nИли удалите его, новый будет создан автоматически.") # сообщение, поверх всех окон с автоматическим закрытием
            exit() # выходим из програмы

def Parameters(): # присвоене значений прочитаных параметров

    global mass_saving # значение делаем глобальным

    mass_saving = parameters.setdefault("mass_saving", False) # путь к файлам на сервере

def Сheck_version(): # проверяем версию компаса

    global iKompasVersion # значение делаем глобальным

    iKompasVersion = iKompasObject.ksGetSystemVersion() # текущая версия компаса
    iKompasVersion = iKompasVersion[1] + iKompasVersion[2]*0.1 # разбиваем полученое значение на обычную запись (19.0)

def Сheck_active_file(): # проверяем открыт ли файл в КОМПАСе

    iKompasDocument = iApplication.ActiveDocument # делаем открытый докумени активным

    if iKompasDocument: # если открыт документ выполнять макрос
        return True # возвращаем значение

    else: # не открыт документ
        return False # возвращаем значение

def Save_file(): # сохраняем файл/файлы

    from sys import exit # для выхода из приложения без ошибки

    ver = iKompasObject.ksReadString("Введите номер версии КОМПАС-3D:", str(iKompasVersion)).replace(",",".") # окно в КОМПАСе с вводом версии КОМПАСа + замена "," на "."

    if ver == "": # если нажали крестик
        exit() # выходим из програмы

    try: # попытаться преобразовать в число с плавающей запятой
        ver = float(ver) # проверяем что это число

    except ValueError: # введено не число
        Message("Введите числовое значение!") # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)
        Save_file() # сохраняем файл

    if mass_saving: # массовое сохранение включено

        saving_file_namber = 0 # счётчик количества обработаных файлов

        while Сheck_active_file(): # пока есть открытый документ
            Saving_file(ver) # сохраняем файл (версия файла)
            saving_file_namber += 1 # считаем каждый сохранённый файл

        Message("Файлы сохранены!\nВсего файлов обработано: " + str(saving_file_namber)) # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)

    else: # одиночное сохранение
        Saving_file(ver) # сохраняем файл (версия файла)

def Saving_file(ver): # сохраняем файл (версия файла)

    from sys import exit # для выхода из приложения без ошибки

    iKompasDocument = iApplication.ActiveDocument # делаем открытый докумени активным

    PathName = iKompasDocument.PathName # полное имя документа

    if PathName == "": # если нет имени вписать
        type_doc = {1:"Чертёж.cdw", 2:"Фрагмент.frw", 3:"Спецификация.spw", 4:"Деталь.m3d", 5:"Сборка.a3d", 6:"Текстовый документ.kdw", 7:"Технологическая сборка.t3d"} # типы документов
        PathName = type_doc[iKompasDocument.DocumentType] # имя в зависимости от типа документа

    file_ver = Сhoosing_ver(ver) # выбор версии сохранения (версия)

    iKompasDocument1 = KompasAPI7.IKompasDocument1(iKompasDocument) # интерфейс документа

    if iKompasDocument1.SaveAsEx(PathName[:-4] + "_v" + str(file_ver[0]) + PathName[-4:], file_ver[1]): # сохранить, добавив к имени версию файла и сменить версию файла
        Kompas_message("Файл сохранён в версию " + str(file_ver[0])) # сообщение в окне КОМПАСа если он открыт

    else:
        Kompas_message("Файл не может быть сохранён в v" + str(file_ver[0])) # сообщение в окне КОМПАСа если он открыт
        exit()

    if mass_saving: # массовое сохранение включено
        iKompasDocument.Close(0) # 0 - закрыть документ без сохранения; 1 - закрыть документ, сохранив  изменения; 2 - выдать запрос на сохранение документа, если он изменен.

def Сhoosing_ver(ver): # выбор версии сохранения (версия)

    from sys import exit # для выхода из приложения без ошибки

    if ver <= iKompasVersion: # если меньше текущей версии открытого компаса

        dictionary = {5.11:1, 6.0:2, 6.1:3, 7.0:4, 7.1:5, 8.0:6, 8.1:7, 9.0:8, 10.0:9, 11.0:10, 12.0:11, 13.0:12, 14.0:13, 14.1:14, 14.2:15, 15.0:16, 15.1:17,
                    16.0:19, 16.1:20, 17.0:21, 17.1:22, 18.0:23, 18.1:24, 19.0:25, 20.0:26, 21.0:27} # версия файла и его номер сохраненияв КОМПАСе

        for key in dictionary: # перебор всех возможных версий из списка

            if ver < key: # если введённое число меньше чем в списке

                if ver < 5.11: # если меньше первой версии
                    ver = 5.11 # сохранять в v5.11
                    break # прекращаем цикл

                else: # если больше
                    ver = key_old # использовать предыдущую версию (идём с большего к меньшему)
                    break # прекращаем цикл

            elif ver == key: # если введённая версия равна версии из списка, использовать её
                break # прекращаем цикл

            else: # записать старый вариан версии
                key_old = key # старый вариан версии

        file_ver = (ver, dictionary[ver]) # создаём список с параметрами версии файла

        return file_ver # возвращаем версию файла (версия, её значение (saveMode))

    else: # если необходимая версия больше чем версия программы, выдать сообщение
        Kompas_message("Сохранение в указанную\nверсию невозможно!") # сообщение в окне КОМПАСа если он открыт
        exit()

#-------------------------------------------------------------------------------

DoubleExe() # проверка на уже запущеное приложение, с отключённым консольным окном "CREATE_NO_WINDOW"

KompasAPI() # подключение API компаса

Сheck_version() # проверяем версию компаса

if Сheck_active_file(): # проверяем открыт ли файл в КОМПАСе

    if Askyesnocancel("Сохранить все документы?"): # вопросительное сообщение, поверх всех окон
        mass_saving = True
        Save_file() # сохраняем файл/файлы

    else: # сохраняем только один
        mass_saving = False
        Save_file() # сохраняем файл/файлы

else: # файл не открыт в КОМПАСе
    Message("Открываем папку с файлами! (В разработке)") # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)