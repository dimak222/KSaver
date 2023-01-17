#-------------------------------------------------------------------------------
# Author:      dimak222
#
# Created:     21.01.2022
# Copyright:   (c) dimak222 2022
# Licence:     No
#-------------------------------------------------------------------------------

title = "Сохранение в старые версии"
ver = "v1.2.4.0"

#------------------------------Настройки!---------------------------------------
use_txt_file = False # использовать txt файл (True - да, False - нет)

model_name = True # брать обозначение и наименование из модели (False - имя файла)
rewrite = False # перезаписывать файлы с одинаковыми именами (False - не перезаписывать)

mass_saving = False # опция масового сохранения (для открытых файлов), ("False" или "" - не задаёт вопрос, сохраняет только текущий открытый файл)

file_version = "16" # версия в какую сохранять файл ("False" или "" - задавать вопрос при сохранении), примеры: "19"; "18,1"; "14.2"
file_version_name = True # запись версии в имя файла ("False" или "" - не записывать, при near_the_source = "source" и rewrite = "True" меняет версию оригинального файла!)

near_the_source = True # путь к папке, куда сохранять файлы (для открытых файлов) ("False" или "" - задавать вопрос при сохранении, "True" - рядом с исходником, при file_version_name = "False" и rewrite = "True" меняет версию оригинального файла!)

source_directory = False # путь к папке, откуда брать файлы (для сохранения из папки) ("False" или "" - задавать вопрос при сохранении)
final_directory = False # путь к папке, куда сохранять файлы (для сохранения из папки) ("False" или "" - задавать вопрос при сохранении)
#-------------------------------------------------------------------------------

def DoubleExe():# проверка на уже запущеное приложение, с отключённым консольным окном "CREATE_NO_WINDOW"

    import subprocess # модуль вывода запущеных процессов
    import os # модуль файловой системы
    from sys import exit # для выхода из приложения без ошибки

    CREATE_NO_WINDOW = 0x08000000 # отключённое консольное окно
    processes = subprocess.Popen('tasklist', stdin=subprocess.PIPE, stderr=subprocess.PIPE, stdout=subprocess.PIPE, creationflags=CREATE_NO_WINDOW).communicate()[0] # список всех процессов
    processes = processes.decode('cp866') # декодировка списка

    filename = os.path.basename(__file__) # имя запускаемого файла

    if str(processes).count(filename[0:25]) > 2: # если найдено название программы (два процесса) с ограничением в 25 символов
        Message("Приложение уже запущено!") # сообщение, поверх всех окон и с автоматическим закрытием
        exit() # выходим из програмы

def KompasAPI(): # подключение API компаса

    import pythoncom # модуль для запуска без IDLE
    from win32com.client import Dispatch, gencache # библиотека API Windows
    from sys import exit # для выхода из приложения без ошибки

    try: # попытаться подключиться к КОМПАСу

        global KompasAPI7 # значение делаем глобальным
        global iApplication # значение делаем глобальным
        global iKompasObject # значение делаем глобальным
        global iDocuments # значение делаем глобальным

        KompasConst3D = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0).constants # константа 3D документов
        KompasConst2D = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants # константа 2D документов
        KompasConst = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants # константа для скрытия вопросов перестроения

        KompasAPI5 = gencache.EnsureModule('{0422828C-F174-495E-AC5D-D31014DBBE87}', 0, 1, 0) # API5 КОМПАСа
        iKompasObject = Dispatch('Kompas.Application.5', None, KompasAPI5.KompasObject.CLSID) # интерфейс API КОМПАС

        KompasAPI7 = gencache.EnsureModule('{69AC2981-37C0-4379-84FD-5DD2F3C0A520}', 0, 1, 0) # API7 КОМПАСа
        iApplication = Dispatch('Kompas.Application.7') # интерфейс приложения КОМПАС-3D.

        iKompasDocument = iApplication.ActiveDocument # делаем активный открытый документ

        iDocuments = iApplication.Documents # интерфейс для открытия документов

        if iApplication.Visible == False: # если компас невидимый
            iApplication.Visible = True # сделать КОМПАС-3D видемым

    except: # если не получилось подключиться к КОМПАСу

        Message("КОМПАС-3D не найден!\nУстановите или переустановите КОМПАС-3D!") # сообщение, поверх всех окон с автоматическим закрытием
        exit() # выходим из програмы

def Kompas_message(text): # сообщение в окне КОМПАСа если он открыт

    if iApplication.Visible == True: # если компас видимый
        iApplication.MessageBoxEx(text, 'Message:', 64) # сообщение в КОМПАСе

def Message(text = "Ошибка!", counter = 4): # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)

    from threading import Thread # библиотека потоков
    import time # модуль времени

    def Resource_path(relative_path): # для сохранения картинки внутри exe файла

        import os # работа с файовой системой

        try: # попытаться определить путь к папке
            base_path = sys._MEIPASS # путь к временной папки PyInstaller

        except Exception: # если ошибка
            base_path = os.path.abspath(".") # абсолютный путь

        return os.path.join(base_path, relative_path) # объеденяем и возващаем полный путь

    def Message_Thread(text, counter): # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)

        import tkinter.messagebox as mb # окно с сообщением
        import tkinter as tk # модуль окон

        if counter == 0: # время до закрытия окна (если 0)
            counter = 1 # закрытие через 1 сек
        window_msg = tk.Tk() # создание окна
        try: # попытаться использовать значёк
            window_msg.iconbitmap(default = Resource_path("cat.ico")) # значёк программы
        except: # если ошибка
            pass # пропустить
        window_msg.attributes("-topmost",True) # окно поверх всех окон
        window_msg.withdraw() # скрываем окно "невидимое"
        time = counter * 1000 # время в милисекундах
        window_msg.after(time, window_msg.destroy) # закрытие окна через n милисекунд
        if mb.showinfo(title, text, parent = window_msg) == "": # информационное окно закрытое по времени
            pass # пропустить
        else: # если не закрыто по времени
            window_msg.destroy() # окно закрыто по кнопке
        window_msg.mainloop() # отображение окна

    msg_th = Thread(target = Message_Thread, args = (text, counter)) # запуск окна в отдельном потоке
    msg_th.start() # запуск потока

    msg_th.join() # ждать завершения процесса, иначе может закрыться следующие окно

def Askyesnocancel(text): # вопросительное сообщение, поверх всех окон

    import tkinter.messagebox as mb # окно с сообщением
    import tkinter as tk # модуль окон
    from sys import exit # для выхода из приложения без ошибки

    def Resource_path(relative_path): # для сохранения картинки внутри exe файла

        import os # работа с файовой системой

        try: # попытаться определить путь к папке
            base_path = sys._MEIPASS # путь к временной папки PyInstaller

        except Exception: # если ошибка
            base_path = os.path.abspath(".") # абсолютный путь

        return os.path.join(base_path, relative_path) # объеденяем и возващаем полный путь

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

    def Resource_path(relative_path): # для сохранения картинки внутри exe файла

        import os # работа с файовой системой

        try: # попытаться определить путь к папке
            base_path = sys._MEIPASS # путь к временной папки PyInstaller

        except Exception: # если ошибка
            base_path = os.path.abspath(".") # абсолютный путь

        return os.path.join(base_path, relative_path) # объеденяем и возващаем полный путь

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

        def Resource_path(relative_path): # для сохранения картинки внутри exe файла

            import os # работа с файовой системой

            try: # попытаться определить путь к папке
                base_path = sys._MEIPASS # путь к временной папки PyInstaller

            except Exception: # если ошибка
                base_path = os.path.abspath(".") # абсолютный путь

            return os.path.join(base_path, relative_path) # объеденяем и возващаем полный путь

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

def Txt_file(): # считываем значения настроек из txt файла

    import os # работа с файовой системой

    def Text_processing(lines, Path): # обработка текста (строки текста, путь к файлу)

        import re # модуль регулярных выражений
        from sys import exit # для выхода из приложения без ошибки

        global parameters # делаем глобальным список с параметрами

        def Сlearing_the_list(lines): # очистка списка строк от "#"

            import re # модуль регулярных выражений

            lines_clean = [] # список строк с чистым текстом (без текста после "#")

            for line in lines: # для каждой строки производим обработку

                if line.isspace(): # если пустая строка пропустить
                    continue

                ignore_grid = re.findall("\".[^\"]+?#.+?\"", line) # проверка строки на содержание текста с решоткой в ковычках "***#***"

                if ignore_grid != []: # если строка содержит текст с решоткой в ковычках "***#***"

                    line = line.replace(ignore_grid[0], "Текст с решоткой в ковычках!=)", 1) # заменяем текст с решоткой в ковычках "***#***" на "|"

                    line_clean = line.split("#", 1)[0] # если в строке есть "#" не записывать всё что после неё

                    line_clean = line_clean.replace("Текст с решоткой в ковычках!=)", ignore_grid[0], 1) # заменяем "|" на текст с решоткой в ковычках "***#***"

                else: # не содержит текст с решоткой в ковычках "***#***"
                    line_clean = line.split("#", 1) # если в строке есть "#" не записывать всё что после неё

                if line_clean[0].strip() == "": # если нет записи до #, пропустиь строку
                    continue

                lines_clean.append(line_clean) # список строк с чистым текстом (без текста после "#")

            return lines_clean # возврящаем чистые строки

        if lines == []: # если в файле нет записи, вписываем в файл опции для редактирования и инструкцию по использованию
            To_create_txt_file(Path) # создать txt файл с записью в него значений

        else: # если есть текст обратобать его

            try: # попытаться обработать значения в txt файле

                lines_clean = Сlearing_the_list(lines) # очистка списка строк от "#"

                for line in lines_clean: # для каждой строки производим обработку
                    parameter = line[0].split("=") # делим по "="
                    parameter[0] = parameter[0].strip() # убираем пробелы по бокам
                    parameter[1] = parameter[1].strip().strip('"') # убираем пробелы и "..." по бокам

                    if parameter[1].find("True") != -1: # если есть параметр со словом True, обрабатываем его
                        parameter[1] = True # присвоем значение "True"

                    elif parameter[1].find("False") != -1 or parameter[1].strip() == "": # если есть параметр со словом False или "", обрабатываем его
                        parameter[1] = False # присвоем значение "False"

                    elif parameter[1].find(";") != -1: # если есть параметр с ";", обрабатываем его
                        parameter[1] = parameter[1].split(";") # разделяем параметр по ";", создаёться список

                    try: # пытаемся добавить параметры в словарь
                        parameters[parameter[0]] = [parameter[1], line[1]] # добавляем в словарь параметры

                    except NameError: # если нет словаря создаём его и добавляем параметры
                        parameters = {} # создаём словарь с параметрами
                        parameters[parameter[0]] = [parameter[1], line[1]] # добавляем в словарь параметры

            except:
                Message("Проверте правильность записи файла: \n\"" + Path + "\"\nИли удалите его, новый будет создан автоматически.") # сообщение, поверх всех окон с автоматическим закрытием
                exit() # выходим из програмы

    def To_create_txt_file(name_txt_file): # создать txt файл с записью в него значений

        import os # работа с файовой системой
        from sys import exit # для выхода из приложения без ошибки

        txt_file = open(name_txt_file, "w+", encoding = "utf-8") # открываем файл записи (w+), для чтения (r), (невидимый режим)

        txt = """model_name = "True" # брать обозначение и наименование из модели (False - имя файла)
rewrite = False # перезаписывать файлы с одинаковыми именами (False - не перезаписывать)
#-------------------------------------------------------------------------------
mass_saving = "True" # опция масового сохранения (для открытых файлов), ("False" или "" - не задаёт вопрос, сохраняет только текущий открытый файл)
#-------------------------------------------------------------------------------
file_version = "False" # версия в какую сохранять файл ("False" или "" - задавать вопрос при каждом сохранении), примеры: "19"; "18,1"; "14.2"
file_version_name = "True" # запись версии в имя файла ("False" или "" - не записывать, при near_the_source = "source" и rewrite = "True" меняет версию оригинального файла!)
#-------------------------------------------------------------------------------
near_the_source = "source" # путь к папке, куда сохранять файлы (для открытых файлов) ("False" или "" - задавать вопрос при сохранении, "True" - рядом с исходником, при file_version_name = "False" и rewrite = "True" меняет версию оригинального файла!)
#-------------------------------------------------------------------------------
source_directory = "False" # путь к папке, откуда брать файлы (для сохранения из папки) ("False" или "" - задавать вопрос при сохранении)
final_directory = "False" # путь к папке, куда сохранять файлы (для сохранения из папки) ("False" или "" - задавать вопрос при сохранении)
""" # текст записываемый в .txt файл

        txt_file.write(txt) # записываем текст в файл
        txt_file.close() # закрываем файл

        os.startfile(name_txt_file) # открываем файл в системе
        Message("Введите необходимые значения! \nИ запустите приложение повторно.") # сообщение с названием файла
        exit() # выходим

    name_txt_file = os.path.join(title + ".txt") # название текстового файла

    if os.path.exists(name_txt_file): # если есть txt файл использовать его

        txt_file = open(name_txt_file, encoding = "utf-8") # открываем файл записи (w+), для чтения (r), (невидимый режим)
        lines = txt_file.readlines()  # прочитать все строки
        txt_file.close() # закрываем файл

        Text_processing(lines, name_txt_file) # обработка текста (строки текста, путь к файлу)

        Parameters() # присвоене значений прочитаных параметров

    else: # если нет файла
        To_create_txt_file(name_txt_file) # создать txt файл с записью в него значений (путь к txt файлу)

def Parameters(): # присвоене значений прочитаных параметров

    global model_name # значение делаем глобальным
    global rewrite # значение делаем глобальным

    global mass_saving # значение делаем глобальным

    global file_version # значение делаем глобальным
    global file_version_name # значение делаем глобальным

    global near_the_source # значение делаем глобальным

    global source_directory # значение делаем глобальным
    global final_directory # значение делаем глобальным

    model_name = parameters.setdefault("model_name", True)[0] # опция масового сохранения
    rewrite = parameters.setdefault("rewrite", True)[0] # опция масового сохранения

    mass_saving = parameters.setdefault("mass_saving", False)[0] # опция масового сохранения

    file_version = parameters.setdefault("file_version", False)[0] # версия в какую сохранять файл
    file_version_name = parameters.setdefault("file_version_name", True)[0] # запись версии в имя файла

    near_the_source = parameters.setdefault("near_the_source", "source")[0] # путь к папке, куда сохранять файлы (для открытых файлов)

    source_directory = parameters.setdefault("source_directory", False)[0] # путь к папке, откуда брать файлы (для сохранения из папки)
    final_directory = parameters.setdefault("final_directory", False)[0] # путь к папке, куда сохранять файлы (для сохранения из папки)

##    if near_the_source == "source" or final_directory == "source": # если сохранёный файл класть рядом с исходником
##        file_version_name = True # запись версии в имя файла

    for parameter, val in parameters.items(): # для каждой строки производим обработку
        print(f"{parameter} = {val[0]}") # выводим прочитание параметры

def Сheck_version(): # проверяем версию компаса

    global iKompasVersion # значение делаем глобальным

    iKompasVersion = iKompasObject.ksGetSystemVersion() # текущая версия компаса
    iKompasVersion = iKompasVersion[1] + iKompasVersion[2]*0.1 # разбиваем полученое значение на обычную запись (19.0)

def Сheck_active_file(): # проверяем открыт ли файл в КОМПАСе

    if iApplication.ActiveDocument: # если открыт документ выполнять макрос
        return True # возвращаем значение

    else: # не открыт документ
        return False # возвращаем значение

def Window_filedialog(text): # создание окна filedialog

    import tkinter as tk # модуль окон
    from tkinter import filedialog # окно с выбором папки
    from sys import exit # для выхода из приложения без ошибки

    def Resource_path(relative_path): # для сохранения картинки внутри exe файла

            import os # работа с файовой системой

            try: # попытаться определить путь к папке
                base_path = sys._MEIPASS # путь к временной папки PyInstaller

            except Exception: # если ошибка
                base_path = os.path.abspath(".") # абсолютный путь

            return os.path.join(base_path, relative_path) # объеденяем и возващаем полный путь

    window = tk.Tk() # создание окна что бы сделать filedialog поверх всех окон
    window.iconbitmap(default = Resource_path("cat.ico")) # значёк программы
    window.attributes("-topmost",True) # окно поверх всех окон
    window.withdraw() # скрываем окно "невидимое"
    Path = filedialog.askdirectory(title = text) # запрос указания папки
    window.destroy() # закрытие окна от filedialog
    window = window.mainloop() # отображение окна

    if Path != "": # если путь выбран
        return Path # возвращаем путь

    else: # если путь не выбран
        print("Путь не выбран!")
        exit()

def Save_file(mass_saving): # сохраняем файл/файлы

    import os # модуль файловой системы

    if near_the_source and isinstance(near_the_source, bool): # если сохранять рядом с исходником
        Path = True # для работы в следующей функции
    else: # если не сохранять рядом с исходником
        if os.path.exists(str(near_the_source)): # если есть такой путь
            Path = near_the_source # путь вписанный в txt
        else: # нет такого пути
            Path = Window_filedialog("Выберите папку сохранения файлов") # создание окна filedialog

    ver = File_version(file_version) # определение в какую версию сохранять файл

    ver = Сhoosing_ver(ver) # округление версии сохранения (версия)

    if mass_saving: # массовое сохранение включено

        saving_file_namber = 0 # счётчик количества обработаных файлов

        while iApplication.ActiveDocument: # пока есть открытый документ
            Saving_file(ver, mass_saving, Path) # сохраняем файл (версия файла, опция массового сохранения, путь к папке)
            saving_file_namber += 1 # считаем каждый сохранённый файл

        Message("Файлы сохранены!\nВсего файлов обработано: " + str(saving_file_namber)) # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)

    else: # одиночное сохранение
        Saving_file(ver, mass_saving, Path) # сохраняем файл (версия файла, опция массового сохранения, путь к папке)

def File_version(file_version): # определение в какую версию сохранять файл

    from sys import exit # для выхода из приложения без ошибки

    if file_version == False: # если версия файла не задана
        ver = iKompasObject.ksReadString("Введите номер версии КОМПАС-3D:", str(iKompasVersion)).replace(",",".") # окно в КОМПАСе с вводом версии КОМПАСа + замена "," на "."

    else: # версия файла задана
        ver = file_version # значение версии из txt файла

    if ver == "": # если нажали крестик
        exit() # выходим из програмы

    try: # попытаться преобразовать в число с плавающей запятой
        ver = float(ver) # проверяем что это число

    except ValueError: # введено не число
        Message("Введите числовое значение!") # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)
        File_version(False) # определение в какую версию сохранять файл

    return ver

def Сhoosing_ver(ver): # округление версии сохранения (версия)

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

def Saving_file(ver, mass_saving, Path): # сохраняем файл (версия файла, опция массового сохранения, путь к папке)

    import os # модуль файловой системы

    from sys import exit # для выхода из приложения без ошибки

    iKompasDocument = iApplication.ActiveDocument # делаем открытый докумени активным

    resultPath = ResultPath(iKompasDocument, model_name, Path, ver) # результирующий путь сохранения файла, по 3D модели или имени файла (интерфейс док., параметр обозн. и наим. из модели)

    iKompasDocument1 = KompasAPI7.IKompasDocument1(iKompasDocument) # интерфейс документа

    if iKompasDocument1.SaveAsEx(resultPath, ver[1]): # сохранить, добавив к имени версию файла и сменить версию файла
        Kompas_message("Файл сохранён в версию " + str(ver[0])) # сообщение в окне КОМПАСа если он открыт

    else:
        Kompas_message("Файл не может быть сохранён в v" + str(ver[0])) # сообщение в окне КОМПАСа если он открыт

        if os.path.exists(resultPath): # проверка существования пути
            print("Нужно удалить! Но проверить, что бы путь не совпадал с оригинальным файлом!")
        exit()

    if mass_saving: # массовое сохранение включено
        iKompasDocument.Close(0) # 0 - закрыть документ без сохранения; 1 - закрыть документ, сохранив  изменения; 2 - выдать запрос на сохранение документа, если он изменен.

def ResultPath(iKompasDocument, model_name, Path, ver): # результирующий путь сохранения файла, по 3D модели или имени файла (интерфейс док., параметр обозн. и наим. из модели)

    import os # работа с файовой системой

    def Read_name_and_obozn(iKompasDocument): # считывание обозначение и наименование 3D модели

        import os # работа с файовой системой

        def Removal(text): # удалние знаков которые не могут быть в наименовании файла

            removal = ["\"", "\\", "@/", "/", ":", "*", "?", "<", ">", "|"] # удаляемые знаки

            for removal in removal: # перебор удалния знаков
                text = text.replace(removal, " ") # замена """ "/" и т.д. на " "
                text = text.strip() # убираем пробелы по бокам

            return text # возврящаем значения

        iKompasDocument3D = KompasAPI7.IKompasDocument3D(iKompasDocument) # базовый класс документов-моделей КОМПАС
        iPart7 = iKompasDocument3D.TopPart # интерфейс компонента 3D документа (сам документ)

        iPropertyKeeper = KompasAPI7.IPropertyKeeper(iPart7) # интерфейс получения/редактирования значения свойств
        iPropertyMng = KompasAPI7.IPropertyMng(iApplication) # менеджера свойств

        readobozn = iPropertyKeeper.GetPropertyValue(iPropertyMng.GetProperty(iKompasDocument, "Обозначение"), 0, True)[1] #  прочитаем обозначение текущего исполнения
        readobozn = Removal(readobozn) # удалние знаков которые не могут быть в наименовании файла

        readname = iPropertyKeeper.GetPropertyValue(iPropertyMng.GetProperty(iKompasDocument, "Наименование"), 0, True)[1] #  прочитаем наименоваание текущего исполнения
        readname = Removal(readname) # удалние знаков которые не могут быть в наименовании файла

        if readobozn == "" and readname == "": # если обозначение и наименование не записанно, сохранить имя файла
            readname = os.path.basename(iPart7.FileName) # получаем имя файла с расширением
            readname = os.path.splitext(readname)[0] # имя файла без расширения
            if readname == "": # если имя файла отсутствует (локальная деталь)
                readname = "Деталь без имени"

        if readobozn != "": # если обозначение есть, добавить разделитель
            separator = "_" # разделитель
        else: # иначе без разделителя
            separator = "" # нет разделителя

        name_and_obozn = readobozn + separator + readname # обозначение и имя файла

        print("Имя файла: " + name_and_obozn)

        return name_and_obozn # возврящаем значения

    def Rename(Path, n): # изменение имени если уже есть такое

        import os # работа с файовой системой

        Path_temp = os.path.splitext(Path)[0] + " (" + str(n) + ")" + os.path.splitext(Path)[1] # добавляем номер файла

        if os.path.exists(Path_temp): # проверка наличия файла

            n = n + 1 # увеличиваем номер

            Path_temp = rename(Path, n) # снова запускаем изменение имени со старым названием, но с новым номером

        return Path_temp # возвращаем последний прошедший проверку путь

    if Path and isinstance(Path, bool): # если сохранять рядом с исходником
        PathName = iKompasDocument.PathName # полное имя документа

        if PathName == "": # если нет имени вписать

            Message("Файл не был сохранён, укажите папку сохранения") # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)

            Path = Window_filedialog("Выберите папку сохранения файлов") # создание окна filedialog

            Saving_file(ver, mass_saving, Path) # сохраняем файл (версия файла, опция массового сохранения, путь к папке)

        if file_version_name:
            resultPath = f"{PathName[:-4]}_v{str(ver[0])}{PathName[-4:]}"

    else: # сохранение в папку

        fileBasename = os.path.basename(iKompasDocument.PathName) # имя документа с расширением
        fileBasename_and_expansion = os.path.splitext(fileBasename) # список имя файла с расширением
        basename = fileBasename_and_expansion[0] # имя файла (без расширения)
        expansion = fileBasename_and_expansion[1] # расширение файла

        if expansion == "": # если нет расширения (не сохранена деталь)
            type_doc = {1:".cdw", 2:".frw", 3:".spw", 4:".m3d", 5:".a3d", 6:".kdw", 7:".t3d"} # типы документов (Чертёж, Фрагмент, Спецификация, Деталь, Сборка, Текстовый документ, Технологическая сборка)
            expansion = type_doc[iKompasDocument.DocumentType] # имя в зависимости от типа документа

        if file_version_name: # если запись версии включена
            version = f"_v{str(ver[0])}" # добавляем запись версии файла
        else: # если запись версии выключена
            version = "" # не вписываем версию

        if model_name: # обозначение и наименование брать из 3D модели
            name_and_obozn = Read_name_and_obozn(iKompasDocument) # считывание обозначение и наименование 3D модели
            resultPath = os.path.join(Path, name_and_obozn + version + expansion) # объеденение пути и имени файла (без расширения)

            if rewrite == False: # если не перезаписывать
                if os.path.exists(resultPath): # проверка наличия файла
                    resultPath = Rename(resultPath, 2) # изменение имени если уже есть такое (полный путь к файлу, номер с какого начинать заменять)

        else: # обозначение и наименование брать из имени файла
            resultPath = os.path.join(Path, basename + version + expansion) # объеденение пути и имени файла (без расширения)

            if rewrite == False: # если не перезаписывать
                if os.path.exists(resultPath): # проверка наличия файла
                    resultPath = Rename(resultPath, 2) # изменение имени если уже есть такое (полный путь к файлу, номер с какого начинать заменять)

    return resultPath # возврящаем путь к файлу

#-------------------------------------------------------------------------------

DoubleExe() # проверка на уже запущеное приложение, с отключённым консольным окном "CREATE_NO_WINDOW"

KompasAPI() # подключение API компаса

if use_txt_file: # использовать txt файл
    Txt_file() # считываем значения настроек из txt файла

Сheck_version() # проверяем версию компаса

if iApplication.ActiveDocument: # проверяем открыт ли файл в КОМПАСе

    if iDocuments.Count > 1: # количество открытых вкладок

        if mass_saving: # если массовое сохранение включено
            if Askyesnocancel("Сохранить все документы?"): # вопросительное сообщение, поверх всех окон
                Save_file(True) # сохраняем файл/файлы

            else: # сохраняем только один
                Save_file(False) # сохраняем файл/файлы

        else: # сохраняем только один
            Save_file(False) # сохраняем файл/файлы

    else: # сохраняем только один
        Save_file(False) # сохраняем файл/файлы

else: # файл не открыт в КОМПАСе
    source_directory = Window_filedialog("Выберите папку с файлами") # создание окна filedialog
    final_directory = Window_filedialog("Выберите папку сохранения файлов") # создание окна filedialog
    print(source_directory)
    Message("Открываем папку с файлами! (В разработке)") # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)