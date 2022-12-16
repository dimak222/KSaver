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

    while msg_th.is_alive(): # ждать запущеный поток
        time.sleep(0.5) # ждём n сек

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

def Save_one_file(): # сохраняем один файл

    from sys import exit # для выхода из приложения без ошибки

    iKompasDocument = iApplication.ActiveDocument # делаем открытый докумени активным

    PathName = iKompasDocument.PathName # полное имя документа

    if PathName == "": # если нет имени вписать
        type_doc = {1:"Чертёж.cdw", 2:"Фрагмент.frw", 3:"Спецификация.spw", 4:"Деталь.m3d", 5:"Сборка.a3d", 6:"Текстовый документ.kdw", 7:"Технологическая сборка.t3d"} # типы документов
        PathName = type_doc[iKompasDocument.DocumentType] # имя в зависимости от типа документа

    ver = iKompasObject.ksReadString("Введите номер версии КОМПАС-3D:", str(iKompasVersion)).replace(",",".") # окно в КОМПАСе с вводом версии КОМПАСа + замена "," на "."

    if ver == "": # если нажали крестик
        exit() # выходим из програмы

    try: # попытаться преобразовать в число с плавающей запятой
        ver = float(ver) # проверяем что это число

    except ValueError: # введено не число
        Message("Введите числовое значение!") # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)
        Save_one_file() # один файл

    Save_file(ver, PathName, iKompasDocument) # сохраняем файл

def Save_file(ver, PathName, iKompasDocument): # сохраняем файл (версия файла, путь к файлу, активный документ)

##    try: # попытаться сохранить

        file_ver = Сhoosing_ver(ver) # выбор версии сохранения (версия)

        if file_ver: # если версия указана (версия, её значение (saveMode)), выполнить

            iKompasDocument1 = KompasAPI7.IKompasDocument1(iKompasDocument) # интерфейс документа

            if iKompasDocument1.SaveAsEx(PathName[:-4] + "_v" + str(file_ver[0]) + PathName[-4:], file_ver[1]): # сохранить, добавив к имени версию файла и сменить версию файла
                iApplication.MessageBoxEx("Файл сохранён в версию " + str(file_ver[0]), "Отчёт:", 64) # выдать сообщение в какую версию сохранено

            else:
                iApplication.MessageBoxEx("Файл не может быть сохранён в v" + str(file_ver[0]), "Отчёт:", 64) # выдать сообщение в какую версию сохранено

##    except: # не получилось сохранить
##        iApplication.MessageBoxEx( "Данные введены некорректно!", "Отчёт:", 64) # выдать сообщение о неправильной записи

def Сhoosing_ver(ver): # выбор версии сохранения (версия)

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
        iApplication.MessageBoxEx( "Сохранение в указанную\nверсию невозможно!", "Отчёт:", 64) # выдать сообщение в компасе
        return False

#-------------------------------------------------------------------------------

DoubleExe() # проверка на уже запущеное приложение, с отключённым консольным окном "CREATE_NO_WINDOW"

KompasAPI() # подключение API компаса

Сheck_version() # проверяем версию компаса

if Сheck_active_file(): # проверяем открыт ли файл в КОМПАСе

    if Askyesnocancel("Сохранить все документы?"): # вопросительное сообщение, поверх всех окон
        Message("Сохраняем все! (В разработке)") # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)

    else: # сохраняем только один
        Save_one_file() # сохраняем файл

else: # файл не открыт в КОМПАСе
    Message("Открываем папку с файлом! (В разработке)") # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)