#-------------------------------------------------------------------------------
# Author:      dimak222
#
# Created:     21.01.2022
# Copyright:   (c) dimak222 2022
# Licence:     No
#-------------------------------------------------------------------------------

title = "KSaver" # наименование приложения
ver = "v1.6.0.0" # версия файла

#------------------------------Настройки!---------------------------------------

use_txt_file = True # использовать txt файл (True - да; False - нет)

dict_settings = {
"check_update" : [True, '# проверять обновление программы ("True" - да; "False" или "" - нет)'],
"beta" : [False, '# скачивать бета версии программы ("True" - да; "False" или "" - нет)'],

"model_name" : [True, '# брать обозначение и наименование из модели ("True" - обозначение и наименование из модели; "False" или "" - имя файла)'],
"rewrite" : [False, '# перезаписывать файлы с одинаковыми именами ("True" - перезаписывать; "False" или "" - не перезаписывать)'],

"mass_saving" : [True, '# опция массового сохранения (для открытых файлов), ("True" - сохраняет все открытые вкладки с файлами; "False" или "" - не задаёт вопрос, сохраняет только текущий открытый файл)'],

"file_version" : [False, '# версия в какую сохранять файл (примеры: "19"; "18,1"; "14.2") ("True" - сохранять в v5.11; "False" или "" - задавать вопрос при сохранении)'],
"file_version_name" : [False, '# запись версии в имя файла ("True" - записывать; "False" или "" - не записывать)'],

"near_the_source" : [True, '# путь к папке (пример: "C:\ASCON"), куда сохранять файлы (открытые файлы) ("True" - рядом с исходником; "False" или "" - задавать вопрос при сохранении)'],

"types_of_documents" : ["1-7", '# типы документов для обработки (примеры: "1-7" - все типы; "1-3,5,7" - тип с 1-го по 3-й, 5-й, 7-ой) (1:".cdw", 2:".frw", 3:".spw", 4:".m3d", 5:".a3d", 6:".kdw", 7:".t3d")'],
"recursion" : [True, '# рекурсия (обработка папок внутри выбраной папки) ("True" - обрабатывать все папки; "False" или "" - обрабатывать только выбранную папку)'],

"source_directory" : [False, '# путь к папке (пример: "C:\ASCON"), откуда брать файлы (сохранение из папки) ("False" или "" - задавать вопрос при сохранении)'],
"final_directory" : [False, '# путь к папке (пример: "C:\ASCON"), куда сохранять файлы (сохранение из папки) ("True" - рядом с исходником; "False" или "" - задавать вопрос при сохранении)'],
} # словарь с настройками программы

#------------------------------Импорт модулей-----------------------------------

import psutil # модуль вывода запущеных процессов
import os # работа с файовой системой

from sys import exit # для выхода из приложения без ошибки
import pythoncom # модуль для запуска без IDE
from win32com.client import Dispatch, gencache # библиотека API Windows

from threading import Thread # библиотека потоков
import tkinter as tk # модуль окон
import tkinter.messagebox as mb # окно с сообщением

import tkinter.ttk as ttk # модуль прогресс-бара

from bisect import bisect_right # модуль бинарного поиска

import re # модуль регулярных выражений

import glob # модуль поиска файлов

from tkinter import filedialog # окно с выбором папки

#-------------------------------------------------------------------------------

def DoubleExe(): # проверка на уже запущеное приложение

    global program_directory # значение делаем глобальным

    list = [] # список найденых названий программы

    filename = psutil.Process().name() # имя запущеного файла
    filename2 = title + ".exe" # имя запущеного файла

    if filename == "python.exe" or filename == "pythonw.exe": # если программа запущена в IDE/консоли
        pass # пропустить

    else: # запущено не в IDE/консоли

        for process in psutil.process_iter(): # перебор всех процессов

            try: # попытаться узнать имя процесса
                proc_name = process.name() # имя процесса

            except psutil.NoSuchProcess: # в случае ошибки
                pass # пропускаем

            else: # если есть имя
                if proc_name == filename or proc_name == filename2: # сравниваем имя
                    list.append(process.cwd()) # добавляем в список название программы
                    if len(list) > 2: # если найдено больше двух названий программы (два процесса)
                        Message("Приложение уже запущено!") # сообщение, поверх всех окон и с автоматическим закрытием
                        exit() # выходим из программы

    if list == []: # если нет найденых названий программы
        program_directory = os.path.dirname(os.path.abspath(__file__)) # директория рядом с программой

    else: # если путь найден
        program_directory = os.path.dirname(psutil.Process().exe()) # директория программы

    program_directory = program_directory.replace("\\", "//", 1) # замена на слеш первого символа (при "\\192.168....")

def Txt_file(): # считываем значения настроек из txt файла

    def Text_processing(lines, path): # обработка текста (строки текста, путь к файлу)

        def Сlearing_the_list(lines): # очистка списка строк от "#"

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
            Сreate_settings_file(path) # создать txt файл с записью в него значений

        else: # если есть текст обратобать его

            lines_clean = Сlearing_the_list(lines) # очистка списка строк от "#"

            for line in lines_clean: # для каждой строки производим обработку

                parameter = line[0].split("=") # делим по "="
                parameter[0] = parameter[0].strip() # убираем пробелы по бокам
                parameter[1] = parameter[1].strip().strip('"') # убираем пробелы и "..." по бокам

                if str(parameter[1]).find("True") != -1: # если есть параметр со словом True, обрабатываем его
                    dict_settings[parameter[0]] = [True, line[1]] # словарь параметров

                elif str(parameter[1]).find("False") != -1 or parameter[1].strip() == "": # если есть параметр со словом False или "", обрабатываем его
                    dict_settings[parameter[0]] = [False, line[1]] # словарь параметров

                elif str(parameter[1]).find(";") != -1: # если есть параметр с ";", обрабатываем его
                    parameter[1] = parameter[1].split(";") # разделяем параметр по ";", создаёться список
                    dict_settings[parameter[0]] = [parameter[1], line[1]] # словарь параметров

                else: # все остальные значение записать
                    dict_settings[parameter[0]] = [parameter[1], line[1]] # словарь параметров

    def Сreate_settings_file(path): # создать txt файл с записью в него значений

        txt = """check_update = "True" # проверять обновление программы ("True" - да; "False" или "" - нет)
beta = "False" # скачивать бета версии программы ("True" - да; "False" или "" - нет)
#-------------------------------------------------------------------------------
model_name = "True" # брать обозначение и наименование из модели ("True" - обозначение и наименование из модели; "False" или "" - имя файла)
rewrite = "False" # перезаписывать файлы с одинаковыми именами ("True" - перезаписывать; "False" или "" - не перезаписывать)
#-------------------------------------------------------------------------------
mass_saving = "True" # опция массового сохранения (для открытых файлов), ("True" - сохраняет все открытые вкладки с файлами; "False" или "" - не задаёт вопрос, сохраняет только текущий открытый файл)
#-------------------------------------------------------------------------------
file_version = "False" # версия в какую сохранять файл (примеры: "19"; "18,1"; "14.2") ("True" - сохранять в v5.11; "False" или "" - задавать вопрос при сохранении)
file_version_name = "False" # запись версии в имя файла ("True" - записывать; "False" или "" - не записывать)
#-------------------------------------------------------------------------------
near_the_source = "True" # путь к папке (пример: "C:\ASCON"), куда сохранять файлы (открытые файлы) ("True" - рядом с исходником; "False" или "" - задавать вопрос при сохранении)
#-------------------------------------------------------------------------------
types_of_documents = "1-7" # типы документов для обработки (примеры: "1-7" - все типы; "1-3,5,7" - тип с 1-го по 3-й, 5-й, 7-ой) (1:".cdw", 2:".frw", 3:".spw", 4:".m3d", 5:".a3d", 6:".kdw", 7:".t3d")
recursion = "True" # рекурсия (обработка папок внутри выбраной папки) ("True" - обрабатывать все папки; "False" или "" - обрабатывать только выбранную папку)
#-------------------------------------------------------------------------------
source_directory = "False" # путь к папке (пример: "C:\ASCON"), откуда брать файлы (сохранение из папки) ("False" или "" - задавать вопрос при сохранении)
final_directory = "False" # путь к папке (пример: "C:\ASCON"), куда сохранять файлы (сохранение из папки) ("True" - рядом с исходником; "False" или "" - задавать вопрос при сохранении)
""" # текст записываемый в .txt файл

        with open(path, "w+", encoding = "utf-8") as f: # открываем файл записи (w+), для чтения (r), (невидимый режим)

            f.write(txt) # записываем текст в файл

        os.startfile(path.replace("//", r"\\", 1)) # открываем файл в системе
        Message("Введите необходимые значения! \nИ запустите приложение повторно.") # сообщение с названием файла
        exit() # выходим

    path = os.path.join(program_directory, title + ".txt") # название текстового файла

    if os.path.exists(path): # если есть txt файл использовать его

        with open(path, encoding = "utf-8") as f: # открываем файл записи (w+), для чтения (r), (невидимый режим)
            lines = f.readlines()  # прочитать все строки

        Text_processing(lines, path) # обработка текста (строки текста, путь к файлу)

        Settings() # присвоене значений прочитаных параметров

    else: # если нет файла
        Сreate_settings_file(path) # создать txt файл с записью в него значений (путь к txt файлу)

def Settings(): # присвоене значений прочитаных параметров

    global check_update # значение делаем глобальным
    global beta # значение делаем глобальным

    global model_name # значение делаем глобальным
    global rewrite # значение делаем глобальным

    global mass_saving # значение делаем глобальным

    global file_version # значение делаем глобальным
    global file_version_name # значение делаем глобальным

    global near_the_source # значение делаем глобальным

    global types_of_documents # значение делаем глобальны
    global recursion # значение делаем глобальны

    global source_directory # значение делаем глобальным
    global final_directory # значение делаем глобальным

    check_update = dict_settings["check_update"][0] # опция проверки обновления программы
    beta = dict_settings["beta"][0] # опция скачивания бета версии программы

    model_name = dict_settings["model_name"][0] # опция масового сохранения
    rewrite = dict_settings["rewrite"][0] # опция масового сохранения

    mass_saving = dict_settings["mass_saving"][0] # опция масового сохранения

    file_version = dict_settings["file_version"][0] # версия в какую сохранять файл
    file_version_name = dict_settings["file_version_name"][0] # запись версии в имя файла

    near_the_source = dict_settings["near_the_source"][0] # путь к папке, куда сохранять файлы (для открытых файлов)

    types_of_documents = dict_settings["types_of_documents"][0] # типы документов для обработки
    recursion = dict_settings["recursion"][0] # рекурсия (обработка папок внутри выбраной папки)

    source_directory = dict_settings["source_directory"][0] # путь к папке, откуда брать файлы (для сохранения из папки)
    final_directory = dict_settings["final_directory"][0] # путь к папке, куда сохранять файлы (для сохранения из папки)

    for parameter, val in dict_settings.items(): # для каждой строки производим обработку
        print(f"{parameter} = \"{val[0]}\"") # выводим прочитание параметры
##        print(f"{parameter} = \"{val[0]}\" #{val[1]}") # выводим прочитание параметры

def CheckUpdate(): # проверить обновление приложение

    global url # значение делаем глобальным

    if check_update: # если проверка обновлений включена

        try: # попытаться импортировать модуль обновления

            from Updater import Updater # импортируем модуль обновления

            if "url" not in globals(): # если нет ссылки на программу
                url = "" # нет ссылки

            Updater.Update(title, ver, beta, url, program_directory, Resource_path("cat.ico")) # проверяем обновление (имя программы, версия программы, скачивать бета версию, ссылка на программу, директория программы, путь к иконке)

        except SystemExit: # если закончили обновление (exit в Update)
            exit() # выходим из программы

        except: # не удалось
            pass # пропустить

def Resource_path(relative_path): # для сохранения картинки внутри exe файла

    try: # попытаться определить путь к папке
        base_path = sys._MEIPASS # путь к временной папки PyInstaller

    except Exception: # если ошибка
        base_path = os.path.abspath(".") # абсолютный путь

    return os.path.join(base_path, relative_path) # объеденяем и возващаем полный путь

def KompasAPI(): # подключение API компаса

    try: # попытаться подключиться к КОМПАСу

        global KompasConst # значение делаем глобальным
        global KompasAPI7 # значение делаем глобальным
        global iApplication # значение делаем глобальным
        global iKompasObject # значение делаем глобальным
        global iDocuments # значение делаем глобальным

        KompasConst3D = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0).constants # константа 3D документов
        KompasConst = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants # константа для скрытия вопросов перестроения и 2D документов

        KompasAPI5 = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0) # API5 КОМПАСа
        iKompasObject = Dispatch("Kompas.Application.5", None, KompasAPI5.KompasObject.CLSID) # интерфейс API КОМПАС

        KompasAPI7 = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0) # API7 КОМПАСа
        iApplication = Dispatch("Kompas.Application.7") # интерфейс приложения КОМПАС-3D.

        iKompasDocument = iApplication.ActiveDocument # получить текущий активный документ

        iDocuments = iApplication.Documents # интерфейс для открытия документов

        if iApplication.Visible == False: # если компас невидимый
            iApplication.Visible = True # сделать КОМПАС-3D видемым

    except: # если не получилось подключиться к КОМПАСу

        Message("КОМПАС-3D не найден!\nУстановите или переустановите КОМПАС-3D!") # сообщение, поверх всех окон с автоматическим закрытием
        exit() # выходим из программы

def Kompas_message(text): # сообщение в окне КОМПАСа если он открыт

    if iApplication.Visible == True: # если компас видимый
        iApplication.MessageBoxEx(text, 'Message:', 64) # сообщение в КОМПАСе

def Message(text = "Ошибка!", counter = 4): # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)

    def Message_Thread(text, counter): # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)

        if counter == 0: # время до закрытия окна (если 0)
            counter = 1 # закрытие через 1 сек
        window_msg = tk.Tk() # создание окна
        try: # попытаться использовать значок
            window_msg.iconbitmap(default = Resource_path("cat.ico")) # значок программы
        except: # если ошибка
            pass # пропустить
        window_msg.attributes("-topmost", True) # окно поверх всех окон
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

def AskYesNoCancel(text): # вопросительное сообщение, поверх всех окон

    def AskYesNoCancel_Thread(text): # сообщение, поверх всех окон в потоке

        global answer # для чтения в потоке

        ask = tk.Tk() # создание окна
        try: # попытаться использовать значёк
            ask.iconbitmap(default = Resource_path("cat.ico")) # значёк программы
        except: # если ошибка
            pass # пропустить
        ask.attributes("-topmost", True) # окно поверх всех окон
        ask.withdraw() # скрываем окно "невидимое"
        answer = mb.askyesnocancel(title, text, parent = ask) # задаём вопрос
        ask.destroy() # закрываем окно
        ask.mainloop() # отображение окна

    msg_th = Thread(target = AskYesNoCancel_Thread, args = (text, )) # запуск окна в отдельном потоке
    msg_th.start() # запуск потока

    msg_th.join() # ждать завершения процесса, иначе может закрыться следующие окно

    if answer == True or False: # если ответ "Да" или "Нет"
        return answer # возвращаем результат вопроса

    if answer == None: # если нажали отмена или крестик
        if iApplication.Visible == False: # если компас невидимый
            iApplication.Quit() # закрываем его
        exit() # выходим из программы

def AskYesNo(text): # вопросительное сообщение, поверх всех окон

    ask = tk.Tk() # создание окна
    try: # попытаться использовать значок
        ask.iconbitmap(default = Resource_path("cat.ico")) # значок программы
    except: # если ошибка
        pass # пропустить
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

def Window_filedialog(text): # создание окна filedialog

    window = tk.Tk() # создание окна что бы сделать filedialog поверх всех окон
    window.iconbitmap(default = Resource_path("cat.ico")) # значок программы
    window.attributes("-topmost",True) # окно поверх всех окон
    window.withdraw() # скрываем окно "невидимое"
    Path = filedialog.askdirectory(title = text) # запрос указания папки
    window.destroy() # закрытие окна от filedialog
    window = window.mainloop() # отображение окна

    if Path != "": # если путь выбран
        Path = Path.replace("/","\\") # меняем косую четру для единообразия
        return Path # возвращаем путь

    else: # если путь не выбран
        print("Путь не выбран!")
        exit()

def File_search(directory): # поиск файлов (путь к папке для поиска)

    allfilesPaths = [] # список всех найденых файлов

    type_doc = {1:".cdw", 2:".frw", 3:".spw", 4:".m3d", 5:".a3d", 6:".kdw", 7:".t3d"} # типы документов (Чертёж, Фрагмент, Спецификация, Деталь, Сборка, Текстовый документ, Технологическая сборка)

    type_doc_numbers = Iteration(types_of_documents) # обработка перечисленых цифр

    for n in type_doc_numbers: # перебор всех типов документов
        expansion = type_doc[n] # расширение в зависимости от типа документа

        if recursion: # если включена обработка всех папок
            expansion = glob.glob(directory + "/**/*" + expansion, recursive = True) # Получаем список всех файлов имеющихся в папке
        else: # только выбраная папка
            expansion = glob.glob(directory + "/*" + expansion, recursive = False) # Получаем список всех файлов имеющихся в папке

        allfilesPaths = allfilesPaths + expansion # добавляем найденый файлы в общий список

    if allfilesPaths == []: # если нет файлов

        Message("В указанной папке нет файлов!") # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)

        Kompas_message("В указанной папке нет документов!") # сообщение в окне КОМПАСа если он открыт

        if iApplication.Visible == False: # если КОМПАС невидимый
            iApplication.Quit() # закрываем КОМПАС

        exit() # выходим

    return allfilesPaths # возвращаем список

def Iteration(numbers): # обработка перечисленых цифр

    list_numbers = [] # список цифр

    numbers = numbers.split(",") # разделяем строку по ","

    for n in numbers: # обработка каждого элемента в списке

        if n.find("-") != -1: # если в элементе найден знак "-" обработать его
            n = n.split("-") # разделяем строку по "-"

            list_numbers = list_numbers + list(range(int(n[0]), int(n[1])+1)) # добавляем к списку цифр список целых числ из диапазона

        else: # без знака "-"
            if n != "0": # если значение не "0"
                list_numbers.append(int(n)) # добавляем целое число к списку цифр

            else: # введены неправельные значения
                Message("Введите правильное значение в Style_line!") # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)
                exit() # выходим из програмы

    return list_numbers # выводим список цифр

def File_processing(allfilesPaths, final_directory): # обработка файлов (путь к обрабатываемым файлам)

    global file_number # для чтения в потоке
    global current_file_name # для чтения в потоке
    global Stop # для чтения в потоке

    all_failes_number = len(allfilesPaths) # количество всех файлов в списке

    MessageСount(all_failes_number, "Идёт обработка файлов!") # выдача сообщений о количестве файлов (количество всех файлов, сообщение) + file_number (номер обрабатываемого файла) + current_file_name (текущее название файла)

    iApplication.HideMessage = KompasConst.ksHideMessageNo # скрыть сообщение перестроения и не перестраивать

    for file in allfilesPaths: # создаём цикл для работы с каждым файлом из списка

        if Stop == False: # если не нажата кнопка "Отмена" или крестик

            file_number += 1 # отчёт количества обработаных файлов

            current_file_name = os.path.basename(file) # имя документа с расширением для вывода названия файла в окно сообщений

            iKompasDocument = iDocuments.Open(file, False, False) # открытие файлов (False - в невидимом режиме, False - с возможностью редактирования)

            if iKompasDocument: # если документ открылся
                save = Save_file(False, final_directory) # сохраняем файл/файлы (True - массовое сохранение; False - одиночное, путь к папке куда сохранять файлы)

            else: # если документ не открылся
                print(file)
                list_not_open_files.append(file) # добавляем дет. в список
                continue

            if isinstance(save, bool) == False: # если файл не сохранён
                list_error_files.append(save) # добавляем дет. в список

            iKompasDocument.Close(0) # 0 - закрыть документ без сохранения; 1 - закрыть документ, сохранив  изменения; 2 - выдать запрос на сохранение документа, если он изменен.

        else: # если нажали кнопку "Отмена" или крестик
            print("Остановили окном!")
            break # прерываем цикл

    Stop = True # триггер остановки обработки и сообщения

    iApplication.HideMessage = KompasConst.ksShowMessage # показывать сообщение перестроения

    file_number = file_number - len(list_error_files) - len(list_not_open_files) # количество обработаных файлов без ошибок

    Message("Всего файлов сохранено: " + str(file_number)) # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)

def MessageСount(all_failes_number, msg = "Идёт обработка файлов!"): # выдача сообщений о количестве файлов (количество всех файлов, сообщение) + file_number (номер обрабатываемого файла) + current_file_name (текущее название файла)

    global Stop # глобальный параметр остановки сообщения
    global file_number # для чтения в потоке
    global current_file_name # для чтения в потоке

    def MessageСountThread(all_failes_number, msg): # сообщений о количестве файлов в потоке

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

        def UpdateText(): # обновление отчёта цифр

            def UpdatingText(): # обновление текста

                if Stop: # если триггер остановки обработки и сообщения включён
                    print("Остановил поток!")
                    window.quit() # закрываем окно

                else: # триггер выключен
                    text.config(text = str(file_number) + "/" + str(all_failes_number)) # обновляем текст
                    text.after(500, UpdatingText) # через милисекунды запускаем функцию заново

            UpdatingText() # обновление текста

        def UpdateProgress(): # обновление прогресса

            def UpdatingProgress(): # обновление прогресса

                percent_file_number = percent_all_failes_number * file_number # процент выполнения

                if Stop: # если триггер остановки обработки и сообщения включён
                    print("Остановил поток прогресса!")
                    window.quit() # закрываем окно

                else: # триггер выключен
                    progress['value'] = percent_file_number # процент выполнения
                    window.update() # (update_idletasks не сбрасывет дпока не дошёл до конца)
                    progress.after(500, UpdatingProgress) # через милисекунды запускаем функцию заново

            UpdatingProgress() # обновление прогресса

        def ButtonExit(): # кнопка "Отмена"
            window.quit() # закрываем окно

        window = tk.Tk() # создание окна
        try: # попытаться использовать значок
            window.iconbitmap(default = Resource_path("cat.ico")) # значок программы
        except: # если ошибка
            pass # пропустить
        window.title(title) # заголовок окна
        window.attributes("-topmost",True) # окно поверх всех окон
        x = (window.winfo_screenwidth() - window.winfo_reqwidth()) / 2 # положение по центру монитора
        y = (window.winfo_screenheight() - window.winfo_reqheight()) / 2 # положение по центру монитора
        window.wm_geometry("+%d+%d" % (x - 25, y + 38)) # положение по центру монитора
        window.resizable(width = False, height = False) # блокировка изменение размера окна

        f_top = tk.Frame(window) # блок окна (вверх)
        f_top.pack(expand = True, fill = "both") # размещение блока (с возможностью расширяться и заполненем окна во всех направлениях)

        text = tk.Label(f_top, justify=tk.LEFT, font = "Verdana 10", text = msg) # текст в окне
        text.pack(padx = 5, pady = 2) # размещение блока

        text = tk.Label(f_top, fg="green", justify=tk.LEFT, padx = 3, pady = 3, font = "Verdana 10") # текст
        ToolTip(text, current_file_name) # имя текущего файла в виде всплывающего окна
        UpdateText() # обновление отчёта цифр
        text.pack() # размещение блока

        progress = ttk.Progressbar(f_top, orient = "horizontal", length = 250, mode = 'determinate') # панель прогресса (положение, длина, вид отображения)
        percent_all_failes_number = 100/all_failes_number # перевод в процент от общего числа
        UpdateProgress() # обновление прогресса
        progress.pack(padx = 4) # размещение блока

        button = tk.Button(f_top, font = "Verdana 11", command = ButtonExit, text = "Отмена") # действие кнопки
        button.pack(side = "bottom", pady = 3) # размещение блока

        window.mainloop() # отображение окна

        Stop = True # триггер остановки обработки и сообщения

    Stop = False # триггер остановки сообщения (для работы сообщений при повторном вызове)
    file_number = 0 # отчёт от 0-го файла
    current_file_name = "" # что бы избежать ошибки окна в потоке

    msg_th = Thread(target = MessageСountThread, args = (all_failes_number, msg, )) # запуск сообщений о количестве файлов в отдельном потоке
    msg_th.start() # запуск потока

    return msg_th # возращаем запущеный поток для определения его завершения

def Save_file(mass_saving, Path): # сохраняем файл/файлы (True - массовое сохранение; False - одиночное, путь к папке куда сохранять файлы)

    global file_number # для чтения в потоке
    global current_file_name # для чтения в потоке
    global Stop # для чтения в потоке

    if mass_saving: # массовое сохранение включено

        all_failes_number = iDocuments.Count # количество открытых файлов

        print(all_failes_number)

        MessageСount(all_failes_number, "Идёт обработка файлов!") # выдача сообщений о количестве файлов (количество всех файлов, сообщение) + file_number (номер обрабатываемого файла) + current_file_name (текущее название файла)

        iApplication.HideMessage = KompasConst.ksHideMessageNo # скрыть сообщение перестроения и не перестраивать

        while iApplication.ActiveDocument: # пока есть открытый документ

            if Stop == False: # если не нажата кнопка "Отмена" или крестик

                file_number += 1 # отчёт количества обработаных файлов

                save = Saving_file(ver, mass_saving, Path) # сохраняем файл (версия файла, опция массового сохранения, путь к папке)

                if save and isinstance(save, bool): # если файл сохранён
                    Kompas_message("Файл сохранён в v" + str(ver[0])) # сообщение в окне КОМПАСа если он открыт

                else: # файл не сохранён
                    list_error_files.append(save) # добавляем дет. в список

            else: # если нажали кнопку "Отмена" или крестик
                print("Остановили окном!")
                break # прерываем цикл

        Stop = True # триггер остановки обработки и сообщения

        iApplication.HideMessage = KompasConst.ksShowMessage # показывать сообщение перестроения

        file_number = file_number - len(list_error_files) # количество обработаных файлов без ошибок

        Message("Всего файлов сохранено: " + str(file_number)) # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)

    else: # одиночное сохранение
        save = Saving_file(ver, mass_saving, Path) # сохраняем файл (версия файла, опция массового сохранения, путь к папке)

    return save # возврвщаем сохранён ли файл, если нет то выведет путь этого файла

def File_version(file_version): # определение в какую версию сохранять файл

    def Сhoosing_ver(ver): # округление версии сохранения (версия)

        if ver <= iKompasVersion: # если меньше текущей версии открытого компаса

            dictionary = {5.11:1, 6.0:2, 6.1:3, 7.0:4, 7.1:5, 8.0:6, 8.1:7, 9.0:8, 10.0:9, 11.0:10, 12.0:11, 13.0:12, 14.0:13, 14.1:14, 14.2:15, 15.0:16, 15.1:17,
                        16.0:19, 16.1:20, 17.0:21, 17.1:22, 18.0:23, 18.1:24, 19.0:25, 20.0:26, 21.0:27, 22.0:28, 23.0:29, 24.0:30} # версия файла и его номер сохраненияв КОМПАСе

            sorted_versions = sorted(dictionary.keys()) # сортиповка словаря
            max_ver = sorted_versions[-1] # максимальная версия
            max_code = dictionary[max_ver] # значение максимальной версии

            if ver > max_ver: # если версия больше максимальной
                major_ver = int(ver) # целая часть версии
                major_max = int(max_ver) # целая часть максимальной версии

                code = max_code + (major_ver - major_max) # для основных версий выше максимальной
                return (ver, code) # возвращаем версию файла (версия, её значение (saveMode))

            index = bisect_right(sorted_versions, ver) - 1 # поиск ближайшей меньшей или равной версии

            if index < 0: # если версия меньше минимальной
                return (sorted_versions[0], dictionary[sorted_versions[0]]) # возвращаем версию файла (версия, её значение (saveMode))

            return (sorted_versions[index], dictionary[sorted_versions[index]]) # возвращаем версию файла (версия, её значение (saveMode))

        else: # если необходимая версия больше чем версия программы, выдать сообщение
            text = "Сохранение в указанную версию невозможно!" # текст сообщения
            Kompas_message(text) # сообщение в окне КОМПАСа если он открыт
            Message(text) # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)
            exit()

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

    ver = Сhoosing_ver(ver) # округление версии сохранения (версия)

    return ver

def Path_or_question(Path, text): # вписан путь или нужно задать путь к папке

    if Path and isinstance(Path, bool): # если сохранять рядом с исходником
        Path = True # для работы в следующей функции
    else: # если не сохранять рядом с исходником
        if os.path.exists(str(Path)) == False: # если нет такого пути
            Path = Window_filedialog(text) # создание окна filedialog

    return Path # возврящаем путь к папке

def Path_or_question_for_directory(Path, text): # вписан путь или нужно задать путь к папке

    if Path == False and isinstance(Path, bool): # если сохранять рядом с исходником
        Path = Window_filedialog(text) # создание окна filedialog
    else: # если не сохранять рядом с исходником
        if os.path.exists(str(Path)) == False: # если нет такого пути
            Path = Window_filedialog(text) # создание окна filedialog

    return Path # возврящаем путь к папке

def Saving_file(ver, mass_saving, Path): # сохраняем файл (версия файла, опция массового сохранения, путь к папке)

    global Stop # для чтения в потоке

    iKompasDocument = iApplication.ActiveDocument # получить текущий активный документ

    resultPath, originalPath = ResultPath(iKompasDocument, model_name, Path, ver) # результирующий путь сохранения файла, по 3D модели или имени файла (интерфейс док., параметр обозн. и наим. из модели)

    iKompasDocument1 = KompasAPI7.IKompasDocument1(iKompasDocument) # интерфейс документа

    try: # попытаться сохранить

        iCount = iDocuments.Count # количество открытых вкладок

        if iKompasDocument1.SaveAsEx(resultPath, ver[1]): # сохранить в старую версию
            save = True # файл сохранён

        else: # если ошибка сохранения
            if os.path.exists(resultPath): # проверка существования пути
                os.remove(resultPath) # удаляем файл который создался но имеет 0 байт нужно для КОМПАС v19

            save = originalPath # файл не сохранён

    except: # не получилось сохранить

        iCount1 = iDocuments.Count # количество открытых вкладок
        if iCount1 > iCount: # если количество открытых вкладок увеличилось
            current_file_name = os.path.basename(originalPath) # имя документа с расширением для вывода названия файла в окно сообщений
            print(f"Ошибка сохранения файла \"{current_file_name}\. Перезапустите КОМПАС!")
            Message(f"Ошибка сохранения файла \"{current_file_name}\".\nИсключите файл из списка сохранения!\nДля правильной работы, перезапустите КОМПАС!", 10) # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)
            Stop = True # триггер остановки обработки и сообщения

        save = originalPath # файл не сохранён

    if isinstance(save, bool) == False: # если файл не сохранён
        Kompas_message("Файл не может быть сохранён в v" + str(ver[0])) # сообщение в окне КОМПАСа если он открыт

    if mass_saving: # массовое сохранение включено
        iKompasDocument.Close(0) # 0 - закрыть документ без сохранения; 1 - закрыть документ, сохранив  изменения; 2 - выдать запрос на сохранение документа, если он изменен.

    return save # возрящаем обработан ли файл

def ResultPath(iKompasDocument, model_name, Path, ver): # результирующий путь сохранения файла, по 3D модели или имени файла (интерфейс док., параметр обозн. и наим. из модели)

    def Read_name_and_obozn(iKompasDocument): # считывание обозначение и наименование 3D модели

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

        return name_and_obozn # возврящаем значения

    def Rename(Path, n): # изменение имени если уже есть такое

        fileBasename_and_expansion = os.path.splitext(Path) # список имя файла с расширением

        Path_temp = fileBasename_and_expansion[0] + " (" + str(n) + ")" + fileBasename_and_expansion[1] # добавляем номер файла

        if os.path.exists(Path_temp): # проверка наличия файла

            n += 1 # увеличиваем номер

            Path_temp = Rename(Path, n) # снова запускаем изменение имени со старым названием, но с новым номером

        return Path_temp # возвращаем последний прошедший проверку путь

    PathName = iKompasDocument.PathName # полное имя документа

    if Path and isinstance(Path, bool): # если сохранять рядом с исходником

        if PathName == "": # если нет имени вписать

            name_and_obozn = Read_name_and_obozn(iKompasDocument) # считывание обозначение и наименование 3D модели

            Message(f"Файл \"{name_and_obozn}\" не был сохранён, укажите папку сохранения") # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)

            Path = Window_filedialog("Выберите папку сохранения файлов") # создание окна filedialog

            Saving_file(ver, mass_saving, Path) # сохраняем файл (версия файла, опция массового сохранения, путь к папке)

        if file_version_name: # если запись версии включена
            resultPath = f"{PathName[:-4]}_v{str(ver[0])}{PathName[-4:]}" # добавляем запись версии файла

        else: # если запись версии выключена
            resultPath = PathName # не записываем версию

    else: # сохранение в папку

        if file_version_name: # если запись версии включена
            version = f"_v{str(ver[0])}" # добавляем запись версии файла
        else: # если запись версии выключена
            version = "" # не записываем версию

        fileBasename = os.path.basename(PathName) # имя документа с расширением
        fileBasename_and_expansion = os.path.splitext(fileBasename) # список имя файла с расширением
        basename = fileBasename_and_expansion[0] # имя файла (без расширения)
        expansion = fileBasename_and_expansion[1] # расширение файла

        if expansion == "": # если нет расширения (не сохранена деталь)
            type_doc = {1:".cdw", 2:".frw", 3:".spw", 4:".m3d", 5:".a3d", 6:".kdw", 7:".t3d"} # типы документов (Чертёж, Фрагмент, Спецификация, Деталь, Сборка, Текстовый документ, Технологическая сборка)
            expansion = type_doc[iKompasDocument.DocumentType] # имя в зависимости от типа документа

        if model_name and expansion in [".m3d", ".a3d", ".t3d"]: # обозначение и наименование брать из 3D модели
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

    if os.path.exists(resultPath): # проверка наличия файла
        if os.path.samefile(PathName, resultPath): # если совпадают пути оригинального и сохраняемого файла
            resultPath = Rename(PathName, 2) # изменение имени если уже есть такое (полный путь к файлу, номер с какого начинать заменять)

    return resultPath, PathName # возврящаем путь к файлу

def Error_files(list_error_files): # если есть не сохранённые файлы, задаём вопрос и открываем их

    global file_number # для чтения в потоке
    global current_file_name # для чтения в потоке
    global Stop # для чтения в потоке

    if len(list_error_files) > 0: # если есть не сохранённые файлы

        if AskYesNoCancel(f"Есть не сохранённые файлы!\nВсего файлов не сохраненно: {len(list_error_files)}\n\"Да\" - открыть файлы в КОМПАСе\n\"Нет\" - сохранить список файлов в txt."): # вопросительное сообщение, поверх всех окон

            if iApplication.Visible == False: # если компас невидимый
                iApplication.Visible = True # сделать компас видимым

            list_error_files.reverse() # список в обратном порядке

            all_failes_number = len(list_error_files) # количество всех файлов в списке

            msg_th = MessageСount(all_failes_number, "Идёт открытие файлов!") # выдача сообщений о количестве файлов (количество всех файлов, сообщение) + file_number (номер обрабатываемого файла) + current_file_name (текущее название файла)

            iApplication.HideMessage = KompasConst.ksHideMessageNo # скрыть сообщение перестроения

            for file in list_error_files: # создаём цикл для работы с каждым файлом

                if Stop == False: # если не нажата кнопка "Отмена" или крестик

                    file_number += 1 # отчёт количества обработаных файлов

                    current_file_name = os.path.basename(file) # имя документа с расширением для вывода названия файла в окно сообщений

                    iKompasDocument = iDocuments.Open(file, True, False) # Открытие файлов (False - в невидимом режиме, False - с возможностью редактирования)

                else: # если нажали кнопку "Отмена" или крестик
                    print("Остановили окном!")
                    break # прерываем цикл

            Stop = True # триггер остановки обработки и сообщения

            iApplication.HideMessage = KompasConst.ksShowMessage # показывать сообщение перестроения

            msg_th.join() # ждать завершения процесса, иначе может закрыться следующие окно

        else: # создаём txt файл

            Create_text_file(list_error_files, "Список не сохранённых файлов") # создать и открыть txt файл (список текста, название файла)

def Create_text_file(list, txt): # создать и открыть txt файл (список текста, название файла)

    unsaved_files = os.path.join(txt + ".txt") # название текстового файла

    with open(unsaved_files, "w+", encoding = "utf-8") as file: # открываем файл (с автоматическим закрытием) для записи (w+), для чтения (r), (невидимый режим)

        for line in list: # каждую строку
            file.write(line + "\n") # записываем текст в файл

    os.startfile(unsaved_files) # открываем файл в системе

def Not_open_files(): # если есть неоткрытые файлы задаём вопрос о создания списка

    if len(list_not_open_files) > 0: # если есть не открытые файлы
        if AskYesNo(f"Есть неоткрытые файлы!\nВсего файлов неоткрыто: {len(list_not_open_files)}\nСоздать список файлов?"): # вопросительное сообщение, поверх всех окон
            Create_text_file(list_not_open_files, "Список неоткрытых файлов") # создать и открыть txt файл (список текста, название файла)

#-------------------------------------------------------------------------------

list_error_files = [] # список не сохранённых файлов
list_not_open_files = [] # список не открытых файлов

DoubleExe() # проверка на уже запущеное приложение, с отключённым консольным окном "CREATE_NO_WINDOW"

if use_txt_file: # использовать txt файл
    Txt_file() # считываем значения настроек из txt файла

CheckUpdate() # обновить приложение

KompasAPI() # подключение API компаса

Сheck_version() # проверяем версию компаса

if iApplication.ActiveDocument: # проверяем открыт ли файл в КОМПАСе

    ver = File_version(file_version) # определение в какую версию сохранять файл

    Path = Path_or_question(near_the_source, "Выберите папку сохранения файлов") # вписан путь или нужно задать путь к папке

    if iDocuments.Count > 1: # количество открытых вкладок

        if mass_saving: # если массовое сохранение включено

            if AskYesNoCancel("Сохранить все документы?"): # вопросительное сообщение, поверх всех окон

                Save_file(True, Path) # сохраняем файл/файлы (True - массовое сохранение; False - одиночное, путь к папке куда сохранять файлы)

                Error_files(list_error_files) # если есть не сохранённые файлы, задаём вопрос и открываем их

            else: # сохраняем только один

                save = Save_file(False, Path) # сохраняем файл/файлы (True - массовое сохранение; False - одиночное, путь к папке куда сохранять файлы)

                if isinstance(save, bool) == False: # если файл сохранён
                    Message(f"Файл не может быть сохранён в необходимую версию.\nДля 3D попробуйте исключить из расчёта дерево постоения.", 6) # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)

        else: # сохраняем только один

            save = Save_file(False, Path) # сохраняем файл/файлы (True - массовое сохранение; False - одиночное, путь к папке куда сохранять файлы)

            if isinstance(save, bool) == False: # если файл сохранён
                Message(f"Файл не может быть сохранён в необходимую версию.\nДля 3D попробуйте исключить из расчёта дерево постоения.", 6) # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)

    else: # сохраняем только один

        save = Save_file(False, Path) # сохраняем файл/файлы (True - массовое сохранение; False - одиночное, путь к папке куда сохранять файлы)

        if isinstance(save, bool) == False: # если файл сохранён
            Message(f"Файл не может быть сохранён в необходимую версию.\nДля 3D попробуйте исключить из расчёта дерево постоения.", 6) # сообщение, поверх всех окон с автоматическим закрытием (текст, время закрытия)

else: # файл не открыт в КОМПАСе

    print("Сохранение из папки!")

    source_directory = Path_or_question_for_directory(source_directory, "Выберите папку с файлами") # вписан путь или нужно задать путь к папке

    allfilesPaths = File_search(source_directory) # поиск файлов (путь к папке для поиска)

    final_directory = Path_or_question_for_directory(final_directory, "Выберите папку куда сохранять файлы") # вписан путь или нужно задать путь к папке

    ver = File_version(file_version) # определение в какую версию сохранять файл

    File_processing(allfilesPaths, final_directory) # обработка файлов (путь к обрабатываемым файлам)

    Error_files(list_error_files) # если есть не сохранённые файлы, задаём вопрос и открываем их

    Not_open_files() # если есть неоткрытые файлы задаём вопрос о создания списка