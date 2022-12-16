#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      Каширских Дмитрий
#
# Created:     17.10.2022
# Copyright:   (c) Каширских Дмитрий 2022
# Licence:     <your licence>
#-------------------------------------------------------------------------------

def Resource_path(relative_path): # для сохранения картинки внутри exe файла

    import os # работа с файовой системой

    try: # попытаться определить путь к папке
        base_path = sys._MEIPASS # путь к временной папки PyInstaller

    except Exception: # если ошибка
        base_path = os.path.abspath(".") # абсолютный путь

    return os.path.join(base_path, relative_path) # объеденяем и возващаем полный путь

def message_count(all_failes_number, msg = "Идёт обработка файлов!"): # выдача сообщений о количестве файлов

    from threading import Thread # библиотека потоков

    global Stop_msg # глобальный параметр остановки сообщения

    Stop_msg = False # тригер остановки

    def message_count_Thread(all_failes_number, msg): # сообщений о количестве файлов в потоке

        import tkinter as tk # модуль окон
        import tkinter.ttk as ttk # модуль окон
        import time # модуль времени

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
        		tw.wm_attributes('-topmost', 1)
        		tw.wm_geometry('+%d+%d' % (self.widget.winfo_rootx(), self.widget.winfo_rooty() + self.widget.winfo_height() + 2))
        		tk.Label(tw, text = current_file_name, justify = 'left', bg = '#f2f2f2', relief = 'solid', bd = 1, font = "Verdana 10").pack()

        	def hidetip(self):
        		tw = self.tipwindow
        		if tw:
        			tw.destroy()
        		self.tipwindow = None

        def Updating_text(text): # обновление отчёта цифр

            def Update(): # обновление текста

                global Stop_msg # глобальный параметр остановки сообщения

                if Stop_msg: # если тригер остановки сообщения включён
                    print("Остановил поток!")
                    window.destroy() # закрываем окно
                    Stop_msg = False # для дальнейшей работы сообщений при повторном вызове

                else: # тригер выключен
                    text.config(text = str(file_number) + "/" + str(all_failes_number)) # обновляем текст
                    text.after(300, Update) # через милисекунды запускаем функцию заново

            Update() # обновление текста

        def Updating_progress(progress): # обновление прогресса

            percent_file_number = percent_all_failes_number * file_number # процент выполнения

            if percent_file_number != 100:
                progress['value'] = percent_file_number
                window.update() # (update_idletasks не сбрасывет дпока не дошёл до конца)
                time.sleep(0.2)
                Updating_progress(progress)
            else:
                print("100%!")
                window.destroy() # окно закрыть окно

        def button_exit(): # кнопка "Отмена"

            global Stop # глобальный параметр остановки

            window.destroy() # закрываем окно
            Stop = True # тригер остановки окна

        window = tk.Tk() # создание окна
        window.iconbitmap(default = Resource_path("cat.ico")) # значёк программы
        window.title(title) # заголовок окна
        window.attributes("-topmost",True) # окно поверх всех окон
        x = (window.winfo_screenwidth() - window.winfo_reqwidth()) / 2 # положение по центру монитора
        y = (window.winfo_screenheight() - window.winfo_reqheight()) / 2 # положение по центру монитора
        window.wm_geometry("+%d+%d" % (x-50, y)) # положение по центру монитора -50 из-за логотипа
##        window.geometry('200x100') # размер окна

##        logo = tk.PhotoImage(file = Resource_path("cat.png")) # картинка
##        logo = logo.subsample(1, 1) # мастаб картинки
##        tk.Label(window, image=logo).pack(side="right") # расположение картинки в окне

        f_top = tk.Frame(window) # блок окна (вверх)
        f_top.pack(expand = True, fill = "both") # размещение блока (с возможностью расширяться и заполненем окна во всех направлениях)

        text = tk.Label(f_top, justify=tk.LEFT, font = "Verdana 10", text = msg) # текст в окне
        text.pack(padx = 5, pady = 2) # размещение блока

        text = tk.Label(f_top, fg="green", justify=tk.LEFT, padx = 3, pady = 3, font = "Verdana 10") # текст
        ToolTip(text, current_file_name) # имя текущего файла в виде всплывающего окна
        Updating_text(text) # обновление отчёта цифр
        text.pack() # размещение блока

        progress = ttk.Progressbar(f_top, orient = "horizontal", length = 250, mode = 'determinate') # панель прогресса (положение, длина, вид отображения)
        percent_all_failes_number = 100/all_failes_number # перевод в процент от общего числа
        progress.pack(padx = 2) # размещение блока
        Updating_progress(progress) # обновление прогресса

        button = tk.Button(f_top, font = "Verdana 11", command = button_exit, text = "Отмена") # действие кнопки
        button.pack(side = "bottom", pady = 3) # размещение блока

##        Updating_progress() # обновление прогресса

        window.mainloop() # отображение окна

    msg_th = Thread(target = message_count_Thread, args = (all_failes_number, msg, )) # запуск определения положения мышки в отдельном потоке
    msg_th.start() # запуск потока
##    msg_th.join() # ждать завершения процесса, иначе может закрыться следующие окно

#-------------------------------------------------------------------------------

import time # модуль времени

title = "Тест панели прогресса"

all_failes_number = 100
file_number = 1
Stop = False # тригер остановки

message_count(all_failes_number, "Идёт обработка файлов!")

while file_number < all_failes_number:

    if Stop == False:
        file_number += 1
        current_file_name = "Имя фала " + str(file_number)
        print(file_number)
        time.sleep(0.1)



