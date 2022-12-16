#-------------------------------------------------------------------------------
# Name:        Сохранение в старые версии
# Purpose:
#
# Author:      dimak222
#
# Created:     21.01.2022
# Copyright:   (c) dimak222 2022
# Licence:     No
#-------------------------------------------------------------------------------

##import ctypes
##ctypes.windll.user32.ShowWindow( ctypes.windll.kernel32.GetConsoleWindow(), 6 )

def KompasAPI(): # подключение API компаса

    import pythoncom
    from win32com.client import Dispatch, gencache
    from sys import exit

    try:
        global KompasAPI7
        global iApplication
        global iKompasObject
        global iKompasDocument

        KompasConst3D = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0).constants

        KompasAPI5 = gencache.EnsureModule('{0422828C-F174-495E-AC5D-D31014DBBE87}', 0, 1, 0)
        iKompasObject = Dispatch('Kompas.Application.5', None, KompasAPI5.KompasObject.CLSID)

        KompasAPI7 = gencache.EnsureModule('{69AC2981-37C0-4379-84FD-5DD2F3C0A520}', 0, 1, 0)
        iApplication = Dispatch('Kompas.Application.7')

        iKompasDocument = iApplication.ActiveDocument

        if iApplication.Visible == False: # если компас невидимый
            iApplication.Visible = True # сделать КОМПАС-3D видемым

    except:
        message(0, "КОМПАС-3D не найден!\nУстановите или переустановите КОМПАС-3D!")
        exit()

#-------------------------------------------------------------------------------

KompasAPI() # подключение API компаса

ver = iKompasObject.ksGetSystemVersion() # текущая версия компаса
ver = ver[1]+ver[2]*0.1

iKompasDocument = iApplication.ActiveDocument
iKompasDocument1 =KompasAPI7.IKompasDocument1(iKompasDocument)

PathName = iKompasDocument.PathName
if PathName == "":
    type_doc = {1:"Чертёж.cdw", 2:"Фрагмент.frw", 3:"Спецификация.spw", 4:"Деталь.m3d", 5:"Сборка.a3d", 6:"Текстовый документ.kdw", 7:"Технологическая сборка.t3d"}
    PathName = type_doc[iKompasDocument.DocumentType]

v = iKompasObject.ksReadString ("Введите номер версии КОМПАС-3D:", str(ver)).replace(",",".").replace(" ","")

try:
    if float(v) >= 5.0 and float(v) <= ver:
        dictionary = {5.11:1, 6.0:2, 6.1:3, 7.0:4, 7.1:5, 8.0:6, 8.1:7, 9.0:8, 10.0:9, 11.0:10, 12.0:11, 13.0:12, 14.0:13, 14.1:14, 14.2:15, 15.0:16, 15.1:17, 16.0:19, 16.1:20, 17.0:21, 17.1:22, 18.0:23, 18.1:24, 19.0:25, 20.0:26}
        list_ver = [5.11, 6.0, 6.1, 7.0, 7.1, 8.0, 8.1, 9.0, 10.0, 11.0, 12.0, 13.0, 14.0, 14.1, 14.2, 15.0, 15.1, 16.0, 16.1, 17.0, 17.1, 18.0, 18.1, 19.0, 20.0]
        for i in range(len(list_ver)):
            if float(v) < list_ver[i]:
                if i == 0:
                    v = 5.11
                    break
                v = list_ver[i-1]
                break
            elif float(v) == list_ver[i]:
                break
    else:
        v = None
        iApplication.MessageBoxEx( "Сохранение в указанную\nверсию невозможно!", "Отчёт:", 64)

    if v:
        iKompasDocument1.SaveAsEx(PathName[:-4] + "_v" + str(float(v)) + PathName[-4:], dictionary[float(v)])
        iApplication.MessageBoxEx( "Файл сохранён в версию " + str(float(v)), "Отчёт:", 64)

except:
    iApplication.MessageBoxEx( "Данные введены некорректно!", "Отчёт:", 64)