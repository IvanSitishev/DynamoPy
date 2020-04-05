#! /usr/bin/env python
# -*- coding: utf-8 -*-

import clr

clr.AddReference("System")
clr.AddReference("Microsoft.Office.Interop.Excel")

import System
from System.Collections.Generic import *
from System.Runtime.InteropServices import Marshal
from System.Reflection import Assembly
from System import Array
import Microsoft
from Microsoft import Office
import sys
pyt_path = r'C:\Program Files (x86)\IronPython 2.7\Lib'
sys.path.append(pyt_path)
import os

clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import *

clr.AddReference("DSCoreNodes")
import DSCore
from DSCore import *

clr.AddReference("RevitNodes")
import Revit
from Revit.Elements import *
clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)

clr.AddReference("RevitAPI")
import Autodesk
from Autodesk.Revit.DB import *

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import *

clr.AddReference('RevitServices')
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

class Param:
    def __init__(self):
        self.name = None
        self.parameterGroup = None
        self.famGroup = None
        self.isInst = None

def addParam(name, famGroup, parameterGroup, isInst):
    app.SharedParametersFilename = file
    def_file = app.OpenSharedParameterFile()
    def_groups = def_file.Groups
    group = def_groups.get_Item(famGroup)
    def_params = group.Definitions
    param = def_params.get_Item(name)
    familyManager.AddParameter(param, parameterGroup, isInst)

#Получение текущего интерфейса пользователья
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
#Получение текущего документа
doc = DocumentManager.Instance.CurrentDBDocument
familyManager = doc.FamilyManager
# Пользовательские параметры
file = UnwrapElement(IN[0]) #Файл общих параметров
paramList_path = IN[1]
paramList = []

# Параметр для работы с Excel
excelapp = Microsoft.Office.Interop.Excel.ApplicationClass()
excelapp.Visible = False
excelapp.DisplayAlerts = False
workbook = excelapp.Workbooks.Open(paramList_path)
worksheet = workbook.ActiveSheet


success = 0
error  = []

TransactionManager.Instance.EnsureInTransaction(doc)
i = 2
while worksheet.Cells[i, 1].Value2 != None:
    param = Param()
    param.name = worksheet.Cells[i, 1].Value2
    param.isInst = worksheet.Cells[i, 2].Value2
    param.famGroup = worksheet.Cells[i, 4].Value2
    exec("paramGoup = BuiltInParameterGroup.{}".format(worksheet.Cells[i, 5].Value2))
    param.parameterGroup = paramGoup
    try:
        addParam(param.name, param.famGroup, param.parameterGroup, param.isInst) #Добавление параметров
        success += 1
    except Exception as e:
        error.append("Имя параметра {0}. Ошибка: {1}".format(param.name, e))

    paramList.append(param)
    i += 1

TransactionManager.Instance.TransactionTaskDone()
# test = worksheet.Cells[10, 1].Value2



if len(error) == 0:
    TaskDialog.Show("Добавление параметров",
                    "Успешно добавлено {0} параметров".format(success))
else:
    TaskDialog.Show("Добавление параметров",
                    "Успешно добавлено {0} параметров. Не добавлено {1} парамметров".format(success, len(error)))

workbook.Close()

OUT = [i.parameterGroup for i in paramList] #familyManager.AddParameter(param, BuiltInParameterGroup.PG_IDENTITY_DATA, True)