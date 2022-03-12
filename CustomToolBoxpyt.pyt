# -*- coding: utf-8 -*-
# Устранение проблем с кодировкой UTF-8

import xlrd
import arcpy
import sys
reload(sys)
sys.setdefaultencoding('utf8')


class Toolbox(object):
    def __init__(self):
        """Define the toolbox (the name of the toolbox is the name of the .pyt file)."""
        self.label = "Toolbox"
        self.alias = ""

        # List of tool classes associated with this toolbox
        self.tools = [XYtoPolygonManagement]


class XYtoPolygonManagement(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "Создание полигонов из таблицы Excel XY"
        self.description = "Создание полигонов выполняется при помощи стандартных инструментов ArcPy"
        self.canRunInBackground = False

    def getParameterInfo(self):
        """Define parameter definitions"""

        in_excel = arcpy.Parameter(
            name='in_excel_file',
            displayName='Входной Excel-файл с координатами XY',
            datatype='DEFile',
            direction='Input',
            parameterType='Required')

        in_x = arcpy.Parameter(
            name='in_x_coord',
            displayName='Поле X',
            datatype='GPString',
            direction='Input',
            parameterType='Required')

        in_y = arcpy.Parameter(
            name='in_y_coord',
            displayName='Поле Y',
            datatype='GPString',
            direction='Input',
            parameterType='Required')

        in_line_field = arcpy.Parameter(
            name='in_line_field_opt',
            displayName='Поле линий',
            datatype='GPString',
            direction='Input',
            parameterType='Optional')

        in_line_close = arcpy.Parameter(
            name=' in_line_close_opt',
            displayName='Замкнуть линию',
            datatype='GPBoolean',
            direction='Input',
            parameterType='Optional')

        in_coord_system = arcpy.Parameter(
            name='in_cs',
            displayName='Система координат входных данных',
            datatype='GPCoordinateSystem',
            direction='Input',
            parameterType='Required')
        
        in_gedatabase = arcpy.Parameter(
            name='in_GDB_opt',
            displayName='Выходная база геоданных',
            datatype='DEWorkspace',
            direction='Input',
            parameterType='Optional')

        in_x .filter.type = "ValueList"
        in_y .filter.type = "ValueList"
        in_line_field .filter.type = "ValueList"
        in_excel.filter.list = ['xls']

        params = [in_excel, in_x, in_y, in_line_field, in_line_close, in_coord_system, in_gedatabase]
        return params

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""
        
        try:
            
            book = xlrd.open_workbook(parameters[0].valueAsText)  # получаем книгу Excel

            sh = book.sheet_by_index(0)  # Страница книги Excel с индексом 0

            listFields = []
            for headCells in range(sh.ncols):
                listFields.append(sh.cell(0, headCells).value)

            parameters[1].filter.list = listFields
            parameters[2].filter.list = listFields
            parameters[3].filter.list = listFields
            
        except Exception as err:

            arcpy.AddMessage("Ошибка: {0}".format(err))

        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""

        if not parameters[0].value:
            parameters[0].setErrorMessage(
                'Необходимо выбрать файл Excel с расширением .xls')

        if not parameters[1].value:
            parameters[1].setErrorMessage(
                'Необходимо выбрать поле с координатой X')

        if not parameters[2].value:
            parameters[2].setErrorMessage(
                'Необходимо выбрать поле с координатой Y')

        if not parameters[5].value:
            parameters[5].setErrorMessage(
                'Необходимо выбрать систему координат входных данных')

        return

    def execute(self, parameters, messages):
        """The source code of the tool."""

        try:

            # получаем текущий документ карты
            mxd = arcpy.mapping.MapDocument("CURRENT")
            
            # получаем фрейм данных с индексом=0
            dataFrame = arcpy.mapping.ListDataFrames(mxd)[0]

            # получаем текущую базу геоданных
            defaultGDB = arcpy.env.workspace
            
            # проверяем, установлена ли база геоданных, если нет устанавливаем по умолчанию
            if parameters[6].valueAsText:
                pathGDB = parameters[6].valueAsText
            else:
                pathGDB = defaultGDB
                
            # устанавливаем текущее рабочее пространство
            arcpy.env.workspace = pathGDB

            # получаем книгу Excel
            book = xlrd.open_workbook(
                parameters[0].valueAsText)
            
            # получаем страницу книги Excel с индексом=0
            sh = book.sheet_by_index(0)

            # формируем путь для загрузки данных из страницы Excel
            inExcelSheetName = parameters[0].valueAsText + '\\' + sh.name + '$'
            
            # формируем название точечных данных
            pointsLayerName = ''.join(sh.name.split()).replace('-', '_') + '_point'
            
            # формируем название точечного класса пространственных объектов
            pointsLayerNameFeature = ''.join(
                sh.name.split()).replace('-', '_') + '_pointFeature'
            
            # формируем название линейного класса пространственных объектов
            linesLayerNameFeature = pathGDB + '\\' + \
                ''.join(sh.name.split()).replace('-', '_') + '_lineFeature'
                
            # формируем название полигонального класса пространственных объектов
            polygonLayerNameFeature = pathGDB + '\\' + \
                ''.join(sh.name.split()).replace('-', '_') + '_polygonFeature'

            arcpy.AddMessage("\n-------------------------\n")
            arcpy.AddMessage("Текущий фрейм: {0}".format(dataFrame.name))
            arcpy.AddMessage("База геоданных: {0}".format(pathGDB))
            arcpy.AddMessage("Книга Excel: {0}".format(parameters[0].value))
            arcpy.AddMessage("Лист Excel: {0}".format(sh.name))
            arcpy.AddMessage("\n-------------------------\n")

            # проверяем, установлен ли параметр - Поле линий
            if parameters[3].valueAsText:
                lineField = parameters[3].valueAsText
            else:
                lineField = ""

            # проверяем, установлен ли параметр - Замкнуть линию
            if parameters[4]:
                lineClose = "CLOSE"
            else:
                lineClose = "NO_CLOSE"

            # преобразовываем точки из таблицы Excel в слой событий
            arcpy.MakeXYEventLayer_management(
                inExcelSheetName, parameters[1].valueAsText, parameters[2].valueAsText, pointsLayerName, parameters[5].valueAsText, "")

            # преобразовываем точечный слой событий в точечный класс объектов
            arcpy.conversion.FeatureClassToFeatureClass(
                pointsLayerName, pathGDB, pointsLayerNameFeature)

            # преобразовываем точечный класс объектов в линейный класс объектов
            arcpy.PointsToLine_management(
                pointsLayerNameFeature, linesLayerNameFeature, lineField, "", lineClose)

            # преобразовываем линейный класс объектов в полигональный класс объектов
            arcpy.FeatureToPolygon_management(
                linesLayerNameFeature, polygonLayerNameFeature)

            # преобразовываем полигональный класс объектов в полигональный слой
            addLayerPolygonToMap = arcpy.mapping.Layer(polygonLayerNameFeature)

            # добавляем полигональный слой на карту
            arcpy.mapping.AddLayer(dataFrame, addLayerPolygonToMap)

            arcpy.AddMessage("Удаляются временные слои... ")
            arcpy.AddMessage("\n-------------------------\n")

            # удаляем из базы геоданных промежуточные слои
            arcpy.Delete_management(pointsLayerNameFeature)
            arcpy.Delete_management(linesLayerNameFeature)

            arcpy.AddMessage(
                "Слой: {0} - добавлен на карту".format(addLayerPolygonToMap.name))
            arcpy.AddMessage("\n-------------------------\n")

            # устанавливаем экстент по добавленному полигональному слою
            dataFrame.extent = arcpy.mapping.ListLayers(
                mxd, addLayerPolygonToMap.name, dataFrame)[0].getExtent()

            # удаляем переменные
            del mxd, dataFrame, defaultGDB, pathGDB, book, sh

        except Exception as err:

            arcpy.AddMessage("Ошибка: {0}".format(err))

        return
