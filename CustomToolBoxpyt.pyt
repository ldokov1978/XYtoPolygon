# -*- coding: utf-8 -*-
# Устранение проблем с кодировкой UTF-8

import xlrd
import xlwt
import arcpy
import os
import datetime
import subprocess
import sys
reload(sys)
sys.setdefaultencoding('utf8')


class Toolbox(object):
    def __init__(self):
        """Define the toolbox (the name of the toolbox is the name of the .pyt file)."""
        self.label = "Toolbox"
        self.alias = ""

        # List of tool classes associated with this toolbox
        self.tools = [XYtoPolygonManagement, XYtoPolygon, DMStoPointsTable]


class XYtoPolygonManagement(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "Создание полигонов из таблицы Excel XY (стандарт)"
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

        params = [in_excel, in_x, in_y, in_line_field,
                  in_line_close, in_coord_system, in_gedatabase]
        return params

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""

        try:

            book = xlrd.open_workbook(
                parameters[0].valueAsText)  # получаем книгу Excel

            sh = book.sheet_by_index(0)  # Страница книги Excel с индексом 0

            list_fields = []
            for head_cells in range(sh.ncols):
                list_fields.append(sh.cell(0, head_cells).value)

            parameters[1].filter.list = list_fields
            parameters[2].filter.list = list_fields
            parameters[3].filter.list = list_fields

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
            data_frame = arcpy.mapping.ListDataFrames(mxd)[0]

            # получаем текущую базу геоданных
            default_gdb = arcpy.env.workspace

            # проверяем, установлена ли база геоданных, если нет устанавливаем по умолчанию
            if parameters[6].valueAsText:
                path_gdb = parameters[6].valueAsText
            else:
                path_gdb = default_gdb

            # устанавливаем текущее рабочее пространство
            arcpy.env.workspace = path_gdb

            # получаем книгу Excel
            book = xlrd.open_workbook(
                parameters[0].valueAsText)

            # получаем страницу книги Excel с индексом=0
            sh = book.sheet_by_index(0)

            # формируем путь для загрузки данных из страницы Excel
            in_excel_sh = parameters[0].valueAsText + '\\' + sh.name + '$'

            # формируем название точечных данных
            points_data = ''.join(sh.name.split()).replace('-', '_') + '_point'

            # формируем название точечного класса пространственных объектов
            points_feature = ''.join(
                sh.name.split()).replace('-', '_') + '_pointFeature'

            # формируем название линейного класса пространственных объектов
            lines_feature = path_gdb + '\\' + \
                ''.join(sh.name.split()).replace('-', '_') + '_lineFeature'

            # формируем название полигонального класса пространственных объектов
            polygons_feature = path_gdb + '\\' + \
                ''.join(sh.name.split()).replace('-', '_') + '_polygonFeature'

            arcpy.AddMessage("\n-------------------------\n")
            arcpy.AddMessage("Текущий фрейм: {0}".format(data_frame.name))
            arcpy.AddMessage("База геоданных: {0}".format(path_gdb))
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
                in_excel_sh, parameters[1].valueAsText, parameters[2].valueAsText, points_data, parameters[5].valueAsText, "")

            # преобразовываем точечный слой событий в точечный класс объектов
            arcpy.conversion.FeatureClassToFeatureClass(
                points_data, path_gdb, points_feature)

            # преобразовываем точечный класс объектов в линейный класс объектов
            arcpy.PointsToLine_management(
                points_feature, lines_feature, lineField, "", lineClose)

            # преобразовываем линейный класс объектов в полигональный класс объектов
            arcpy.FeatureToPolygon_management(
                lines_feature, polygons_feature)

            # преобразовываем полигональный класс объектов в полигональный слой
            layer_to_map = arcpy.mapping.Layer(polygons_feature)

            arcpy.AddMessage("Удаляются временные слои... ")
            arcpy.AddMessage("\n-------------------------\n")

            # удаляем из базы геоданных промежуточные слои
            arcpy.Delete_management(points_data)
            arcpy.Delete_management(points_feature)
            arcpy.Delete_management(lines_feature)

            # добавляем полигональный слой на карту
            arcpy.mapping.AddLayer(data_frame, layer_to_map)

            # устанавливаем экстент по добавленному полигональному слою
            data_frame.extent = arcpy.mapping.ListLayers(
                mxd, layer_to_map.name, data_frame)[0].getExtent()

            arcpy.AddMessage(
                "Слой: {0} - добавлен на карту".format(layer_to_map.name))
            arcpy.AddMessage("\n-------------------------\n")

            # удаляем переменные
            del mxd, data_frame, default_gdb, path_gdb, book, sh

        except Exception as err:

            arcpy.AddMessage("Ошибка: {0}".format(err))

        return

# ---------------------------------------------------------------------------------------


class XYtoPolygon(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "Создание полигонов из таблицы Excel XY (таблица -> полигон)"
        self.description = "Создание полигонов выполняется напрямую из списка точек"
        self.canRunInBackground = False

    def getParameterInfo(self):
        """Define parameter definitions"""

        arcpy.env.overwriteOutput = True

        # 0 - входной файл Excel
        in_excel = arcpy.Parameter(
            name='in_excel_file',
            displayName='Входной Excel-файл с координатами XY',
            datatype='DEFile',
            direction='Input',
            parameterType='Required')

        # 1 - координата X
        in_x = arcpy.Parameter(
            name='in_x_coord',
            displayName='Поле X',
            datatype='GPString',
            direction='Input',
            parameterType='Required')

        # 2 - координата Y
        in_y = arcpy.Parameter(
            name='in_y_coord',
            displayName='Поле Y',
            datatype='GPString',
            direction='Input',
            parameterType='Required')

        # 3 - сформировать полигоны по значению
        in_part_field = arcpy.Parameter(
            name='in_line_field_opt',
            displayName='Сформировать полигоны по значению',
            datatype='GPString',
            direction='Input',
            parameterType='Optional')

        # 4 - замыкание координат
        in_close = arcpy.Parameter(
            name=' in_line_close_opt',
            displayName='Замкнуть координаты',
            datatype='GPBoolean',
            direction='Input',
            parameterType='Optional')

        # 5 - система координат входного файла
        in_coord_system = arcpy.Parameter(
            name='in_cs',
            displayName='Система координат входных данных',
            datatype='GPCoordinateSystem',
            direction='Input',
            parameterType='Required')

        # 6 - выходная база геоданных
        in_gedatabase = arcpy.Parameter(
            name='in_GDB_opt',
            displayName='Выходная база геоданных',
            datatype='DEWorkspace',
            direction='Input',
            parameterType='Optional')

        in_x.filter.type = "ValueList"
        in_y.filter.type = "ValueList"
        in_part_field.filter.type = "ValueList"
        in_excel.filter.list = ['xls', 'xlsx']

        params = [in_excel, in_x, in_y, in_part_field,
                  in_close, in_coord_system, in_gedatabase]

        return params

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""

        try:
            book = xlrd.open_workbook(
                parameters[0].valueAsText)  # получаем книгу Excel
            sh = book.sheet_by_index(0)  # страница книги Excel с индексом 0

            list_fields = []  # список полей в таблице

            # заполняем список полей
            for head_cells in range(sh.ncols):
                list_fields.append(sh.cell(0, head_cells).value)

            # передаём в параметры для выбора поля, полученный список полей
            parameters[1].filter.list = list_fields
            parameters[2].filter.list = list_fields
            parameters[3].filter.list = list_fields

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
            data_frame = arcpy.mapping.ListDataFrames(mxd)[0]

            # получаем текущую базу геоданных
            default_gdb = arcpy.env.workspace

            # проверяем, установлена ли база геоданных, если нет устанавливаем по умолчанию
            if parameters[6].valueAsText:
                path_gdb = parameters[6].valueAsText
            else:
                path_gdb = default_gdb

            # устанавливаем текущее рабочее пространство
            arcpy.env.workspace = path_gdb

            # получаем книгу Excel
            book = xlrd.open_workbook(parameters[0].valueAsText)

            # получаем страницу книги Excel с индексом=0
            sh = book.sheet_by_index(0)

            # формируем название полигонального класса пространственных объектов
            polygons_feature = path_gdb + '\\' + \
                ''.join(sh.name.split()).replace('-', '_') + '_polygonFeature'

            arcpy.AddMessage("\n-------------------------\n")
            arcpy.AddMessage("Текущий фрейм: {0}".format(data_frame.name))
            arcpy.AddMessage("База геоданных: {0}".format(path_gdb))
            arcpy.AddMessage("Книга Excel: {0}".format(parameters[0].value))
            arcpy.AddMessage("Лист Excel: {0}".format(sh.name))
            arcpy.AddMessage("\n-------------------------\n")

            zero_row = []  # создаём список для строки с названиями полей
            coord_list = []
            parts_list = []
            uniq_parts_list = []
            itog_coord_list = []

            # получаем список названий полей
            for zr in range(sh.ncols):
                zero_row.append(sh.cell_value(rowx=0, colx=zr))

            # в зависимости от параметра "сформировать полигоны..." запиываем
            # в список все значения координат из таблицы
            if parameters[3].valueAsText:
                # получаем индексы полей x, y и  part
                arcpy.AddMessage("Получаем индексы полей...")

                field_x = zero_row.index(parameters[1].valueAsText)
                field_y = zero_row.index(parameters[2].valueAsText)
                field_part = zero_row.index(parameters[3].valueAsText)

                # записываем все значения координат и частей полигонов в список
                arcpy.AddMessage(
                    "Записываем все значения координат и частей полигонов в список...")

                for item in range(1, sh.nrows):
                    coord_list.append([sh.cell_value(rowx=item, colx=int(field_x)), sh.cell_value(
                        rowx=item, colx=int(field_y)), sh.cell_value(rowx=item, colx=int(field_part))])

                # записываем все значения из частей полигонов
                arcpy.AddMessage(
                    "Записываем все значения из частей полигонов...")

                for item in coord_list:
                    parts_list.append(item[-1])

                # получаем уникальные значения частей полигонов
                arcpy.AddMessage(
                    "Получаем уникальные значения частей полигонов...")

                uniq_parts_list = list(set(parts_list))

                # формируем полигоны по уникальным значениям частей полигонов
                arcpy.AddMessage(
                    "Формируем полигоны по уникальным значениям частей полигонов...")
                for uniq_parts_item in uniq_parts_list:
                    total_list_coord = []

                    for coord_list_item in coord_list:
                        if coord_list_item[-1] == uniq_parts_item:
                            total_list_coord.append(coord_list_item)

                    # если параметр "замкнуть координаты" = true, добавляем в конец списка первую координату
                    if parameters[4]:
                        total_list_coord.append(total_list_coord[0])

                    # добавляем значения координат в итоговый список
                    itog_coord_list.append(total_list_coord)
                arcpy.AddMessage(
                    "Добавляем значения координат в итоговый список...")
            else:
                # иначе записываем только координаты X и Y

                # получаем индексы полей x, y
                arcpy.AddMessage("Получаем индексы полей...")

                field_x = zero_row.index(parameters[1].valueAsText)
                field_y = zero_row.index(parameters[2].valueAsText)

                # записываем все значения координат в список
                arcpy.AddMessage(
                    "Записываем все значения координат в список...")

                for item in range(1, sh.nrows):
                    coord_list.append([sh.cell_value(rowx=item, colx=int(
                        field_x)), sh.cell_value(rowx=item, colx=int(field_y))])

                # если параметр "замкнуть линию" = true, добавляем в конец списка первую координату
                if parameters[4]:
                    coord_list.append(coord_list[0])

                # добавляем значения координат в итоговый список
                itog_coord_list.append(coord_list)
                arcpy.AddMessage(
                    "Добавляем значения координат в итоговый список...")

            # строим полигоны и добавляем на карту
            arcpy.AddMessage("Строим полигоны и добавляем на карту...")

            point = arcpy.Point()
            array = arcpy.Array()
            featureList = []

            for feature in itog_coord_list:
                for coordPair in feature:
                    point.X = coordPair[0]
                    point.Y = coordPair[1]
                    array.add(point)

                polygon = arcpy.Polygon(array, parameters[5].valueAsText)
                array.removeAll()
                featureList.append(polygon)

            # создаём основной класс объектов
            arcpy.AddMessage(
                "Количество полигонов: {0}".format(len(featureList)))
            arcpy.CopyFeatures_management(featureList, polygons_feature)
            arcpy.AddMessage("Выходной объект: {0}".format(polygons_feature))

            if parameters[4].valueAsText:

                # создаём класс пересекающих объектов
                arcpy.Intersect_analysis(
                    polygons_feature, polygons_feature + "_intersect", "NO_FID", "", "INPUT")
                arcpy.AddMessage("Пересечение...")

                # агрегируем пересекающиеся объекты
                arcpy.AggregatePolygons_cartography(
                    polygons_feature + "_intersect", polygons_feature + "_aggregate", "1 Meters")
                arcpy.AddMessage("Агрегирование...")

                # вырезаем агрегированные объекты из основного класса
                arcpy.Erase_analysis(
                    polygons_feature, polygons_feature + "_aggregate", polygons_feature + "_erase")
                arcpy.AddMessage("Вырезание...")

                # делаем слияние основного класса объектов и вырезанного класса объектов
                arcpy.Union_analysis(
                    [polygons_feature + "_erase", polygons_feature + "_aggregate"], polygons_feature + "_result", "NO_FID", "", "")
                arcpy.AddMessage("Слияние...")

            else:
                arcpy.CopyFeatures_management(
                    polygons_feature, polygons_feature + "_result")

            # создаём слой из результирующего класса объектов

            layer_to_map = arcpy.mapping.Layer(polygons_feature + "_result")
            arcpy.AddMessage("Слой на карту - {0}".format(layer_to_map.name))

            # удаляем из базы промежутчные классы объектов
            arcpy.Delete_management(polygons_feature)
            arcpy.Delete_management(polygons_feature + "_intersect")
            arcpy.Delete_management(polygons_feature + "_aggregate")
            arcpy.Delete_management(polygons_feature + "_erase")

            arcpy.AddMessage("Удаляются временные слои... ")

            # добавляем слой на карту
            arcpy.mapping.AddLayer(data_frame, layer_to_map, "TOP")

            # устанавливаем экстент по добавленному полигональному слою
            data_frame.extent = arcpy.mapping.ListLayers(
                mxd, layer_to_map.name, data_frame)[0].getExtent()

            arcpy.AddMessage(
                "Слой: {0} - добавлен на карту".format(layer_to_map.name))
            arcpy.AddMessage("\n-------------------------\n")

            # удаляем переменные
            del mxd, data_frame, default_gdb, path_gdb

        except Exception as err:
            arcpy.AddMessage("Ошибка: {0}".format(err))
        return

# -------------------------------------------------


class DMStoPointsTable(object):

    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "Добавление полей ГМС в точечный слой"
        self.description = "Добавление полей ГМС в точечный слой"
        self.canRunInBackground = False

    def getParameterInfo(self):
        """Define parameter definitions"""

        # 0 - входной точечный слой
        points_layer = arcpy.Parameter(
            name='in_points_layer',
            displayName='Входной точечный слой',
            datatype='GPFeatureLayer',
            direction='Input',
            parameterType='Optional')

        # 1 - папка для выходного файла Excel
        excel_path = arcpy.Parameter(
            name='in_excel_path',
            displayName='Место сохранения файла Excel',
            datatype='DEFolder',
            direction='Input',
            parameterType='Optional')

        excel_path.value = os.path.dirname(
            arcpy.mapping.MapDocument("CURRENT").filePath)

        if not excel_path.valueAsText:
            excel_path.value = os.path.join(os.path.join(
                os.environ['USERPROFILE']), 'Documents')

        params = [points_layer, excel_path]

        return params

    def isLicensed(self):
        """Set whether tool is licensed to execute."""

        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""

        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""

        fc = parameters[0].valueAsText

        if fc:
            desc = arcpy.Describe(parameters[0].valueAsText)
            if desc.shapeType not in ('Point'):
                parameters[0].setErrorMessage(
                    'Необходимо выбрать точечный слой')

            if desc.spatialReference.type not in ('Geographic'):
                parameters[0].setErrorMessage(
                    'Необходимо выбрать слой с географической системой координат')

        else:
            parameters[0].setErrorMessage(
                'Необходимо выбрать точечный слой')

        return

    def execute(self, parameters, messages):
        """The source code of the tool."""

        desc = arcpy.Describe(parameters[0].valueAsText)

        arcpy.AddMessage("Name: {}".format(desc.name))
        arcpy.AddMessage("Shape type: {}".format(desc.shapeType))
        arcpy.AddMessage("Feature type: {}".format(desc.featureType))
        arcpy.AddMessage("Spatial Reference: {}".format(
            desc.spatialReference.name))
        arcpy.AddMessage("Spatial Reference Type: {}".format(
            desc.spatialReference.type))

        fields = ['SHAPE@', '_x', '_y', '_gSh', '_mSh',
                  '_sSh', '_gDl', '_mDl', '_sDl']

        try:
            arcpy.DeleteField_management(parameters[0].value, fields)

        except Exception as err:
            pass

        finally:
            for f in fields:
                if f in ('SHAPE@'):
                    pass
                elif f in ('_x', '_y'):
                    arcpy.management.AddField(
                        parameters[0].value, f, 'DOUBLE', '255', 12)
                elif f in ('_sSh', '_sDl'):
                    arcpy.management.AddField(
                        parameters[0].value, f, 'DOUBLE', '255', 4)
                else:
                    arcpy.management.AddField(
                        parameters[0].value, f, 'SHORT')

        with arcpy.da.UpdateCursor(parameters[0].value, fields) as cursor:
            for row in cursor:
                for pnt in row[0]:
                    row[1] = pnt.X
                    row[2] = pnt.Y
                    row[3] = self.coord(pnt.Y)[0]
                    row[4] = self.coord(pnt.Y)[1]
                    row[5] = self.coord(pnt.Y)[2]
                    row[6] = self.coord(pnt.X)[0]
                    row[7] = self.coord(pnt.X)[1]
                    row[8] = self.coord(pnt.X)[2]

                cursor.updateRow(row)

        # Обновление атрибутивной таблицы
        layer = parameters[0].value
        definition_query = layer.definitionQuery
        if definition_query == '':
            oid = arcpy.ListFields(dataset=layer, field_type='OID')[0]
            layer.definitionQuery = '{} > 0'.format(oid.name)
        else:
            layer.definitionQuery = ''

        arcpy.RefreshActiveView()

        layer.definitionQuery = definition_query
        arcpy.RefreshActiveView()

        # Excel-файл

        # Создание стилей для ячеек Excel
        alignment = xlwt.Alignment()
        alignment.horz = xlwt.Alignment.HORZ_CENTER
        alignment.vert = xlwt.Alignment.VERT_CENTER

        border = xlwt.Borders()
        border.left = xlwt.Borders.THIN
        border.right = xlwt.Borders.THIN
        border.top = xlwt.Borders.THIN
        border.bottom = xlwt.Borders.THIN

        font_head = xlwt.Font()
        font_head.name = 'Calibri'
        font_head.height = 220  # высота - 11 (20 * 11 = 220)
        font_head.colour_index = 0
        font_head.bold = True
        font_head.italic = True

        font_cells = xlwt.Font()
        font_cells.name = 'Calibri'
        font_cells.height = 220  # высота - 11 (20 * 11 = 220)
        font_cells.colour_index = 0
        font_cells.bold = False
        font_cells.italic = False

        style_head = xlwt.XFStyle()
        style_head.font = font_head
        style_head.alignment = alignment
        style_head.borders = border

        style_cells = xlwt.XFStyle()
        style_cells.font = font_cells
        style_cells.alignment = alignment
        style_cells.borders = border

        # Создание книги и страницы Excel
        wb = xlwt.Workbook()
        ws = wb.add_sheet('coord_' + desc.spatialReference.name,
                          cell_overwrite_ok=True)

        # Создание шапки на странице Excel
        ws .write_merge(
            0, 0, 0, 6, 'Система координат - ' + desc.spatialReference.name, style_head)
        ws.write_merge(1, 2, 0, 0, u'№ п/п', style_head)
        ws.write_merge(1, 1, 1, 3, u'Северная широта', style_head)
        ws.write_merge(1, 1, 4, 6, u'Восточная долгота', style_head)
        ws.write(2, 1, u'гр.', style_head)
        ws.write(2, 2, u'мин.', style_head)
        ws.write(2, 3, u'сек.', style_head)
        ws.write(2, 4, u'гр.', style_head)
        ws.write(2, 5, u'мин.', style_head)
        ws.write(2, 6, u'сек.', style_head)

        # Заполнение таблицы Excel
        fields_excel = ['_gSh', '_mSh', '_sSh', '_gDl', '_mDl', '_sDl']

        row_idx = 3
        num = 1

        with arcpy.da.SearchCursor(parameters[0].value, fields_excel) as excel_cursor:
            for row in excel_cursor:
                for col in range(len(fields_excel)):
                    # Запись в первое поле номера точки
                    ws.write(row_idx, 0, num, style_cells)
                    # Запись значений координат в ячейки
                    ws.write(row_idx, col + 1, row[col], style_cells)
                row_idx += 1
                num += 1

        # путь к файлу Excel
        path_excel = desc.name + '_' + datetime.datetime.now().strftime('%d_%m_%Y_%H_%M_%S') + \
            '_' + desc.spatialReference.name + '.xls'

        wb.save(os.path.join(parameters[1].valueAsText, path_excel))
        subprocess.Popen(os.path.join(
            parameters[1].valueAsText, path_excel), shell=True)

        return

    def coord(self, dec_coord):
        _deg = int(dec_coord)
        _min = int((dec_coord - int(dec_coord)) * 60)
        _sec = (((dec_coord - int(dec_coord)) * 60) -
                int((dec_coord - int(dec_coord)) * 60)) * 60

        return [_deg, _min, _sec]
