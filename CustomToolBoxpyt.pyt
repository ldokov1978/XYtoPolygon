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
        self.tools = [XYtoPolygonManagement, XYtoPolygon]


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
                arcpy.AddMessage("\n-------------------------\n")

                field_x = zero_row.index(parameters[1].valueAsText)
                field_y = zero_row.index(parameters[2].valueAsText)
                field_part = zero_row.index(parameters[3].valueAsText)

                # записываем все значения координат и частей полигонов в список
                arcpy.AddMessage(
                    "Записываем все значения координат и частей полигонов в список...")
                arcpy.AddMessage("\n-------------------------\n")

                for item in range(1, sh.nrows):
                    coord_list.append([sh.cell_value(rowx=item, colx=int(field_x)), sh.cell_value(
                        rowx=item, colx=int(field_y)), sh.cell_value(rowx=item, colx=int(field_part))])

                # записываем все значения из частей полигонов
                arcpy.AddMessage(
                    "Записываем все значения из частей полигонов...")
                arcpy.AddMessage("\n-------------------------\n")

                for item in coord_list:
                    parts_list.append(item[-1])

                # получаем уникальные значения частей полигонов
                arcpy.AddMessage(
                    "Получаем уникальные значения частей полигонов...")
                arcpy.AddMessage("\n-------------------------\n")

                uniq_parts_list = list(set(parts_list))

                # формируем полигоны по уникальным значениям частей полигонов
                arcpy.AddMessage(
                    "Формируем полигоны по уникальным значениям частей полигонов...")
                arcpy.AddMessage("\n-------------------------\n")
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
                arcpy.AddMessage("\n-------------------------\n")
            else:
                # иначе записываем только координаты X и Y

                # получаем индексы полей x, y
                arcpy.AddMessage("Получаем индексы полей...")
                arcpy.AddMessage("\n-------------------------\n")

                field_x = zero_row.index(parameters[1].valueAsText)
                field_y = zero_row.index(parameters[2].valueAsText)

                # записываем все значения координат в список
                arcpy.AddMessage(
                    "Записываем все значения координат в список...")
                arcpy.AddMessage("\n-------------------------\n")

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
                arcpy.AddMessage("\n-------------------------\n")

            #arcpy.AddMessage (itog_coord_list)

            # строим полигоны и добавляем на карту
            arcpy.AddMessage("Строим полигоны и добавляем на карту...")
            arcpy.AddMessage("\n-------------------------\n")

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
            arcpy.CopyFeatures_management(featureList, polygons_feature)
            
            # создаём класс пересекающих объектов
            arcpy.Intersect_analysis(
                polygons_feature, polygons_feature + "_intersect", "NO_FID", "", "INPUT")
            
            # агрегируем пересекающиеся объекты
            arcpy.AggregatePolygons_cartography(polygons_feature + "_intersect", polygons_feature + "_aggregate", "1 Meters")
            
            # вырезаем агрегированные объекты из основного класса
            arcpy.Erase_analysis(
                polygons_feature, polygons_feature + "_aggregate", polygons_feature + "_erase")

            # делаем слияние основного класса объектов и вырезанного класса объектов
            arcpy.Union_analysis(
                [polygons_feature + "_erase", polygons_feature + "_aggregate"], polygons_feature + "_result", "NO_FID", "", "")

            # создаём слой из результирующего класса объектов
            layer_to_map = arcpy.mapping.Layer(polygons_feature + "_result")

            # удаляем из базы промежутчные классы объектов
            arcpy.Delete_management(polygons_feature)
            arcpy.Delete_management(polygons_feature + "_intersect")
            arcpy.Delete_management(polygons_feature + "_aggregate")
            arcpy.Delete_management(polygons_feature + "_erase")

            arcpy.AddMessage("Удаляются временные слои... ")
            arcpy.AddMessage("\n-------------------------\n")

            # добавляем слой на карту
            arcpy.mapping.AddLayer(data_frame, layer_to_map, "BOTTOM")

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
