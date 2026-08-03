Imports Topomatic.Dwg.Layer
Imports Topomatic.Cad.View
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports System.IO
Public Class FormInsertPointPPT
    Public ActivDocument As Topomatic.Dwg.Drawing = Nothing
    Public boolShow As Boolean = False
    Public arrayPutch As String() = Nothing
    Public templateXML As String = ""
    Public templateDWG As String = ""
    Public generalDir As String = ""
    Public Sub New(ByRef cadWiever As CadView)
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        'получаем активный документ
        If IsNothing(cadWiever) Then Return
        'получаем активный слой
        Dim DrawLayer As DrawingLayer = DrawingLayer.GetDrawingLayer(cadWiever)
        If IsNothing(DrawLayer) = True Then Return
        'получаем активный документ
        ActivDocument = DrawLayer.Drawing
        If IsNothing(ActivDocument) = True Then Return
        'сброс выделения
        cadWiever.SelectionSet.Clear()
        'сброс выделения
        '1.\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim arrayDirSupport As String() = Nothing
        Dim boolFindDirSupport As Boolean = FuncFiles.funcReadDirectoriesSupport(arrayDirSupport)
        If IsArray(arrayDirSupport) = True Then
            For i As Integer = 0 To arrayDirSupport.Length - 1
                Dim nameFolder As String = arrayDirSupport(i)
                Dim pos As Integer = nameFolder.IndexOf("InfrastradaToolsUtility")
                If pos > -1 Then
                    generalDir = Mid(nameFolder, 1, pos)
                    generalDir = generalDir & "InfrastradaToolsUtility"
                    Exit For
                End If
            Next
        End If
        'шаблон оформления
        Dim directorySupport As String = generalDir & "\FileResources\Sample\"
        If Directory.Exists(directorySupport) = True Then
            Dim xmlFiles As String() = Directory.GetFiles(directorySupport, "*.xml")
            If IsArray(xmlFiles) = True Then
                For i As Integer = 0 To xmlFiles.Length - 1
                    templateXML = xmlFiles(i)
                    Dim shortNameFile As String = IO.Path.GetFileNameWithoutExtension(templateXML)
                    Dim tempNameFileDwg As String = directorySupport & shortNameFile & ".dwg"
                    If IO.File.Exists(tempNameFileDwg) = True Then
                        templateDWG = tempNameFileDwg
                        Exit For
                    End If
                Next
            End If
        End If
        If IsNothing(templateXML) = True Or templateXML.Trim.Length = 0 Then
            MsgBox("Файл с шаблонами =Стандарт_ADSK.xml= не найден. Пожалуйста, добавьте его в пути доступа к вспомогательным файлам!!!")
            Exit Sub
        End If
        arrayPutch = {templateXML, templateDWG}
        'заполням поле с категорией
        Dim arrayCategory As String() = Nothing
        Dim boolFindCategory As Boolean = FuncXML.FuncReadCategoryToXML(templateXML, arrayCategory)
        If IsArray(arrayCategory) = True Then
            ComboBox4.DataSource = arrayCategory
            For i As Integer = 0 To arrayCategory.Length - 1
                Dim categoryName As String = arrayCategory(i)
                If categoryName Like "Проект планировки и межевания территории" Then
                    ComboBox4.Text = "Проект планировки и межевания территории"
                End If
            Next i
            'заполняем таблицы
            Dim arrayTables As String() = Nothing
            Dim boolFindTablesXML As Boolean = FuncXML.FuncReadTablesXMLByCategory(templateXML, ComboBox4.Text.Trim, arrayTables)
            If IsArray(arrayTables) = True Then
                ComboBox1.DataSource = arrayTables
                For i As Integer = 0 To arrayTables.Length - 1
                    Dim tableName As String = arrayTables(i)
                    If tableName Like "Земельные участки образуемые" Then
                        ComboBox1.Text = "Земельные участки образуемые"
                    End If
                Next i
                'заполняем блоки
                Dim arrayBlkDesk As String() = Nothing
                Dim boolFindBlk As Boolean = FuncXML.FuncReadDeskObjectByCategoryAndTablesToXML(templateXML, ComboBox4.Text.Trim, ComboBox1.Text.Trim, "Блок", arrayBlkDesk)
                If IsArray(arrayBlkDesk) = True Then
                    ComboBox5.DataSource = arrayBlkDesk
                    For i As Integer = 0 To arrayBlkDesk.Length - 1
                        Dim deskBlk As String = arrayBlkDesk(i)
                        If deskBlk Like "Характерная точка границы образуемого земельного участка" Then
                            ComboBox5.Text = "Характерная точка границы образуемого земельного участка"
                        End If
                    Next i
                End If
            End If
        End If
        '==========================================================================================================================
        'ищем характеристики
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'смотрим характеристики блока
        Dim categoryPoint As String = ComboBox4.Text.Trim
        Dim tablePoint As String = ComboBox1.Text.Trim
        Dim deskPoint As String = ComboBox5.Text.Trim
        Dim arrayPropertiesObject As String(,) = Nothing
        Dim boolPropObject As Boolean = FuncXML.FuncReadDataObjectToXMLbyDesk(templateXML, categoryPoint, tablePoint, "Блок", deskPoint, arrayPropertiesObject)
        'забираем в массив все блоки с чертежа
        If IsArray(arrayPropertiesObject) = True Then
            'имя блока
            Dim nameBlk As String = FuncXML.FuncReturnValueByArrayProperties("Name", arrayPropertiesObject, 1)
            'вставляем его в чертеж
            Dim blkPK As DwgBlock = RoburFunc.FuncInsertBlockByDWGFile(ActivDocument, templateDwg, nameBlk)
            Dim arrayBLK As String() = Nothing
            Dim boolFindBlkProperties1 As Boolean = RoburFunc.FuncReadNameInsertBlocks(ActivDocument, arrayBLK)
            If IsArray(arrayBLK) = True Then
                ComboBox2.DataSource = arrayBLK
                For i As Integer = 0 To arrayBLK.Length - 1
                    If arrayBLK(i) Like nameBlk Then
                        ComboBox2.Text = nameBlk
                    End If
                Next i
            End If
        End If
        'слой
        Dim arrayLayers As String() = Nothing
        If IsArray(arrayPropertiesObject) = True Then
            Dim nameLayer As String = FuncXML.FuncReturnValueByArrayProperties("Layer", arrayPropertiesObject, 1)
            If nameLayer.Trim.Length > 0 Then
                ReDim arrayLayers(0)
                arrayLayers(0) = nameLayer
            End If
        End If
        Dim boolFindBlkProperties2 As Boolean = RoburFunc.FuncReadLayers(ActivDocument, arrayLayers)
        ComboBox3.DataSource = arrayLayers
        'цвет
        Dim colorBlk As String = FuncXML.FuncReturnValueByArrayProperties("Color", arrayPropertiesObject, 1)
        TextBox2.Text = colorBlk
        'размер блока
        Dim ScaleBlk As Double = Val(FuncXML.FuncReturnValueByArrayProperties("Scale", arrayPropertiesObject, 1))
        If ScaleBlk <= 0 Then ScaleBlk = 1
        NumericUpDown1.Value = ScaleBlk
        'размер атрибута
        Dim attribHeight As Double = Val(FuncXML.FuncReturnValueByArrayProperties("AttributeHeight", arrayPropertiesObject, 1))
        If attribHeight <= 0 Then attribHeight = 1.5
        NumericUpDown2.Value = attribHeight
        'расширенные данные
        Dim arrayPs As String() = Nothing
        Dim boolFindPs As Boolean = FuncXML.FuncReadPSTablesToXML(templateXML, arrayPs)
        If IsArray(arrayPs) = True Then
            ComboBox7.DataSource = arrayPs
            Dim PSTable As String = FuncXML.FuncReturnValueByArrayProperties("PS", arrayPropertiesObject, 1)
            If PSTable.Trim.Length > 0 Then
                For i As Integer = 0 To arrayPs.Length - 1
                    Dim PsBlk As String = arrayPs(i)
                    If PsBlk Like PSTable Then
                        ComboBox7.Text = PSTable
                        Exit For
                    End If
                Next i
            End If
        End If
        '======================================================================================================================
        'загружаем блок
        Dim namePoint As String = ComboBox2.Text
        Dim blkPK1 As DwgBlock = ActivDocument.Blocks.Item(namePoint.Trim)
        'читаем атрибуты блока
        Dim arrayAttribute As String() = Nothing
        Dim boolReadBlk As Boolean = RoburFunc.FuncReadNameAttribBlk(ActivDocument, blkPK1, arrayAttribute)
        If IsArray(arrayAttribute) = True Then
            ComboBox6.DataSource = arrayAttribute
            For i As Integer = 0 To arrayAttribute.Length - 1
                Dim nameAttrib As String = arrayAttribute(i)
                If nameAttrib Like "NUM_GEOPOINT" Then
                    ComboBox6.Text = "NUM_GEOPOINT"
                    Exit For
                End If
            Next i
        Else
            ReDim arrayAttribute(0)
            arrayAttribute(0) = ""
            ComboBox6.DataSource = arrayAttribute
        End If
        '======================================================================================
        'префикс номера
        Dim arrayPreficsNimber As String() = {"", "н", "т"}
        ComboBox8.DataSource = arrayPreficsNimber
        '===============================================================================================================
        'закрепление точки
        Dim PointZakr As String() = {"Закрепление отсутствует", "Временный межевой знак", "Долговременный межевой знак"}
        ComboBox10.DataSource = PointZakr
        'метод определения координат
        Dim MetodCoord As String() = {"Геодезический метод", "Метод спутниковых геодезических измерений", "Аналитический метод", "Картометрический метод", "Фотограмметрический метод", "Иное описание"}
        ComboBox9.DataSource = MetodCoord
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'отмена
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        boolShow = False
        Me.Close()
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'OK
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        boolShow = True
        Me.Hide()
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'изменить категорию
    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        Dim templateXML As String = arrayPutch(0)
        Dim templateDwg As String = arrayPutch(1)
        'заполняем таблицы
        Dim arrayTables As String() = Nothing
        Dim boolFindTablesXML As Boolean = FuncXML.FuncReadTablesXMLByCategory(templateXML, ComboBox4.Text.Trim, arrayTables)
        If IsArray(arrayTables) = True Then
            ComboBox1.DataSource = arrayTables
            'заполняем блоки
            Dim arrayBlkDesk As String() = Nothing
            Dim boolFindBlk As Boolean = FuncXML.FuncReadDeskObjectByCategoryAndTablesToXML(templateXML, ComboBox4.Text.Trim, ComboBox1.Text.Trim, "Блок", arrayBlkDesk)
            If IsArray(arrayBlkDesk) = True Then
                ComboBox5.DataSource = arrayBlkDesk
            End If
        End If
        '==========================================================================================================================
        'ищем характеристики
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'смотрим характеристики блока
        Dim categoryPoint As String = ComboBox4.Text.Trim
        Dim tablePoint As String = ComboBox1.Text.Trim
        Dim deskPoint As String = ComboBox5.Text.Trim
        Dim arrayPropertiesObject As String(,) = Nothing
        Dim boolPropObject As Boolean = FuncXML.FuncReadDataObjectToXMLbyDesk(templateXML, categoryPoint, tablePoint, "Блок", deskPoint, arrayPropertiesObject)
        'забираем в массив все блоки с чертежа
        If IsArray(arrayPropertiesObject) = True Then
            'имя блока
            Dim nameBlk As String = FuncXML.FuncReturnValueByArrayProperties("Name", arrayPropertiesObject, 1)
            'вставляем его в чертеж
            Dim blkPK As DwgBlock = RoburFunc.FuncInsertBlockByDWGFile(ActivDocument, templateDwg, nameBlk)
            Dim arrayBLK As String() = Nothing
            Dim boolFindBlkProperties1 As Boolean = RoburFunc.FuncReadNameInsertBlocks(ActivDocument, arrayBLK)
            If IsArray(arrayBLK) = True Then
                ComboBox2.DataSource = arrayBLK
                For i As Integer = 0 To arrayBLK.Length - 1
                    If arrayBLK(i) Like nameBlk Then
                        ComboBox2.Text = nameBlk
                    End If
                Next i
            End If
        End If
        'слой
        Dim arrayLayers As String() = Nothing
        If IsArray(arrayPropertiesObject) = True Then
            Dim nameLayer As String = FuncXML.FuncReturnValueByArrayProperties("Layer", arrayPropertiesObject, 1)
            If nameLayer.Trim.Length > 0 Then
                ReDim arrayLayers(0)
                arrayLayers(0) = nameLayer
            End If
        End If
        Dim boolFindBlkProperties2 As Boolean = RoburFunc.FuncReadLayers(ActivDocument, arrayLayers)
        ComboBox3.DataSource = arrayLayers
        'цвет
        Dim colorBlk As String = FuncXML.FuncReturnValueByArrayProperties("Color", arrayPropertiesObject, 1)
        TextBox2.Text = colorBlk
        'размер блока
        Dim ScaleBlk As Double = Val(FuncXML.FuncReturnValueByArrayProperties("Scale", arrayPropertiesObject, 1))
        If ScaleBlk <= 0 Then ScaleBlk = 1
        NumericUpDown1.Value = ScaleBlk
        'размер атрибута
        Dim attribHeight As Double = Val(FuncXML.FuncReturnValueByArrayProperties("AttributeHeight", arrayPropertiesObject, 1))
        If attribHeight <= 0 Then attribHeight = 1.5
        NumericUpDown2.Value = attribHeight
        'расширенные данные
        Dim arrayPs As String() = Nothing
        Dim boolFindPs As Boolean = FuncXML.FuncReadPSTablesToXML(templateXML, arrayPs)
        If IsArray(arrayPs) = True Then
            ComboBox7.DataSource = arrayPs
            Dim PSTable As String = FuncXML.FuncReturnValueByArrayProperties("PS", arrayPropertiesObject, 1)
            If PSTable.Trim.Length > 0 Then
                For i As Integer = 0 To arrayPs.Length - 1
                    Dim PsBlk As String = arrayPs(i)
                    If PsBlk Like PSTable Then
                        ComboBox7.Text = PSTable
                        Exit For
                    End If
                Next i
            End If
        End If
        '======================================================================================================================
        'загружаем блок
        Dim namePoint As String = ComboBox2.Text
        Dim blkPK1 As DwgBlock = RoburFunc.FuncInsertBlockByDWGFile(ActivDocument, templateDwg, namePoint)
        'читаем атрибуты блока
        Dim arrayAttribute As String() = Nothing
        Dim boolReadBlk As Boolean = RoburFunc.FuncReadNameAttribBlk(ActivDocument, blkPK1, arrayAttribute)
        If IsArray(arrayAttribute) = True Then
            ComboBox6.DataSource = arrayAttribute
            For i As Integer = 0 To arrayAttribute.Length - 1
                Dim nameAttrib As String = arrayAttribute(i)
                If nameAttrib Like "NUM_GEOPOINT" Then
                    ComboBox6.Text = "NUM_GEOPOINT"
                    Exit For
                End If
            Next i
        Else
            ReDim arrayAttribute(0)
            arrayAttribute(0) = ""
            ComboBox6.DataSource = arrayAttribute
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'смена таблицы
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        'заполням поле с категорией
        Dim templateXML As String = arrayPutch(0)
        Dim templateDwg As String = arrayPutch(1)
        'заполняем блоки
        Dim arrayBlkDesk As String() = Nothing
        Dim boolFindBlk As Boolean = FuncXML.FuncReadDeskObjectByCategoryAndTablesToXML(templateXML, ComboBox4.Text.Trim, ComboBox1.Text.Trim, "Блок", arrayBlkDesk)
        If IsArray(arrayBlkDesk) = True Then
            ComboBox5.DataSource = arrayBlkDesk
        End If
        '==========================================================================================================================
        'ищем характеристики
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'смотрим характеристики блока
        Dim categoryPoint As String = ComboBox4.Text.Trim
        Dim tablePoint As String = ComboBox1.Text.Trim
        Dim deskPoint As String = ComboBox5.Text.Trim
        Dim arrayPropertiesObject As String(,) = Nothing
        Dim boolPropObject As Boolean = FuncXML.FuncReadDataObjectToXMLbyDesk(templateXML, categoryPoint, tablePoint, "Блок", deskPoint, arrayPropertiesObject)
        'забираем в массив все блоки с чертежа
        If IsArray(arrayPropertiesObject) = True Then
            'имя блока
            Dim nameBlk As String = FuncXML.FuncReturnValueByArrayProperties("Name", arrayPropertiesObject, 1)
            'вставляем его в чертеж
            Dim blkPK As DwgBlock = RoburFunc.FuncInsertBlockByDWGFile(ActivDocument, templateDwg, nameBlk)
            Dim arrayBLK As String() = Nothing
            Dim boolFindBlkProperties1 As Boolean = RoburFunc.FuncReadNameInsertBlocks(ActivDocument, arrayBLK)
            If IsArray(arrayBLK) = True Then
                ComboBox2.DataSource = arrayBLK
                For i As Integer = 0 To arrayBLK.Length - 1
                    If arrayBLK(i) Like nameBlk Then
                        ComboBox2.Text = nameBlk
                    End If
                Next i
            End If
        End If
        'слой
        Dim arrayLayers As String() = Nothing
        If IsArray(arrayPropertiesObject) = True Then
            Dim nameLayer As String = FuncXML.FuncReturnValueByArrayProperties("Layer", arrayPropertiesObject, 1)
            If nameLayer.Trim.Length > 0 Then
                ReDim arrayLayers(0)
                arrayLayers(0) = nameLayer
            End If
        End If
        Dim boolFindBlkProperties2 As Boolean = RoburFunc.FuncReadLayers(ActivDocument, arrayLayers)
        ComboBox3.DataSource = arrayLayers
        'цвет
        Dim colorBlk As String = FuncXML.FuncReturnValueByArrayProperties("Color", arrayPropertiesObject, 1)
        TextBox2.Text = colorBlk
        'размер блока
        Dim ScaleBlk As Double = Val(FuncXML.FuncReturnValueByArrayProperties("Scale", arrayPropertiesObject, 1))
        If ScaleBlk <= 0 Then ScaleBlk = 1
        NumericUpDown1.Value = ScaleBlk
        'размер атрибута
        Dim attribHeight As Double = Val(FuncXML.FuncReturnValueByArrayProperties("AttributeHeight", arrayPropertiesObject, 1))
        If attribHeight <= 0 Then attribHeight = 1.5
        NumericUpDown2.Value = attribHeight
        'расширенные данные
        Dim arrayPs As String() = Nothing
        Dim boolFindPs As Boolean = FuncXML.FuncReadPSTablesToXML(templateXML, arrayPs)
        If IsArray(arrayPs) = True Then
            ComboBox7.DataSource = arrayPs
            Dim PSTable As String = FuncXML.FuncReturnValueByArrayProperties("PS", arrayPropertiesObject, 1)
            If PSTable.Trim.Length > 0 Then
                For i As Integer = 0 To arrayPs.Length - 1
                    Dim PsBlk As String = arrayPs(i)
                    If PsBlk Like PSTable Then
                        ComboBox7.Text = PSTable
                        Exit For
                    End If
                Next i
            End If
        End If
        '======================================================================================================================
        'загружаем блок
        Dim namePoint As String = ComboBox2.Text
        Dim blkPK1 As DwgBlock = RoburFunc.FuncInsertBlockByDWGFile(ActivDocument, templateDwg, namePoint)
        'читаем атрибуты блока
        Dim arrayAttribute As String() = Nothing
        Dim boolReadBlk As Boolean = RoburFunc.FuncReadNameAttribBlk(ActivDocument, blkPK1, arrayAttribute)
        If IsArray(arrayAttribute) = True Then
            ComboBox6.DataSource = arrayAttribute
            For i As Integer = 0 To arrayAttribute.Length - 1
                Dim nameAttrib As String = arrayAttribute(i)
                If nameAttrib Like "NUM_GEOPOINT" Then
                    ComboBox6.Text = "NUM_GEOPOINT"
                    Exit For
                End If
            Next i
        Else
            ReDim arrayAttribute(0)
            arrayAttribute(0) = ""
            ComboBox6.DataSource = arrayAttribute
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'смена блока
    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        'заполням поле с категорией
        Dim templateXML As String = arrayPutch(0)
        Dim templateDwg As String = arrayPutch(1)
        '==========================================================================================================================
        'ищем характеристики
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'смотрим характеристики блока
        Dim categoryPoint As String = ComboBox4.Text.Trim
        Dim tablePoint As String = ComboBox1.Text.Trim
        Dim deskPoint As String = ComboBox5.Text.Trim
        Dim arrayPropertiesObject As String(,) = Nothing
        Dim boolPropObject As Boolean = FuncXML.FuncReadDataObjectToXMLbyDesk(templateXML, categoryPoint, tablePoint, "Блок", deskPoint, arrayPropertiesObject)
        'забираем в массив все блоки с чертежа
        If IsArray(arrayPropertiesObject) = True Then
            'имя блока
            Dim nameBlk As String = FuncXML.FuncReturnValueByArrayProperties("Name", arrayPropertiesObject, 1)
            'вставляем его в чертеж
            Dim blkPK As DwgBlock = RoburFunc.FuncInsertBlockByDWGFile(ActivDocument, templateDwg, nameBlk)
            Dim arrayBLK As String() = Nothing
            Dim boolFindBlkProperties1 As Boolean = RoburFunc.FuncReadNameInsertBlocks(ActivDocument, arrayBLK)
            If IsArray(arrayBLK) = True Then
                ComboBox2.DataSource = arrayBLK
                For i As Integer = 0 To arrayBLK.Length - 1
                    If arrayBLK(i) Like nameBlk Then
                        ComboBox2.Text = nameBlk
                    End If
                Next i
            End If
        End If
        'слой
        Dim arrayLayers As String() = Nothing
        If IsArray(arrayPropertiesObject) = True Then
            Dim nameLayer As String = FuncXML.FuncReturnValueByArrayProperties("Layer", arrayPropertiesObject, 1)
            If nameLayer.Trim.Length > 0 Then
                ReDim arrayLayers(0)
                arrayLayers(0) = nameLayer
            End If
        End If
        Dim boolFindBlkProperties2 As Boolean = RoburFunc.FuncReadLayers(ActivDocument, arrayLayers)
        ComboBox3.DataSource = arrayLayers
        'цвет
        Dim colorBlk As String = FuncXML.FuncReturnValueByArrayProperties("Color", arrayPropertiesObject, 1)
        TextBox2.Text = colorBlk
        'размер блока
        Dim ScaleBlk As Double = Val(FuncXML.FuncReturnValueByArrayProperties("Scale", arrayPropertiesObject, 1))
        If ScaleBlk <= 0 Then ScaleBlk = 1
        NumericUpDown1.Value = ScaleBlk
        'размер атрибута
        Dim attribHeight As Double = Val(FuncXML.FuncReturnValueByArrayProperties("AttributeHeight", arrayPropertiesObject, 1))
        If attribHeight <= 0 Then attribHeight = 1.5
        NumericUpDown2.Value = attribHeight
        'расширенные данные
        Dim arrayPs As String() = Nothing
        Dim boolFindPs As Boolean = FuncXML.FuncReadPSTablesToXML(templateXML, arrayPs)
        If IsArray(arrayPs) = True Then
            ComboBox7.DataSource = arrayPs
            Dim PSTable As String = FuncXML.FuncReturnValueByArrayProperties("PS", arrayPropertiesObject, 1)
            If PSTable.Trim.Length > 0 Then
                For i As Integer = 0 To arrayPs.Length - 1
                    Dim PsBlk As String = arrayPs(i)
                    If PsBlk Like PSTable Then
                        ComboBox7.Text = PSTable
                        Exit For
                    End If
                Next i
            End If
        End If
        '======================================================================================================================
        'загружаем блок
        Dim namePoint As String = ComboBox2.Text
        Dim blkPK1 As DwgBlock = RoburFunc.FuncInsertBlockByDWGFile(ActivDocument, templateDwg, namePoint)
        'читаем атрибуты блока
        Dim arrayAttribute As String() = Nothing
        Dim boolReadBlk As Boolean = RoburFunc.FuncReadNameAttribBlk(ActivDocument, blkPK1, arrayAttribute)
        If IsArray(arrayAttribute) = True Then
            ComboBox6.DataSource = arrayAttribute
            For i As Integer = 0 To arrayAttribute.Length - 1
                Dim nameAttrib As String = arrayAttribute(i)
                If nameAttrib Like "NUM_GEOPOINT" Then
                    ComboBox6.Text = "NUM_GEOPOINT"
                    Exit For
                End If
            Next i
        Else
            ReDim arrayAttribute(0)
            arrayAttribute(0) = ""
            ComboBox6.DataSource = arrayAttribute
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'поменять имя блока
    Private Sub ComboBox2_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        'заполням поле с категорией
        Dim templateXML As String = arrayPutch(0)
        Dim templateDwg As String = arrayPutch(1)
        '======================================================================================================================
        'загружаем блок
        Dim namePoint As String = ComboBox2.Text
        Dim blkPK As DwgBlock = RoburFunc.FuncInsertBlockByDWGFile(ActivDocument, templateDwg, namePoint)
        'читаем атрибуты блока
        Dim arrayAttribute As String() = Nothing
        Dim boolReadBlk As Boolean = RoburFunc.FuncReadNameAttribBlk(ActivDocument, blkPK, arrayAttribute)
        If IsArray(arrayAttribute) = True Then
            ComboBox6.DataSource = arrayAttribute
            For i As Integer = 0 To arrayAttribute.Length - 1
                Dim nameAttrib As String = arrayAttribute(i)
                If nameAttrib Like "NUM_GEOPOINT" Then
                    ComboBox6.Text = "NUM_GEOPOINT"
                    Exit For
                End If
            Next i
        Else
            ReDim arrayAttribute(0)
            arrayAttribute(0) = ""
            ComboBox6.DataSource = arrayAttribute
        End If
    End Sub

    Private Sub FormInsertPointPPT_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub GroupBox4_Enter(sender As Object, e As EventArgs) Handles GroupBox4.Enter

    End Sub
End Class