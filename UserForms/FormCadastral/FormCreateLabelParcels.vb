Imports System.IO
Imports System.Windows.Forms
Imports Topomatic.Cad.Foundation
Imports Topomatic.Cad.View
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Layer

Public Class FormCreateLabelParcels
    Public ActivDocument As Topomatic.Dwg.Drawing = Nothing
    Public boolShow As Boolean = False
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
        'категория объекта
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
                'заполняем мультивыноски
                Dim arrayMLeaderDesk As String() = Nothing
                Dim boolFindBlk As Boolean = FuncXML.FuncReadDeskObjectByCategoryAndTablesToXML(templateXML, ComboBox4.Text.Trim, ComboBox1.Text.Trim, "Мультивыноска", arrayMLeaderDesk)
                If IsArray(arrayMLeaderDesk) = True Then
                    ComboBox5.DataSource = arrayMLeaderDesk
                    For i As Integer = 0 To arrayMLeaderDesk.Length - 1
                        Dim deskBlk As String = arrayMLeaderDesk(i)
                        If deskBlk Like "Земельный участок (обозначение и площадь)" Then
                            ComboBox5.Text = "Земельный участок (обозначение и площадь)"
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
        Dim boolPropObject As Boolean = FuncXML.FuncReadDataObjectToXMLbyDesk(templateXML, categoryPoint, tablePoint, "Мультивыноска", deskPoint, arrayPropertiesObject)
        'имя сстиля
        Dim nameTextStyle As String = FuncXML.FuncReturnValueByArrayProperties("TextStyle", arrayPropertiesObject, 1)
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
        Dim arrowSize As Double = Val(FuncXML.FuncReturnValueByArrayProperties("ArrowSize", arrayPropertiesObject, 1))
        If arrowSize <= 0 Then arrowSize = 1
        NumericUpDown1.Value = arrowSize
        'размер атрибута
        Dim attribHeight As Double = Val(FuncXML.FuncReturnValueByArrayProperties("TextHeight", arrayPropertiesObject, 1))
        If attribHeight <= 0 Then attribHeight = 2
        NumericUpDown2.Value = attribHeight
        '======================================================================================================================
        'префикс номера
        Dim ArrayPref As String() = {"00:00:0000000:0000:ЗУ", "Участок №", ":0000:ЗУ", ":0000/чзу", "Публ.сервитут №", "Сервитут №", "00:00:0000000:0000/чзу", "Часть участка №", "№", ""}
        ComboBox2.DataSource = ArrayPref
        'тип стрелки выноски
        Dim arrayArrowName As String() = {"нет", "Стрелка", "Точка", "Малая точка"}
        ComboBox6.DataSource = arrayArrowName
        'текстовый стиль
        Dim arrayTextStyle As String() = Nothing
        'Dim boolTextStyle As Boolean = FuncStyles.FuncReturnTextStyles(ActivDocument, arrayTextStyle)
        Dim boolfindTextStyles As Boolean = FuncXML.FuncReadTextStylesToXML(templateXML, arrayTextStyle)
        If IsArray(arrayTextStyle) = True Then
            ComboBox7.DataSource = arrayTextStyle
            For i As Integer = 0 To arrayTextStyle.Length - 1
                If IsNothing(arrayTextStyle(i)) = False Then
                    If arrayTextStyle(i).Trim Like nameTextStyle Then
                        ComboBox7.Text = nameTextStyle
                        Exit For
                    End If
                End If
            Next i
        End If
    End Sub
    Private Sub FormCreateLabelParcels_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
    'Отмена
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        boolShow = False
        Me.Hide()
    End Sub
    'OK
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        boolShow = True
        Me.Hide()
    End Sub

    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton5.CheckedChanged

    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'вызвать цвет
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim userColor As CadColor = New CadColor(7)
        Dim boolColorDlg As Boolean = Topomatic.Cad.View.ColorsBrowserDlg.Execute(userColor)
        If boolColorDlg = True Then
            If userColor.ColorIndex = 0 Then
                TextBox2.Text = "ПоБлоку"
            ElseIf userColor.ColorIndex = 256 Then
                TextBox2.Text = "ПоСлою"
            Else
                TextBox2.Text = userColor.ColorIndex
            End If
        End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

    End Sub
End Class