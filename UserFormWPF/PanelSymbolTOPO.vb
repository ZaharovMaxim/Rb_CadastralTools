Imports System.ComponentModel
Imports System.Windows.Forms
Imports System.Xml
Imports Topomatic.Dwg.Entities
Imports Topomatic.ApplicationPlatform.Plugins
Imports Topomatic.Cad.View
Imports Topomatic.Dwg.Layer
Imports Topomatic.Dwg
Imports Topomatic.Cad.Foundation
Imports Topomatic.Sfc
Imports Topomatic.FoundationClasses
Imports Topomatic.Alg.Road.Urb.UserStrips
Imports Topomatic.Alg.Road.Urb.Border
Imports System.Windows.Shapes
Imports Topomatic.Visualization
Imports Topomatic.Visualization.Runtime
Imports Topomatic.Cad.Foundation.Brep
Imports Topomatic.Visualization.Geometry
Imports Topomatic.Visualization.Runtime.Tools
Imports Topomatic.Alg.Road.Crossing
Imports Topomatic.Alg.Road
Imports Topomatic.Alg.Runtime
Imports Topomatic.Alg.Runtime.ServiceClasses
Imports Topomatic.Alg
Imports Topomatic.Alg.Road.Urb
Imports Topomatic.Alg.Road.Urb.Descent
Imports Topomatic.Alg.Plan
Imports Topomatic.Alg.Stationing
'Imports Topomatic.Alg.Kilometres
Imports Topomatic.Alg.Kilometres.AlgAlignmentKilometres
Imports Topomatic.Cad.View.Hints
Imports System.Windows.Controls
Imports System.IO
Imports Topomatic.ApplicationPlatform.Core
Imports System.Security.Policy
Imports System.Windows.Media.Media3D
Imports Topomatic.Alg.Crs
Imports Topomatic.Planchet.Entities
Imports Topomatic.Alg.Kilometres
Imports Topomatic.ApplicationPlatform
Imports Topomatic.Arrangements
Imports Topomatic.Sfc.Layer
Imports Topomatic.Controls.Dialogs

Public Class PanelSymbolTOPO
    Public templateXML As String = ""
    Public templateDwg As String = ""
    Public boolSelectObject As Boolean = False
    Private old_node As TreeNode = Nothing
    Public userCadView As CadView = Nothing
    Public generalDir As String = ""
    Public Sub New()
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        ' Добавить код инициализации после вызова InitializeComponent().
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim arrayDirSupport As String() = Nothing
        Dim boolFindDirSupport As Boolean = FuncFiles.funcReadDirectoriesSupport(arrayDirSupport)
        If IsArray(arrayDirSupport) = True Then
            For i As Integer = 0 To arrayDirSupport.Length - 1
                Dim nameFolder As String = arrayDirSupport(i)
                Dim pos As Integer = nameFolder.IndexOf("Civil3DToolsUtility")
                If pos > -1 Then
                    generalDir = Mid(nameFolder, 1, pos)
                    generalDir = generalDir & "Civil3DToolsUtility"
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
                        templateDwg = tempNameFileDwg
                        Exit For
                    End If
                Next
            End If
        End If
        'templateXML = FuncFiles.funcFindFileToDirectorySupport("Стандарт_INFRASTRADA.xml")
        'templateDwg = FuncFiles.funcFindFileToDirectorySupport("Стандарт_ADSK.dwg")
        If IsNothing(templateXML) = True Then
            MsgBox("Файл с шаблонами =Стандарт_ADSK.xml= не найден. Пожалуйста, добавьте его в пути доступа к вспомогательным файлам!!!")
            Exit Sub
        End If
        Dim xDoc As XmlDocument = New XmlDocument()
        If IO.File.Exists(templateXML) = True Then
            Dim userNode As TreeNode = Nothing
            Dim userNode0 As TreeNode = Nothing 'категория
            Dim reader As XmlTextReader = New XmlTextReader(templateXML)
            Dim arrayTag As String(,) = Nothing
            Dim countArrayTag As Integer = 0
            While reader.Read()
                Select Case reader.NodeType
                    Case XmlNodeType.Element
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "Classifier_ADSK" Then
                            userNode = New TreeNode("Оформление плана")
                            userNode.Tag = "Classifier_ADSK"
                            TreeView1.Nodes.Add(userNode)
                        ElseIf NameBlock Like "LibrarySymbols" Then
                            If IsNothing(userNode) = False Then
                                If reader.HasAttributes = True Then
                                    Dim arrayData As String(,) = Nothing
                                    Dim countAttr As Integer = 1
                                    ReDim arrayData(1, reader.AttributeCount)
                                    arrayData(0, 0) = "LibrarySymbols"
                                    arrayData(1, 0) = "NameLibrarySymbols"
                                    While (reader.MoveToNextAttribute())
                                        Dim name As String = reader.Name
                                        Dim valN As String = reader.Value
                                        arrayData(0, countAttr) = name
                                        arrayData(1, countAttr) = valN
                                        countAttr += 1
                                        If name Like "NameLibrarySymbols" Then
                                            userNode0 = New TreeNode(valN)
                                            userNode.Nodes.Add(userNode0)
                                        End If
                                    End While
                                    If IsArray(arrayData) = True Then
                                        userNode0.Tag = arrayData
                                    End If
                                End If
                            End If
                        End If
                End Select
            End While
            TreeView1.ExpandAll()
            reader.Close()
        End If
    End Sub

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'выбор узла дерева
    Private Sub TreeView1_NodeMouseClick(sender As Object, e As Windows.Forms.TreeNodeMouseClickEventArgs) Handles TreeView1.NodeMouseClick
        Dim selectNode As TreeNode = e.Node
        Dim tagSelectNode As String(,) = Nothing
        If IsArray(selectNode.Tag) = True Then
            tagSelectNode = selectNode.Tag
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            If selectNode.Nodes.Count = 0 And tagSelectNode(0, 0) Like "LibrarySymbols" Then
                Button1.Tag = ""
                'читаем xml
                Dim xDoc As XmlDocument = New XmlDocument()
                Dim userNode1 As TreeNode = Nothing 'категория
                Dim userNode2 As TreeNode = Nothing 'имя таблицы
                Dim userNode3 As TreeNode = Nothing 'тип данных
                Dim userNode4 As TreeNode = Nothing 'данные
                If IO.File.Exists(templateXML) = True Then
                    Dim boolCategory As Boolean = False
                    Dim reader As XmlTextReader = New XmlTextReader(templateXML)
                    Dim arrayTag As String(,) = Nothing
                    Dim countArrayTag As Integer = 0
                    While reader.Read()
                        Select Case reader.NodeType
                            Case XmlNodeType.Element
                                Dim NameBlock As String = reader.Name 'читаем ветку
                                If NameBlock Like "Category" Then
                                    If reader.HasAttributes = True Then
                                        While (reader.MoveToNextAttribute())
                                            Dim nameN As String = reader.Name
                                            Dim valN As String = reader.Value
                                            If nameN Like "CategoryName" Then
                                                Dim newNode1 As TreeNode = New TreeNode(valN)
                                                ReDim arrayTag(1, 1)
                                                arrayTag(0, 0) = "Category"
                                                arrayTag(0, 1) = "CategoryName"
                                                newNode1.Tag = arrayTag
                                                selectNode.Nodes.Add(newNode1)
                                                Exit While
                                            End If
                                        End While
                                    End If
                                End If
                        End Select
                    End While
                    reader.Close()
                End If
                If selectNode.Nodes.Count > 0 Then
                    selectNode.Expand()
                End If
                'выбор категории
            ElseIf selectNode.Nodes.Count = 0 And tagSelectNode(0, 0) Like "Category" Then
                Dim nameNode As String = selectNode.Text
                Button1.Tag = ""
                'читаем xml
                Dim xDoc As XmlDocument = New XmlDocument()
                If IO.File.Exists(templateXML) = True Then
                    Dim boolCategory As Boolean = False
                    Dim arrayTag As String(,) = Nothing
                    Dim reader As XmlTextReader = New XmlTextReader(templateXML)
                    While reader.Read()
                        Select Case reader.NodeType
                            Case XmlNodeType.Element
                                Dim NameBlock As String = reader.Name 'читаем ветку
                                If NameBlock Like "Category" Then
                                    If reader.HasAttributes = True Then
                                        While (reader.MoveToNextAttribute())
                                            Dim valN As String = reader.Value
                                            If valN Like nameNode Then
                                                boolCategory = True
                                                Exit While
                                            Else
                                                boolCategory = False
                                            End If
                                        End While
                                    End If
                                ElseIf NameBlock Like "Table" Then
                                    If boolCategory = True Then
                                        If reader.HasAttributes = True Then
                                            While (reader.MoveToNextAttribute())
                                                Dim valN As String = reader.Value
                                                Dim newNode2 As TreeNode = New TreeNode(valN)
                                                ReDim arrayTag(1, 1)
                                                arrayTag(0, 0) = "Table"
                                                arrayTag(0, 1) = reader.Name
                                                arrayTag(1, 1) = valN
                                                newNode2.Tag = arrayTag
                                                selectNode.Nodes.Add(newNode2)
                                                Exit While
                                            End While
                                        End If
                                    End If
                                End If
                        End Select
                    End While
                    reader.Close()
                End If
                If selectNode.Nodes.Count > 0 Then
                    selectNode.Expand()
                End If
                'выбор таблицы
            ElseIf selectNode.Nodes.Count = 0 And tagSelectNode(0, 0) Like "Table" Then
                Dim nameNode As String = selectNode.Text
                Dim parentNode As String = selectNode.Parent.Text
                Button1.Tag = ""
                'читаем xml
                Dim xDoc As XmlDocument = New XmlDocument()
                If IO.File.Exists(templateXML) = True Then
                    Dim boolCategory As Boolean = False
                    Dim boolTable As Boolean = False
                    Dim arrayTag As String(,) = Nothing
                    Dim reader As XmlTextReader = New XmlTextReader(templateXML)
                    While reader.Read()
                        Select Case reader.NodeType
                            Case XmlNodeType.Element
                                Dim NameBlock As String = reader.Name 'читаем ветку
                                If NameBlock Like "Category" Then
                                    If reader.HasAttributes = True Then
                                        While (reader.MoveToNextAttribute())
                                            Dim valN As String = reader.Value
                                            If valN Like parentNode Then
                                                boolCategory = True
                                                Exit While
                                            Else
                                                boolCategory = False
                                            End If
                                        End While
                                    End If
                                ElseIf NameBlock Like "Table" Then
                                    If reader.HasAttributes = True Then
                                        While (reader.MoveToNextAttribute())
                                            Dim valN As String = reader.Value
                                            If valN Like nameNode Then
                                                boolTable = True
                                                Exit While
                                            Else
                                                boolTable = False
                                            End If
                                        End While
                                    End If
                                ElseIf NameBlock Like "TypeObject" Then
                                    If boolCategory = True And boolTable = True Then
                                        If reader.HasAttributes = True Then
                                            While (reader.MoveToNextAttribute())
                                                Dim valN As String = reader.Value
                                                Dim newNode As TreeNode = New TreeNode(valN)
                                                ReDim arrayTag(1, 1)
                                                arrayTag(0, 0) = "TypeObject"
                                                arrayTag(0, 1) = reader.Name
                                                arrayTag(1, 1) = valN
                                                newNode.Tag = arrayTag
                                                selectNode.Nodes.Add(newNode)
                                                Exit While
                                            End While
                                        End If
                                    End If
                                End If
                        End Select
                    End While
                    reader.Close()
                End If
                If selectNode.Nodes.Count > 0 Then
                    selectNode.Expand()
                End If
            ElseIf selectNode.Nodes.Count = 0 And tagSelectNode(0, 0) Like "TypeObject" Then
                Dim nameNode As String = selectNode.Text
                Dim parentNode As String = selectNode.Parent.Text
                Dim parentParentNode As String = selectNode.Parent.Parent.Text
                Button1.Tag = ""
                'читаем xml
                Dim xDoc As XmlDocument = New XmlDocument()
                If IO.File.Exists(templateXML) = True Then
                    Dim boolCategory As Boolean = False
                    Dim boolTable As Boolean = False
                    Dim boolType As Boolean = False
                    Dim arrayTag As String(,) = Nothing
                    Dim reader As XmlTextReader = New XmlTextReader(templateXML)
                    While reader.Read()
                        Select Case reader.NodeType
                            Case XmlNodeType.Element
                                Dim NameBlock As String = reader.Name 'читаем ветку
                                If NameBlock Like "Category" Then
                                    If reader.HasAttributes = True Then
                                        While (reader.MoveToNextAttribute())
                                            Dim valN As String = reader.Value
                                            If valN Like parentParentNode Then
                                                boolCategory = True
                                                Exit While
                                            Else
                                                boolCategory = False
                                            End If
                                        End While
                                    End If
                                ElseIf NameBlock Like "Table" Then
                                    If reader.HasAttributes = True Then
                                        While (reader.MoveToNextAttribute())
                                            Dim valN As String = reader.Value
                                            If valN Like parentNode Then
                                                boolTable = True
                                                Exit While
                                            Else
                                                boolTable = False
                                            End If
                                        End While
                                    End If
                                ElseIf NameBlock Like "TypeObject" Then
                                    If reader.HasAttributes = True Then
                                        While (reader.MoveToNextAttribute())
                                            Dim valN As String = reader.Value
                                            If valN Like nameNode Then
                                                boolType = True
                                                Exit While
                                            Else
                                                boolType = False
                                            End If
                                        End While
                                    End If
                                ElseIf NameBlock Like "Rows" Then
                                    If boolCategory = True And boolTable = True And boolType Then
                                        If reader.HasAttributes = True Then
                                            While (reader.MoveToNextAttribute())
                                                Dim name As String = reader.Name
                                                If name Like "Description" Then
                                                    Dim valN As String = reader.Value
                                                    Dim newNode As TreeNode = New TreeNode(valN)
                                                    ReDim arrayTag(1, 1)
                                                    arrayTag(0, 0) = "Rows"
                                                    arrayTag(0, 1) = reader.Name
                                                    arrayTag(1, 1) = valN
                                                    newNode.Tag = arrayTag
                                                    selectNode.Nodes.Add(newNode)
                                                    Exit While
                                                End If
                                            End While
                                        End If
                                    End If
                                End If
                        End Select
                    End While
                    reader.Close()
                End If
                If selectNode.Nodes.Count > 0 Then
                    selectNode.Expand()
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'выбор объекта для редактирования
            ElseIf selectNode.Nodes.Count = 0 And tagSelectNode(0, 0) Like "Rows" Then
                Dim nameNode As String = selectNode.Text
                Dim parentNode As String = selectNode.Parent.Text
                Dim parentParentNode As String = selectNode.Parent.Parent.Text
                Dim parentParentParentNode As String = selectNode.Parent.Parent.Parent.Text
                'читаем xml
                Dim xDoc As XmlDocument = New XmlDocument()
                If IO.File.Exists(templateXML) = True Then
                    Dim boolCategory As Boolean = False
                    Dim boolTable As Boolean = False
                    Dim boolType As Boolean = False
                    Dim reader As XmlTextReader = New XmlTextReader(templateXML)
                    While reader.Read()
                        Select Case reader.NodeType
                            Case XmlNodeType.Element
                                Dim NameBlock As String = reader.Name 'читаем ветку
                                If NameBlock Like "Category" Then
                                    If reader.HasAttributes = True Then
                                        While (reader.MoveToNextAttribute())
                                            Dim valN As String = reader.Value
                                            If valN Like parentParentParentNode Then
                                                boolCategory = True
                                                Exit While
                                            Else
                                                boolCategory = False
                                            End If
                                        End While
                                    End If
                                ElseIf NameBlock Like "Table" Then
                                    If reader.HasAttributes = True Then
                                        While (reader.MoveToNextAttribute())
                                            Dim valN As String = reader.Value
                                            If valN Like parentParentNode Then
                                                boolTable = True
                                                Exit While
                                            Else
                                                boolTable = False
                                            End If
                                        End While
                                    End If
                                ElseIf NameBlock Like "TypeObject" Then
                                    If reader.HasAttributes = True Then
                                        While (reader.MoveToNextAttribute())
                                            Dim valN As String = reader.Value
                                            If valN Like parentNode Then
                                                boolType = True
                                                Exit While
                                            Else
                                                boolType = False
                                            End If
                                        End While
                                    End If
                                ElseIf NameBlock Like "Rows" Then
                                    If boolCategory = True And boolTable = True And boolType Then
                                        Dim arrayProperties As String(,) = Nothing
                                        Dim countArrayProperties As Integer = 0
                                        Dim boolElement As Boolean = False
                                        If reader.HasAttributes = True Then
                                            While (reader.MoveToNextAttribute())
                                                Dim name As String = reader.Name
                                                Dim valN As String = reader.Value
                                                ReDim Preserve arrayProperties(1, countArrayProperties)
                                                arrayProperties(0, countArrayProperties) = name
                                                arrayProperties(1, countArrayProperties) = valN
                                                countArrayProperties += 1
                                                If valN Like nameNode Then
                                                    boolElement = True
                                                End If
                                            End While
                                            If boolElement = True Then
                                                Button1.Tag = arrayProperties
                                                Exit While
                                            End If
                                        End If
                                    End If
                                End If
                        End Select
                    End While
                    reader.Close()
                End If
            End If
        End If
    End Sub

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'нажата кнопка нарисовать объект
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim userNode As TreeNode = TreeView1.SelectedNode
        If IsArray(Button1.Tag) = True Then
            'характеристики объекта
            Dim arraytag As String(,) = Button1.Tag
            'масштаб чертежа
            Dim MDraw As Double = Val(Label2.Tag)
            If MDraw <= 0 Then MDraw = 1
            'имя таблицы
            Dim nameTables As String = userNode.Parent.Parent.Text
            Dim nameCategoryTable As String = userNode.Parent.Parent.Parent.Text
            Dim nameModel As String = ComboBox1.Text
            If IsNothing(userCadView) = False Then
                Dim userProject As ModelProject = ApplicationHost.Current.ActiveProject
                Dim userSubObjectBearm As ArrangementModel = New ArrangementModel()
                'получаем активный документ
                'Dim DrawLayer As DrawingLayer = Nothing 'DrawLayer.Drawing
                Dim ActivDocument As Topomatic.Dwg.Drawing = Nothing
                If IsNothing(nameModel) = False Then
                    nameModel = nameModel.Trim
                    Dim childs As IProjectModel() = userProject.Model.GetChilds()
                    Try
                        For Each child As IProjectModel In childs
                            Dim modelUri As URI = child.Uri
                            Dim fileNameSfcx As String = IO.Path.GetFileNameWithoutExtension(modelUri.LastPathComponent)
                            If fileNameSfcx.Trim Like nameModel.Trim Then
                                Dim real_model As IDrawingContainer = child.LockReadContainer(Of IDrawingContainer)()
                                If IsNothing(real_model) = False Then
                                    ApplicationHost.Current.Plugins.Execute("activate", New Object() {child})
                                    ActivDocument = real_model.Drawing
                                    Exit For
                                End If
                            End If
                        Next
                    Catch ex As System.Exception
                    End Try
                End If
                If IsNothing(ActivDocument) = True Then Exit Sub
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '=======================================================================================================
                'идентификатор объекта по treeView
                Dim nameUserObject As String = userNode.Text 'дескриптор объекта
                Dim typeUserObject As String = userNode.Parent.Text 'тип объекта
                Dim codeUserObject As String = ""
                'имя объекта для вставки
                Dim nameCreateObject As String = nameUserObject
                'слой
                Dim nameObjectLayer As String = "0" 'слой
                Dim objectLayer As DwgLayer = ActivDocument.ActiveLayer
                'линии
                Dim nameUserTypeLine As String = "Continuous"
                Dim objectUserTypeLine As DwgLinetype = ActivDocument.ActiveLinetype
                Dim deskUserTypeLine As String = "Непрерывный"
                Dim scaleUserTypeLine As Double = 1 'масштаб
                Dim widthUserTypeLine As Double = 20 'ширина
                Dim gWidthUserTypeLine As Double = 0 'толщина
                Dim arrayMaskTypeLine As String() = Nothing 'массив для создания типа линии
                'цвет
                Dim IndexColor As Integer = 256
                Dim objectColor As CadColor = New CadColor(IndexColor)
                'текстовый стиль
                Dim userTextStyle As DwgStyle = ActivDocument.ActiveStyle
                Dim textStyle As String = ""
                Dim hText As Double = 2
                Dim backGround As Boolean = True
                Dim indexColorBaskground As Integer = 256
                'выноска
                Dim arrowSize As Double = 2
                Dim nameArrow As String = "_Closed"
                'прочие характеристики
                Dim scaleUserObject As Double = 1
                Dim angleUserObject As Double = 0
                Dim attributeHeight As Double = 2.5
                Dim nameBlk As String = "нет"
                'примечание
                Dim note As String = ""
                Dim PSTable As String = "OBJECT"
                Dim rng As Random = New Random
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                If arraytag.GetUpperBound(1) > 0 Then
                    If typeUserObject Like "Полилиния" Or typeUserObject Like "Линия" Or typeUserObject Like "Характерная линия" Or typeUserObject Like "Откос" Or typeUserObject Like "Дополнительная ось" Or typeUserObject Like "Бордюр" Or typeUserObject Like "Бордюр" Or typeUserObject Like "Структурная линия" Or typeUserObject Like "Дорожная разметка" Or typeUserObject Like "Ось трассы" Then
                        Dim boolCreateStyle As Boolean = FuncDecorPlane.FuncCreateStyleLinearObject(ActivDocument, templateXML, arraytag, objectLayer, objectUserTypeLine, objectColor, scaleUserTypeLine, widthUserTypeLine, PSTable, note, gWidthUserTypeLine)
                    ElseIf typeUserObject Like "Участок" Or typeUserObject Like "Штриховка" Then
                        FuncDecorPlane.FuncCreateStyleHatchObject(ActivDocument, templateXML, arraytag, nameUserObject, objectLayer, objectUserTypeLine, objectColor, scaleUserTypeLine, widthUserTypeLine, scaleUserObject, angleUserObject, PSTable, note)
                    ElseIf typeUserObject Like "Блок" Then
                        FuncDecorPlane.FuncCreateStylePointObject(ActivDocument, templateXML, arraytag, nameCreateObject, objectLayer, objectUserTypeLine, objectColor, scaleUserTypeLine, widthUserTypeLine, scaleUserObject, angleUserObject, attributeHeight, PSTable, note)
                    ElseIf typeUserObject Like "Мультивыноска" Then
                        FuncDecorPlane.FuncCreateStyleLeaderObject(ActivDocument, templateXML, arraytag, nameCreateObject, objectLayer, objectUserTypeLine, objectColor, scaleUserTypeLine, widthUserTypeLine, textStyle, hText, nameArrow, arrowSize, PSTable, note)
                    End If
                End If
                '===========================================================================================================
                'получаем PS таблицу для записи в объект
                Dim arrayDSDataRecord As String(,) = Nothing
                ReDim Preserve arrayDSDataRecord(1, 0)
                arrayDSDataRecord(0, 0) = "LABEL"
                arrayDSDataRecord(1, 0) = nameUserObject
                Dim dictionaryDataobject As Dictionary(Of String, String(,)) = New Dictionary(Of String, String(,))
                If PSTable.Trim.Length > 0 Then
                    Dim boolFindFields As Boolean = FuncXRecords.FuncReadDataFieldsByTablePSToXML(templateXML, PSTable, dictionaryDataobject)
                End If
                'сброс выделения
                userCadView.SelectionSet.Clear()
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'создать новый тип объекта
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                If RadioButton1.Checked = True Then
                    'поставить блок
                    If typeUserObject Like "Блок" Then
                        If nameCreateObject.IndexOf("A-Уклон-Расстояние") > -1 Then
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            'выбираем трассу
                            Dim boolSelectAlign As Boolean = False
                            Dim align As Alignment = Nothing
                            Dim userWrapperAlign As Topomatic.Alg.Layers.Wrappers.AlignmentWrapper = Nothing
                            Dim poly As IPolyline3D = Nothing
                            Dim ID_Rng As String = ""
                            Dim IuserProjectModels As Topomatic.ApplicationPlatform.Core.IProjectModel() = PluginCoreOps.FindModels(New String() {"road"})
                            If IuserProjectModels.Count > 0 Then
                                If IuserProjectModels.Count = 1 Then
                                    Dim MProj As Topomatic.ApplicationPlatform.Core.IProjectModel = IuserProjectModels.ElementAt(0)
                                    Dim Icont As IAlignmentContainer = MProj.LockRead
                                    align = Icont.Alignment
                                Else
                                    userCadView.SelectionSet.SelectOneObjectAtScreen(Function(obj) TypeOf obj Is Topomatic.Alg.Layers.Wrappers.AlignmentWrapper, "Выборите трассу: ")
                                    For Each acEnt As Object In userCadView.SelectionSet
                                        If TypeOf acEnt Is Topomatic.Alg.Layers.Wrappers.AlignmentWrapper Then
                                            Dim acAlignWrapper As Topomatic.Alg.Layers.Wrappers.AlignmentWrapper = acEnt
                                            align = acAlignWrapper.Alignment
                                            userWrapperAlign = acEnt
                                        End If
                                    Next
                                End If
                                boolSelectAlign = True
                            End If
                            If IsNothing(align) = False Then
                                Dim insblkRef As DwgBlock = RoburFunc.FuncInsertBlockByDWGFile(ActivDocument, templateDwg, nameCreateObject)
                                If IsNothing(insblkRef) = False Then
                                    Dim guidStr As Guid = align.GetId
                                    ID_Rng = guidStr.ToString
                                    Dim profs As Topomatic.Alg.Road.RoadTransitions = align.Transitions()
                                    For Each pr As Prf.Transition In profs
                                        Dim nameProf As String = pr.Name
                                        If nameProf Like "Левый кювет" Then
                                            Dim prRedProf As Prf.ProjectProfile = pr.RedProfile
                                            If prRedProf.Count > 1 Then
                                                Dim prefStation As Double = prRedProf.ElementAt(0).Station
                                                Dim prefElev As Double = prRedProf.ElementAt(0).Elevation
                                                For j As Integer = 1 To prRedProf.Count - 1
                                                    Dim elemFrof As Prf.ProjectNode = prRedProf.ElementAt(j)
                                                    Dim station As Double = elemFrof.Station
                                                    Dim elev As Double = elemFrof.Elevation
                                                    Dim dist As Double = station - prefStation
                                                    Dim i As Double = (elev - prefElev) / dist
                                                    i = Math.Round(i * 1000, 2)
                                                    Dim pkInsert As Double = (station + prefStation) / 2
                                                    Dim insertPoint As Vector2D = New Vector2D
                                                    Dim boolPk As Boolean = align.Plan.CompoundLine.StaOffsetToPos(pkInsert, 0, insertPoint)
                                                    If boolPk = True Then
                                                        'If IsNothing(blk) = False Then
                                                        '    Dim arrayattib As String(,) = {}
                                                        '    ReDim arrayattib(1, 1)
                                                        '    arrayattib(0, 0) = "УКЛОН"
                                                        '    arrayattib(1, 0) = i
                                                        '    arrayattib(0, 1) = "РАССТОЯНИЕ"
                                                        '    arrayattib(1, 1) = dist
                                                        '    Dim promptAttrib As IDictionary(Of String, String(,)) = New Dictionary(Of String, String(,))
                                                        '    promptAttrib.Add("УКЛОН", arrayattib)
                                                        'End If
                                                    End If
                                                Next j
                                            End If
                                        End If
                                    Next
                                End If

                            End If
                        Else
                            Dim insertPoint As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(0, 0, 0)
                            Dim blk As DwgBlock = RoburFunc.FuncInsertBlockByDWGFile(ActivDocument, templateDwg, nameCreateObject)
                            Dim arrayattib As String(,) = Nothing
                            If IsNothing(blk) = False Then
                                ReDim arrayattib(1, 2)
                                arrayattib(0, 0) = "height"
                                arrayattib(1, 0) = attributeHeight
                                arrayattib(0, 1) = "flag"
                                arrayattib(1, 1) = 4
                                arrayattib(0, 2) = "value"
                                arrayattib(1, 2) = ""
                                Dim promptAttrib As IDictionary(Of String, String(,)) = New Dictionary(Of String, String(,))
                                promptAttrib.Add("NAME", arrayattib)
                                Do
                                    userCadView.Lock()
                                    Dim boolClick As Boolean = CadCursors.GetPoint(userCadView, insertPoint, "Выберите положение блока: ")
                                    If boolClick = True Then
                                        Dim insBlk As DwgInsert = RoburFunc.FuncInsertBlk(ActivDocument, nameCreateObject, insertPoint, promptAttrib)
                                        If IsNothing(insBlk) = False Then
                                            If IsNothing(objectLayer) = False Then
                                                insBlk.Layer = objectLayer
                                            End If
                                            If IsNothing(objectUserTypeLine) = False Then
                                                insBlk.Linetype = objectUserTypeLine
                                            End If
                                            insBlk.Color = objectColor
                                            insBlk.Lineweight = widthUserTypeLine
                                            insBlk.Scale = New Topomatic.Cad.Foundation.Vector3D(MDraw * scaleUserObject, MDraw * scaleUserObject, MDraw * scaleUserObject)
                                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                            'записываем расширенные данные
                                            If IsArray(arrayDSDataRecord) = True Then
                                                Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXData(insBlk, PSTable, arrayDSDataRecord, templateXML, dictionaryDataobject)
                                            End If
                                        End If
                                    Else
                                        Exit Do
                                    End If
                                Loop
                                userCadView.Unlock()
                                userCadView.Invalidate()
                            End If
                        End If
                    ElseIf typeUserObject Like "Мультивыноска" Then
                        '==========================================================================================
                        'проверяем стиль текста и загружаем в случае необходимости
                        Dim ID_Rng As String = ""
                            userTextStyle = FuncStyles.FuncReturnTextStyle(ActivDocument, textStyle)
                            If IsNothing(userTextStyle) = True Then
                                Dim arrayPropertiesTextStyle As String(,) = Nothing
                                Dim boolFindTextStyleXML As Boolean = FuncXML.FuncReadTextStylePropertiesToXML(templateXML, textStyle, arrayPropertiesTextStyle)
                                If IsArray(arrayPropertiesTextStyle) = True Then
                                    Dim createTextStyle As DwgStyle = FuncStyles.FuncCreateTextStyleByArray(ActivDocument, textStyle, arrayPropertiesTextStyle)
                                    If IsNothing(createTextStyle) = False Then
                                        userTextStyle = createTextStyle
                                    End If
                                End If
                            End If
                            'создаем блок-стрелку для выноски
                            If Not (nameArrow Like "Нет") Then
                                Dim blk As DwgBlock = RoburFunc.FuncInsertBlockByDWGFile(ActivDocument, templateDwg, nameArrow)
                            End If
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            'выбираем трассу
                            Dim boolSelectAlign As Boolean = False
                            Dim align As Alignment = Nothing
                            Dim userWrapperAlign As Topomatic.Alg.Layers.Wrappers.AlignmentWrapper = Nothing
                            Dim poly As IPolyline3D = Nothing
                            If note.Trim.Length > 0 Then
                                Dim IuserProjectModels As Topomatic.ApplicationPlatform.Core.IProjectModel() = PluginCoreOps.FindModels(New String() {"road"})
                                If IuserProjectModels.Count > 0 Then
                                    If IuserProjectModels.Count = 1 Then
                                        Dim MProj As Topomatic.ApplicationPlatform.Core.IProjectModel = IuserProjectModels.ElementAt(0)
                                        Dim Icont As IAlignmentContainer = MProj.LockRead
                                        align = Icont.Alignment
                                    Else
                                        userCadView.SelectionSet.SelectOneObjectAtScreen(Function(obj) TypeOf obj Is Topomatic.Alg.Layers.Wrappers.AlignmentWrapper, "Выборите трассу: ")
                                        For Each acEnt As Object In userCadView.SelectionSet
                                            If TypeOf acEnt Is Topomatic.Alg.Layers.Wrappers.AlignmentWrapper Then
                                                Dim acAlignWrapper As Topomatic.Alg.Layers.Wrappers.AlignmentWrapper = acEnt
                                                align = acAlignWrapper.Alignment
                                                userWrapperAlign = acEnt
                                            End If
                                        Next
                                    End If
                                    boolSelectAlign = True
                                End If
                            End If
                            If IsNothing(align) = False Then
                                Dim guidStr As Guid = align.GetId
                                ID_Rng = guidStr.ToString
                            End If
                            Dim profs As Topomatic.Alg.Road.RoadTransitions = align.Transitions()
                            For Each pr As Prf.Transition In profs
                                Dim nameProf As String = pr.Name
                                If nameProf Like "Левый кювет" Then
                                    Dim prRedProf As Prf.ProjectProfile = pr.RedProfile
                                    Dim a = 1
                                End If
                            Next
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            'выбираем откос
                            Dim posSlope As Integer = note.IndexOf("<Slope>")
                            Dim boolSlope As Boolean = False
                            Dim userSlope As DwgSlope = Nothing
                            Dim arrayPointBr As Double(,) = Nothing
                            Dim arrayPointPod As Double(,) = Nothing
                            If posSlope > -1 Then
                                userCadView.SelectionSet.Clear()
                                userCadView.SelectionSet.SelectOneObjectAtScreen(Function(obj) TypeOf obj Is Topomatic.Planchet.Entities.DwgSlope, "Выборите откос: ")
                                For Each acEnt As Object In userCadView.SelectionSet
                                    If TypeOf acEnt Is Topomatic.Planchet.Entities.DwgSlope Then
                                        boolSlope = True
                                        userSlope = acEnt
                                        Dim sPoly As Topomatic.Cad.Foundation.Polyline3D = userSlope.SourcePolyline 'верх откоса
                                        Dim countArrayCoord As Integer = 0
                                        For i As Integer = 0 To sPoly.Count - 1
                                            ReDim Preserve arrayPointBr(2, countArrayCoord)
                                            arrayPointBr(0, countArrayCoord) = sPoly.Item(i).Vertex.Pos.X
                                            arrayPointBr(1, countArrayCoord) = sPoly.Item(i).Vertex.Pos.Y
                                            arrayPointBr(2, countArrayCoord) = sPoly.Item(i).Vertex.Z
                                            countArrayCoord += 1
                                        Next i
                                        Dim dPoly As Topomatic.Cad.Foundation.Polyline3D = userSlope.DestPolyline 'низ откоса
                                        countArrayCoord = 0
                                        For i As Integer = 0 To dPoly.Count - 1
                                            ReDim Preserve arrayPointPod(2, countArrayCoord)
                                            arrayPointPod(0, countArrayCoord) = dPoly.Item(i).Vertex.Pos.X
                                            arrayPointPod(1, countArrayCoord) = dPoly.Item(i).Vertex.Pos.Y
                                            arrayPointPod(2, countArrayCoord) = dPoly.Item(i).Vertex.Z
                                            countArrayCoord += 1
                                        Next i
                                        Exit For
                                    End If
                                Next
                            End If
                            If IsNothing(userSlope) = False Then
                                ID_Rng = userSlope.ObjectID
                            End If
                            '================================================================================================================
                            'трасса найдена
                            If boolSelectAlign = True Then
                                If IsNothing(align) = False Then
                                    'расставляем элементы кривой
                                    If nameUserObject Like "Элементы кривой трассы" Then
                                        Dim strNoteAlign As String = note
                                        Dim plLine As PlanLine = align.Plan
                                        If plLine.Count > 0 Then
                                            Dim dist As Double = userWrapperAlign.PlanLength
                                            Dim beta As Double = 0
                                            Dim T1 As Double = 0
                                            Dim T2 As Double = 0
                                            Dim B As Double = 0
                                            Dim D As Double = 0
                                            Dim R As Double = 0
                                            Dim L1 As Double = 0
                                            Dim L2 As Double = 0
                                            Dim K As Double = 0
                                            Dim Kpoln As Double = 0
                                            Dim numberVU As String = "0"
                                            Dim startPoint As Vector2D = New Vector2D
                                            Dim endPoint As Vector2D = New Vector2D
                                            Dim PKVU As Double = 0
                                            Dim off As Double = 0
                                            For i As Integer = 0 To plLine.Count - 1
                                                Dim plVertex As PlanLine.Vertex = plLine.Item(i)
                                                numberVU = plVertex.Name
                                                Dim v1 = plVertex.Position
                                                align.Plan.CompoundLine.PosToStaOffset(v1, PKVU, off)
                                                If plVertex.Count = 0 Then
                                                    Continue For
                                                ElseIf plVertex.Count = 1 Then
                                                    Dim sdata As PlanLine.SinglePlanData = plVertex.PlanData
                                                    beta = sdata.Beta
                                                    T1 = sdata.T1
                                                    T2 = sdata.T2
                                                    B = sdata.B
                                                    D = sdata.D
                                                    L1 = sdata.BCC - sdata.BTC
                                                    L2 = sdata.ETC - sdata.ECC
                                                    K = sdata.ECC - sdata.BCC
                                                    Kpoln = sdata.ETC - sdata.BTC
                                                    Dim boolSt As Boolean = align.Plan.CompoundLine.StaOffsetToPos(sdata.BCC, 0, startPoint)
                                                    Dim boolEn As Boolean = align.Plan.CompoundLine.StaOffsetToPos(sdata.ECC, 0, endPoint)
                                                    If boolEn = True And boolSt = True Then
                                                        R = MathFunction.GetArcRadius(startPoint, endPoint, K)
                                                    End If
                                                Else
                                                    Dim sdata As PlanLine.MultiPlanData = plVertex.PlanData
                                                    Dim Vtdata As PlanLine.MultiPlanData.DataItem = sdata.Item(0)
                                                    beta = sdata.Beta
                                                    T1 = sdata.T1
                                                    T2 = sdata.T2
                                                    D = Vtdata.D
                                                End If
                                                'номер вершины угла
                                                If strNoteAlign.IndexOf("<NumberVU>") > -1 Then
                                                    If Not (numberVU Like "0") Then
                                                        strNoteAlign = strNoteAlign.Replace("ВУ=<NumberVU>", numberVU)
                                                    End If
                                                End If
                                                'угод дельта
                                                If strNoteAlign.IndexOf("<deg>") > -1 Then
                                                    Dim strDeg As String = FuncFormatZn.RadiansToDMSString(beta, 6)
                                                    strNoteAlign = strNoteAlign.Replace("<deg>", strDeg)
                                                End If
                                                'радиус
                                                If strNoteAlign.IndexOf("<R>") > -1 Then
                                                    If R > 0 Then
                                                        Dim strRes As String = FuncFormatZn.FuncFormatCoordinate(R, 2)
                                                        strNoteAlign = strNoteAlign.Replace("<R>", strRes)
                                                    End If
                                                End If
                                                'бисектриса
                                                If strNoteAlign.IndexOf("<B>") > -1 Then
                                                    Dim strRes As String = FuncFormatZn.FuncFormatCoordinate(B, 2)
                                                    strNoteAlign = strNoteAlign.Replace("<B>", strRes)
                                                End If
                                                'полная длина кривой
                                                If strNoteAlign.IndexOf("<K>") > -1 Then
                                                    Dim strRes As String = FuncFormatZn.FuncFormatCoordinate(Kpoln, 2)
                                                    strNoteAlign = strNoteAlign.Replace("<K>", strRes)
                                                End If
                                                'тангенсы
                                                If strNoteAlign.IndexOf("<T1>") > -1 Then
                                                    Dim strRes As String = FuncFormatZn.FuncFormatCoordinate(T1, 2)
                                                    strNoteAlign = strNoteAlign.Replace("<T1>", strRes)
                                                End If
                                                If strNoteAlign.IndexOf("<T2>") > -1 Then
                                                    Dim strRes As String = FuncFormatZn.FuncFormatCoordinate(T2, 2)
                                                    strNoteAlign = strNoteAlign.Replace("<T2>", strRes)
                                                End If
                                                'переходные кривые
                                                If strNoteAlign.IndexOf("<L1>") > -1 Then
                                                    If Math.Round(L1, 2) > 0 Then
                                                        Dim strRes As String = FuncFormatZn.FuncFormatCoordinate(L1, 2)
                                                        strNoteAlign = strNoteAlign.Replace("<L1>", strRes)
                                                    End If
                                                End If
                                                If strNoteAlign.IndexOf("<L2>") > -1 Then
                                                    If Math.Round(L2, 2) > 0 Then
                                                        Dim strRes As String = FuncFormatZn.FuncFormatCoordinate(L2, 2)
                                                        strNoteAlign = strNoteAlign.Replace("<L2>", strRes)
                                                    End If
                                                End If
                                                'домер
                                                If strNoteAlign.IndexOf("<Д>") > -1 Then
                                                    Dim strRes As String = FuncFormatZn.FuncFormatCoordinate(D, 2)
                                                    strNoteAlign = strNoteAlign.Replace("<Д>", strRes)
                                                End If
                                                'пикет
                                                If strNoteAlign.IndexOf("<StationVU>") > -1 Then
                                                    Dim strRes As String = FuncFormatZn.FuncFormatPK(PKVU, 2)
                                                    strNoteAlign = strNoteAlign.Replace("<StationVU>", strRes)
                                                End If
                                                strNoteAlign = strNoteAlign.Replace("numerator=", "")
                                                strNoteAlign = strNoteAlign.Replace("denominator=", "\P")
                                                strNoteAlign = strNoteAlign.Replace(";", "")
                                                Dim leaderObj As DwgLeader = RoburFunc.FuncDrawLeader(ActivDocument, strNoteAlign, v1, New Vector2D(v1.X + 10, v1.Y + 10), MDraw, hText, userTextStyle, nameArrow, arrowSize)
                                                If IsNothing(leaderObj) = False Then
                                                    leaderObj.Layer = objectLayer
                                                    leaderObj.Linetype = objectUserTypeLine
                                                    leaderObj.Color = objectColor
                                                    leaderObj.Lineweight = widthUserTypeLine
                                                    'leaderObj.Rotation = angleMleader
                                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                                    'записываем расширенные данные
                                                    Dim nameTablePS As String = "IS_SYSTEM"
                                                    Dim arrayRec As String(,) = Nothing
                                                    ReDim Preserve arrayRec(1, 3)
                                                    arrayRec(0, 0) = "LABEL"
                                                    arrayRec(1, 0) = nameUserObject
                                                    arrayRec(0, 1) = "LOCALID"
                                                    arrayRec(1, 1) = rng.Next
                                                    arrayRec(0, 2) = "GLOBALID"
                                                    arrayRec(0, 3) = "PARENTID"
                                                    arrayRec(1, 3) = ID_Rng
                                                    If IsArray(arrayRec) = True Then
                                                        Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXData(leaderObj, nameTablePS, arrayRec)
                                                    End If
                                                End If
                                            Next i
                                            Exit Sub
                                        End If
                                    End If
                                    Do
                                        userCadView.Lock()
                                        'если это не площадной объект, предлагаем выбрать точку вставки мультивыноски
                                        Dim posAlign As Integer = note.IndexOf("<Area>")
                                        If posAlign = -1 Then
                                            Dim insertPoint As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(0, 0, 0)
                                            Dim boolClick As Boolean = CadCursors.GetPoint(userCadView, insertPoint, "Укажите точку: ")
                                            If boolClick = True Then
                                                'ищем пикетаж
                                                Dim strNote1 As String = note
                                                Dim PK As Double = -1
                                                Dim off As Double = -1
                                                Dim insPT2d As Vector2D = New Vector2D(insertPoint.X, insertPoint.Y)
                                                insPT2d.X = insertPoint.X
                                                insPT2d.Y = insertPoint.Y
                                                Dim boolFind As Boolean = align.Plan.CompoundLine.PosToStaOffset(insPT2d, PK, off)
                                                'определяем пикет трассы
                                                If boolFind = True Then
                                                    'форматируем пикет
                                                    Dim posStation As Integer = strNote1.IndexOf("<Station>")
                                                    If posStation > -1 Then
                                                        Dim StrRes As String = FuncFormatZn.FuncFormatPK(PK, 2)
                                                        strNote1 = strNote1.Replace("<Station>", StrRes)
                                                    End If
                                                End If
                                                'X
                                                Dim posX As Integer = strNote1.IndexOf("<X>")
                                                If posX > -1 Then
                                                    Dim StrRes As String = FuncFormatZn.FuncFormatCoordinate(insertPoint.Y, 2)
                                                    strNote1 = strNote1.Replace("<X>", StrRes)
                                                End If
                                                'Y
                                                Dim posY As Integer = strNote1.IndexOf("<Y>")
                                                If posY > -1 Then
                                                    Dim StrRes As String = FuncFormatZn.FuncFormatCoordinate(insertPoint.X, 2)
                                                    strNote1 = strNote1.Replace("<Y>", StrRes)
                                                End If
                                                'характеристики трассы
                                                If strNote1.IndexOf("<NumberVU>") > -1 Or strNote1.IndexOf("<deg>") > -1 Or strNote1.IndexOf("<R>") > -1 Or strNote1.IndexOf("<K>") > -1 Or strNote1.IndexOf("<T1>") > -1 Or strNote1.IndexOf("<T2>") > -1 Or strNote1.IndexOf("<Б>") > -1 Or strNote1.IndexOf("<Д>") > -1 Or strNote1.IndexOf("<StationVU>") > -1 Or strNote1.IndexOf("<L1>") > -1 Or strNote1.IndexOf("<L2>") > -1 Then
                                                    Dim plLine As PlanLine = align.Plan
                                                    If plLine.Count > 0 Then
                                                        Dim dist As Double = userWrapperAlign.PlanLength
                                                        Dim beta As Double = 0
                                                        Dim T1 As Double = 0
                                                        Dim T2 As Double = 0
                                                        Dim B As Double = 0
                                                        Dim D As Double = 0
                                                        Dim R As Double = 0
                                                        For i As Integer = 0 To plLine.Count - 1
                                                            Dim plVertex As PlanLine.Vertex = plLine.Item(i)
                                                            Dim v1 = plVertex.Position
                                                            Dim tempLenght As Double = MathFunction.funcCalcDistanceByToPoints2d(v1, insPT2d, 2)
                                                            If plVertex.Count = 0 Then
                                                                Continue For
                                                            ElseIf plVertex.Count = 1 Then
                                                                If tempLenght < dist Then
                                                                    Dim sdata As PlanLine.SinglePlanData = plVertex.PlanData
                                                                    beta = sdata.Beta
                                                                    T1 = sdata.T1
                                                                    T2 = sdata.T2
                                                                    B = sdata.B
                                                                    D = sdata.D
                                                                End If
                                                            Else
                                                                Dim sdata As PlanLine.MultiPlanData = plVertex.PlanData
                                                                Dim Vtdata As PlanLine.MultiPlanData.DataItem = sdata.Item(0)
                                                                beta = sdata.Beta
                                                                T1 = sdata.T1
                                                                T2 = sdata.T2
                                                                D = Vtdata.D
                                                            End If
                                                        Next
                                                        If strNote1.IndexOf("<T1>") > -1 Then
                                                            Dim strRes As String = FuncFormatZn.FuncFormatCoordinate(T1, 2)
                                                            strNote1 = strNote1.Replace("<T1>", strRes)
                                                        End If
                                                        If strNote1.IndexOf("<T2>") > -1 Then
                                                            Dim strRes As String = FuncFormatZn.FuncFormatCoordinate(T2, 2)
                                                            strNote1 = strNote1.Replace("<T2>", strRes)
                                                        End If
                                                        If strNote1.IndexOf("<Beta>") > -1 Then
                                                            Dim strRes As String = FuncFormatZn.FuncFormatCoordinate(beta, 2)
                                                            strNote1 = strNote1.Replace("<Beta>", strRes)
                                                        End If
                                                        If strNote1.IndexOf("<Б>") > -1 Then
                                                            Dim strRes As String = FuncFormatZn.FuncFormatCoordinate(B, 2)
                                                            strNote1 = strNote1.Replace("<Б>", strRes)
                                                        End If
                                                        If strNote1.IndexOf("<Д>") > -1 Then
                                                            Dim strRes As String = FuncFormatZn.FuncFormatCoordinate(D, 2)
                                                            strNote1 = strNote1.Replace("<Д>", strRes)
                                                        End If
                                                    End If
                                                End If
                                                'откос
                                                Dim posSelectSlope As Integer = strNote1.IndexOf("<Slope>")
                                                If posSelectSlope > -1 Then
                                                    Dim z1 As Topomatic.Cad.Foundation.Vector3D = MathFunction.funcGetNearestPointOnLine(arrayPointBr, insertPoint)
                                                    Dim z2 As Topomatic.Cad.Foundation.Vector3D = MathFunction.funcGetNearestPointOnLine(arrayPointPod, insertPoint)
                                                    Dim d As Double = Math.Sqrt((z1.X - z2.X) ^ 2 + (z1.Y - z2.Y) ^ 2)
                                                    If d > 0 Then
                                                        Dim h As Double = Math.Abs(z1.Elevation - z2.Elevation)
                                                        Dim i As Double = Math.Round(d / h, 0)
                                                        strNote1 = strNote1.Replace("<Slope>", h)
                                                    End If
                                                End If
                                                'отметка в точке
                                                Dim posSelectElev As Integer = strNote1.IndexOf("<Elevation>")
                                                If posSelectElev > -1 Then
                                                    Dim cont = align.FindOwner(Of ISurfaceContainer)()
                                                    Dim surfContainer As ISurfaceContainer = align.FindOwner(Of ISurfaceContainer)()
                                                    Dim userSurf As Surface = surfContainer.Surface
                                                    Dim elev As Double = 0
                                                    Try
                                                        elev = userSurf.GetElevation(insPT2d)
                                                    Catch ex As System.InvalidOperationException
                                                    End Try
                                                    Dim StrRes As String = FuncFormatZn.FuncFormatCoordinate(elev, 2)
                                                    strNote1 = strNote1.Replace("<Elevation>", StrRes)
                                                End If
                                                Dim distContr As Double = 15 * MDraw
                                                Dim angleMleader As Double = 0
                                                If PK >= 0 Then
                                                    Dim nextDist As Double = PK + distContr
                                                    If nextDist < (align.Plan.CompoundLine.EndPos - align.Plan.CompoundLine.StartPos).Length Then
                                                        Dim stPoint As Vector2D
                                                        Dim nextPoint As Vector2D
                                                        Dim boolPt1 As Boolean = align.Plan.CompoundLine.StaOffsetToPos(PK, 0, stPoint)
                                                        Dim boolPt2 As Boolean = align.Plan.CompoundLine.StaOffsetToPos(nextDist, 0, nextPoint)
                                                        If boolPt1 = True And boolPt2 = True Then
                                                            angleMleader = (stPoint - nextPoint).Angle
                                                            Select Case angleMleader
                                                                Case Math.PI / 2 To Math.PI
                                                                    angleMleader = angleMleader + Math.PI
                                                                Case Math.PI To (3 * Math.PI) / 2
                                                                    angleMleader = angleMleader - Math.PI
                                                            End Select
                                                        End If
                                                    Else
                                                        nextDist = PK - distContr
                                                        If nextDist < (align.Plan.CompoundLine.EndPos - align.Plan.CompoundLine.StartPos).Length Then
                                                            Dim stPoint As Vector2D
                                                            Dim nextPoint As Vector2D
                                                            Dim boolPt1 As Boolean = align.Plan.CompoundLine.StaOffsetToPos(PK, 0, stPoint)
                                                            Dim boolPt2 As Boolean = align.Plan.CompoundLine.StaOffsetToPos(nextDist, 0, nextPoint)
                                                            If boolPt1 = True And boolPt2 = True Then
                                                                angleMleader = (stPoint - nextPoint).Angle
                                                                If angleMleader > Math.PI Then
                                                                    angleMleader = angleMleader - Math.PI
                                                                Else
                                                                    angleMleader = angleMleader + Math.PI
                                                                End If
                                                                Select Case angleMleader
                                                                    Case Math.PI / 2 To Math.PI
                                                                        angleMleader = angleMleader + Math.PI
                                                                    Case Math.PI To (3 * Math.PI) / 2
                                                                        angleMleader = angleMleader - Math.PI
                                                                End Select
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                                Dim strNote As String = strNote1.Replace("numerator=", "")
                                                strNote = strNote.Replace("denominator=", "\P")
                                                strNote = strNote.Replace(";", "")
                                                Dim leaderObj As DwgLeader = RoburFunc.FuncDrawLeader(ActivDocument, strNote, insPT2d, New Vector2D(insPT2d.X + 10, insPT2d.Y + 10), MDraw, hText, userTextStyle, nameArrow, arrowSize)
                                                If IsNothing(leaderObj) = False Then
                                                    leaderObj.Layer = objectLayer
                                                    leaderObj.Linetype = objectUserTypeLine
                                                    leaderObj.Color = objectColor
                                                    leaderObj.Lineweight = widthUserTypeLine
                                                    leaderObj.Rotation = angleMleader
                                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                                    'записываем расширенные данные
                                                    Dim nameTablePS As String = "IS_SYSTEM"
                                                    Dim arrayRec As String(,) = Nothing
                                                    ReDim Preserve arrayRec(1, 3)
                                                    arrayRec(0, 0) = "LABEL"
                                                    arrayRec(1, 0) = nameUserObject
                                                    arrayRec(0, 1) = "LOCALID"
                                                    arrayRec(1, 1) = rng.Next
                                                    arrayRec(0, 2) = "GLOBALID"
                                                    arrayRec(0, 3) = "PARENTID"
                                                    arrayRec(1, 3) = ID_Rng
                                                    If IsArray(arrayRec) = True Then
                                                        Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXData(leaderObj, nameTablePS, arrayRec)
                                                    End If
                                                End If
                                            Else
                                                Exit Do
                                            End If
                                        Else
                                            'выбираем площадной объект
                                            userCadView.SelectionSet.Clear()
                                            Dim strNote As String = ""
                                            Dim ptSelect As Object = userCadView.SelectionSet.SelectOneObjectAtScreen(Function(obj) TypeOf obj Is Object, "Выборите Участок или Штриховку: ")
                                            If userCadView.SelectionSet.Count = 0 Then
                                                Exit Do
                                            Else
                                                For Each acEnt As Object In userCadView.SelectionSet
                                                    Dim areaStr As String = "0"
                                                    If TypeOf acEnt Is Topomatic.Sfc.Layer.Wrappers.SurfacePatchWrapper Then
                                                        Dim UserPatch As Topomatic.Sfc.Layer.Wrappers.SurfacePatchWrapper = acEnt
                                                        Dim AreaRez As Double = UserPatch.Aread2d
                                                        areaStr = FuncFormatZn.FuncFormatCoordinate(AreaRez, 0)
                                                        strNote = note.Replace("<Area>", areaStr)
                                                    ElseIf TypeOf acEnt Is DwgHatch Then
                                                        Dim UserHatch As DwgHatch = acEnt
                                                        Dim listB As BoundaryPathList = UserHatch.BoundaryPath
                                                        Dim areaHatch As Double = 0
                                                        If listB.Count > 0 Then
                                                            For i As Integer = 0 To listB.Count - 1
                                                                Dim boundPat As BoundaryPath = listB.Item(i)
                                                                Dim vertList As List(Of Vector2D) = boundPat.Contur
                                                                Dim tempArea As Double = MathFunction.FuncSQRTPolylineByList(vertList)
                                                                If boundPat.IsExternal = True Then
                                                                    areaHatch += tempArea
                                                                Else
                                                                    areaHatch -= tempArea
                                                                End If
                                                            Next i
                                                        End If
                                                        areaStr = FuncFormatZn.FuncFormatCoordinate(areaHatch)
                                                        strNote = note.Replace("<Area>", areaStr)
                                                    End If
                                                    'Dim insertPoint As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(0, 0, 0)
                                                    Dim insertPoint As Topomatic.Cad.Foundation.Vector3D = userCadView.LastPoint
                                                    'Dim boolClick As Boolean = CadCursors.GetPoint(userCadView, insertPoint, "Укажите точку: ")
                                                    'If boolClick = True Then
                                                    Dim insPT2d As Vector2D = New Vector2D(insertPoint.X, insertPoint.Y)
                                                    'ищем пикетаж
                                                    Dim PK As Double = -1
                                                    Dim off As Double = -1
                                                    Dim boolFind As Boolean = align.Plan.CompoundLine.PosToStaOffset(insPT2d, PK, off)
                                                    Dim distContr As Double = 15 * MDraw
                                                    Dim angleMleader As Double = 0
                                                    If PK >= 0 Then
                                                        Dim nextDist As Double = PK + distContr
                                                        If nextDist < (align.Plan.CompoundLine.EndPos - align.Plan.CompoundLine.StartPos).Length Then
                                                            Dim stPoint As Vector2D
                                                            Dim nextPoint As Vector2D
                                                            Dim boolPt1 As Boolean = align.Plan.CompoundLine.StaOffsetToPos(PK, 0, stPoint)
                                                            Dim boolPt2 As Boolean = align.Plan.CompoundLine.StaOffsetToPos(nextDist, 0, nextPoint)
                                                            If boolPt1 = True And boolPt2 = True Then
                                                                angleMleader = (stPoint - nextPoint).Angle
                                                                Select Case angleMleader
                                                                    Case Math.PI / 2 To Math.PI
                                                                        angleMleader = angleMleader + Math.PI
                                                                    Case Math.PI To (3 * Math.PI) / 2
                                                                        angleMleader = angleMleader - Math.PI
                                                                End Select
                                                            End If
                                                        Else
                                                            nextDist = PK - distContr
                                                            If nextDist < (align.Plan.CompoundLine.EndPos - align.Plan.CompoundLine.StartPos).Length Then
                                                                Dim stPoint As Vector2D
                                                                Dim nextPoint As Vector2D
                                                                Dim boolPt1 As Boolean = align.Plan.CompoundLine.StaOffsetToPos(PK, 0, stPoint)
                                                                Dim boolPt2 As Boolean = align.Plan.CompoundLine.StaOffsetToPos(nextDist, 0, nextPoint)
                                                                If boolPt1 = True And boolPt2 = True Then
                                                                    angleMleader = (stPoint - nextPoint).Angle
                                                                    Select Case angleMleader
                                                                        Case Math.PI / 2 To Math.PI
                                                                            angleMleader = angleMleader + Math.PI
                                                                        Case Math.PI To (3 * Math.PI) / 2
                                                                            angleMleader = angleMleader - Math.PI
                                                                    End Select
                                                                End If
                                                            End If
                                                        End If
                                                    End If

                                                    strNote = strNote.Replace("numerator=", "")
                                                    strNote = strNote.Replace("denominator=", "\P")
                                                    strNote = strNote.Replace(";", "")
                                                    Dim leaderObj As DwgLeader = RoburFunc.FuncDrawLeader(ActivDocument, strNote, insPT2d, New Vector2D(insPT2d.X + 10, insPT2d.Y + 10), MDraw, hText, userTextStyle, nameArrow, arrowSize)
                                                    If IsNothing(leaderObj) = False Then
                                                        leaderObj.Layer = objectLayer
                                                        leaderObj.Linetype = objectUserTypeLine
                                                        leaderObj.Color = objectColor
                                                        leaderObj.Lineweight = widthUserTypeLine
                                                        leaderObj.Rotation = angleMleader
                                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                                        'записываем расширенные данные
                                                        Dim nameTablePS As String = "IS_SYSTEM"
                                                        Dim arrayRec As String(,) = Nothing
                                                        ReDim Preserve arrayRec(1, 3)
                                                        arrayRec(0, 0) = "LABEL"
                                                        arrayRec(1, 0) = nameUserObject
                                                        arrayRec(0, 1) = "LOCALID"
                                                        arrayRec(1, 1) = rng.Next
                                                        arrayRec(0, 2) = "GLOBALID"
                                                        arrayRec(0, 3) = "PARENTID"
                                                        arrayRec(1, 3) = ID_Rng
                                                        If IsArray(arrayRec) = True Then
                                                            Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXData(leaderObj, nameTablePS, arrayRec)
                                                        End If
                                                    End If
                                                Next
                                            End If
                                        End If
                                    Loop
                                End If
                            End If
                        ElseIf typeUserObject Like "Полилиния" Or typeUserObject Like "Линия" Or typeUserObject Like "Характерная линия" Or typeUserObject Like "Откос" Or typeUserObject Like "Дополнительная ось" Or typeUserObject Like "Бордюр" Or typeUserObject Like "Бордюр" Or typeUserObject Like "Структурная линия" Or typeUserObject Like "Дорожная разметка" Or typeUserObject Like "Ось трассы" Then
                            If userCadView IsNot Nothing Then
                            Dim positions As New List(Of Vector2D)()

                            Dim dynamic_draw As DrawCursorEvent = Sub(pen As CadPen, vertex As Vector2D)
                                                                      If positions.Count > 0 Then
                                                                          pen.Color = objectColor.Win32Color
                                                                          pen.ComplexLinetype = objectUserTypeLine.LinetypePattern
                                                                          pen.BeginDraw()
                                                                          Try
                                                                              For i As Integer = 1 To positions.Count - 1
                                                                                  'Используем рефлексию для вызова метода
                                                                                  Dim method = pen.GetType().GetMethod("DrawLine",
                                            New Type() {GetType(Vector2D), GetType(Vector2D)})
                                                                                  method.Invoke(pen, New Object() {
                                            New Vector2D(positions(i - 1)),
                                            New Vector2D(positions(i))
                                        })
                                                                              Next
                                                                              Dim method2 = pen.GetType().GetMethod("DrawLine",
                                        New Type() {GetType(Vector2D), GetType(Vector2D)})
                                                                              method2.Invoke(pen, New Object() {
                                        New Vector2D(positions(positions.Count - 1)),
                                        vertex
                                    })
                                                                          Finally
                                                                              pen.EndDraw()
                                                                          End Try
                                                                      End If
                                                                  End Sub

                            AddHandler userCadView.DynamicDraw, dynamic_draw
                            Try
                                Dim pos As Topomatic.Cad.Foundation.Vector3D
                                While CadCursors.GetPoint(userCadView, pos, "Укажите точку")
                                    positions.Add(pos.Pos)
                                End While
                                If positions.Count > 1 Then
                                    Dim newPolyline As DwgPolyline = RoburFunc.FuncDrawPolylineFromList2d(ActivDocument, positions)
                                    If IsNothing(newPolyline) = False Then
                                        newPolyline.Layer = objectLayer
                                        newPolyline.Linetype = objectUserTypeLine
                                        newPolyline.Color = objectColor
                                        newPolyline.Lineweight = widthUserTypeLine
                                        newPolyline.LinetypeScale = scaleUserTypeLine
                                        newPolyline.Width = gWidthUserTypeLine
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'записываем расширенные данные
                                        Dim nameTablePS As String = "IS_SYSTEM"
                                        Dim arrayRec As String(,) = Nothing
                                        ReDim Preserve arrayRec(1, 3)
                                        arrayRec(0, 0) = "LABEL"
                                        arrayRec(1, 0) = nameUserObject
                                        arrayRec(0, 1) = "LOCALID"
                                        arrayRec(1, 1) = rng.Next
                                        arrayRec(0, 2) = "GLOBALID"
                                        arrayRec(0, 3) = "PARENTID"
                                        arrayRec(1, 3) = newPolyline.ObjectID
                                        If IsArray(arrayRec) = True Then
                                            Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXData(newPolyline, nameTablePS, arrayRec)
                                        End If
                                    End If
                                End If
                            Finally
                                RemoveHandler userCadView.DynamicDraw, dynamic_draw
                            End Try
                        End If
                    End If
                Else
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'выбор объектов типа Robur для преобразования
                    Dim countParcels As Integer = 1
                    Do
                        Try
                            userCadView.SelectionSet.Clear()
                            userCadView.SelectionSet.SelectOneObjectAtScreen(Function(obj) TypeOf obj Is Object, "Выборите объект: ")
                            If userCadView.SelectionSet.Count = 0 Then
                                Exit Do
                            End If
                        Catch ex As System.NullReferenceException
                            MsgBox("Активный документ не найден. Попробуйте заново выбрать модель в списке!!!")
                            Exit Do
                        End Try
                        'код объекта
                        For Each acEnt As Object In userCadView.SelectionSet
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            'обрабатываем откос
                            Dim strType As String = acEnt.GetType.ToString
                            'If TypeOf acEnt Is Topomatic.Planchet.Entities.Then Then

                            'End If
                            If TypeOf acEnt Is Topomatic.Planchet.Entities.DwgSlope Then

                                If typeUserObject.Trim Like "Откос" Then
                                    Dim UserSlope As Topomatic.Planchet.Entities.DwgSlope = acEnt
                                    If nameUserObject.Trim Like "Откос проектируемый автомобильной дороги" Then
                                        '===========================================================================================
                                        Dim IpolyDest As Polyline3D = UserSlope.DestPolyline
                                        Dim IpolySourse As Polyline3D = UserSlope.SourcePolyline
                                        Dim newUserSlope As DwgSlope = RoburFunc.FuncDrawSlope(ActivDocument, IpolySourse, IpolyDest)
                                        If IsNothing(newUserSlope) = False Then
                                            newUserSlope.Color = objectColor
                                            newUserSlope.Lineweight = widthUserTypeLine
                                            newUserSlope.Layer = objectLayer
                                            newUserSlope.LinetypeScale = scaleUserTypeLine
                                            newUserSlope.Linetype = objectUserTypeLine
                                            newUserSlope.DisplaySourceLinetype = False
                                            newUserSlope.DisplaySourceLinetype = False
                                            newUserSlope.SlopeType = UserSlope.SlopeType
                                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                            'записываем расширенные данные
                                            Dim nameTablePS As String = "IS_SYSTEM"
                                            Dim arrayRec As String(,) = Nothing
                                            ReDim Preserve arrayRec(1, 3)
                                            arrayRec(0, 0) = "LABEL"
                                            arrayRec(1, 0) = nameUserObject
                                            arrayRec(0, 1) = "LOCALID"
                                            arrayRec(1, 1) = rng.Next
                                            arrayRec(0, 2) = "GLOBALID"
                                            arrayRec(0, 3) = "PARENTID"
                                            arrayRec(1, 3) = UserSlope.ObjectID
                                            If IsArray(arrayRec) = True Then
                                                Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXData(newUserSlope, nameTablePS, arrayRec)
                                            End If
                                        End If
                                    ElseIf nameUserObject.Trim Like "Бровка проектируемого откоса автомобильной дороги" Then
                                        '===========================================================================================
                                        'рисуем бровку откоса
                                        Dim sPoly As Topomatic.Cad.Foundation.Polyline3D = UserSlope.SourcePolyline 'верх откоса
                                        Dim ArrayCoord As Double(,) = Nothing
                                        Dim countArrayCoord As Integer = 0
                                        For i As Integer = 0 To sPoly.Count - 1
                                            ReDim Preserve ArrayCoord(1, countArrayCoord)
                                            ArrayCoord(0, countArrayCoord) = sPoly.Item(i).Vertex.Pos.X
                                            ArrayCoord(1, countArrayCoord) = sPoly.Item(i).Vertex.Pos.Y
                                            countArrayCoord += 1
                                        Next i
                                        If IsArray(ArrayCoord) = True Then
                                            If ArrayCoord.GetUpperBound(1) > 0 Then
                                                Dim TempPoly As DwgPolyline = RoburFunc.FuncDrawPolylineToArrayCoord(ActivDocument, ArrayCoord)
                                                If IsNothing(TempPoly) = False Then
                                                    If IsNothing(objectLayer) = False Then
                                                        TempPoly.Layer = objectLayer
                                                    End If
                                                    If IsNothing(objectUserTypeLine) = False Then
                                                        TempPoly.Linetype = objectUserTypeLine
                                                    End If
                                                    TempPoly.Color = objectColor
                                                    TempPoly.Lineweight = widthUserTypeLine
                                                    TempPoly.LinetypeScale = MDraw * scaleUserObject
                                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                                    'записываем расширенные данные
                                                    Dim nameTablePS As String = "IS_SYSTEM"
                                                    Dim arrayRec As String(,) = Nothing
                                                    ReDim Preserve arrayRec(1, 3)
                                                    arrayRec(0, 0) = "LABEL"
                                                    arrayRec(1, 0) = nameUserObject
                                                    arrayRec(0, 1) = "LOCALID"
                                                    arrayRec(1, 1) = rng.Next
                                                    arrayRec(0, 2) = "GLOBALID"
                                                    arrayRec(0, 3) = "PARENTID"
                                                    arrayRec(1, 3) = UserSlope.ObjectID
                                                    If IsArray(arrayRec) = True Then
                                                        Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXData(TempPoly, nameTablePS, arrayRec)
                                                    End If
                                                End If
                                            End If
                                        End If
                                    ElseIf nameUserObject.Trim Like "Подошва проектируемого откоса автомобильной дороги" Then
                                        '===========================================================================================
                                        'рисуем подошву откоса
                                        Dim dPoly As Topomatic.Cad.Foundation.Polyline3D = UserSlope.DestPolyline 'верх откоса
                                        Dim ArrayCoord As Double(,) = Nothing
                                        Dim countArrayCoord As Integer = 0
                                        For i As Integer = 0 To dPoly.Count - 1
                                            ReDim Preserve ArrayCoord(1, countArrayCoord)
                                            ArrayCoord(0, countArrayCoord) = dPoly.Item(i).Vertex.Pos.X
                                            ArrayCoord(1, countArrayCoord) = dPoly.Item(i).Vertex.Pos.Y
                                            countArrayCoord += 1
                                        Next i
                                        If IsArray(ArrayCoord) = True Then
                                            If ArrayCoord.GetUpperBound(1) > 0 Then
                                                Dim TempPoly As DwgPolyline = RoburFunc.FuncDrawPolylineToArrayCoord(ActivDocument, ArrayCoord)
                                                If IsNothing(TempPoly) = False Then
                                                    If IsNothing(objectLayer) = False Then
                                                        TempPoly.Layer = objectLayer
                                                    End If
                                                    If IsNothing(objectUserTypeLine) = False Then
                                                        TempPoly.Linetype = objectUserTypeLine
                                                    End If
                                                    TempPoly.Color = objectColor
                                                    TempPoly.Lineweight = widthUserTypeLine
                                                    TempPoly.LinetypeScale = MDraw * scaleUserObject
                                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                                    'записываем расширенные данные
                                                    Dim nameTablePS As String = "IS_SYSTEM"
                                                    Dim arrayRec As String(,) = Nothing
                                                    ReDim Preserve arrayRec(1, 3)
                                                    arrayRec(0, 0) = "LABEL"
                                                    arrayRec(1, 0) = nameUserObject
                                                    arrayRec(0, 1) = "LOCALID"
                                                    arrayRec(1, 1) = rng.Next
                                                    arrayRec(0, 2) = "GLOBALID"
                                                    arrayRec(0, 3) = "PARENTID"
                                                    arrayRec(1, 3) = UserSlope.ObjectID
                                                    If IsArray(arrayRec) = True Then
                                                        Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXData(TempPoly, nameTablePS, arrayRec)
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                'структурные линии
                            ElseIf TypeOf acEnt Is Topomatic.Sfc.Layer.Wrappers.SurfaceStructureLineWrapper Then
                                If typeUserObject.Trim Like "Структурная линия" Or typeUserObject.Trim Like "Характерная линия" Then
                                    Dim UserStructureLine As Topomatic.Sfc.Layer.Wrappers.SurfaceStructureLineWrapper = acEnt
                                    Dim keyPairs As KeyValuePair(Of Integer, Surface)
                                    'код структурной линии
                                    keyPairs = UserStructureLine.LinearObjectCode
                                    Dim UserKey As Integer = keyPairs.Key
                                    If nameCategoryTable Like "Цифровая модель рельефа" Then
                                        Dim numberCode As Integer = 0
                                        For j As Integer = 0 To arraytag.GetUpperBound(1)
                                            Dim nameCode As String = arraytag(0, j)
                                            If nameCode Like "Object_code" Then
                                                Dim codeStructLine As Integer = arraytag(1, j)
                                                If IsNumeric(codeStructLine) = True Then
                                                    numberCode = Val(codeStructLine)
                                                    Exit For
                                                End If
                                            End If
                                        Next
                                        Dim dict As Dictionary(Of Integer, Surface) = New Dictionary(Of Integer, Surface)
                                        dict.Add(numberCode, Nothing)
                                        UserStructureLine.LinearObjectCode = dict.ElementAt(0)
                                    Else
                                        '/////////////////////////////////////////////////////////////////////////////////////////////////////
                                        Dim ArrayCoord As Double(,) = Nothing
                                        Dim countarrayPos As Integer = 0
                                        For i As Integer = 0 To UserStructureLine.Count - 1
                                            UserStructureLine.CurrentVertexIndex = i
                                            Dim pos = UserStructureLine.Vertex
                                            ReDim Preserve ArrayCoord(1, countarrayPos)
                                            ArrayCoord(0, countarrayPos) = pos.X
                                            ArrayCoord(1, countarrayPos) = pos.Y
                                            countarrayPos += 1
                                        Next
                                        If IsArray(ArrayCoord) = True Then
                                            If ArrayCoord.GetUpperBound(1) > 0 Then
                                                Dim TempPoly As DwgPolyline = RoburFunc.FuncDrawPolylineToArrayCoord(ActivDocument, ArrayCoord)
                                                If IsNothing(TempPoly) = False Then
                                                    If IsNothing(objectLayer) = False Then
                                                        TempPoly.Layer = objectLayer
                                                    End If
                                                    If IsNothing(objectUserTypeLine) = False Then
                                                        TempPoly.Linetype = objectUserTypeLine
                                                    End If
                                                    TempPoly.Color = objectColor
                                                    TempPoly.Lineweight = widthUserTypeLine
                                                    TempPoly.LinetypeScale = MDraw * scaleUserObject
                                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                                    'записываем расширенные данные
                                                    Dim nameTablePS As String = "IS_SYSTEM"
                                                    Dim arrayRec As String(,) = Nothing
                                                    ReDim Preserve arrayRec(1, 3)
                                                    arrayRec(0, 0) = "LABEL"
                                                    arrayRec(1, 0) = nameUserObject
                                                    arrayRec(0, 1) = "LOCALID"
                                                    arrayRec(1, 1) = rng.Next
                                                    arrayRec(0, 2) = "GLOBALID"
                                                    arrayRec(0, 3) = "PARENTID"
                                                    arrayRec(1, 3) = ""
                                                    If IsArray(arrayRec) = True Then
                                                        Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXData(TempPoly, nameTablePS, arrayRec)
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                'участок
                            ElseIf TypeOf acEnt Is Topomatic.Sfc.Layer.Wrappers.SurfacePatchWrapper Then
                                If typeUserObject.Trim Like "Участок" Then
                                    Dim UserPatch As Topomatic.Sfc.Layer.Wrappers.SurfacePatchWrapper = acEnt
                                    Dim keyPairs As KeyValuePair(Of Integer, Surface)
                                    'код структурной линии
                                    keyPairs = UserPatch.ObjectCode
                                    Dim UserKey As Integer = keyPairs.Key
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'рисуем штриховку
                                    Dim boolDrawHatch As Boolean = RoburFunc.FuncDrawHatchByParcels(ActivDocument, UserPatch, nameUserObject, objectLayer, objectColor, scaleUserTypeLine * MDraw, widthUserTypeLine, angleUserObject, scaleUserObject * MDraw)
                                End If
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                'дополнительные оси, бордюры и т.п
                            ElseIf TypeOf acEnt Is IWrapped Then
                                Dim objectIWrapped As IWrapped = acEnt
                                Dim objWrapped As Object = objectIWrapped.WrappedObject
                                Dim acPoly3d As Polyline3D = New Polyline3D
                                'ищем характеристики в мссиве
                                Dim ArrayRez As String(,) = Nothing
                                If TypeOf objWrapped Is AxisStrip Then
                                    If typeUserObject.Trim Like "Дополнительная ось" Then
                                        Dim acObjectAxisStrip As AxisStrip = objWrapped
                                        acObjectAxisStrip.GetPolyline(acPoly3d)
                                    End If
                                ElseIf TypeOf objWrapped Is BorderStrip Then
                                    If typeUserObject.Trim Like "Бордюр" Then
                                        Dim acObjectBorderStrip As BorderStrip = objWrapped
                                        acObjectBorderStrip.GetPolyline(acPoly3d)
                                    End If
                                ElseIf TypeOf acEnt Is Topomatic.Rsf.Layers.Wrappers.FenceWrapper Then
                                    If typeUserObject.Trim Like "Дорожное ограждение" Then
                                        Dim userFence As Topomatic.Rsf.Layers.Wrappers.FenceWrapper = acEnt
                                        userFence.GetPolyline(acPoly3d)
                                    End If
                                ElseIf TypeOf acEnt Is Topomatic.Alg.Layers.Wrappers.AlignmentWrapper Then
                                    If typeUserObject.Trim Like "Ось трассы" Then
                                        Dim userWrapperAlign As Topomatic.Alg.Layers.Wrappers.AlignmentWrapper = acEnt
                                        Dim userAlign As Topomatic.Alg.Alignment = userWrapperAlign.Alignment
                                        Using receiver = ActiveAlignmentReciver(Of RoadAlignment).CreateReciver(True)
                                            Dim ralignment As RoadAlignment = receiver.Alignment
                                            Dim urb As UrbParams = ralignment.Urb
                                            Dim desk As DescentStrip = urb.Item(15)
                                        End Using
                                        Dim cross As Topomatic.Alg.Road.Crossing.Crossings = userAlign.GetCrossings
                                        If cross.Count > 0 Then
                                            For i As Integer = 0 To cross.Count - 1
                                                Dim g As Intersection = cross.Item(i)
                                            Next i
                                        End If
                                        userWrapperAlign.GetPolyline(acPoly3d)
                                    ElseIf typeUserObject.Trim Like "Тангенс" Then
                                        Dim userWrapperAlign As Topomatic.Alg.Layers.Wrappers.AlignmentWrapper = acEnt
                                        Dim userAlign As Topomatic.Alg.Alignment = userWrapperAlign.Alignment
                                        'достаем тангенсы и вершины углов
                                        Dim arrayPropertiesObject As String(,) = Nothing
                                        '============================================================================
                                        Dim tnLayer As DwgLayer = ActivDocument.ActiveLayer
                                        Dim tnTypeLine As DwgLinetype = ActivDocument.ActiveLinetype
                                        Dim tnColor As CadColor = New CadColor(7)
                                        Dim tnScaleTypeLine As Double = 1
                                        Dim tnWidthLine As Double = 20
                                        Dim boolFindData As Boolean = FuncXML.FuncReadDataObjectToXMLbyDesk(templateXML, nameCategoryTable, nameTables, typeUserObject, "Тангенциальный ход", arrayPropertiesObject)
                                        If IsArray(arrayPropertiesObject) = True Then
                                            Dim boolCreateStyle As Boolean = FuncDecorPlane.FuncCreateStyleLinearObject(ActivDocument, templateXML, arrayPropertiesObject, tnLayer, tnTypeLine, tnColor, tnScaleTypeLine, tnWidthLine)
                                        End If
                                        Dim plLine As PlanLine = userAlign.Plan
                                        Dim vert As Topomatic.Alg.Plan.PlanLine.Vertex
                                        Dim tn As Topomatic.Alg.Plan.PlanLine.PlanData
                                        If plLine.Count > 0 Then
                                            For i As Integer = 0 To plLine.Count - 1
                                                Dim plData As PlanLine.PlanData = plLine.Item(i).PlanData
                                                Dim stStation As Double = plData.StartStation
                                                Dim enStation As Double = plData.EndStation
                                                Dim posStart As Vector2D = New Vector2D(-1, -1)
                                                Dim boolPosition As Boolean = userAlign.Plan.CompoundLine.StaOffsetToPos(stStation, 0, posStart)
                                                Dim posEnd As Vector2D = New Vector2D(-1, -1)
                                                Dim boolPosition1 As Boolean = userAlign.Plan.CompoundLine.StaOffsetToPos(enStation, 0, posEnd)
                                                Dim VU As Topomatic.Cad.Foundation.Vector3D = plLine.Item(i).Position
                                                Dim posStart3d As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(posStart.X, posStart.Y, 0)
                                                Dim posEnd3d As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(posEnd.X, posEnd.Y, 0)
                                                Dim t1 As DwgLine = RoburFunc.FuncDrawLine(ActivDocument, posStart3d, VU)
                                                If IsNothing(t1) = False Then
                                                    If IsNothing(tnLayer) = False Then
                                                        t1.Layer = tnLayer
                                                    End If
                                                    If IsNothing(tnTypeLine) = False Then
                                                        t1.Linetype = tnTypeLine
                                                    End If
                                                    t1.Color = tnColor
                                                    t1.Lineweight = tnWidthLine
                                                    t1.LinetypeScale = tnScaleTypeLine * MDraw
                                                End If
                                                Dim t2 As DwgLine = RoburFunc.FuncDrawLine(ActivDocument, VU, posEnd3d)
                                                If IsNothing(t2) = False Then
                                                    If IsNothing(tnLayer) = False Then
                                                        t2.Layer = tnLayer
                                                    End If
                                                    If IsNothing(tnTypeLine) = False Then
                                                        t2.Linetype = tnTypeLine
                                                    End If
                                                    t2.Color = tnColor
                                                    t2.Lineweight = tnWidthLine
                                                    t2.LinetypeScale = tnScaleTypeLine * MDraw
                                                End If
                                            Next
                                        End If
                                    ElseIf typeUserObject.Trim Like "Блок" Then
                                        If nameUserObject Like "Вершина угла поворота трассы" Then
                                            Dim blk As DwgBlock = RoburFunc.FuncInsertBlockByDWGFile(ActivDocument, templateDwg, nameCreateObject)
                                            If IsNothing(blk) = False Then
                                                Dim userWrapperAlign As Topomatic.Alg.Layers.Wrappers.AlignmentWrapper = acEnt
                                                Dim userAlign As Topomatic.Alg.Alignment = userWrapperAlign.Alignment
                                                Dim plLine As PlanLine = userAlign.Plan
                                                If plLine.Count > 0 Then
                                                    For i As Integer = 0 To plLine.Count - 1
                                                        Dim plVertex As PlanLine.Vertex = plLine.Item(i)
                                                        Dim nameVertex As String = plVertex.Name
                                                        Dim posVertex As Vector2D = plVertex.Position

                                                        Dim arrayattib As String(,) = Nothing
                                                        ReDim arrayattib(1, 2)
                                                        arrayattib(0, 0) = "height"
                                                        arrayattib(1, 0) = attributeHeight
                                                        arrayattib(0, 1) = "flag"
                                                        arrayattib(1, 1) = 4
                                                        arrayattib(0, 2) = "value"
                                                        arrayattib(1, 2) = nameVertex
                                                        Dim promptAttrib As IDictionary(Of String, String(,)) = New Dictionary(Of String, String(,))
                                                        promptAttrib.Add("NAME", arrayattib)

                                                        Dim insBlk As DwgInsert = RoburFunc.FuncInsertBlk(ActivDocument, nameCreateObject, posVertex, promptAttrib)
                                                        If IsNothing(insBlk) = False Then
                                                            If IsNothing(objectLayer) = False Then
                                                                insBlk.Layer = objectLayer
                                                            End If
                                                            If IsNothing(objectUserTypeLine) = False Then
                                                                insBlk.Linetype = objectUserTypeLine
                                                            End If
                                                            insBlk.Color = objectColor
                                                            insBlk.Lineweight = widthUserTypeLine
                                                            insBlk.Scale = New Topomatic.Cad.Foundation.Vector3D(MDraw * scaleUserObject, MDraw * scaleUserObject, MDraw * scaleUserObject)
                                                        End If
                                                    Next i
                                                End If
                                            End If
                                        ElseIf nameUserObject Like "Пикет" Then
                                            Dim blk As DwgBlock = RoburFunc.FuncInsertBlockByDWGFile(ActivDocument, templateDwg, nameCreateObject)
                                            If IsNothing(blk) = False Then
                                                Dim userWrapperAlign As Topomatic.Alg.Layers.Wrappers.AlignmentWrapper = acEnt
                                                Dim userAlign As Topomatic.Alg.Alignment = userWrapperAlign.Alignment
                                                Dim plLine As PlanLine = userAlign.Plan
                                                Dim listStations As IEnumerable(Of Double) = userAlign.Stationing.Stations
                                                If listStations.Count > 0 Then
                                                    For i As Integer = 0 To listStations.Count - 1
                                                        Dim userStation As Double = listStations.ElementAt(i)
                                                        Dim pos As Vector2D = New Vector2D(0, 0)
                                                        userAlign.Plan.CompoundLine.StaOffsetToPos(userStation, 0, pos)
                                                    Next i
                                                End If
                                                Dim st As List(Of Double) = New List(Of Double)
                                                userAlign.StationsToString(st)
                                                'Dim km As AlgAlignmentKilometres = userAlign.Kilometres

                                                If plLine.Count > 0 Then
                                                    For i As Integer = 0 To plLine.Count - 1
                                                        Dim plVertex As PlanLine.Vertex = plLine.Item(i)
                                                        Dim nameVertex As String = plVertex.Name
                                                        Dim posVertex As Vector2D = plVertex.Position

                                                        Dim arrayattib As String(,) = Nothing
                                                        ReDim arrayattib(1, 2)
                                                        arrayattib(0, 0) = "height"
                                                        arrayattib(1, 0) = attributeHeight
                                                        arrayattib(0, 1) = "flag"
                                                        arrayattib(1, 1) = 4
                                                        arrayattib(0, 2) = "value"
                                                        arrayattib(1, 2) = nameVertex
                                                        Dim promptAttrib As IDictionary(Of String, String(,)) = New Dictionary(Of String, String(,))
                                                        promptAttrib.Add("NAME", arrayattib)

                                                        Dim insBlk As DwgInsert = RoburFunc.FuncInsertBlk(ActivDocument, nameCreateObject, posVertex, promptAttrib)
                                                        If IsNothing(insBlk) = False Then
                                                            If IsNothing(objectLayer) = False Then
                                                                insBlk.Layer = objectLayer
                                                            End If
                                                            If IsNothing(objectUserTypeLine) = False Then
                                                                insBlk.Linetype = objectUserTypeLine
                                                            End If
                                                            insBlk.Color = objectColor
                                                            insBlk.Lineweight = widthUserTypeLine
                                                            insBlk.Scale = New Topomatic.Cad.Foundation.Vector3D(MDraw * scaleUserObject, MDraw * scaleUserObject, MDraw * scaleUserObject)
                                                        End If
                                                    Next i
                                                End If
                                            End If
                                        End If
                                    End If

                                ElseIf TypeOf acEnt Is Topomatic.Alg.Road.Urb.Islands.MinorDropShapedIsland Then
                                    Dim ostrov As Topomatic.Alg.Road.Urb.Islands.MinorDropShapedIsland = acEnt
                                    Dim a = 1
                                End If
                                If acPoly3d.Count > 0 Then
                                    'рисуем полилинию
                                    Dim ArrayCoord As Double(,) = Nothing
                                    Dim countarrayPos As Integer = 0
                                    If acPoly3d.Count > 1 Then
                                        For i As Integer = 0 To acPoly3d.Count - 1
                                            Dim pos As BugleVector3D = acPoly3d.Item(i)
                                            ReDim Preserve ArrayCoord(3, countarrayPos)
                                            ArrayCoord(0, countarrayPos) = pos.Vertex.X
                                            ArrayCoord(1, countarrayPos) = pos.Vertex.Y
                                            ArrayCoord(2, countarrayPos) = pos.Bugle
                                            countarrayPos += 1
                                        Next
                                    End If
                                    If ArrayCoord.GetUpperBound(1) > 0 Then
                                        Dim TempPoly As DwgPolyline = RoburFunc.FuncDrawPolylineToArrayCoord(ActivDocument, ArrayCoord)
                                        If IsNothing(TempPoly) = False Then
                                            If IsNothing(objectLayer) = False Then
                                                TempPoly.Layer = objectLayer
                                            End If
                                            If IsNothing(objectUserTypeLine) = False Then
                                                TempPoly.Linetype = objectUserTypeLine
                                            End If
                                            TempPoly.Color = objectColor
                                            TempPoly.Lineweight = widthUserTypeLine
                                            TempPoly.LinetypeScale = scaleUserTypeLine * MDraw
                                        End If
                                    End If
                                End If
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                '3д объекты
                            ElseIf TypeOf acEnt Is Topomatic.Visualization.Runtime.DwgModel3DElement Then
                                Dim ac3dObject As DwgModel3DElement = acEnt
                                Dim mat As Matrix = ac3dObject.Matrix
                                Dim angl As Double = ac3dObject.Angle
                                Dim userElement As ImElement = ac3dObject.Element 'доступ к элементу

                                Dim userPropertiesObject As ImProperties = userElement.GetAllProperties()
                                Dim userImProperties As Topomatic.Visualization.ImProperty = Nothing
                                Dim arrayCoord As Double(,) = Nothing
                                Dim countArrayCoord As Integer = 0
                                If userPropertiesObject.Count > 0 Then
                                    For i As Integer = 0 To userPropertiesObject.Count - 1
                                        userImProperties = userPropertiesObject.Item(i)
                                        If userImProperties.Tag Like "axe" Then
                                            Exit For
                                        Else
                                            userImProperties = Nothing
                                        End If
                                    Next
                                End If
                                If IsNothing(userPropertiesObject) = False Then
                                    Dim valueImAgg As Topomatic.Visualization.ImAggregates = userImProperties.Value
                                    Dim objProperties As Object() = valueImAgg.Properties.Item(0).Value
                                    If IsArray(objProperties) = True Then
                                        For j As Integer = 0 To objProperties.Length - 1
                                            Dim obj2 As ImAggregates = objProperties.ElementAt(j)
                                            Dim obj3 As Double = Val(obj2.Properties.Item(0).Value) * -1
                                            Dim obj4 As Double = Val(obj2.Properties.Item(1).Value)
                                            Dim obj5 As Double = Val(obj2.Properties.Item(2).Value)
                                            '==================================================================================
                                            'перемножение матриц
                                            Dim a1 As Double = obj3 * mat.M11 + obj4 * mat.M12
                                            Dim a2 As Double = obj3 * mat.M21 + obj4 * mat.M22
                                            ReDim Preserve arrayCoord(2, countArrayCoord)
                                            arrayCoord(0, countArrayCoord) = a1 * -1 + ac3dObject.Position.Item(0)
                                            arrayCoord(1, countArrayCoord) = a2 + ac3dObject.Position.Item(1)
                                            arrayCoord(2, countArrayCoord) = 0
                                            countArrayCoord += 1
                                        Next
                                    End If
                                    Dim TempPoly As DwgPolyline = RoburFunc.FuncDrawPolylineToArrayCoord(ActivDocument, arrayCoord)
                                    If IsNothing(TempPoly) = False Then
                                        If IsNothing(objectLayer) = False Then
                                            TempPoly.Layer = objectLayer
                                        End If
                                        If IsNothing(objectUserTypeLine) = False Then
                                            TempPoly.Linetype = objectUserTypeLine
                                        End If
                                        TempPoly.Color = objectColor
                                        TempPoly.Lineweight = widthUserTypeLine
                                        TempPoly.LinetypeScale = scaleUserTypeLine * MDraw
                                    End If
                                End If
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                'дорожная разметка
                            ElseIf TypeOf acEnt Is Topomatic.RoadMarking.Entities.DwgLinearSmdxRoadMarking Then
                                Dim userRoadMark As Topomatic.RoadMarking.Entities.DwgLinearSmdxRoadMarking = acEnt
                                Dim imUserElement As ImElement = userRoadMark.Element
                                Dim userProperties As ImProperties = imUserElement.GetAllProperties()
                                Dim bColor As CadColor = userRoadMark.BufferColors(0)
                                'If prop.Count > 0 Then
                                '    For i As Integer = 0 To prop.Count - 1
                                '        Dim im = prop.Item(i)
                                '        Dim name = im.Name
                                '        Dim val = im.Value
                                '        Str = Str() & vbLf & "name: " & val
                                '    Next
                                'End If
                                'MsgBox(str)
                                Dim acPoly3d As Polyline3D = New Polyline3D
                                userRoadMark.GetPolyline(acPoly3d)
                                'рисуем полилинию
                                Dim ArrayCoord As Double(,) = Nothing
                                Dim countarrayPos As Integer = 0
                                If acPoly3d.Count > 1 Then
                                    For i As Integer = 0 To acPoly3d.Count - 1
                                        Dim pos As BugleVector3D = acPoly3d.Item(i)
                                        ReDim Preserve ArrayCoord(3, countarrayPos)
                                        ArrayCoord(0, countarrayPos) = pos.Vertex.X
                                        ArrayCoord(1, countarrayPos) = pos.Vertex.Y *
                                        ArrayCoord(2, countarrayPos) = pos.Bugle
                                        countarrayPos += 1
                                    Next
                                End If
                                If ArrayCoord.GetUpperBound(1) > 0 Then
                                    Dim TempPoly As DwgPolyline = RoburFunc.FuncDrawPolylineToArrayCoord(ActivDocument, ArrayCoord)
                                    If IsNothing(TempPoly) = False Then
                                        If IsNothing(objectLayer) = False Then
                                            TempPoly.Layer = objectLayer
                                        End If
                                        If IsNothing(objectUserTypeLine) = False Then
                                            TempPoly.Linetype = objectUserTypeLine
                                        End If
                                        TempPoly.Color = bColor
                                        TempPoly.Lineweight = widthUserTypeLine
                                        TempPoly.LinetypeScale = scaleUserTypeLine * MDraw
                                    End If
                                End If
                            ElseIf TypeOf acEnt Is DwgPolyline Then
                                Dim acPoly As DwgPolyline = acEnt
                                If IsNothing(ActivDocument) = True Then Return

                                Dim actSfcLayer = SurfaceLayer.GetSurfaceLayer(userCadView)
                                Dim userSurf As Surface = FuncSurface.FuncReturnSurfaceToProjectByName(ActivDocument, nameModel)
                                Dim numberCode As Integer = 0
                                Dim offsetElev As Double = 0
                                For j As Integer = 0 To arraytag.GetUpperBound(1)
                                    Dim nameCode As String = arraytag(0, j)
                                    If nameCode Like "Object_code" Then
                                        Dim codeStructLine As Integer = arraytag(1, j)
                                        If IsNumeric(codeStructLine) = True Then
                                            numberCode = Val(codeStructLine)
                                            Exit For
                                        End If
                                    End If
                                Next
                                Select Case numberCode
                                    Case 1014
                                        offsetElev = 10
                                    Case 1016
                                        offsetElev = -0.9
                                    Case 1017
                                        offsetElev = -2
                                    Case 1021
                                        offsetElev = 8
                                    Case 1048
                                        offsetElev = -3
                                    Case 1049
                                        offsetElev = -1.5
                                    Case 1050
                                        offsetElev = -2
                                    Case 1052
                                        offsetElev = -0.9
                                    Case 1051
                                        offsetElev = -2
                                End Select
                                If nameCategoryTable Like "Цифровая модель рельефа" Then
                                    Dim structLine As StructureLine = RoburFunc.FuncStructuresLineByPolyLine(userSurf, acPoly, numberCode, offsetElev)
                                Else
                                    If IsNothing(objectLayer) = False Then
                                        acPoly.Layer = objectLayer
                                    End If
                                    If IsNothing(objectUserTypeLine) = False Then
                                        acPoly.Linetype = objectUserTypeLine
                                    End If
                                    acPoly.Color = objectColor
                                    acPoly.Lineweight = widthUserTypeLine
                                    acPoly.LinetypeScale = scaleUserTypeLine * MDraw

                                    Dim countData As Integer = 0
                                    If dictionaryDataobject.Count > 0 Then
                                        For i As Integer = 0 To dictionaryDataobject.Count - 1
                                            Dim nameField As KeyValuePair(Of String, String(,)) = dictionaryDataobject.ElementAt(i)
                                            If nameField.Key Like "LOCALID" Then
                                                ReDim Preserve arrayDSDataRecord(1, countData)
                                                arrayDSDataRecord(0, countData) = "LOCALID"
                                                arrayDSDataRecord(1, countData) = rng.Next
                                                countData += 1
                                            ElseIf nameField.Key Like "ConditionalNumber" Then
                                                ReDim Preserve arrayDSDataRecord(1, countData)
                                                arrayDSDataRecord(0, countData) = "ConditionalNumber"
                                                arrayDSDataRecord(1, countData) = countParcels
                                                countData += 1
                                                countParcels += 1
                                            ElseIf nameField.Key Like "Area" Then
                                                ReDim Preserve arrayDSDataRecord(1, countData)
                                                arrayDSDataRecord(0, countData) = "Area"
                                                Dim sqrtParcels As Double = MathFunction.GetArea(acPoly)
                                                sqrtParcels = Math.Abs(sqrtParcels)
                                                If sqrtParcels >= 1 Then
                                                    arrayDSDataRecord(1, countData) = Math.Round(sqrtParcels, 0)
                                                Else
                                                    arrayDSDataRecord(1, countData) = Math.Round(sqrtParcels, 2)
                                                End If
                                                countData += 1
                                            ElseIf nameField.Key Like "Lenght" Then
                                                ReDim Preserve arrayDSDataRecord(1, countData)
                                                arrayDSDataRecord(0, countData) = "Lenght"
                                                Dim lenghtParcels As Double = Math.Round(acPoly.Length, 2)
                                                arrayDSDataRecord(1, countData) = lenghtParcels
                                                countData += 1
                                            End If
                                        Next
                                    End If
                                    If IsArray(arrayDSDataRecord) = True Then
                                        Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXData(acPoly, PSTable, arrayDSDataRecord, templateXML, dictionaryDataobject)
                                    End If
                                End If
                            ElseIf TypeOf acEnt Is DwgPolyline3D Then
                                Dim acPoly3d As DwgPolyline3D = acEnt
                                'рисуем полилинию
                                Dim ArrayCoord As Double(,) = Nothing
                                Dim countarrayPos As Integer = 0
                                If acPoly3d.Count > 1 Then
                                    For i As Integer = 0 To acPoly3d.Count - 1
                                        Dim pos As Topomatic.Cad.Foundation.Vector3D = acPoly3d.Item(i)
                                        ReDim Preserve ArrayCoord(3, countarrayPos)
                                        ArrayCoord(0, countarrayPos) = pos.X
                                        ArrayCoord(1, countarrayPos) = pos.Y
                                        ArrayCoord(2, countarrayPos) = 0
                                        countarrayPos += 1
                                    Next
                                End If
                                If ArrayCoord.GetUpperBound(1) > 0 Then
                                    Dim TempPoly As DwgPolyline = RoburFunc.FuncDrawPolylineToArrayCoord(ActivDocument, ArrayCoord)
                                    If IsNothing(TempPoly) = False Then
                                        If IsNothing(objectLayer) = False Then
                                            TempPoly.Layer = objectLayer
                                        End If
                                        If IsNothing(objectUserTypeLine) = False Then
                                            TempPoly.Linetype = objectUserTypeLine
                                        End If
                                        TempPoly.Color = objectColor
                                        TempPoly.Lineweight = widthUserTypeLine
                                        TempPoly.LinetypeScale = scaleUserTypeLine * MDraw
                                    End If
                                End If
                            End If
                        Next
                    Loop
                End If
            End If
        End If
    End Sub
    Private Sub ComboBox1_MouseClick(sender As Object, e As MouseEventArgs) Handles ComboBox1.MouseClick
        If IsNothing(userCadView) = False Then
            Dim arrayDrawLayer As String() = Nothing
            Dim DrawShowLayers As IEnumerable(Of DrawingLayer) = DrawingLayer.GetDrawingLayers(userCadView)
            If DrawShowLayers.Count > 0 Then
                For k1 As Integer = 0 To DrawShowLayers.Count - 1
                    Dim DrawLayerTemp As DrawingLayer = DrawShowLayers.ElementAt(k1)
                    Dim nameUserFile As String = DrawLayerTemp.Drawing.Filename
                    nameUserFile = IO.Path.GetFileNameWithoutExtension(nameUserFile)
                    ReDim Preserve arrayDrawLayer(k1)
                    arrayDrawLayer(k1) = nameUserFile
                Next k1
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            If IsArray(arrayDrawLayer) = True Then
                ComboBox1.DataSource = arrayDrawLayer
            End If
        End If
    End Sub

    Private Sub TreeView1_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeView1.AfterSelect

    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'изменнеие размера формы
    Private Sub PanelSymbolTOPO_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        Dim sizeForm As System.Drawing.Size = Me.Size
        Dim formWight As Double = sizeForm.Width
        Dim formHeight As Double = sizeForm.Height
        If formWight < 200 Then formWight = 200
        If formHeight < 200 Then formHeight = 200
        'размер combo
        Dim razmComb As Double = ComboBox1.Size.Width
        Dim widthButton As Double = formWight - razmComb - 32
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'рамка
        Dim sizeGroupBox As System.Drawing.Size = New System.Drawing.Size(formWight - 10, formHeight - 10)
        GroupBox1.Size = sizeGroupBox
        'листищч
        Dim sizeListBox As System.Drawing.Size = New System.Drawing.Size(formWight - 25, formHeight - 25)
        TreeView1.Size = sizeListBox
        'кнопка
        If widthButton > 5 Then
            Dim sizeButton As System.Drawing.Size = New System.Drawing.Size(widthButton, 60)
            Button1.Size = sizeButton
        End If        '
        Label2.Location = New System.Drawing.Point(formWight - 150, 15)
    End Sub

    'смена активной модели
    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged
        Dim nameSelectModel As String = Me.ComboBox1.Text
        Dim arrayNameModel As String() = Me.ComboBox1.DataSource
        Dim countArray As Integer = 0
        Dim oldCountArray As Integer = -1
        If IsArray(arrayNameModel) = True Then
            countArray = arrayNameModel.Length
            oldCountArray = arrayNameModel.Length
        End If
        Dim userProject As ModelProject = ApplicationHost.Current.ActiveProject
        Dim userSubObjectBearm As ArrangementModel = New ArrangementModel()
        If IsNothing(nameSelectModel) = False Then
            nameSelectModel = nameSelectModel.Trim
            Dim childs As IProjectModel() = userProject.Model.GetChilds()
            Try
                For Each child As IProjectModel In childs
                    Dim modelUri As URI = child.Uri
                    Dim fileNameSfcx As String = IO.Path.GetFileNameWithoutExtension(modelUri.LastPathComponent)
                    If fileNameSfcx.Trim Like nameSelectModel.Trim Then
                        ApplicationHost.Current.Plugins.Execute("activate", New Object() {child})
                        Exit Sub
                    End If
                Next
            Catch ex As System.Exception
            End Try
        End If
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        Button1.Text = "Нарисовать"
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        Button1.Text = "Преобразовать"
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub PanelSymbolTOPO_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'обновить список моделей
    Private Sub ОбновитьСписокМоделейToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОбновитьСписокМоделейToolStripMenuItem.Click
        Dim arrayModel As String() = Me.ComboBox1.DataSource
        Dim count As Integer = arrayModel.Length
        Dim userProject As ModelProject = ApplicationHost.Current.ActiveProject
        Dim userSubObjectBearm As ArrangementModel = New ArrangementModel()
        Dim childs As IProjectModel() = userProject.Model.GetChilds()
        Try
            For Each child As IProjectModel In childs
                Dim modelUri As URI = child.Uri
                Dim fileNameSfcx As String = IO.Path.GetFileNameWithoutExtension(modelUri.LastPathComponent)
                Dim boolFindModel As Boolean = MathFunction.FuncFindValueToFArray(fileNameSfcx.Trim, arrayModel)
                If boolFindModel = True Then
                    ReDim Preserve arrayModel(count)
                    arrayModel(count) = fileNameSfcx.Trim
                    count += 1
                End If
            Next
        Catch ex As System.Exception
        End Try
        If count > 0 Then
            Me.ComboBox1.DataSource = arrayModel
        End If
    End Sub
End Class
