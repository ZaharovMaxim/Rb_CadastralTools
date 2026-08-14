Imports System.IO
Imports Topomatic.Alg
Imports Topomatic.Alg.Offsets
Imports Topomatic.Cad.Foundation
Imports Topomatic.Cad.View
Imports Topomatic.Cad.View.Hints
Imports Topomatic.Dwg
Imports System.Text.RegularExpressions
Imports Topomatic.Dwg.Entities
Imports Topomatic.FoundationClasses
Imports Topomatic.FoundationClasses.Undo
Imports System.Runtime.InteropServices
Imports System.Drawing.Text
Imports System.Drawing
Imports Topomatic.Alg.Layers.BaseSectionLayer
Imports System.Windows

Public Class StyleCadastralObject
    Public Enum typeEntity
        Линия
        Полилиния
        Модель
    End Enum
    Private _nameStyle As String
    Private _indexColor As CadColor
    Private _layer As DwgLayer
    Private _lineType As DwgLinetype
    Private _lineTypeScale As Double
    Private _lineWeight As Double
    Private _lineGWidth As Double
    Private _drawing As Drawing
    Public Sub New(drawing As Drawing)
        If IsNothing(drawing) = False Then
            _nameStyle = "Standart"
            _indexColor = New CadColor(7)
            _layer = drawing.ActiveLayer
            _lineType = drawing.ActiveLinetype
            _lineTypeScale = 1
            _lineWeight = 0.2
            _lineGWidth = 0
            _drawing = drawing
        End If
    End Sub
    Public Sub New(drawing As Drawing, IndexColor As CadColor, Layer As DwgLayer, LineType As DwgLinetype, LineTypeScale As Double, LineWeight As Double, LineGWidth As Double, ByVal NameStyle As String)
        If IsNothing(drawing) = False Then
            _nameStyle = NameStyle
            _indexColor = IndexColor
            _layer = Layer
            _lineType = LineType
            _lineTypeScale = LineTypeScale
            _lineWeight = LineWeight
            _lineGWidth = LineGWidth
            _drawing = drawing
        End If
    End Sub
    Public Property NameStyle As String
        Get
            Return _nameStyle
        End Get
        Set(value As String)
            _nameStyle = value
        End Set
    End Property
    Public Property IndexColor As CadColor
        Get
            Return _indexColor
        End Get
        Set(value As CadColor)
            _indexColor = value
        End Set
    End Property

    'Получить или установить слой чертежа</summary>
    Public Property Layer As DwgLayer
        Get
            Return _layer
        End Get
        Set(value As DwgLayer)
            _layer = value
        End Set
    End Property

    'Получить или установить тип линии</summary>
    Public Property LineType As DwgLinetype
        Get
            Return _lineType
        End Get
        Set(value As DwgLinetype)
            _lineType = value
        End Set
    End Property

    'Получить или установить масштаб типа линии</summary>
    Public Property LineTypeScale As Double
        Get
            Return _lineTypeScale
        End Get
        Set(value As Double)
            _lineTypeScale = value
        End Set
    End Property

    'Получить или установить вес линии</summary>
    Public Property LineWeight As Double
        Get
            Return _lineWeight
        End Get
        Set(value As Double)
            _lineWeight = value
        End Set
    End Property

    'Получить или установить глобальную ширину линии</summary>
    Public Property LineGWidth As Double
        Get
            Return _lineGWidth
        End Get
        Set(value As Double)
            _lineGWidth = value
        End Set
    End Property

    'Получить или установить чертеж</summary>
    Public Property ActiveProjectDrawing As Drawing
        Get
            Return _drawing
        End Get
        Set(value As Drawing)
            _drawing = value
        End Set
    End Property
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'получаем шаблон оформления в формате xml
    Public Shared Function getTemplateXml() As Dictionary(Of String, String)
        Dim dictPutchXml As Dictionary(Of String, String) = New Dictionary(Of String, String)
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim arrayDirSupport As String() = Nothing
        Dim boolFindDirSupport As Boolean = FuncFiles.funcReadDirectoriesSupport(arrayDirSupport)
        Dim parentFolder As String = ""
        If IsArray(arrayDirSupport) = True Then
            For i As Integer = 0 To arrayDirSupport.Length - 1
                Dim nameFolder As String = arrayDirSupport(i)
                Dim pos As Integer = nameFolder.IndexOf("Civil3DToolsUtility")
                If pos > -1 Then
                    parentFolder = Mid(nameFolder, 1, pos)
                    parentFolder = parentFolder & "Civil3DToolsUtility"
                    Exit For
                End If
            Next
        End If
        'шаблон оформления
        Dim directorySupport As String = parentFolder & "\FileResources\Sample\"
        If Directory.Exists(directorySupport) = True Then
            Dim xmlFiles As String() = Directory.GetFiles(directorySupport, "*.xml")
            If IsArray(xmlFiles) = True Then
                For i As Integer = 0 To xmlFiles.Length - 1
                    Dim putchXml As String = xmlFiles(i)
                    Dim nameTemplate As String = IO.Path.GetFileNameWithoutExtension(putchXml)
                    If dictPutchXml.ContainsKey(nameTemplate) = False Then
                        dictPutchXml.Add(nameTemplate, putchXml)
                    End If
                Next
            End If
        End If
        Return dictPutchXml
    End Function
    'устанавливает стиль
    Public Sub setObjectStyle(ByVal putchTemplateXML As String, ByVal categoryTables As String, ByVal nameTable As String, ByVal typeEntity As typeEntity, ByVal nameObject As String)
        Dim arrayBeamProperties As String(,) = Nothing
        Dim boolFindParcelsStyle As Boolean = FuncXML.FuncReadDataObjectToXMLbyDesk(putchTemplateXML, categoryTables, nameTable, typeEntity.ToString, nameObject, arrayBeamProperties)
        If IsArray(arrayBeamProperties) = True Then
            Dim boolCreateStyle As Boolean = createStyleLinearObject(_drawing, putchTemplateXML, arrayBeamProperties, _layer, _lineType, _indexColor, _lineTypeScale, _lineWeight,,, _lineGWidth)
            NameStyle = nameObject
        End If
    End Sub
    Public Function setObjectStyle(ByRef entity As DwgEntity) As Boolean
        If IsNothing(entity) = False Then
            entity.Color = IndexColor
            entity.Layer = Layer
            entity.Linetype = LineType
            entity.LinetypeScale = LineTypeScale
            entity.Lineweight = LineWeight
            If TypeOf entity Is DwgPolyline And LineGWidth > 0 Then
                Dim pPoly As DwgPolyline = entity
                pPoly.Width = LineGWidth
            End If
        End If
        Return True
    End Function
    '===========================================================================================================================
    'фунция создает все необходимые стили для линейных объектов
    Public Function createStyleLinearObject(ByVal ActivDocument As Drawing, ByVal namePutchXML As String, ByVal ArrayRez As String(,), ByRef layerObj As DwgLayer, ByRef nameTypeLineObj As DwgLinetype, ByRef colorObj As CadColor, Optional ByRef ScaleLineObj As Double = 1, Optional ByRef WidthLineObj As Double = 20, Optional ByRef PSTable As String = "OBJECT", Optional ByRef note As String = "", Optional ByRef lineGWidth As Double = 0) As Boolean
        createStyleLinearObject = False
        '===================================================================================================================
        'оформляем откосы
        If IsArray(ArrayRez) = True Then
            Dim nameObject As String = ""
            Dim typeUserObject As String = ""
            Dim deskObject As String = ""
            Dim nameTypeLine As String = ""
            Dim lineTypeCode1 As String = ""
            Dim arrayMaskLine As String() = Nothing
            For i As Integer = 0 To ArrayRez.GetUpperBound(1)
                '=========================================================================================================
                'создаем слой
                If IsNothing(ArrayRez(0, i)) = True Then Continue For
                If ArrayRez(0, i).Trim Like "Layer" Then
                    Dim NameLayer As String = ArrayRez(1, i)
                    If IsNothing(NameLayer) = False Then
                        NameLayer = NameLayer.Trim
                        'ищем слой
                        Dim tempObjectLayer As DwgLayer = getLayerDwg(NameLayer)
                        'слой не найден, ищем его в шаблоне и создаем его
                        If tempObjectLayer.Name Like "0" Then
                            Dim ArrayRezLayer As String(,) = Nothing
                            If IO.File.Exists(namePutchXML) = True Then
                                Dim boolFindLayerXML As Boolean = FuncXML.FuncReadPropertiesLayerToXML(namePutchXML, NameLayer, ArrayRezLayer)
                            End If
                            Dim indexColor As Integer = 7
                            Dim nameLayerTypeLine As String = "Continuous"
                            Dim strDeskTypeLine As String = ""
                            Dim lineTypeCode As String = ""
                            Dim widthLineLayer As Integer = 20
                            Dim TypeLineLayer As DwgLinetype = Nothing
                            If IsArray(ArrayRezLayer) = True Then
                                If ArrayRezLayer.GetUpperBound(1) > 0 Then
                                    For k As Integer = 0 To ArrayRezLayer.GetUpperBound(1)
                                        If ArrayRezLayer(0, k).Trim Like "Color" Then
                                            indexColor = Val(ArrayRezLayer(1, k).Trim)
                                            If indexColor <= 0 Then
                                                indexColor = 7
                                            ElseIf indexColor >= 255 Then
                                                indexColor = 7
                                            End If
                                        ElseIf ArrayRezLayer(0, k).Trim Like "LineType" Then
                                            nameLayerTypeLine = ArrayRezLayer(1, k).Trim
                                            lineTypeCode = FuncXML.FuncFindLineTypeCodeByNameTypeLineToXML(namePutchXML, nameLayerTypeLine, strDeskTypeLine)
                                        ElseIf ArrayRezLayer(0, k).Trim Like "LineWidth" Then
                                            If ArrayRezLayer(1, k).Trim Like "ПоСлою" Then
                                                widthLineLayer = -1
                                            ElseIf ArrayRezLayer(1, k).Trim Like "ПоБлоку" Then
                                                widthLineLayer = -2
                                            ElseIf ArrayRezLayer(1, k).Trim Like "ПоУмолчанию" Then
                                                widthLineLayer = -3
                                            Else
                                                widthLineLayer = Val(ArrayRezLayer(1, k).Trim)
                                            End If
                                        End If
                                    Next k
                                    'штриовой тип линии
                                    Dim arrayMask As String() = lineTypeCode.Split(",")
                                    If IsArray(arrayMask) = True Then
                                        If arrayMask.Length > 2 Then
                                            TypeLineLayer = FuncCreateTypeLineByPattern(ActivDocument, nameLayerTypeLine, strDeskTypeLine, arrayMask, namePutchXML)
                                        End If
                                    End If
                                End If
                            End If
                            layerObj = RoburFunc.FuncAddLayer(ActivDocument, NameLayer, indexColor, nameLayerTypeLine, widthLineLayer)
                        Else
                            layerObj = tempObjectLayer
                        End If
                        'Имя объекта стиля
                    End If
                ElseIf ArrayRez(0, i).Trim Like "Name" Then
                    If IsNothing(ArrayRez(1, i)) = True Then
                        nameObject = "Continuous"
                    Else
                        If nameObject Like "ПоСлою" Then
                            nameObject = "ByLayer"
                        ElseIf nameObject Like "ПоБлоку" Then
                            nameObject = "ByBlock"
                        Else
                            nameObject = ArrayRez(1, i).Trim
                            lineTypeCode1 = FuncXML.FuncFindLineTypeCodeByNameTypeLineToXML(namePutchXML, nameObject, deskObject)
                        End If
                    End If

                ElseIf ArrayRez(0, i).Trim Like "Color" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        Dim strColor As String = ArrayRez(1, i).Trim
                        If strColor Like "ПоСлою" Then
                            colorObj = New CadColor(256)
                        ElseIf strColor Like "ПоБлоку" Then
                            colorObj = New CadColor(0)
                        Else
                            Dim ColorLine As Integer = Val(ArrayRez(1, i).Trim)
                            If ColorLine < 0 Then ColorLine = 0
                            If ColorLine > 255 Then ColorLine = 255
                            colorObj = New CadColor(ColorLine)
                        End If
                    End If
                ElseIf ArrayRez(0, i).Trim Like "LineWidth" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        If ArrayRez(1, i).Trim Like "ПоСлою" Then
                            WidthLineObj = -1
                        ElseIf ArrayRez(1, i).Trim Like "ПоБлоку" Then
                            WidthLineObj = -2
                        ElseIf ArrayRez(1, i).Trim Like "ByLineWeightDefault" Then
                            WidthLineObj = -3
                        Else
                            WidthLineObj = Val(ArrayRez(1, i).Trim)
                            WidthLineObj = Math.Abs(WidthLineObj)
                        End If
                    End If
                ElseIf ArrayRez(0, i).Trim Like "LineScale" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        ScaleLineObj = Val(ArrayRez(1, i).Trim)
                        If ScaleLineObj <= 0 Then
                            ScaleLineObj = 1
                        End If
                    End If
                ElseIf ArrayRez(0, i).Trim Like "Description" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        deskObject = ArrayRez(1, i).Trim
                    End If
                ElseIf ArrayRez(0, i).Trim Like "PS" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        PSTable = ArrayRez(1, i).Trim
                    End If
                ElseIf ArrayRez(0, i).Trim Like "LineGWidth" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        lineGWidth = Val(ArrayRez(1, i))
                        lineGWidth = Math.Abs(lineGWidth)
                    End If
                ElseIf ArrayRez(0, i).Trim Like "Note" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        note = ArrayRez(1, i).Trim
                    End If
                End If
            Next i
            '===========================================================================================================
            arrayMaskLine = lineTypeCode1.Split(",")
            If IsArray(arrayMaskLine) = True Then
                If arrayMaskLine.Length > 2 Then
                    nameTypeLineObj = FuncCreateTypeLineByPattern(ActivDocument, nameObject, deskObject, arrayMaskLine, namePutchXML)
                End If
            End If
            Return True
        End If
    End Function

    '===========================================================================================================================
    'функция создает все необходимые стили для штриховок
    Public Shared Function FuncCreateStyleHatchObject(ByVal ActivDocument As Drawing, ByVal namePutchXML As String, ByVal ArrayRez As String(,), ByRef namePattern As String, ByRef layerObj As DwgLayer, ByRef nameTypeLineObj As DwgLinetype, ByRef colorObj As CadColor, Optional ByRef ScaleLineObj As Double = 1, Optional ByRef WidthLineObj As Double = 20, Optional ByRef ScaleObject As Double = 1, Optional ByRef AngleObject As Double = 0, Optional ByRef PSTable As String = "OBJECT", Optional ByRef note As String = "") As Boolean
        FuncCreateStyleHatchObject = False
        Dim nameTypeLine As String = "Continuous"
        Dim deskObject As String = "Непрерывный"
        Dim arrayMaskLine As String() = Nothing
        If IsArray(ArrayRez) = False Then Return False
        For i As Integer = 0 To ArrayRez.GetUpperBound(1)
            'создаем слой
            If ArrayRez(0, i).Trim Like "Layer" Then
                If IsNothing(ArrayRez(1, i)) = False Then
                    Dim NameLayer As String = ArrayRez(1, i).Trim
                    Dim tempObjectLayer As DwgLayer = RoburFunc.FuncFindLayerDwg(ActivDocument, NameLayer)
                    'слой не найден, ищем его в шаблоне и создаем его
                    If IsNothing(tempObjectLayer) = True Then
                        Dim ArrayRezLayer As String(,) = Nothing
                        If IO.File.Exists(namePutchXML) = True Then
                            Dim boolFindLayerXML As Boolean = FuncXML.FuncReadPropertiesLayerToXML(namePutchXML, NameLayer, ArrayRezLayer)
                        End If
                        Dim indexColor As Integer = 7
                        Dim nameLayerTypeLine As String = "Continuous"
                        Dim strDeskTypeLine As String = "Непрерывный"
                        Dim lineTypeCode As String = ""
                        Dim widthLineLayer As Integer = 20
                        Dim TypeLineLayer As DwgLinetype = Nothing
                        If IsArray(ArrayRezLayer) = True Then
                            If ArrayRezLayer.GetUpperBound(1) > 0 Then
                                For k As Integer = 0 To ArrayRezLayer.GetUpperBound(1)
                                    If ArrayRezLayer(0, k).Trim Like "Color" Then
                                        indexColor = Val(ArrayRezLayer(1, k).Trim)
                                        If indexColor <= 0 Then
                                            indexColor = 7
                                        ElseIf indexColor >= 255 Then
                                            indexColor = 7
                                        End If
                                    ElseIf ArrayRezLayer(0, k).Trim Like "LineType" Then
                                        nameLayerTypeLine = ArrayRezLayer(1, k).Trim
                                    ElseIf ArrayRezLayer(0, k).Trim Like "DeskLineType" Then
                                        strDeskTypeLine = ArrayRezLayer(1, k).Trim
                                    ElseIf ArrayRezLayer(0, k).Trim Like "LineTypeCode" Then
                                        lineTypeCode = ArrayRezLayer(1, k).Trim
                                    ElseIf ArrayRezLayer(0, k).Trim Like "LineWidth" Then
                                        If ArrayRezLayer(1, k).Trim Like "ПоСлою" Then
                                            widthLineLayer = -1
                                        ElseIf ArrayRezLayer(1, k).Trim Like "ПоБлоку" Then
                                            widthLineLayer = -2
                                        ElseIf ArrayRezLayer(1, k).Trim Like "ПоУмолчанию" Then
                                            widthLineLayer = -3
                                        Else
                                            widthLineLayer = Val(ArrayRezLayer(1, k).Trim)
                                        End If
                                    End If
                                Next k
                                Dim arrayMask As String() = lineTypeCode.Split(",")
                                If IsArray(arrayMask) = True Then
                                    If arrayMask.Length > 2 Then
                                        TypeLineLayer = FuncCreateTypeLineByPattern(ActivDocument, nameLayerTypeLine, strDeskTypeLine, arrayMask, namePutchXML)
                                    End If
                                End If
                            End If
                        End If
                        layerObj = RoburFunc.FuncAddLayer(ActivDocument, NameLayer, indexColor, nameLayerTypeLine, widthLineLayer)
                    Else
                        layerObj = tempObjectLayer
                    End If
                End If
                'Имя объекта стиля
            ElseIf ArrayRez(0, i).Trim Like "Name" Then
                If IsNothing(ArrayRez(1, i)) = False Then
                    namePattern = ArrayRez(1, i).Trim
                End If
                'дескриптор объекта
            ElseIf ArrayRez(0, i).Trim Like "Description" Then
                If IsNothing(ArrayRez(1, i)) = False Then
                    deskObject = ArrayRez(1, i).Trim
                End If
                'тип линии
            ElseIf ArrayRez(0, i).Trim Like "LineType" Then
                If IsNothing(ArrayRez(1, i)) = False Then
                    nameTypeLine = ArrayRez(1, i).Trim
                    If nameTypeLine Like "ПоСлою" Then
                        nameTypeLine = "ByLayer"
                    ElseIf nameTypeLine Like "ПоБлоку" Then
                        nameTypeLine = "ByBlock"
                    End If
                End If
                'код типа линии
            ElseIf ArrayRez(0, i).Trim Like "LineTypeCode" Then
                If IsNothing(ArrayRez(1, i)) = False Then
                    arrayMaskLine = ArrayRez(1, i).Trim.Split(",")
                End If
                'цвет
            ElseIf ArrayRez(0, i).Trim Like "Color" Then
                If IsNothing(ArrayRez(1, i)) = False Then
                    Dim ColorLine As Integer = Val(ArrayRez(1, i).Trim)
                    If ColorLine < 0 Then ColorLine = 0
                    If ColorLine > 255 Then ColorLine = 255
                    colorObj = New CadColor(ColorLine)
                End If
            ElseIf ArrayRez(0, i).Trim Like "LineWidth" Then
                If IsNothing(ArrayRez(1, i)) = False Then
                    If ArrayRez(1, i).Trim Like "ПоСлою" Then
                        WidthLineObj = -1
                    ElseIf ArrayRez(1, i).Trim Like "ПоБлоку" Then
                        WidthLineObj = -2
                    ElseIf ArrayRez(1, i).Trim Like "ПоУмолчанию" Then
                        WidthLineObj = -3
                    Else
                        WidthLineObj = Val(ArrayRez(1, i).Trim)
                        WidthLineObj = Math.Abs(WidthLineObj)
                    End If
                End If
            ElseIf ArrayRez(0, i).Trim Like "LineScale" Then
                If IsNothing(ArrayRez(1, i)) = False Then
                    ScaleLineObj = Val(ArrayRez(1, i).Trim)
                    If ScaleLineObj <= 0 Then
                        ScaleLineObj = 1
                    End If
                End If
            ElseIf ArrayRez(0, i).Trim Like "Scale" Then
                If IsNothing(ArrayRez(1, i)) = False Then
                    ScaleObject = Val(ArrayRez(1, i).Trim)
                End If
            ElseIf ArrayRez(0, i).Trim Like "Angle" Then
                If IsNothing(ArrayRez(1, i)) = False Then
                    AngleObject = Val(ArrayRez(1, i).Trim)
                End If
            ElseIf ArrayRez(0, i).Trim Like "PS" Then
                If IsNothing(ArrayRez(1, i)) = False Then
                    PSTable = ArrayRez(1, i).Trim
                End If
            ElseIf ArrayRez(0, i).Trim Like "NOTE" Then
                If IsNothing(ArrayRez(1, i)) = False Then
                    note = ArrayRez(1, i).Trim
                End If
            End If
        Next i
        'создаем тип линии
        If IsArray(arrayMaskLine) = True Then
            If arrayMaskLine.Length > 2 Then
                nameTypeLineObj = FuncCreateTypeLineByPattern(ActivDocument, nameTypeLine, deskObject, arrayMaskLine, namePutchXML)
            End If
        End If
        Return True
    End Function

    '===========================================================================================================================
    'функция создает все необходимые стили для блоков
    Public Shared Function FuncCreateStylePointObject(ByVal ActivDocument As Drawing, ByVal namePutchXML As String, ByVal ArrayRez As String(,), ByRef nameObject As String, ByRef layerObj As DwgLayer, ByRef nameTypeLineObj As DwgLinetype, ByRef colorObj As CadColor, Optional ByRef ScaleLineObj As Double = 1, Optional ByRef WidthLineObj As Double = 20, Optional ByRef ScaleObject As Double = 1, Optional ByRef AngleObject As Double = 0, Optional ByRef attributeHeight As Double = 2.5, Optional ByRef PSTable As String = "OBJECT", Optional ByRef note As String = "") As Boolean
        FuncCreateStylePointObject = False
        Dim nameTypeLine As String = "Continuous"
        Dim deskObject As String = ""
        Dim arrayMaskLine As String() = Nothing
        If IsArray(ArrayRez) = False Then Return False
        For i As Integer = 0 To ArrayRez.GetUpperBound(1)
            'создаем слой
            Try
                If ArrayRez(0, i).Trim Like "Layer" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        Dim NameLayer As String = ArrayRez(1, i).Trim
                        Dim tempObjectLayer As DwgLayer = RoburFunc.FuncFindLayerDwg(ActivDocument, NameLayer)
                        'слой не найден, ищем его в шаблоне и создаем его
                        If IsNothing(tempObjectLayer) = True Then
                            Dim ArrayRezLayer As String(,) = Nothing
                            If IO.File.Exists(namePutchXML) = True Then
                                Dim boolFindLayerXML As Boolean = FuncXML.FuncReadPropertiesLayerToXML(namePutchXML, NameLayer, ArrayRezLayer)
                            End If
                            Dim indexColor As Integer = 7
                            Dim nameLayerTypeLine As String = "Continuous"
                            Dim strDeskTypeLine As String = "Непрерывный"
                            Dim lineTypeCode As String = ""
                            Dim widthLineLayer As Integer = 20
                            Dim TypeLineLayer As DwgLinetype = Nothing
                            If IsArray(ArrayRezLayer) = True Then
                                If ArrayRezLayer.GetUpperBound(1) > 0 Then
                                    For k As Integer = 0 To ArrayRezLayer.GetUpperBound(1)
                                        If ArrayRezLayer(0, k).Trim Like "Color" Then
                                            indexColor = Val(ArrayRezLayer(1, k).Trim)
                                            If indexColor <= 0 Then
                                                indexColor = 7
                                            ElseIf indexColor >= 255 Then
                                                indexColor = 7
                                            End If
                                        ElseIf ArrayRezLayer(0, k).Trim Like "LineType" Then
                                            nameLayerTypeLine = ArrayRezLayer(1, k).Trim
                                        ElseIf ArrayRezLayer(0, k).Trim Like "DeskLineType" Then
                                            strDeskTypeLine = ArrayRezLayer(1, k).Trim
                                        ElseIf ArrayRezLayer(0, k).Trim Like "LineTypeCode" Then
                                            lineTypeCode = ArrayRezLayer(1, k).Trim
                                        ElseIf ArrayRezLayer(0, k).Trim Like "LineWidth" Then
                                            If ArrayRezLayer(1, k).Trim Like "ПоСлою" Then
                                                widthLineLayer = -1
                                            ElseIf ArrayRezLayer(1, k).Trim Like "ПоБлоку" Then
                                                widthLineLayer = -2
                                            ElseIf ArrayRezLayer(1, k).Trim Like "ПоУмолчанию" Then
                                                widthLineLayer = -3
                                            Else
                                                widthLineLayer = Val(ArrayRezLayer(1, k).Trim)
                                            End If
                                        End If
                                    Next k
                                    Dim arrayMask As String() = lineTypeCode.Split(",")
                                    If IsArray(arrayMask) = True Then
                                        If arrayMask.Length > 2 Then
                                            TypeLineLayer = FuncCreateTypeLineByPattern(ActivDocument, nameLayerTypeLine, strDeskTypeLine, arrayMask, namePutchXML)
                                        End If
                                    End If
                                End If
                            End If
                            layerObj = RoburFunc.FuncAddLayer(ActivDocument, NameLayer, indexColor, nameLayerTypeLine, widthLineLayer)
                        Else
                            layerObj = tempObjectLayer
                        End If
                    End If
                    'Имя объекта стиля
                ElseIf ArrayRez(0, i).Trim Like "Name" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        nameObject = ArrayRez(1, i).Trim
                    End If
                    'дескриптор объекта
                ElseIf ArrayRez(0, i).Trim Like "Description" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        deskObject = ArrayRez(1, i).Trim
                    End If
                    'тип линии
                ElseIf ArrayRez(0, i).Trim Like "LineType" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        nameTypeLine = ArrayRez(1, i).Trim
                        If nameTypeLine Like "ПоСлою" Then
                            nameTypeLine = "ByLayer"
                        ElseIf nameTypeLine Like "ПоБлоку" Then
                            nameTypeLine = "ByBlock"
                        End If
                    End If
                    'код типа линии
                ElseIf ArrayRez(0, i).Trim Like "LineTypeCode" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        arrayMaskLine = ArrayRez(1, i).Trim.Split(",")
                    End If
                    'цвет
                ElseIf ArrayRez(0, i).Trim Like "Color" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        Dim ColorLine As Integer = Val(ArrayRez(1, i).Trim)
                        If ColorLine < 0 Then ColorLine = 0
                        If ColorLine > 255 Then ColorLine = 255
                        colorObj = New CadColor(ColorLine)
                    End If
                ElseIf ArrayRez(0, i).Trim Like "LineWidth" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        If ArrayRez(1, i).Trim Like "ПоСлою" Then
                            WidthLineObj = -1
                        ElseIf ArrayRez(1, i).Trim Like "ПоБлоку" Then
                            WidthLineObj = -2
                        ElseIf ArrayRez(1, i).Trim Like "ПоУмолчанию" Then
                            WidthLineObj = -3
                        Else
                            WidthLineObj = Val(ArrayRez(1, i).Trim)
                            WidthLineObj = Math.Abs(WidthLineObj)
                        End If
                    End If
                ElseIf ArrayRez(0, i).Trim Like "LineScale" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        ScaleLineObj = Val(ArrayRez(1, i).Trim)
                        If ScaleLineObj <= 0 Then
                            ScaleLineObj = 1
                        End If
                    End If
                ElseIf ArrayRez(0, i).Trim Like "Scale" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        ScaleObject = Val(ArrayRez(1, i).Trim)
                    End If
                ElseIf ArrayRez(0, i).Trim Like "Angle" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        AngleObject = Val(ArrayRez(1, i).Trim)
                    End If
                ElseIf ArrayRez(0, i).Trim Like "AttributeHeight" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        attributeHeight = Val(ArrayRez(1, i).Trim)
                    End If
                ElseIf ArrayRez(0, i).Trim Like "PS" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        PSTable = ArrayRez(1, i).Trim
                    End If
                ElseIf ArrayRez(0, i).Trim Like "NOTE" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        note = ArrayRez(1, i).Trim
                    End If
                End If
            Catch ex As System.Exception
            End Try
        Next i
        'создаем тип линии
        If IsArray(arrayMaskLine) = True Then
            If arrayMaskLine.Length > 2 Then
                nameTypeLineObj = FuncCreateTypeLineByPattern(ActivDocument, nameTypeLine, deskObject, arrayMaskLine, namePutchXML)
            End If
        End If
        Return True
    End Function

    '===========================================================================================================================
    'функция создает все необходимые стили для выносок
    Public Shared Function FuncCreateStyleLeaderObject(ByVal ActivDocument As Drawing, ByVal namePutchXML As String, ByVal ArrayRez As String(,), ByRef nameObject As String, ByRef layerObj As DwgLayer, ByRef nameTypeLineObj As DwgLinetype, ByRef colorObj As CadColor, Optional ByRef ScaleLineObj As Double = 1, Optional ByRef WidthLineObj As Double = 20, Optional ByRef textStyle As String = "Standard", Optional ByRef hText As Double = 2, Optional ByRef nameArrow As String = "Нет", Optional ByRef sizeArrow As Double = 2, Optional ByRef PSTable As String = "OBJECT", Optional ByRef note As String = "") As Boolean

        FuncCreateStyleLeaderObject = False
        Dim nameTypeLine As String = "Continuous"
        Dim deskObject As String = ""
        Dim arrayMaskLine As String() = Nothing
        If IsArray(ArrayRez) = False Then Return False
        For i As Integer = 0 To ArrayRez.GetUpperBound(1)
            'создаем слой
            Try
                If ArrayRez(0, i).Trim Like "Layer" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        Dim NameLayer As String = ArrayRez(1, i).Trim
                        Dim tempObjectLayer As DwgLayer = RoburFunc.FuncFindLayerDwg(ActivDocument, NameLayer)
                        'слой не найден, ищем его в шаблоне и создаем его
                        If IsNothing(tempObjectLayer) = True Then
                            Dim ArrayRezLayer As String(,) = Nothing
                            If IO.File.Exists(namePutchXML) = True Then
                                Dim boolFindLayerXML As Boolean = FuncXML.FuncReadPropertiesLayerToXML(namePutchXML, NameLayer, ArrayRezLayer)
                            End If
                            Dim indexColor As Integer = 7
                            Dim nameLayerTypeLine As String = "Continuous"
                            Dim strDeskTypeLine As String = "Непрерывный"
                            Dim lineTypeCode As String = ""
                            Dim widthLineLayer As Integer = 20
                            Dim TypeLineLayer As DwgLinetype = Nothing
                            If IsArray(ArrayRezLayer) = True Then
                                If ArrayRezLayer.GetUpperBound(1) > 0 Then
                                    For k As Integer = 0 To ArrayRezLayer.GetUpperBound(1)
                                        If ArrayRezLayer(0, k).Trim Like "Color" Then
                                            indexColor = Val(ArrayRezLayer(1, k).Trim)
                                            If indexColor <= 0 Then
                                                indexColor = 7
                                            ElseIf indexColor >= 255 Then
                                                indexColor = 7
                                            End If
                                        ElseIf ArrayRezLayer(0, k).Trim Like "LineType" Then
                                            nameLayerTypeLine = ArrayRezLayer(1, k).Trim
                                        ElseIf ArrayRezLayer(0, k).Trim Like "DeskLineType" Then
                                            strDeskTypeLine = ArrayRezLayer(1, k).Trim
                                        ElseIf ArrayRezLayer(0, k).Trim Like "LineTypeCode" Then
                                            lineTypeCode = ArrayRezLayer(1, k).Trim
                                        ElseIf ArrayRezLayer(0, k).Trim Like "LineWidth" Then
                                            If ArrayRezLayer(1, k).Trim Like "ПоСлою" Then
                                                widthLineLayer = -1
                                            ElseIf ArrayRezLayer(1, k).Trim Like "ПоБлоку" Then
                                                widthLineLayer = -2
                                            ElseIf ArrayRezLayer(1, k).Trim Like "ПоУмолчанию" Then
                                                widthLineLayer = -3
                                            Else
                                                widthLineLayer = Val(ArrayRezLayer(1, k).Trim)
                                            End If
                                        End If
                                    Next k
                                    Dim arrayMask As String() = lineTypeCode.Split(",")
                                    If IsArray(arrayMask) = True Then
                                        If arrayMask.Length > 2 Then
                                            TypeLineLayer = FuncCreateTypeLineByPattern(ActivDocument, nameLayerTypeLine, strDeskTypeLine, arrayMask, namePutchXML)
                                        End If
                                    End If
                                End If
                            End If
                            layerObj = RoburFunc.FuncAddLayer(ActivDocument, NameLayer, indexColor, nameLayerTypeLine, widthLineLayer)
                        Else
                            layerObj = tempObjectLayer
                        End If
                    End If
                ElseIf ArrayRez(0, i).Trim Like "Name" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        nameObject = ArrayRez(1, i).Trim
                    End If
                    'дескриптор объекта
                ElseIf ArrayRez(0, i).Trim Like "Description" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        deskObject = ArrayRez(1, i).Trim
                    End If
                    'тип линии
                ElseIf ArrayRez(0, i).Trim Like "LineType" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        nameTypeLine = ArrayRez(1, i).Trim
                        If nameTypeLine Like "ПоСлою" Then
                            nameTypeLine = "ByLayer"
                        ElseIf nameTypeLine Like "ПоБлоку" Then
                            nameTypeLine = "ByBlock"
                        End If
                    End If
                    'код типа линии
                ElseIf ArrayRez(0, i).Trim Like "LineTypeCode" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        arrayMaskLine = ArrayRez(1, i).Trim.Split(",")
                    End If
                    'цвет
                ElseIf ArrayRez(0, i).Trim Like "Color" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        Dim ColorLine As Integer = Val(ArrayRez(1, i).Trim)
                        If ColorLine < 0 Then ColorLine = 0
                        If ColorLine > 255 Then ColorLine = 255
                        colorObj = New CadColor(ColorLine)
                    End If
                ElseIf ArrayRez(0, i).Trim Like "LineWidth" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        If ArrayRez(1, i).Trim Like "ПоСлою" Then
                            WidthLineObj = -1
                        ElseIf ArrayRez(1, i).Trim Like "ПоБлоку" Then
                            WidthLineObj = -2
                        ElseIf ArrayRez(1, i).Trim Like "ПоУмолчанию" Then
                            WidthLineObj = -3
                        Else
                            WidthLineObj = Val(ArrayRez(1, i).Trim)
                            WidthLineObj = Math.Abs(WidthLineObj)
                        End If
                    End If
                ElseIf ArrayRez(0, i).Trim Like "LineScale" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        ScaleLineObj = Val(ArrayRez(1, i).Trim)
                        If ScaleLineObj <= 0 Then
                            ScaleLineObj = 1
                        End If
                    End If
                ElseIf ArrayRez(0, i).Trim Like "TextStyle" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        textStyle = ArrayRez(1, i).Trim
                    End If
                ElseIf ArrayRez(0, i).Trim Like "TextHeight" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        hText = Val(ArrayRez(1, i).Trim)
                    End If
                ElseIf ArrayRez(0, i).Trim Like "ArrowSize" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        sizeArrow = Val(ArrayRez(1, i).Trim)
                    End If
                ElseIf ArrayRez(0, i).Trim Like "PS" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        PSTable = ArrayRez(1, i).Trim
                    End If
                ElseIf ArrayRez(0, i).Trim Like "NOTE" Or ArrayRez(0, i).Trim Like "Note" Or ArrayRez(0, i).Trim Like "note" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        note = ArrayRez(1, i).Trim
                    End If
                End If
            Catch ex As System.Exception
            End Try
        Next i
        'создаем тип линии
        If IsArray(arrayMaskLine) = True Then
            If arrayMaskLine.Length > 2 Then
                nameTypeLineObj = FuncCreateTypeLineByPattern(ActivDocument, nameTypeLine, deskObject, arrayMaskLine, namePutchXML)
            End If
        End If
        If nameObject.Trim.Length > 0 Then
            Dim arrayPropertiesMLeader As String(,) = Nothing
            Dim boolStyleMleader As Boolean = FuncXML.FuncReadMLederStylePropertiesToXML(namePutchXML, nameObject, arrayPropertiesMLeader)
            If IsArray(arrayPropertiesMLeader) = True Then
                For i As Integer = 0 To arrayPropertiesMLeader.GetUpperBound(1)
                    Dim field As String = arrayPropertiesMLeader(0, i)
                    Dim valStr As String = arrayPropertiesMLeader(1, i)
                    If field Like "NameArrow" Then
                        nameArrow = valStr
                        Exit For
                    End If
                Next i
            End If
        End If
        Return True
    End Function

    '===========================================================================================================================
    'функция создает слой из XML
    Public Shared Function FuncCreateLayerByXML(ByVal ActivDocument As Drawing, ByVal namePutchXML As String, ByRef nameLayer As String) As DwgLayer
        FuncCreateLayerByXML = Nothing
        Dim layerObj As DwgLayer = Nothing
        'создаем слой
        Try
            Dim tempObjectLayer As DwgLayer = RoburFunc.FuncFindLayerDwg(ActivDocument, nameLayer)
            'слой не найден, ищем его в шаблоне и создаем его
            If IsNothing(tempObjectLayer) = True Then
                Dim ArrayRezLayer As String(,) = Nothing
                If IO.File.Exists(namePutchXML) = True Then
                    Dim boolFindLayerXML As Boolean = FuncXML.FuncReadPropertiesLayerToXML(namePutchXML, nameLayer, ArrayRezLayer)
                End If
                Dim indexColor As Integer = 7
                Dim nameLayerTypeLine As String = "Continuous"
                Dim strDeskTypeLine As String = "Непрерывный"
                Dim lineTypeCode As String = ""
                Dim widthLineLayer As Integer = 20
                Dim TypeLineLayer As DwgLinetype = Nothing
                If IsArray(ArrayRezLayer) = True Then
                    If ArrayRezLayer.GetUpperBound(1) > 0 Then
                        For k As Integer = 0 To ArrayRezLayer.GetUpperBound(1)
                            If ArrayRezLayer(0, k).Trim Like "Color" Then
                                indexColor = Val(ArrayRezLayer(1, k).Trim)
                                If indexColor <= 0 Then
                                    indexColor = 7
                                ElseIf indexColor >= 255 Then
                                    indexColor = 7
                                End If
                            ElseIf ArrayRezLayer(0, k).Trim Like "LineType" Then
                                nameLayerTypeLine = ArrayRezLayer(1, k).Trim
                            ElseIf ArrayRezLayer(0, k).Trim Like "DeskLineType" Then
                                strDeskTypeLine = ArrayRezLayer(1, k).Trim
                            ElseIf ArrayRezLayer(0, k).Trim Like "LineTypeCode" Then
                                lineTypeCode = ArrayRezLayer(1, k).Trim
                            ElseIf ArrayRezLayer(0, k).Trim Like "LineWidth" Then
                                If ArrayRezLayer(1, k).Trim Like "ПоСлою" Then
                                    widthLineLayer = -1
                                ElseIf ArrayRezLayer(1, k).Trim Like "ПоБлоку" Then
                                    widthLineLayer = -2
                                ElseIf ArrayRezLayer(1, k).Trim Like "ПоУмолчанию" Then
                                    widthLineLayer = -3
                                Else
                                    widthLineLayer = Val(ArrayRezLayer(1, k).Trim)
                                End If
                            End If
                        Next k
                        Dim arrayMask As String() = lineTypeCode.Split(",")
                        If IsArray(arrayMask) = True Then
                            If arrayMask.Length > 2 Then
                                TypeLineLayer = FuncCreateTypeLineByPattern(ActivDocument, nameLayerTypeLine, strDeskTypeLine, arrayMask, namePutchXML)
                            End If
                        End If
                    End If
                End If
                layerObj = RoburFunc.FuncAddLayer(ActivDocument, nameLayer, indexColor, nameLayerTypeLine, widthLineLayer)
                Return layerObj
            Else
                Return tempObjectLayer
            End If
        Catch ex As System.Exception
        End Try
    End Function
    '////////////////////////////////////////////////////////////////////////////////////////////////////
    'функция возвращает слой
    Public Function getLayerDwg(ByVal NameLayer As String) As DwgLayer
        Dim result As DwgLayer = ActiveProjectDrawing.Layers.Item("0")
        If IsNothing(NameLayer) = False Then
            If NameLayer.Trim.Length > 0 Then
                Try
                    Dim DwgLayers As DwgLayers = ActiveProjectDrawing.Layers
                    For Each dwglayer As DwgLayer In DwgLayers
                        If dwglayer.Name Like NameLayer.Trim Then
                            result = dwglayer
                        End If
                    Next
                Catch ex As System.Exception
                End Try
            End If
        End If
        Return result
    End Function
    '========================================================================================================================================
    'функция создает тип линии по маске
    Public Shared Function FuncCreateTypeLineByPattern(ByVal ActivDocument As Drawing, ByVal NameTypeLine As String, ByVal DeskTypeLine As String, ByRef ArrayMask() As String, Optional templateXML As String = "") As DwgLinetype
        FuncCreateTypeLineByPattern = Nothing
        Dim strMess As String = ""
        Try
            'делаем проверку на наличие типа линии
            For i As Integer = 0 To ActivDocument.Linetypes.Count - 1
                If ActivDocument.Linetypes.Item(i).Name Like NameTypeLine.Trim Then
                    FuncCreateTypeLineByPattern = ActivDocument.Linetypes.Item(i)
                    Exit Function
                End If
            Next i
            Dim lenghtFirst As Double = -1
            'читаем массив
            If IsArray(ArrayMask) = True Then
                If ArrayMask.Length > 1 Then
                    Dim mask As Topomatic.Cad.Foundation.LinetypePattern = New Topomatic.Cad.Foundation.LinetypePattern
                    For i As Integer = 0 To ArrayMask.Length - 1
                        Dim TempSumb As String = ArrayMask(i).Trim
                        If TempSumb.Length = 0 Then Continue For
                        'пропускаем первый символ
                        If TempSumb Like "A" Then
                            Continue For
                            'числовое значение записываем
                        ElseIf IsNumeric(TempSumb) = True Then
                            Dim Zn As Double = Val(TempSumb)
                            mask.Add.DashDotLength = Zn
                            If lenghtFirst = -1 And Zn > 0 Then
                                lenghtFirst = Zn
                            End If
                        Else
                            '==================================================================
                            'сложный тип линии
                            If Left(TempSumb, 1) Like "[[]" Then
                                Dim sumbolElementMask As LinetypePatternItem = mask.Add 'создаем новый элемент
                                Dim nameSHPForm As String = "" 'переменная нужна для храния имени формы
                                'перебираем элементы пока не дойдем дл обратной скобки
                                For j As Integer = i To ArrayMask.Length - 1
                                    Dim tempSumb2 As String = ArrayMask(j).Trim
                                    If TempSumb.Length = 0 Then Continue For
                                    'имя формы или текста
                                    If i = j Then
                                        tempSumb2 = TempSumb.Replace("[", "") 'избавляемся от первой кавычки
                                        If Left(tempSumb2, 1) Like """" Then 'значит это обычная надпись, не форма
                                            'тексты в кавычках (убираем кавычки слева и справа)
                                            sumbolElementMask.ElementType = Topomatic.Cad.Foundation.LinetypePatternElementType.etTextString
                                            tempSumb2 = tempSumb2.Replace("""", "")
                                            sumbolElementMask.TextString = tempSumb2
                                            strMess = strMess & vbLf & "Текст линии: " & tempSumb2
                                        Else
                                            'иначе это форма (если кавычек нет)
                                            sumbolElementMask.ElementType = Topomatic.Cad.Foundation.LinetypePatternElementType.etShape
                                            tempSumb2 = tempSumb2.Replace("""", "") 'на всякий случай, кавычек вообще быть не должно
                                            nameSHPForm = tempSumb2
                                            strMess = strMess & vbLf & "Имя формы: " & nameSHPForm
                                        End If
                                        'читаем имя шрифта (второй элемент в квадратных скобках)
                                    ElseIf j = i + 1 Then
                                        'это форма
                                        If Right(tempSumb2, 4) Like ".shx" Or Right(tempSumb2, 4) Like ".shp" Then
                                            sumbolElementMask.Filename = tempSumb2
                                            Dim findFile As String = tempSumb2.Replace("shx", "shp")
                                            'ищем файл формы во вспомогательных файлах поддержки
                                            Dim fullNameFileShp As String = FuncFiles.funcFindFileToDirectorySupport(findFile)
                                            If IO.File.Exists(fullNameFileShp) = True Then
                                                strMess = strMess & vbLf & "Файл формы (имя) " & findFile & ", найден. Полный путь: " & fullNameFileShp
                                                Dim Numbshp As Integer = -1
                                                Dim inputshp As StreamReader = New StreamReader(fullNameFileShp, True)
                                                Dim line2 As String
                                                Dim regMatch As String = "\*"
                                                Do Until inputshp.EndOfStream
                                                    line2 = inputshp.ReadLine().Trim 'считываем строку
                                                    If line2.Length > 0 Then
                                                        Dim flagread As String = Left(line2, 1)
                                                        If Regex.IsMatch(flagread, regMatch) Then 'chr(*)
                                                            Dim arrayShp As String() = line2.Split(",")
                                                            If IsArray(arrayShp) = True Then
                                                                If arrayShp.GetUpperBound(0) > 1 Then
                                                                    If arrayShp(2).Trim Like nameSHPForm Then
                                                                        Dim StrNumb As String = arrayShp(0).Trim
                                                                        If StrNumb.Length > 1 Then
                                                                            Numbshp = Val(Mid(StrNumb, 2))
                                                                            If Numbshp > 0 Then
                                                                                strMess = strMess & vbLf & "Номер формы: " & Numbshp
                                                                                Exit Do
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                Loop
                                                inputshp.Close()
                                                If Numbshp > 0 Then
                                                    sumbolElementMask.ShapeNumber = Numbshp
                                                Else
                                                    MsgBox("Не удалос найти форму в указанном файле shp!!! Макро превано")
                                                    Exit Function
                                                End If
                                            Else
                                                MsgBox("Не удалось найти файл описысывающий форму данной линии!!!. Макро прервано.")
                                                Exit Function
                                            End If
                                        Else
                                            'ищем такой текстовый стиль
                                            Dim userTextStyle As DwgStyle = FuncReturnTextStyle(ActivDocument, tempSumb2)
                                            If IsNothing(userTextStyle) = True Then
                                                'создаем новый текстовый стиль
                                                Dim ArrayProperties As String(,) = Nothing
                                                Dim boolFindTextStyle As Boolean = FuncXML.FuncReadTextStylePropertiesToXML(templateXML, tempSumb2, ArrayProperties)
                                                If IsArray(ArrayProperties) = True Then
                                                    userTextStyle = FuncCreateTextStyleByArray(ActivDocument, tempSumb2, ArrayProperties)
                                                End If
                                            End If
                                            If IsNothing(userTextStyle) Then
                                                userTextStyle = FuncCreateTextStyle(ActivDocument, tempSumb2, "Arial.ttf")
                                            End If
                                            If IsNothing(userTextStyle) Then
                                                userTextStyle = ActivDocument.ActiveStyle
                                                tempSumb2 = userTextStyle.Name
                                            End If
                                            sumbolElementMask.Style = tempSumb2
                                            strMess = strMess & vbLf & "Имя стиля текста: " & tempSumb2
                                        End If
                                    Else
                                        'ели элемент последний, то избавляемся от закрывающей последней скобки
                                        Dim flagEnd As Boolean = False
                                        If Right(tempSumb2, 1) Like "[]]" Then
                                            tempSumb2 = tempSumb2.Replace("]", "")
                                            flagEnd = True
                                        End If
                                        'разбиваем элементы по знаку равно
                                        Dim ArS As String() = tempSumb2.Split("=")
                                        Dim F As String = ""
                                        Dim v As Double = 0
                                        If IsArray(ArS) = True Then
                                            For k As Integer = 0 To ArS.Length - 1
                                                If ArS(k).Trim.Length > 0 Then
                                                    If IsNumeric(ArS(k).Trim) Then
                                                        v = Val(ArS(k).Trim)
                                                        v = Math.Round(v, 3)
                                                    Else
                                                        F = ArS(k).Trim
                                                    End If
                                                End If
                                            Next k
                                        End If
                                        If F Like "S" Or F Like "s" Then
                                            sumbolElementMask.Scale = v
                                            strMess = strMess & vbLf & "Масштаб формы: " & sumbolElementMask.Scale
                                        ElseIf F Like "R" Or F Like "r" Then
                                            sumbolElementMask.Rotation = v
                                        ElseIf F Like "U" Or F Like "u" Then
                                            sumbolElementMask.Rotation = v
                                        ElseIf F Like "X" Or F Like "x" Then
                                            sumbolElementMask.XOffset = v
                                        ElseIf F Like "Y" Or F Like "y" Then
                                            sumbolElementMask.YOffset = v
                                        End If
                                        'если элемент последний то выходим из цикла
                                        If flagEnd = True Then
                                            i = j
                                            Exit For
                                        End If
                                    End If
                                Next j
                            End If
                        End If
                    Next i
                    If mask.Item(mask.Count - 1).IsSimpleDashDot = True Then
                        Dim lenEndElement As Double = mask.Item(mask.Count - 1).DashDotLength
                        If lenEndElement < 0 Then
                            If lenghtFirst < 0.1 Then
                                lenghtFirst = lenEndElement / 2
                            ElseIf lenghtFirst < 10 Then
                                lenghtFirst = lenghtFirst / 10
                            Else
                                lenghtFirst = lenghtFirst / 20
                            End If
                            mask.Add.DashDotLength = lenghtFirst
                        End If
                    End If
                    Dim newUserLineType As DwgLinetype = ActivDocument.Linetypes.FindEqualityOrCreate(mask, NameTypeLine, DeskTypeLine)

                    Dim createPattern As LinetypePattern = newUserLineType.LinetypePattern
                    If createPattern.Count > 0 Then
                        Dim strpattern As String = "A,"
                        For i As Integer = 0 To createPattern.Count - 1
                            Dim el As LinetypePatternItem = createPattern.Item(i)
                            If el.ElementType = LinetypePatternElementType.etSimple Then
                                strpattern = strpattern & el.DashDotLength & ","
                            ElseIf el.ElementType = LinetypePatternElementType.etShape Then
                                strpattern = strpattern & "[" & el.ShapeNumber & "," & el.Filename & ",S=" & el.Scale & ",R=" & el.Rotation & ",X=" & el.XOffset & ",Y=" & el.YOffset & "]"
                            ElseIf el.ElementType = LinetypePatternElementType.etTextString Then
                                strpattern = strpattern & "[" & el.TextString & "," & el.Style & ",S=" & el.Scale & ",R=" & el.Rotation & ",X=" & el.XOffset & ",Y=" & el.YOffset & "]"
                            End If
                        Next i
                        strMess = strMess & vbLf & strpattern
                    End If
                    Return newUserLineType
                Else
                    'MsgBox("Массив описывающий выбранный тип линии имеет менее 3 элементов!!! Макро прервано!")
                    Exit Function
                End If
            Else
                'MsgBox("Массив описывающий выбранный тип линии пустой!!! Макро прервано!")
                Exit Function
            End If
        Catch ex As System.Exception
        End Try
    End Function

    '============================================================================================================================
    'функция возвращает имя формы по его коду
    Public Shared Function FuncReturnNameSHPByIndex(ByVal indexShape As Integer, ByVal nameFile As String) As String
        FuncReturnNameSHPByIndex = ""
        Dim putshFileSHP As String = FuncFiles.funcFindFileToDirectorySupport(nameFile)
        If IO.File.Exists(putshFileSHP) = True Then
            'читаем файл с формой
            Dim sr As StreamReader = New StreamReader(putshFileSHP)
            Dim regMatch As String = "\*"
            While sr.Peek <> -1
                Dim st As String = sr.ReadLine().Trim
                Dim flagread As String = Left(st, 1)
                If Regex.IsMatch(flagread, regMatch) Then 'chr(42)
                    Dim rez As String() = st.Split(",")
                    If rez.GetUpperBound(0) > 1 Then
                        Dim indForm As String = Mid(rez(0), 2)
                        If indForm Like indexShape Then
                            FuncReturnNameSHPByIndex = rez(2)
                            Exit While
                        End If
                    End If
                End If
            End While
            sr.Close()
        End If
    End Function

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция проверяет наличие определения блока в чертеже
    Public Shared Function containsBlock(ByVal ActivDocument As Drawing, ByVal nameBlock As String) As DwgBlock
        Dim result As DwgBlock = Nothing
        Try
            If IsNothing(nameBlock) = False And IsNothing(ActivDocument) = False Then
                If nameBlock.Trim.Length > 0 Then
                    Dim DwgBlock As DwgBlocks = ActivDocument.Blocks
                    If DwgBlock.Count > 1 Then
                        For Each dwgblk As DwgBlock In DwgBlock
                            If nameBlock.Trim Like dwgblk.Name Then
                                result = dwgblk
                            End If
                        Next
                    End If
                End If
            End If
        Catch ex As System.Exception
        End Try
        Return result
    End Function


    '=======================================================================================================================================
    'функция возвращает код типа линии по его стилю
    Public Shared Function FuncReturnPatternLineTypeCode(ByVal userLineType As DwgLinetype) As String
        FuncReturnPatternLineTypeCode = "A,"
        If IsNothing(userLineType) = False Then
            Dim linePattern As LinetypePattern = userLineType.LinetypePattern
            If linePattern.Count > 0 Then
                Try
                    For i As Integer = 0 To linePattern.Count - 1
                        Dim element As LinetypePatternItem = linePattern.Item(i)
                        Dim typeElement As LinetypePatternElementType = element.ElementType
                        If typeElement = 0 Then
                            FuncReturnPatternLineTypeCode = FuncReturnPatternLineTypeCode & Math.Round(element.DashDotLength, 3)
                        ElseIf typeElement = 2 Then
                            Dim valStr As String = element.TextString
                            Dim styleName As String = element.Style
                            Dim rot As Double = Math.Round(element.Rotation, 3)
                            Dim scale As Double = Math.Round(element.Scale, 3)
                            Dim dx As Double = Math.Round(element.XOffset, 3)
                            Dim dy As Double = Math.Round(element.YOffset, 3)
                            Dim htext As Double = Math.Round(element.Height, 3)
                            FuncReturnPatternLineTypeCode = FuncReturnPatternLineTypeCode & "[" & Chr(34) & valStr & Chr(34) & "," & styleName & ",S=" & scale & ",R=" & rot & ",X=" & dx & ",Y=" & dy & "]"
                        ElseIf typeElement = 4 Then
                            Dim fontName As String = element.Font.FileName
                            fontName = IO.Path.GetFileNameWithoutExtension(fontName)
                            fontName = fontName & ".shp"
                            Dim valStr As String = FuncReturnNameSHPByIndex(element.ShapeNumber, fontName)
                            Dim rot As Double = Math.Round(element.Rotation, 3)
                            Dim scale As Double = Math.Round(element.Scale, 3)
                            Dim dx As Double = Math.Round(element.XOffset, 3)
                            Dim dy As Double = Math.Round(element.YOffset, 3)
                            Dim htext As Double = Math.Round(element.Height, 3)
                            FuncReturnPatternLineTypeCode = FuncReturnPatternLineTypeCode & "[" & valStr & "," & fontName & ",S=" & scale & ",R=" & rot & ",X=" & dx & ",Y=" & dy & "]"
                        End If
                        If i < linePattern.Count - 1 Then
                            FuncReturnPatternLineTypeCode = FuncReturnPatternLineTypeCode & ","
                        End If
                    Next i
                Catch ex As System.Exception
                End Try
            End If
        End If
    End Function

    '=======================================================================================================================================
    'функция создает текстовый стиль
    Public Shared Function FuncCreateTextStyle(ByVal ActivDocument As Drawing, ByVal NameTextStyle As String, ByVal NameFonts As String, Optional HText As Double = 2.0, Optional rot As Double = 0) As DwgStyle
        FuncCreateTextStyle = Nothing
        Try
            For i As Integer = 0 To ActivDocument.Styles.Count - 1
                If ActivDocument.Styles.Item(i).Name Like NameTextStyle.Trim Then
                    FuncCreateTextStyle = ActivDocument.Styles.Item(i)
                    Exit Function
                End If
            Next i

            FuncCreateTextStyle = ActivDocument.Styles.Add(NameTextStyle.Trim, NameFonts)
            FuncCreateTextStyle.Height = HText
            FuncCreateTextStyle.Ratio = rot
        Catch ex As System.Exception
        End Try
    End Function
    '=======================================================================================================================================
    'функция создает текстовый стиль
    Public Shared Function FuncCreateTextStyleByArray(ByVal ActivDocument As Drawing, ByVal NameTextStyle As String, ByVal arrayProperties As String(,)) As DwgStyle
        FuncCreateTextStyleByArray = Nothing
        Try
            For i As Integer = 0 To ActivDocument.Styles.Count - 1
                If ActivDocument.Styles.Item(i).Name Like NameTextStyle.Trim Then
                    FuncCreateTextStyleByArray = ActivDocument.Styles.Item(i)
                    Exit Function
                End If
            Next i
            Dim font As String = ""
            Dim ital As Boolean = False
            Dim anno As Boolean = False
            Dim textheight As Double = 0
            Dim rot As Double = 0
            Dim textprior As Double = 0
            If IsArray(arrayProperties) = True Then
                For i As Integer = 0 To arrayProperties.GetUpperBound(1)
                    Dim field As String = arrayProperties(0, i)
                    Dim valText As String = arrayProperties(1, i)
                    If field Like "NameFont" Then
                        font = valText.Trim
                        Dim boolFont As Boolean = False
                        If Right(font, 4) Like ".ttf" Then
                            Dim userfontttf As CadFont = FontManager.Current.SearchTTF(font)
                            If IsNothing(userfontttf) = False Then
                                boolFont = True
                            Else
                                boolFont = False
                            End If
                            If boolFont = False Then
                                Dim col As InstalledFontCollection = New InstalledFontCollection()
                                Dim ffile As String = Path.GetFileNameWithoutExtension(font)
                                For Each fontt As FontFamily In col.Families
                                    If fontt.Name Like ffile Then
                                        boolFont = True
                                    End If
                                Next
                            End If
                        ElseIf Right(font, 4) Like ".shx" Or Right(font, 4) Like ".SHX" Then
                            Dim userfontttf As CadFont = FontManager.Current.SearchSHX(font)
                            If IsNothing(userfontttf) = True Then
                                boolFont = True
                            Else
                                boolFont = False
                            End If
                        End If
                        If boolFont = False Then
                            font = FontManager.Current.DefaultFont.FileName
                        End If
                    ElseIf field Like "Italic" Then
                        If valText Like "True" Then
                            ital = True
                        End If
                    ElseIf field Like "Annotative" Then
                        If valText Like "True" Then
                            anno = True
                        End If
                    ElseIf field Like "TextStyleHeight" Then
                        textheight = Val(valText)
                    ElseIf field Like "TextStyleRotate" Then
                        rot = Val(valText)
                    ElseIf field Like "TextStylePrior" Then
                        textprior = Val(valText)
                    End If
                Next i
            End If
            If font.Length > 0 Then
                FuncCreateTextStyleByArray = ActivDocument.Styles.Add(NameTextStyle.Trim, font)
                FuncCreateTextStyleByArray.Height = textheight
            End If
        Catch ex As System.Exception
        End Try
    End Function
    '=======================================================================================================================================
    'функция возвращает текстовый стиль
    Public Shared Function FuncReturnTextStyle(ByVal ActivDocument As Drawing, ByVal NameTextStyle As String) As DwgStyle
        FuncReturnTextStyle = Nothing
        Try
            For i As Integer = 0 To ActivDocument.Styles.Count - 1
                If ActivDocument.Styles.Item(i).Name Like NameTextStyle.Trim Then
                    FuncReturnTextStyle = ActivDocument.Styles.Item(i)
                    Exit Function
                End If
            Next i
        Catch ex As System.Exception
        End Try
    End Function

    '=======================================================================================================================================
    'функция возвращает все текстовые стили с чертежаь
    Public Shared Function FuncReturnTextStyles(ByVal ActivDocument As Drawing, ByRef arrayData As String()) As Boolean
        FuncReturnTextStyles = Nothing
        Dim countArray As Integer = 0
        If IsArray(arrayData) = True Then
            countArray = arrayData.Length
        End If
        Try
            For i As Integer = 0 To ActivDocument.Styles.Count - 1
                ReDim arrayData(countArray)
                arrayData(countArray) = ActivDocument.Styles.Item(i).Name
                countArray += 1
            Next i
        Catch ex As System.Exception
        End Try
    End Function


End Class
