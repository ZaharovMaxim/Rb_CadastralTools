Imports System.IO
Imports Topomatic.Alg.Offsets
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg
Imports System.Text.RegularExpressions
Imports Topomatic.Dwg.Entities
Imports System.Drawing.Text
Imports System.Drawing
Imports Topomatic.Sfc.Style
Imports Topomatic.FoundationClasses
Imports System.Windows.Media
Imports Drawing = Topomatic.Dwg.Drawing
Imports FontFamily = System.Drawing.FontFamily
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop.Word
Imports System.ComponentModel
Imports System.Globalization
Imports System.Windows.Media.Media3D
Imports Topomatic.FoundationClasses.Undo

Public Class StyleCadastralObject
    Private _nameStyle As String
    Private _indexColor As CadColor
    Private _layer As DwgLayer
    Private _lineType As DwgLinetype
    Private _lineTypeScale As Double
    Private _lineWeight As Double
    Private _drawing As Drawing
    Private _templateXml As String
    Private _templateDwg As String
    Public Sub New()
        _nameStyle = "Standart"
        _indexColor = New CadColor(7)
        _layer = Nothing
        _lineType = Nothing
        _lineTypeScale = 1
        _lineWeight = 0.2
        _drawing = Nothing
        _templateXml = ""
        _templateDwg = ""
    End Sub
    Public Sub New(drawing As Drawing)
        If IsNothing(drawing) = False Then
            _nameStyle = "Standart"
            _indexColor = New CadColor(7)
            _layer = drawing.ActiveLayer
            _lineType = drawing.ActiveLinetype
            _lineTypeScale = 1
            _lineWeight = 0.2
            _drawing = drawing
            _templateXml = ""
            _templateDwg = ""
        End If
    End Sub
    Public Sub New(drawing As Drawing, IndexColor As CadColor, Layer As DwgLayer, LineType As DwgLinetype, LineTypeScale As Double, LineWeight As Double, LineGWidth As Double, ByVal NameStyle As String, ByVal TemplateXml As String, ByVal TemplateDwg As String)
        If IsNothing(drawing) = False Then
            _nameStyle = NameStyle
            _indexColor = IndexColor
            _layer = Layer
            _lineType = LineType
            _lineTypeScale = LineTypeScale
            _lineWeight = LineWeight
            _drawing = drawing
            _templateXml = TemplateXml
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
    'Получить или установить чертеж</summary>
    Public Property ActiveProjectDrawing As Drawing
        Get
            Return _drawing
        End Get
        Set(value As Drawing)
            _drawing = value
        End Set
    End Property

    Public Property TemplateXML As String
        Get
            Return _templateXml
        End Get
        Set(value As String)
            _templateXml = value
        End Set
    End Property
    Public Property TemplateDwg As String
        Get
            Return _templateDwg
        End Get
        Set(value As String)
            _templateDwg = value
        End Set
    End Property

    Public Function setObjectStyle(ByRef entity As DwgEntity) As Boolean
        If IsNothing(entity) = False Then
            entity.Color = IndexColor
            entity.Layer = Layer
            entity.Linetype = LineType
            entity.LinetypeScale = LineTypeScale
            entity.Lineweight = LineWeight
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
    'работа со слоями
    '===========================================================================================================================
    '////////////////////////////////////////////////////////////////////////////////////////////////////
    'функция проверяет наличие слоя на чертеже
    Public Function HacLayerDwg(ByVal NameLayer As String) As Boolean
        Dim result As Boolean = False
        If IsNothing(NameLayer) = False Then
            If NameLayer.Trim.Length > 0 Then
                If IsNothing(ActiveProjectDrawing) = False Then
                    Try
                        Dim DwgLayers As DwgLayers = ActiveProjectDrawing.Layers
                        For Each dwglayer As DwgLayer In DwgLayers
                            If dwglayer.Name Like NameLayer.Trim Then
                                result = True
                                Exit For
                            End If
                        Next
                    Catch ex As System.Exception
                    End Try
                End If
            End If
        End If
        Return result
    End Function
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
                            Exit For
                        End If
                    Next
                Catch ex As System.Exception
                End Try
            End If
        End If
        Return result
    End Function
    'функция создает слой
    Public Function createDwgLayer(ByVal NameLayer As String, Optional ByVal Indcolor As Integer = 7, Optional ByVal NameTypeLine As String = "", Optional ValLineWight As Integer = 20, Optional ByVal boolVisible As Boolean = True) As DwgLayer
        Dim result As DwgLayer = Nothing
        If IsNothing(NameLayer) = True Then
            Return result
        End If
        If NameLayer.Trim.Length = 0 Then
            Return result
        End If

        result = getLayerDwg(NameLayer)
        If result.Name Like NameLayer Then
            Return result
        End If
        Try
            result = ActiveProjectDrawing.Layers.Add(NameLayer.Trim)
            'назначаем цвет
            Indcolor = Math.Abs(Indcolor)
            If Indcolor > 256 Then
                Indcolor = 256
            End If
            Dim layerColor As CadColor = New CadColor(Indcolor)
            result.Color = layerColor
            'назначаем тип линии слою
            If IsNothing(NameTypeLine) = False Then
                If NameTypeLine.Trim.Length > 0 Then

                End If
            End If
            'тип линии
            If NameTypeLine.Trim.Length > 0 Then
                Dim LineTypeDwg As DwgLinetype = getLineTypeDwg(NameTypeLine)
                result.Linetype = LineTypeDwg
            End If
            'назначаем вес линиям
            result.Lineweight = ValLineWight
            'назначаем видимость слою
            result.Visible = boolVisible
        Catch ex As System.ArgumentOutOfRangeException
        Catch ex As System.Exception
        End Try
        Return result
    End Function
    'функция создает слой
    Public Function createDwgLayerByProperties(ByRef arrayProperties As String(,)) As DwgLayer
        Dim result As DwgLayer = Nothing
        If IsArray(arrayProperties) = True Then
            Dim nameLayer As String = "0"
            Dim colorLayer As Integer = 7
            Dim lineTypeLayer As String = "Continuous"
            Dim lineWightLayer As Integer = 20
            For i As Integer = 0 To arrayProperties.GetUpperBound(1)
                Dim nameFields As String = arrayProperties(0, i)
                Dim valFields As String = arrayProperties(1, i)
                If IsNothing(nameFields) = True Then Continue For
                nameFields = nameFields.Trim
                If IsNothing(valFields) = True Then Continue For
                valFields = valFields.Trim
                If nameFields Like "Name" Then
                    nameLayer = valFields
                ElseIf nameFields Like "Color" Then
                    colorLayer = Val(valFields)
                    If colorLayer < 0 Then colorLayer = 0
                    If colorLayer > 255 Then colorLayer = 7
                ElseIf nameFields Like "LineType" Then
                    lineTypeLayer = valFields
                ElseIf nameFields Like "LineWidth" Then
                    lineWightLayer = Val(valFields)
                End If
            Next i
            If nameLayer.Trim.Length > 0 Then
                result = getLayerDwg(nameLayer)
                If result.Name Like nameLayer Then
                    Return result
                End If
                Try
                    result = ActiveProjectDrawing.Layers.Add(nameLayer.Trim)
                    'назначаем цвет
                    Dim layerColor As CadColor = New CadColor(colorLayer)
                    result.Color = layerColor
                    'назначаем тип линии слою
                    If lineTypeLayer.Trim.Length > 0 Then
                        Dim LineTypeDwg As DwgLinetype = getLineTypeDwg(lineTypeLayer)
                        result.Linetype = LineTypeDwg
                    End If
                    'назначаем вес линиям
                    result.Lineweight = lineWightLayer

                Catch ex As System.ArgumentOutOfRangeException
                Catch ex As System.Exception
                End Try
            End If
        End If
        Return result
    End Function



    'функция создает слой (старая версия)
    Public Shared Function createLayer(ByVal ActivDocument As Drawing, ByVal Name As String, Optional ByVal Indcolor As Integer = 7, Optional ByVal NameTypeLine As String = "", Optional ValLineWight As Integer = 20, Optional ByVal boolVisible As Boolean = True) As DwgLayer
        createLayer = Nothing
        Try
            For i As Integer = 0 To ActivDocument.Layers.Count - 1
                If ActivDocument.Layers.Item(i).Name Like Name.Trim Then
                    createLayer = ActivDocument.Layers.Item(i)
                    Exit For
                End If
            Next i

            createLayer = ActivDocument.Layers.Add(Name.Trim)
            'назначаем цвет
            Indcolor = Math.Abs(Indcolor)
            If Indcolor > 256 Then
                Indcolor = 256
            End If
            Dim CColor As CadColor = New CadColor(Indcolor)
            createLayer.Color = CColor
            'назначаем тип линии слою
            Dim LineTypes As DwgLinetypes = ActivDocument.Linetypes
            Dim LineType As DwgLinetype = Nothing
            For Each Ltype As DwgLinetype In LineTypes
                Dim NameLineType As String = Ltype.Name
                If NameLineType Like NameTypeLine Then
                    createLayer.Linetype = Ltype
                    Exit For
                End If
            Next
            'назначаем вес линиям
            createLayer.Lineweight = ValLineWight
            'назначаем видимость слою
            If boolVisible = True Then
                createLayer.Visible = True
            Else
                createLayer.Visible = False
            End If
            Return createLayer
        Catch ex As System.ArgumentOutOfRangeException
        End Try
    End Function
    '===========================================================================================================================
    'работа с типами линий
    '===========================================================================================================================
    '////////////////////////////////////////////////////////////////////////////////////////////////////
    'проверка наличия типа линии
    Public Function HacLineTypeDwg(ByVal NameLineType As String) As Boolean
        Dim result As Boolean = False
        If IsNothing(NameLineType) = False Then
            If NameLineType.Trim.Length > 0 Then
                If IsNothing(ActiveProjectDrawing) = False Then
                    Try
                        Dim DwgTypeLine As DwgLinetypes = ActiveProjectDrawing.Linetypes
                        For Each dwgLineType As DwgLinetype In DwgTypeLine
                            If dwgLineType.Name Like NameLineType.Trim Then
                                result = True
                                Exit For
                            End If
                        Next
                    Catch ex As System.Exception
                    End Try
                End If
            End If
        End If
        Return result
    End Function
    'возвращает тип линии
    Public Function getLineTypeDwg(ByVal NameLineType As String) As DwgLinetype
        Dim result As DwgLinetype = ActiveProjectDrawing.Linetypes.Continuous
        If IsNothing(NameLineType) = False Then
            If NameLineType.Trim.Length > 0 Then
                If IsNothing(ActiveProjectDrawing) = False Then
                    Try
                        Dim DwgTypeLine As DwgLinetypes = ActiveProjectDrawing.Linetypes
                        For Each dwgLineType As DwgLinetype In DwgTypeLine
                            If dwgLineType.Name Like NameLineType.Trim Then
                                result = dwgLineType
                                Exit For
                            End If
                        Next
                    Catch ex As System.Exception
                    End Try
                End If
            End If
        End If
        Return result
    End Function
    'функция создает тип линии по маске
    Public Function createLineTypeDwgByPattern(ByVal NameTypeLine As String, ByVal DeskTypeLine As String, ByRef ArrayMask() As String) As DwgLinetype
        Dim result As DwgLinetype = ActiveProjectDrawing.Linetypes.Continuous
        If IsNothing(NameTypeLine) = True Then
            Return result
        End If
        If NameTypeLine.Trim.Length = 0 Then
            Return result
        End If
        If IsNothing(ArrayMask) = True Then
            Return result
        End If
        If IsArray(ArrayMask) = False Then
            Return result
        End If
        If ArrayMask.Length < 2 Then
            Return result
        End If
        Dim strMess As String = ""
        Try
            'делаем проверку на наличие типа линии
            result = getLineTypeDwg(NameTypeLine)
            If result.Name Like NameTypeLine Then Return result
            Dim lenghtFirst As Double = -1
            'читаем массив
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
                                                MsgBox("Не удалось найти форму в указанном файле shp!!! Макро превано")
                                                Exit Function
                                            End If
                                        Else
                                            MsgBox("Не удалось найти файл описысывающий форму данной линии!!!. Макро прервано.")
                                            Exit Function
                                        End If
                                    Else
                                        'ищем такой текстовый стиль
                                        Dim userTextStyle As DwgStyle = FuncReturnTextStyle(ActiveProjectDrawing, tempSumb2)
                                        If IsNothing(userTextStyle) = True Then
                                            'создаем новый текстовый стиль
                                            Dim ArrayProperties As String(,) = Nothing
                                            Dim boolFindTextStyle As Boolean = CadastralTools.TemplateXML.FuncReadTextStylePropertiesInfoToXML(TemplateXML, tempSumb2, ArrayProperties)
                                            If IsArray(ArrayProperties) = True Then
                                                userTextStyle = FuncCreateTextStyleByArray(ActiveProjectDrawing, tempSumb2, ArrayProperties)
                                            End If
                                        End If
                                        If IsNothing(userTextStyle) Then
                                            userTextStyle = FuncCreateTextStyle(ActiveProjectDrawing, tempSumb2, "Arial.ttf")
                                        End If
                                        If IsNothing(userTextStyle) Then
                                            userTextStyle = ActiveProjectDrawing.ActiveStyle
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
                Dim newUserLineType As DwgLinetype = ActiveProjectDrawing.Linetypes.FindEqualityOrCreate(mask, NameTypeLine, DeskTypeLine)

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
        Catch ex As System.Exception
        End Try
    End Function
    'функция создает тип линии по маске (старая версия)
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
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'класс выноски
Public Class CadastralLeaderStyle
    Inherits StyleCadastralObject
    Private _textStyle As DwgStyle
    Private _heightText As Double
    Private _arrowStyle As AcDimArrowheadType
    Private _sizeArraow As Double
    Private _nameBlockArrow As String

    Public Sub New()
        MyBase.New()
        _textStyle = Nothing
        _heightText = 2
        _arrowStyle = AcDimArrowheadType.Default
        _sizeArraow = 2
        _nameBlockArrow = ""
    End Sub
    'стиль текста
    Public Property TextStyleLeader As DwgStyle
        Get
            Return _textStyle
        End Get
        Set(value As DwgStyle)
            _textStyle = value
        End Set
    End Property

    'высота текста
    Public Property HeightTextLeader As Double
        Get
            Return _heightText
        End Get
        Set(value As Double)
            _heightText = value
        End Set
    End Property

    'имя блока для стрелки
    Public Property ArrowStyle As AcDimArrowheadType
        Get
            Return _arrowStyle
        End Get
        Set(value As AcDimArrowheadType)
            _arrowStyle = value
        End Set
    End Property

    Public Property NameBlockArrow As String
        Get
            Return _nameBlockArrow
        End Get
        Set(value As String)
            _nameBlockArrow = value
        End Set
    End Property

    'размер стрелки
    Public Property SizeArraow As Double
        Get
            Return _sizeArraow
        End Get
        Set(value As Double)
            _sizeArraow = value
        End Set
    End Property

    '===========================================================================================================================
    'функция создает все необходимые стили для выносок
    Public Function createStyleLeader(ByVal arrayPropertiesLeader As String(,), Optional ByVal parentLayer As String = "") As Boolean
        createStyleLeader = False
        Dim nameTypeLine As String = "Continuous"
        Dim nameObject As String = ""
        Dim deskObject As String = ""
        Dim arrayMaskLine As String() = Nothing
        If IsArray(arrayPropertiesLeader) = False Then Return False
        For i As Integer = 0 To arrayPropertiesLeader.GetUpperBound(1)
            Dim nameField As String = arrayPropertiesLeader(0, i)
            Dim value As String = arrayPropertiesLeader(1, i)
            If IsNothing(nameField) = True Then Continue For
            If nameField.Trim.Length = 0 Then Continue For
            If IsNothing(value) = True Then Continue For
            If value.Trim.Length = 0 Then Continue For
            nameField = nameField.Trim
            value = value.Trim
            Try
                'создаем слой
                If nameField Like "Layer" Then
                    Dim layerLeader As DwgLayer = RoburFunc.FuncFindLayerDwg(ActiveProjectDrawing, value)
                    If parentLayer.Trim.Length > 0 Then
                        Dim tempLayer As DwgLayer = RoburFunc.FuncFindLayerDwg(ActiveProjectDrawing, parentLayer)
                        If IsNothing(tempLayer) = True Then
                            tempLayer = ActiveProjectDrawing.Layers.Layer0
                        End If
                        Layer = createLayer(ActiveProjectDrawing, parentLayer & value, tempLayer.Color.ColorIndex, tempLayer.Linetype.Name, tempLayer.Lineweight)
                    Else
                        'слой не найден, ищем его в шаблоне и создаем его
                        If IsNothing(layerLeader) = True Then
                            Dim arrayPropertiesLayer As String(,) = {}
                            If IO.File.Exists(TemplateXML) = True Then
                                Dim boolFindLayerXML As Boolean = CadastralTools.TemplateXML.FuncReadLayerPropertiesInfoToXML(TemplateXML, value, arrayPropertiesLayer)
                            End If
                            Dim indexColor As Integer = 7
                            Dim nameLayerTypeLine As String = "Continuous"
                            Dim strDeskTypeLine As String = "Непрерывный"
                            Dim lineTypeCode As String = ""
                            Dim widthLineLayer As Integer = 20
                            Dim TypeLineLayer As DwgLinetype = Nothing
                            If arrayPropertiesLayer.Length > 0 Then
                                For k As Integer = 0 To arrayPropertiesLayer.GetUpperBound(1)
                                    Dim nameLayerField As String = arrayPropertiesLayer(0, k)
                                    Dim valueLayer As String = arrayPropertiesLayer(1, k)
                                    If IsNothing(nameLayerField) = True Then Continue For
                                    If nameLayerField.Trim.Length = 0 Then Continue For
                                    If IsNothing(valueLayer) = True Then Continue For
                                    If valueLayer.Trim.Length = 0 Then Continue For
                                    nameLayerField = nameField.Trim
                                    valueLayer = valueLayer.Trim
                                    Select Case nameLayerField
                                        Case "Color"
                                            indexColor = Val(valueLayer)
                                            If indexColor <= 0 Then
                                                indexColor = 7
                                            ElseIf indexColor >= 255 Then
                                                indexColor = 7
                                            End If
                                        Case "LineType"
                                            Dim tempTypeLine As DwgLinetype = ActiveProjectDrawing.Linetypes.Item(value)
                                            If IsNothing(tempTypeLine) = True Then
                                                Dim arrayPropertiesTypeLine As String(,) = {}
                                                Dim boolReadPropertiesTypeLine As Boolean = CadastralTools.TemplateXML.FuncReadLineTypeInfoToXML(TemplateXML, value, arrayPropertiesTypeLine)
                                                If arrayPropertiesTypeLine.Length > 0 Then
                                                    strDeskTypeLine = CadastralTools.MathFunction.findValueToArray2d(arrayPropertiesTypeLine, "DeskLineType")
                                                    lineTypeCode = CadastralTools.MathFunction.findValueToArray2d(arrayPropertiesTypeLine, "CodeLineType")
                                                    Dim arrayMask As String() = lineTypeCode.Split(",")
                                                    If IsArray(arrayMask) = True Then
                                                        If arrayMask.Length > 2 Then
                                                            TypeLineLayer = FuncCreateTypeLineByPattern(ActiveProjectDrawing, nameLayerTypeLine, strDeskTypeLine, arrayMask, TemplateXML)
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                TypeLineLayer = ActiveProjectDrawing.Linetypes.Continuous
                                            End If
                                        Case "LineWidth"
                                            If value Like "ПоСлою" Then
                                                widthLineLayer = -1
                                            ElseIf value Like "ПоБлоку" Then
                                                widthLineLayer = -2
                                            ElseIf value Like "ПоУмолчанию" Then
                                                widthLineLayer = -3
                                            Else
                                                widthLineLayer = Val(value)
                                            End If
                                    End Select
                                Next k
                            End If
                            Layer = RoburFunc.FuncAddLayer(ActiveProjectDrawing, value, indexColor, nameLayerTypeLine, widthLineLayer)
                        Else
                            Layer = layerLeader
                        End If
                    End If
                ElseIf nameField Like "Name" Then
                    nameObject = value
                    'дескриптор объекта
                ElseIf nameField Like "Description" Then
                    deskObject = value
                    'тип линии
                ElseIf nameField Like "LineType" Then
                    If value Like "ПоСлою" Then
                        nameTypeLine = "ByLayer"
                    ElseIf nameTypeLine Like "ПоБлоку" Then
                        nameTypeLine = "ByBlock"
                    Else
                        Dim tempTypeLine As DwgLinetype = ActiveProjectDrawing.Linetypes.Item(value)
                        If IsNothing(tempTypeLine) = True Then
                            Dim arrayPropertiesTypeLine As String(,) = {}
                            Dim boolReadPropertiesTypeLine As Boolean = CadastralTools.TemplateXML.FuncReadLineTypeInfoToXML(TemplateXML, value, arrayPropertiesTypeLine)
                            If arrayPropertiesTypeLine.Length > 0 Then
                                Dim strDeskTypeLine = CadastralTools.MathFunction.findValueToArray2d(arrayPropertiesTypeLine, "DeskLineType")
                                Dim lineTypeCode = CadastralTools.MathFunction.findValueToArray2d(arrayPropertiesTypeLine, "CodeLineType")
                                Dim arrayMask As String() = lineTypeCode.Split(",")
                                If IsArray(arrayMask) = True Then
                                    If arrayMask.Length > 2 Then
                                        LineType = FuncCreateTypeLineByPattern(ActiveProjectDrawing, value, strDeskTypeLine, arrayMask, TemplateXML)
                                    End If
                                End If
                            End If
                        Else
                            LineType = ActiveProjectDrawing.Linetypes.Continuous
                        End If
                    End If
                    'цвет
                ElseIf nameField Like "Color" Then
                    Dim ColorLine As Integer = Val(value)
                    If ColorLine < 0 Then ColorLine = 0
                    If ColorLine > 255 Then ColorLine = 255
                    IndexColor = New CadColor(ColorLine)

                ElseIf nameField Like "LineWidth" Then
                    If nameField Like "ПоСлою" Then
                        LineWeight = -1
                    ElseIf nameField Like "ПоБлоку" Then
                        LineWeight = -2
                    ElseIf nameField Like "ПоУмолчанию" Then
                        LineWeight = -3
                    Else
                        LineWeight = Val(value)
                    End If
                ElseIf nameField Like "LineScale" Then
                    LineTypeScale = Val(value)
                    If LineTypeScale <= 0 Then
                        LineTypeScale = 1
                    End If
                ElseIf nameField Like "TextStyle" Then
                    Dim textCadastralStyle As DwgStyle = FuncReturnTextStyle(ActiveProjectDrawing, value)
                    If IsNothing(textCadastralStyle) = True Then
                        Dim arrayPropertiesTextStyle As String(,) = {}
                        Dim boolFindStyle As Boolean = CadastralTools.TemplateXML.FuncReadTextStylePropertiesInfoToXML(TemplateXML, value, arrayPropertiesTextStyle)
                        textCadastralStyle = FuncCreateTextStyleByArray(ActiveProjectDrawing, value, arrayPropertiesTextStyle)
                    End If
                    If IsNothing(textCadastralStyle) = False Then
                        TextStyleLeader = textCadastralStyle
                    End If
                ElseIf nameField Like "TextHeight" Then
                    HeightTextLeader = Val(value)
                ElseIf nameField Like "ArrowSize" Then
                    Dim nameArrow As String = value
                    If nameArrow Like "" Then
                        ArrowStyle = AcDimArrowheadType.None
                    ElseIf nameArrow Like "_ArchTick" Then
                        ArrowStyle = AcDimArrowheadType.ArchTick
                    ElseIf nameArrow Like "_BoxBlank" Then
                        ArrowStyle = AcDimArrowheadType.BoxBlank
                    ElseIf nameArrow Like "_BoxFilled" Then
                        ArrowStyle = AcDimArrowheadType.BoxFilled
                    ElseIf nameArrow Like "_Closed" Then
                        ArrowStyle = AcDimArrowheadType.Closed
                    ElseIf nameArrow Like "_ClosedBlank" Then
                        ArrowStyle = AcDimArrowheadType.ClosedBlank
                    ElseIf nameArrow Like "_DatumBlank" Then
                        ArrowStyle = AcDimArrowheadType.DatumBlank
                    ElseIf nameArrow Like "_DatumFilled" Then
                        ArrowStyle = AcDimArrowheadType.DatumFilled
                    ElseIf nameArrow Like "_Default" Then
                        ArrowStyle = AcDimArrowheadType.Default
                    ElseIf nameArrow Like "_Dot" Then
                        ArrowStyle = AcDimArrowheadType.Dot
                    ElseIf nameArrow Like "_DotBlank" Then
                        ArrowStyle = AcDimArrowheadType.DotBlank
                    ElseIf nameArrow Like "_DotSmall" Then
                        ArrowStyle = AcDimArrowheadType.DotSmall
                    ElseIf nameArrow Like "_Integral" Then
                        ArrowStyle = AcDimArrowheadType.Integral
                    ElseIf nameArrow Like "_None" Then
                        ArrowStyle = AcDimArrowheadType.None
                    ElseIf nameArrow Like "_Oblique" Then
                        ArrowStyle = AcDimArrowheadType.Oblique
                    ElseIf nameArrow Like "_Open" Then
                        ArrowStyle = AcDimArrowheadType.Open
                    ElseIf nameArrow Like "_Open30" Then
                        ArrowStyle = AcDimArrowheadType.Open30
                    ElseIf nameArrow Like "_Open90" Then
                        ArrowStyle = AcDimArrowheadType.Open90
                    ElseIf nameArrow Like "_Origin" Then
                        ArrowStyle = AcDimArrowheadType.Origin
                    ElseIf nameArrow Like "_Origin2" Then
                        ArrowStyle = AcDimArrowheadType.Origin2
                    ElseIf nameArrow Like "_Small" Then
                        ArrowStyle = AcDimArrowheadType.Small
                    Else
                        Dim userBlock As DwgBlock = ActiveProjectDrawing.Blocks.Item(nameArrow)
                        If IsNothing(userBlock) = True Then
                            userBlock = RoburFunc.FuncInsertBlockByDWGFile(ActiveProjectDrawing, TemplateXML, nameArrow)
                        End If
                        If IsNothing(userBlock) = False Then
                            ArrowStyle = AcDimArrowheadType.UserDefined
                        Else
                            ArrowStyle = AcDimArrowheadType.None
                        End If
                    End If
                End If
            Catch ex As System.Exception
            End Try
        Next i
        Return True
    End Function

    Public Function setLeaderStyle(ByRef entityLeader As DwgLeader) As Boolean
        If IsNothing(entityLeader) = False Then
            If IsNothing(IndexColor) = False Then
                entityLeader.Color = IndexColor
            End If
            If IsNothing(Layer) = False Then
                entityLeader.Layer = Layer
            End If
            If IsNothing(LineType) = False Then
                entityLeader.Linetype = LineType
            End If
            If LineTypeScale > 0 Then
                entityLeader.LinetypeScale = LineTypeScale
            End If
            If LineWeight > 0 Then
                entityLeader.Lineweight = LineTypeScale
            End If
            If SizeArraow > 0 Then
                entityLeader.ArrowheadSize = SizeArraow
            End If
            If NameBlockArrow.Trim.Length > 0 Then
                entityLeader.ArrowbodyName = NameBlockArrow
            End If
            If IsNothing(TextStyleLeader) = False Then
                entityLeader.Style = TextStyleLeader
            End If
        End If
        Return True
    End Function
End Class
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'класс полилиний
Public Class CadastralPolylineStyle
    Inherits StyleCadastralObject
    Private _lineGWidth As Double

    Public Sub New()
        MyBase.New()
        _lineGWidth = 0
    End Sub
    Public Sub New(parentPolyline As DwgPolyline)
        IndexColor = parentPolyline.Color
        Layer = parentPolyline.Layer
        LineType = parentPolyline.Linetype
        LineTypeScale = parentPolyline.LinetypeScale
        LineWeight = parentPolyline.Lineweight
        LineGWidth = parentPolyline.Width
    End Sub
    'Получить или установить глобальную ширину линии</summary>
    Public Property LineGWidth As Double
        Get
            Return _lineGWidth
        End Get
        Set(value As Double)
            _lineGWidth = value
        End Set
    End Property

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

    Public Function setPolylineStyle(ByRef entityPolyline As DwgPolyline) As Boolean
        If IsNothing(entityPolyline) = False Then
            If IsNothing(IndexColor) = False Then
                entityPolyline.Color = IndexColor
            End If
            If IsNothing(Layer) = False Then
                entityPolyline.Layer = Layer
            End If
            If IsNothing(LineType) = False Then
                entityPolyline.Linetype = LineType
            End If
            If LineTypeScale > 0 Then
                entityPolyline.LinetypeScale = LineTypeScale
            End If
            If LineWeight > 0 Then
                entityPolyline.Lineweight = LineWeight
            End If
            If LineGWidth > 0 Then
                entityPolyline.Width = LineGWidth
            End If
        End If
        Return True
    End Function
    Public Function setlineStyle(ByRef entity As DwgEntity) As Boolean
        If IsNothing(entity) = False Then
            If IsNothing(IndexColor) = False Then
                entity.Color = IndexColor
            End If
            If IsNothing(Layer) = False Then
                entity.Layer = Layer
            End If
            If IsNothing(LineType) = False Then
                entity.Linetype = LineType
            End If
            If LineTypeScale > 0 Then
                entity.LinetypeScale = LineTypeScale
            End If
            If LineWeight > 0 Then
                entity.Lineweight = LineWeight
            End If
        End If
        Return True
    End Function
    'устанавливает стиль для полилиний
    Public Sub setObjectStyle(ByVal putchTemplateXML As String, ByVal categoryTables As String, ByVal nameTable As String, ByVal typeEntity As CadastralTools.StructureElement.typeGeometryObject, ByVal nameObject As String)
        Dim arrayBeamProperties As String(,) = Nothing
        Dim boolFindParcelsStyle As Boolean = CadastralTools.TemplateXML.FuncReadPropertiesObjectToXML(putchTemplateXML, categoryTables, nameTable, typeEntity, nameObject, arrayBeamProperties)
        If IsArray(arrayBeamProperties) = True Then
            Dim boolCreateStyle As Boolean = createStyleLinearObject(ActiveProjectDrawing, putchTemplateXML, arrayBeamProperties, Layer, LineType, IndexColor, LineTypeScale, LineWeight,,, LineGWidth)
            NameStyle = nameObject
        End If
    End Sub
End Class
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'класс блоки
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Public Class CadastralBlockStyle
    Inherits StyleCadastralObject
    Private _name As String
    Private _scale As Double
    Private _rotation As Double
    Private _annotative As Boolean
    Private _atribyteProperties As List(Of AttributeCadastralBlock)
    Public Sub New()
        MyBase.New()
        _name = ""
        _scale = 1
        _rotation = 0
        _annotative = False
        _atribyteProperties = New List(Of AttributeCadastralBlock)
    End Sub
    Public Sub New(parentBlock As DwgInsert)
        NameBlock = parentBlock.Block.Name
        IndexColor = parentBlock.Color
        Layer = parentBlock.Layer
        LineType = parentBlock.Linetype
        LineTypeScale = parentBlock.LinetypeScale
        LineWeight = parentBlock.Lineweight
        ScaleBlock = parentBlock.Scale.X
        RotationBlock = parentBlock.Rotation
        AnnotativeBlock = parentBlock.Annotative
    End Sub
    'Получить или установить имя блока
    Public Property NameBlock As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property
    'Получить или установить масштаб
    Public Property ScaleBlock As Double
        Get
            Return _scale
        End Get
        Set(value As Double)
            _scale = value
        End Set
    End Property
    'Получить или установить поворот
    Public Property RotationBlock As Double
        Get
            Return _rotation
        End Get
        Set(value As Double)
            _rotation = value
        End Set
    End Property
    'Получить или установить аннотативность
    Public Property AnnotativeBlock As Boolean
        Get
            Return _annotative
        End Get
        Set(value As Boolean)
            _annotative = value
        End Set
    End Property

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция проверяет наличие определения блока в чертеже
    Public Function HacBlockDwg(ByVal nameBlock As String) As Boolean
        Dim result As Boolean = False
        Try
            If IsNothing(nameBlock) = False And IsNothing(ActiveProjectDrawing) = False Then
                If nameBlock.Trim.Length > 0 Then
                    Dim Dwgblks As DwgBlocks = ActiveProjectDrawing.Blocks
                    If Dwgblks.Count > 1 Then
                        For Each dwgblk As DwgBlock In Dwgblks
                            If nameBlock.Trim Like dwgblk.Name Then
                                result = True
                                Exit For
                            End If
                        Next
                    End If
                End If
            End If
        Catch ex As System.Exception
        End Try
        Return result
    End Function
    'функция возвращает определения блока в чертеже
    Public Function getBlockDwg(ByVal nameBlock As String) As DwgBlock
        Dim result As DwgBlock = Nothing
        Try
            If IsNothing(nameBlock) = False And IsNothing(ActiveProjectDrawing) = False Then
                If nameBlock.Trim.Length > 0 Then
                    Dim Dwgblks As DwgBlocks = ActiveProjectDrawing.Blocks
                    If Dwgblks.Count > 1 Then
                        For Each dwgblk As DwgBlock In Dwgblks
                            If nameBlock.Trim Like dwgblk.Name Then
                                result = dwgblk
                                Exit For
                            End If
                        Next
                    End If
                End If
            End If
        Catch ex As System.Exception
        End Try
        Return result
    End Function
    'вставка определения блока из внешнего файла
    Public Function CopyBlockReferenceByDestDraw(ByVal nameBlock As String) As DwgBlock
        Dim result As DwgBlock = Nothing
        If File.Exists(TemplateDwg) = False Then Exit Function
        If IsNothing(nameBlock) = True Then Exit Function
        If nameBlock.Trim.Length = 0 Then Exit Function
        If IsNothing(ActiveProjectDrawing) = True Then Exit Function
        Try
            If IsNothing(ActiveProjectDrawing.Blocks.Item(nameBlock)) = True Then
                Dim sourceDraw As Drawing = New Drawing()
                Topomatic.Acax.Import.Dxf.AcaxImporter.Import(sourceDraw, TemplateDwg)
                Dim acBlk As DwgBlock = sourceDraw.Blocks.Item(nameBlock)
                If IsNothing(acBlk) = False Then
                    ActiveProjectDrawing.BeginUpdate()
                    Try
                        Dim newBlk As DwgBlock = ActiveProjectDrawing.Blocks.Add(acBlk.Name)
                        newBlk.Entities.CopyFrom(acBlk, New ReferencesContext(sourceDraw))
                        For i As Integer = 0 To newBlk.Count - 1
                            Dim dwgEnt As DwgEntity = newBlk.ElementAt(i)
                            If TypeOf dwgEnt Is DwgAttdef Then
                                Dim userAttrDef As DwgAttdef = dwgEnt
                                If userAttrDef.Content = "####" Then
                                    'ActivDocument.ActiveSpace.Entities.Remove(userAttrDef)
                                End If
                            End If
                        Next i
                        ActiveProjectDrawing.EndUpdate()
                        Return newBlk
                    Catch ex As Exception
                    End Try
                End If
            Else
                Return ActiveProjectDrawing.Blocks.Item(nameBlock)
            End If
        Catch ex As System.Exception
        End Try
    End Function
    '===========================================================================================================================
    'фунция создает все необходимые стили для блоков
    Public Function createBlockStyle(ByRef arrayPropertyStyle As String(,)) As Boolean
        If IsNothing(ActiveProjectDrawing) = True Then Return False
        If IsNothing(arrayPropertyStyle) = True Then Return False
        If arrayPropertyStyle.Length = 0 Then Return False
        If IsNothing(TemplateXML) = True Then Return False
        If IO.File.Exists(TemplateXML) = False Then Return False
        Try
            If IsArray(arrayPropertyStyle) = True Then
                If arrayPropertyStyle.GetUpperBound(1) > 0 Then
                    For i As Integer = 0 To arrayPropertyStyle.GetUpperBound(1)
                        Dim nameField As String = arrayPropertyStyle(0, i)
                        If IsNothing(nameField) = True Then Continue For
                        If nameField.Trim.Length = 0 Then Continue For
                        Dim value As String = arrayPropertyStyle(1, i)
                        If IsNothing(value) = True Then Continue For
                        If value.Trim.Length = 0 Then Continue For
                        If nameField Like "Layer" Then
                            'создаем новый слой
                            Dim layerObject As String = value
                            Dim boolFindLayers As Boolean = HacLayerDwg(layerObject)
                            If boolFindLayers = False Then
                                'ищем настройки слоя
                                Dim userPropertiesLayer As String(,) = {}
                                Dim boolFindData As Boolean = CadastralTools.TemplateXML.FuncReadLayerPropertiesInfoToXML(TemplateXML, layerObject, userPropertiesLayer)
                                If IsArray(userPropertiesLayer) = True Then
                                    'назначаем тип линии для слоя
                                    Dim nameTypeLineLayer As String = CadastralTools.MathFunction.findValueToArray2d(userPropertiesLayer, "LineType", 0, 1)
                                    If IsNothing(nameTypeLineLayer) = True Then
                                        nameTypeLineLayer = "Continuous"
                                        Dim boolWriteData As Boolean = CadastralTools.MathFunction.setValueToArray2d(userPropertiesLayer, "DeskLineType", "Continuous", 0, 1)
                                    ElseIf nameTypeLineLayer.Trim.Length = 0 Then
                                        nameTypeLineLayer = "Continuous"
                                        Dim boolWriteData As Boolean = CadastralTools.MathFunction.setValueToArray2d(userPropertiesLayer, "DeskLineType", "Continuous", 0, 1)
                                    Else
                                        'загружаем тип линии
                                        Dim findTypeLine As Boolean = HacLineTypeDwg(nameTypeLineLayer)
                                        If findTypeLine = False Then
                                            Dim userPropertiesTypeLine As String(,) = {}
                                            boolFindData = CadastralTools.TemplateXML.FuncReadLineTypeInfoToXML(TemplateXML, nameTypeLineLayer, userPropertiesTypeLine)
                                            If IsArray(userPropertiesTypeLine) = True Then
                                                Dim deskTypeLineLayer As String = CadastralTools.MathFunction.findValueToArray2d(userPropertiesLayer, "DeskLineType", 0, 1)
                                                Dim patternTypeLineLayer As String = CadastralTools.MathFunction.findValueToArray2d(userPropertiesLayer, "CodeLineType", 0, 1)
                                                If IsNothing(deskTypeLineLayer) = True Then
                                                    deskTypeLineLayer = nameTypeLineLayer
                                                ElseIf deskTypeLineLayer.Trim.Length = 0 Then
                                                    deskTypeLineLayer = nameTypeLineLayer
                                                ElseIf IsNothing(patternTypeLineLayer) = True Then
                                                    Dim boolWriteData As Boolean = CadastralTools.MathFunction.setValueToArray2d(userPropertiesLayer, "DeskLineType", "Continuous", 0, 1)
                                                ElseIf patternTypeLineLayer.Trim.Length = 0 Then
                                                    Dim boolWriteData As Boolean = CadastralTools.MathFunction.setValueToArray2d(userPropertiesLayer, "DeskLineType", "Continuous", 0, 1)
                                                Else
                                                    Dim arrayMaskLine As String() = patternTypeLineLayer.Split(",")
                                                    Dim userLayerLineType As DwgLinetype = createLineTypeDwgByPattern(nameTypeLineLayer, deskTypeLineLayer, arrayMaskLine)
                                                    If IsNothing(userLayerLineType) = True Then
                                                        Dim boolWriteData As Boolean = CadastralTools.MathFunction.setValueToArray2d(userPropertiesLayer, "DeskLineType", "Continuous", 0, 1)
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                                Dim createNewLayer As DwgLayer = createDwgLayerByProperties(userPropertiesLayer)
                                If IsNothing(createNewLayer) = False Then
                                    Layer = createNewLayer
                                End If
                            Else
                                Layer = getLayerDwg(layerObject)
                            End If
                            'создать блок
                        ElseIf nameField Like "Name" Then
                            Dim nameBlk As String = arrayPropertyStyle(1, i)
                            If nameBlk.Trim.Length > 0 Then
                                Dim boolFindBlk As Boolean = HacBlockDwg(nameBlk)
                                If boolFindBlk = False Then
                                    Dim blk As DwgBlock = CopyBlockReferenceByDestDraw(nameBlk)
                                    If IsNothing(blk) = False Then
                                        NameBlock = nameBlk
                                    Else
                                        NameBlock = ""
                                    End If
                                Else
                                    NameBlock = nameBlk
                                End If
                            End If
                        ElseIf nameField Like "LineType" Then
                            If value Like "ПоСлою" Then
                                LineType = ActiveProjectDrawing.Linetypes.ByLayer
                            ElseIf value Like "ПоБлоку" Then
                                LineType = ActiveProjectDrawing.Linetypes.ByBlock
                            Else
                                'загружаем тип линии
                                Dim findTypeLine As Boolean = HacLineTypeDwg(value)
                                If findTypeLine = False Then
                                    Dim userPropertiesTypeLine As String(,) = {}
                                    Dim boolFindLineTypeData = CadastralTools.TemplateXML.FuncReadLineTypeInfoToXML(TemplateXML, value, userPropertiesTypeLine)
                                    If IsArray(userPropertiesTypeLine) = True Then
                                        Dim deskTypeLineLayer As String = CadastralTools.MathFunction.findValueToArray2d(userPropertiesTypeLine, "DeskLineType", 0, 1)
                                        Dim patternTypeLineLayer As String = CadastralTools.MathFunction.findValueToArray2d(userPropertiesTypeLine, "CodeLineType", 0, 1)
                                        If IsNothing(deskTypeLineLayer) = True Then
                                            deskTypeLineLayer = value
                                        ElseIf deskTypeLineLayer.Trim.Length = 0 Then
                                            deskTypeLineLayer = value
                                        ElseIf IsNothing(patternTypeLineLayer) = True Then
                                            Dim boolWriteData As Boolean = CadastralTools.MathFunction.setValueToArray2d(userPropertiesTypeLine, "DeskLineType", "Continuous", 0, 1)
                                        ElseIf patternTypeLineLayer.Trim.Length = 0 Then
                                            Dim boolWriteData As Boolean = CadastralTools.MathFunction.setValueToArray2d(userPropertiesTypeLine, "DeskLineType", "Continuous", 0, 1)
                                        Else
                                            Dim arrayMaskLine As String() = patternTypeLineLayer.Split(",")
                                            Dim userLayerLineType As DwgLinetype = createLineTypeDwgByPattern(value, deskTypeLineLayer, arrayMaskLine)
                                            If IsNothing(userLayerLineType) = True Then
                                                Dim boolWriteData As Boolean = CadastralTools.MathFunction.setValueToArray2d(userPropertiesTypeLine, "DeskLineType", "Continuous", 0, 1)
                                                LineType = ActiveProjectDrawing.Linetypes.Continuous
                                            End If
                                        End If
                                    End If
                                Else
                                    LineType = getLineTypeDwg(value)
                                End If
                            End If
                        ElseIf nameField Like "Color" Then
                            Dim userColor As Integer = Val(value)
                            If userColor < 0 Then userColor = 0
                            If userColor > 256 Then userColor = 256
                            Dim cadColor As CadColor = New CadColor(userColor)
                            IndexColor = cadColor
                        ElseIf nameField Like "LineWidth" Then
                            LineWeight = Val(value)
                        ElseIf nameField Like "LineScale" Then
                            LineTypeScale = Val(value)
                        ElseIf nameField Like "Scale" Then
                            Dim userScale As Double = Val(value)
                            ScaleBlock = userScale
                        ElseIf nameField Like "Annotative" Then
                            AnnotativeBlock = CBool(value)
                        End If
                    Next i
                End If
            End If
            Return True
        Catch ex As System.Exception
            Return False
        End Try

    End Function
    '===========================================================================================================================
    'применить стиль к блоку
    Function setStyle(ByRef userBlock As DwgInsert) As Boolean
        If IsNothing(ActiveProjectDrawing) = True Then Return False
        If IsNothing(userBlock) = False Then
            Try
                'цвет
                userBlock.Color = IndexColor
                'слой
                userBlock.Layer = Layer
                'тип линии
                userBlock.Linetype = LineType
                'масштаб типа линии
                userBlock.LinetypeScale = LineTypeScale
                'вес линии
                userBlock.Lineweight = LineWeight
                'сасштаб блока
                If ScaleBlock > 0 Then
                    Dim tempScale As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(ScaleBlock, ScaleBlock, ScaleBlock)
                    userBlock.Scale = tempScale
                End If
                userBlock.Rotation = RotationBlock
            Catch ex As system.Exception
            End Try
        End If
        Return True
    End Function

End Class
Public Class AttributeCadastralBlock
    Private _nameField As String
    Private _valueField As String
    Private _definition As String
    Private _rotation As Double
    Private _height As Double
    Private _X As Double
    Private _Y As Double
    Private _alignmentH As Integer
    Private _alignmentY As Integer
    Private _visible As Boolean
    Private _textStyle As String
    Private _color As Integer
    Public Sub New()
        _nameField = ""
        _valueField = ""
        _definition = ""
        _rotation = 0
        _height = 2
        _X = 0
        _Y = 0
        _alignmentH = 0
        _alignmentY = 0
        _visible = True
        _textStyle = ""
        _color = 7
    End Sub
    'Получить или установить имя поля
    Public Property TagAttribute As String
        Get
            Return _nameField
        End Get
        Set(value As String)
            _nameField = value
        End Set
    End Property

    'Получить или установить значение поля
    Public Property TextStringAttribute As String
        Get
            Return _valueField
        End Get
        Set(value As String)
            _valueField = value
        End Set
    End Property
    'Получить или установить значение поля
    Public Property DefinitionAttribute As String
        Get
            Return _definition
        End Get
        Set(value As String)
            _definition = value
        End Set
    End Property
    'Получить или установить поворот
    Public Property RotationAttribute As Double
        Get
            Return _rotation
        End Get
        Set(value As Double)
            _rotation = value
        End Set
    End Property

    'Получить или установить высоту
    Public Property HeightAttribute As Double
        Get
            Return _height
        End Get
        Set(value As Double)
            _height = value
        End Set
    End Property

    'Получить или установить смещение по X
    Public Property XAttribute As Double
        Get
            Return _X
        End Get
        Set(value As Double)
            _X = value
        End Set
    End Property

    'Получить или установить смещение по Y
    Public Property YAttribute As Double
        Get
            Return _Y
        End Get
        Set(value As Double)
            _Y = value
        End Set
    End Property

    'Получить или установить смещение по X
    Public Property HorizontalAlignment As Integer
        Get
            Return _alignmentH
        End Get
        Set(value As Integer)
            _alignmentH = value
        End Set
    End Property

    'Получить или установить смещение по Y
    Public Property VerticalAlignment As Integer
        Get
            Return _alignmentY
        End Get
        Set(value As Integer)
            _alignmentY = value
        End Set
    End Property

    'Получить или установить видимость
    Public Property VisibleAttribute As Boolean
        Get
            Return _visible
        End Get
        Set(value As Boolean)
            _visible = value
        End Set
    End Property
    'Получить или установить значение поля
    Public Property TextStyleAttribute As String
        Get
            Return _textStyle
        End Get
        Set(value As String)
            _textStyle = value
        End Set
    End Property
    'Получить или установить смещение по Y
    Public Property ColorAttribute As Integer
        Get
            Return _color
        End Get
        Set(value As Integer)
            _color = value
        End Set
    End Property
    '////////////////////////////////////////////////////////////////////////////////////////////////////
    'функция читиет атрибуты блока, возвращает массив (Tag,Prompt,Content)
    Public Shared Function getAttributesInsertBlock(ByVal dwgBlk As DwgInsert) As List(Of AttributeCadastralBlock)
        Dim result As List(Of AttributeCadastralBlock) = New List(Of AttributeCadastralBlock)
        If IsNothing(dwgBlk) = False Then
            Dim attribute As IEnumerable(Of IAttrib) = dwgBlk.Attribs
            If attribute.Count > 0 Then
                For i As Integer = 0 To attribute.Count - 1
                    Dim attr As IAttrib = attribute.ElementAt(i)
                    Dim propertiesAttrib As AttributeCadastralBlock = New AttributeCadastralBlock
                    propertiesAttrib.TagAttribute = attr.TagString
                    propertiesAttrib.DefinitionAttribute = attr.PromtString
                    propertiesAttrib.TextStringAttribute = attr.Content
                    propertiesAttrib.XAttribute = attr.Position.X
                    propertiesAttrib.YAttribute = attr.Position.Y
                    result.Add(propertiesAttrib)
                Next i
            End If
        End If
        Return result
    End Function
    '////////////////////////////////////////////////////////////////////////////////////////////////////
    'функция читиет определенный атрибут блока
    Public Shared Function getTextStringAttrabuteBlock(ByVal insertBlock As DwgInsert, ByVal tagAttribute As String) As String
        Dim result As String = ""
        If IsNothing(insertBlock) = True Then Return result
        If IsNothing(tagAttribute) = True Then Return result
        Dim attribute As IEnumerable(Of IAttrib) = insertBlock.Attribs
        If attribute.Count > 0 Then
            For i As Integer = 0 To attribute.Count - 1
                Dim attr As IAttrib = attribute.ElementAt(i)
                If attr.TagString Like tagAttribute Then
                    result = attr.Content
                    Exit For
                End If
            Next
        End If
        Return result
    End Function
    '////////////////////////////////////////////////////////////////////////////////////////////////////
    'функция перезаписыват значения атрибутов блока, входной массив (Tag,Content)
    Public Shared Function OverwriteAttributeBlk(ByVal dwgBlk As DwgInsert, ByVal listAttribs As List(Of AttributeCadastralBlock)) As Boolean
        OverwriteAttributeBlk = False
        If IsNothing(listAttribs) = True Then
            Return False
        End If
        If listAttribs.Count > 0 Then
            Dim attribute As IEnumerable(Of IAttrib) = dwgBlk.Attribs
            If attribute.Count > 0 Then
                For i As Integer = 0 To attribute.Count - 1
                    Dim attr As IAttrib = attribute.ElementAt(i)
                    Dim tagAttrib As String = attr.TagString
                    For Each userAttrib As AttributeCadastralBlock In listAttribs
                        Dim tagUser As String = userAttrib.TagAttribute
                        If IsNothing(tagUser) = False Then
                            If tagUser.Trim Like tagAttrib Then
                                attr.Content = userAttrib.TextStringAttribute
                                Exit For
                            End If
                        End If
                    Next
                Next
            End If
        End If
        OverwriteAttributeBlk = True
    End Function
    '-===================================================================================
    'Функция возвращает координаты вставки атрибута и параметры точки выравнивания(массив-X,Y,ГорВыр 1-лево 2-право,ВертВыр 1-низ 2-верх)
    Public Function calculatePositionAttributes(ByVal Vert1 As Vector2D, ByVal Vert2 As Vector2D, ByVal Vert3 As Vector2D, ByVal Dist As Double) As Boolean
        calculatePositionAttributes = False
        Try
            'определяем дирекционный угол на первую линию
            Dim Angle1 As Double = (Vert2 - Vert1).Angle
            'определяем дирекционый угол на вторую линию
            Dim Angle2 As Double = (Vert2 - Vert3).Angle
            'определяем внутренний угол
            Dim Angle As Double = Angle2 - Angle1
            If Angle < 0 Then Angle = Angle + Math.PI * 2 'если угол меньше 0 тогда добавим 360
            Dim Angle2D As Double = Angle / 2 'разделим угол пополам
            Dim Napravlenie As Double = Angle1 + Angle2D + Math.PI  'вычисляем дир.напраление
            If Napravlenie > Math.PI * 2 Then Napravlenie = Napravlenie - Math.PI * 2 'если угол больше 360 градусов
            'вычисляем координаты точки вставки атрибута
            Dim Dx As Double = Dist * Math.Cos(Napravlenie) 'вычисляем приращения координат
            Dim Dy As Double = Dist * Math.Sin(Napravlenie)
            XAttribute = Dx
            YAttribute = Dy
            Return True
        Catch ex As System.Exception
            XAttribute = Dist
            YAttribute = 0
        End Try
    End Function
End Class