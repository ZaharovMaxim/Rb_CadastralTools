Imports Topomatic.ApplicationPlatform
Imports System.ComponentModel
Imports System.Windows.Forms
Imports Topomatic.ApplicationPlatform.Plugins
Imports Topomatic.Controls.Dialogs
'Imports Topomatic.Srv
Imports Topomatic.Dwg.Layer
Imports Topomatic.Cad.View
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.Cad.Foundation
Imports Topomatic.Cad.View.Hints
Imports Topomatic.Sfc
Imports Topomatic.Alg.Prf
Imports Topomatic.Visualization.Runtime
Imports Topomatic.Visualization.Libx
Imports Topomatic.Visualization
Imports Topomatic.Pipes.Runtime.Plt.Profile.DescrEnumConverters
Imports Topomatic.Sfc.Layer
Imports Topomatic.Visualization.Geometry
Imports Topomatic.Crs.Ast
Imports Topomatic.Alg.Road.Urb.UserStrips
Imports Topomatic.Controls.Core
Imports Topomatic.ApplicationPlatform.Core
Imports Topomatic.Dtm
Imports Topomatic.Planchet.Entities
Imports Topomatic.Pipes.Layers.Plan
Imports System.IO
Imports System.Drawing

Public Class FuncDecorPlane
    '===========================================================================================================================
    'фунция создает все необходимые стили для линейных объектов
    Public Shared Function FuncCreateStyleLinearObject(ByVal ActivDocument As Drawing, ByVal namePutchXML As String, ByVal ArrayRez As String(,), ByRef layerObj As DwgLayer, ByRef nameTypeLineObj As DwgLinetype, ByRef colorObj As CadColor, Optional ByRef ScaleLineObj As Double = 1, Optional ByRef WidthLineObj As Double = 20) As Boolean
        FuncCreateStyleLinearObject = False
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
                If ArrayRez(0, i).Trim Like "Layer" Then
                    Dim NameLayer As String = ArrayRez(1, i)
                    If IsNothing(NameLayer) = False Then
                        NameLayer = NameLayer.Trim
                        'ищем слой
                        Dim tempObjectLayer As DwgLayer = RoburFunc.FuncFindLayerDwg(ActivDocument, NameLayer)
                        'слой не найден, ищем его в шаблоне и создаем его
                        If IsNothing(tempObjectLayer) = True Then
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
                                            TypeLineLayer = FuncStyles.FuncCreateTypeLineByPattern(ActivDocument, nameLayerTypeLine, strDeskTypeLine, arrayMask, namePutchXML)
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
                    deskObject = ArrayRez(1, i).Trim
                End If
            Next i
            '===========================================================================================================
            arrayMaskLine = lineTypeCode1.Split(",")
            If IsArray(arrayMaskLine) = True Then
                If arrayMaskLine.Length > 2 Then
                    nameTypeLineObj = FuncStyles.FuncCreateTypeLineByPattern(ActivDocument, nameObject, deskObject, arrayMaskLine, namePutchXML)
                End If
            End If
            Return True
        End If
    End Function

    '===========================================================================================================================
    'функция создает все необходимые стили для штриховок
    Public Shared Function FuncCreateStyleHatchObject(ByVal ActivDocument As Drawing, ByVal namePutchXML As String, ByVal ArrayRez As String(,), ByRef namePattern As String, ByRef layerObj As DwgLayer, ByRef nameTypeLineObj As DwgLinetype, ByRef colorObj As CadColor, Optional ByRef ScaleLineObj As Double = 1, Optional ByRef WidthLineObj As Double = 20, Optional ByRef ScaleObject As Double = 1, Optional ByRef AngleObject As Double = 0) As Boolean
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
                                        TypeLineLayer = FuncStyles.FuncCreateTypeLineByPattern(ActivDocument, nameLayerTypeLine, strDeskTypeLine, arrayMask, namePutchXML)
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
            End If
        Next i
        'создаем тип линии
        If IsArray(arrayMaskLine) = True Then
            If arrayMaskLine.Length > 2 Then
                nameTypeLineObj = FuncStyles.FuncCreateTypeLineByPattern(ActivDocument, nameTypeLine, deskObject, arrayMaskLine, namePutchXML)
            End If
        End If
        Return True
    End Function

    '===========================================================================================================================
    'функция создает все необходимые стили для блоков
    Public Shared Function FuncCreateStylePointObject(ByVal ActivDocument As Drawing, ByVal namePutchXML As String, ByVal ArrayRez As String(,), ByRef nameObject As String, ByRef layerObj As DwgLayer, ByRef nameTypeLineObj As DwgLinetype, ByRef colorObj As CadColor, Optional ByRef ScaleLineObj As Double = 1, Optional ByRef WidthLineObj As Double = 20, Optional ByRef ScaleObject As Double = 1, Optional ByRef AngleObject As Double = 0, Optional ByRef attributeHeight As Double = 2.5) As Boolean
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
                                            TypeLineLayer = FuncStyles.FuncCreateTypeLineByPattern(ActivDocument, nameLayerTypeLine, strDeskTypeLine, arrayMask, namePutchXML)
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
                End If
            Catch ex As System.Exception
            End Try
        Next i
        'создаем тип линии
        If IsArray(arrayMaskLine) = True Then
            If arrayMaskLine.Length > 2 Then
                nameTypeLineObj = FuncStyles.FuncCreateTypeLineByPattern(ActivDocument, nameTypeLine, deskObject, arrayMaskLine, namePutchXML)
            End If
        End If
        Return True
    End Function

    '===========================================================================================================================
    'функция создает все необходимые стили для выносок
    Public Shared Function FuncCreateStyleLeaderObject(ByVal ActivDocument As Drawing, ByVal namePutchXML As String, ByVal ArrayRez As String(,), ByRef nameObject As String, ByRef layerObj As DwgLayer, ByRef nameTypeLineObj As DwgLinetype, ByRef colorObj As CadColor, Optional ByRef ScaleLineObj As Double = 1, Optional ByRef WidthLineObj As Double = 20, Optional ByRef textStyle As String = "Standard", Optional ByRef hText As Double = 2, Optional ByRef nameArrow As String = "Нет", Optional ByRef sizeArrow As Double = 2, Optional ByRef note As String = "") As Boolean
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
                                            TypeLineLayer = FuncStyles.FuncCreateTypeLineByPattern(ActivDocument, nameLayerTypeLine, strDeskTypeLine, arrayMask, namePutchXML)
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
                ElseIf ArrayRez(0, i).Trim Like "Note" Then
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
                nameTypeLineObj = FuncStyles.FuncCreateTypeLineByPattern(ActivDocument, nameTypeLine, deskObject, arrayMaskLine, namePutchXML)
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

End Class
