Imports System.Windows.Media
Imports System.Windows.Media.Media3D
Imports System.Windows.Shapes
Imports Microsoft.Office.Interop.Word
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.FoundationClasses
Imports Topomatic.FoundationClasses.Undo
Imports Drawing = Topomatic.Dwg.Drawing

Public Class CreateDWGObject
    Private _activeDocument As Drawing
    Public Sub New()
        _activeDocument = Nothing
    End Sub
    Public Property ActiveDrawing As Drawing
        Get
            Return _activeDocument
        End Get
        Set(value As Drawing)
            _activeDocument = value
        End Set
    End Property

    '================================================================================================
    'создание выноски с текстом
    Public Function createLeader(ByVal value As String, ByVal InsertFirstVertex As Vector2D, ByVal InserSecondVertex As Vector2D) As DwgLeader
        Dim result As DwgLeader = Nothing
        If IsNothing(ActiveDrawing) = True Then
            Return result
        End If
        result = New Topomatic.Dwg.Entities.DwgLeader
        result.Add(InsertFirstVertex)
        result.Position = InserSecondVertex
        result.Content = value
        ActiveDrawing.ActiveSpace.Add(result)
        Return result
    End Function
    '================================================================================================
    'создание выноски с текстом
    Public Function createLeader(ByVal value As String, ByVal InsertFirstVertex As Vector2D, ByVal InserSecondVertex As Vector2D, ByRef arrayObjectProperties As String(,), ByRef arrayMleaderProperties As String(,), ByVal templateXML As String, ByVal templateDWG As String, Optional annoScale As Double = 1) As DwgLeader
        Dim result As DwgLeader = Nothing
        If IsNothing(ActiveDrawing) = True Then
            Return result
        End If

        Try
            result = New Topomatic.Dwg.Entities.DwgLeader
            result.Add(InsertFirstVertex)
            result.Position = InserSecondVertex
            result.Content = value
            'назначаем свойства
            If IsArray(arrayObjectProperties) = True Then
                If arrayObjectProperties.Length > 0 Then
                    For i As Integer = 0 To arrayObjectProperties.GetUpperBound(1)
                        Dim nameField As String = arrayObjectProperties(0, i)
                        Dim valueField As String = arrayObjectProperties(1, i)
                        If nameField Like "Layer" Then
                            Dim mleaderLayer As DwgLayer = RoburFunc.FuncAddLayer(ActiveDrawing, valueField)
                            If IsNothing(mleaderLayer) = True Then
                                mleaderLayer = ActiveDrawing.ActiveLayer
                            End If
                            result.Layer = mleaderLayer
                        ElseIf nameField Like "Color" Then
                            Dim color As CadColor = New CadColor(256)
                            If valueField Like "ПоБлоку" Then
                                color = New CadColor(0)
                            Else
                                If IsNumeric(valueField) = True Then
                                    Dim indColor As Integer = CInt(valueField)
                                    If indColor < 0 Then indColor = 0
                                    If indColor > 256 Then indColor = 256
                                    color = New CadColor(CInt(indColor))
                                End If
                            End If
                            result.Color = color
                        ElseIf nameField Like "ArrowSize" Then
                            If IsNumeric(valueField) = True Then
                                result.ArrowheadSize = CInt(valueField) * annoScale
                            End If
                        ElseIf nameField Like "TextStyle" Then
                            Dim textStyle As DwgStyle = ActiveDrawing.ActiveStyle
                            If ActiveDrawing.Styles.Names.ContainsKey(valueField) = True Then
                                textStyle = ActiveDrawing.Styles.Item(valueField)
                            Else
                                Dim arrayPropertiesTextStyle As String(,) = {}
                                Dim boolReadTextStyleProperties As Boolean = CadastralTools.TemplateXML.FuncReadTextStylePropertiesInfoToXML(templateXML, valueField, arrayPropertiesTextStyle)
                                If arrayPropertiesTextStyle.Length > 0 Then
                                    textStyle = StyleCadastralObject.FuncCreateTextStyleByArray(ActiveDrawing, valueField, arrayPropertiesTextStyle)
                                End If
                            End If
                            result.Style = textStyle
                        ElseIf nameField Like "TextHeight" Then
                            result.Height = Val(valueField)
                        End If
                    Next i
                End If
            End If
            result.Annotative = True
            result.ArrowheadType = AcDimArrowheadType.ClosedBlank
            If IsNothing(arrayMleaderProperties) = False Then
                If IsArray(arrayMleaderProperties) = True Then
                    If arrayMleaderProperties.Length > 0 Then
                        For i As Integer = 0 To arrayMleaderProperties.GetUpperBound(1)
                            Dim nameField As String = arrayMleaderProperties(0, i)
                            Dim nameArrow As String = arrayMleaderProperties(1, i)
                            If nameField Like "NameArrow" Then
                                If nameArrow Like "" Then
                                    result.ArrowheadType = AcDimArrowheadType.None
                                ElseIf nameArrow Like "_ArchTick" Then
                                    result.ArrowheadType = AcDimArrowheadType.ArchTick
                                ElseIf nameArrow Like "_BoxBlank" Then
                                    result.ArrowheadType = AcDimArrowheadType.BoxBlank
                                ElseIf nameArrow Like "_BoxFilled" Then
                                    result.ArrowheadType = AcDimArrowheadType.BoxFilled
                                ElseIf nameArrow Like "_Closed" Then
                                    result.ArrowheadType = AcDimArrowheadType.Closed
                                ElseIf nameArrow Like "_ClosedBlank" Then
                                    result.ArrowheadType = AcDimArrowheadType.ClosedBlank
                                ElseIf nameArrow Like "_DatumBlank" Then
                                    result.ArrowheadType = AcDimArrowheadType.DatumBlank
                                ElseIf nameArrow Like "_DatumFilled" Then
                                    result.ArrowheadType = AcDimArrowheadType.DatumFilled
                                ElseIf nameArrow Like "_Default" Then
                                    result.ArrowheadType = AcDimArrowheadType.Default
                                ElseIf nameArrow Like "_Dot" Then
                                    result.ArrowheadType = AcDimArrowheadType.Dot
                                ElseIf nameArrow Like "_DotBlank" Then
                                    result.ArrowheadType = AcDimArrowheadType.DotBlank
                                ElseIf nameArrow Like "_DotSmall" Then
                                    result.ArrowheadType = AcDimArrowheadType.DotSmall
                                ElseIf nameArrow Like "_Integral" Then
                                    result.ArrowheadType = AcDimArrowheadType.Integral
                                ElseIf nameArrow Like "_None" Then
                                    result.ArrowheadType = AcDimArrowheadType.None
                                ElseIf nameArrow Like "_Oblique" Then
                                    result.ArrowheadType = AcDimArrowheadType.Oblique
                                ElseIf nameArrow Like "_Open" Then
                                    result.ArrowheadType = AcDimArrowheadType.Open
                                ElseIf nameArrow Like "_Open30" Then
                                    result.ArrowheadType = AcDimArrowheadType.Open30
                                ElseIf nameArrow Like "_Open90" Then
                                    result.ArrowheadType = AcDimArrowheadType.Open90
                                ElseIf nameArrow Like "_Origin" Then
                                    result.ArrowheadType = AcDimArrowheadType.Origin
                                ElseIf nameArrow Like "_Origin2" Then
                                    result.ArrowheadType = AcDimArrowheadType.Origin2
                                ElseIf nameArrow Like "_Small" Then
                                    result.ArrowheadType = AcDimArrowheadType.Small
                                Else
                                    Dim userBlock As DwgBlock = ActiveDrawing.Blocks.Item(nameArrow)
                                    If IsNothing(userBlock) = True Then
                                        userBlock = RoburFunc.FuncInsertBlockByDWGFile(ActiveDrawing, templateDWG, nameArrow)
                                    End If
                                    If IsNothing(userBlock) = False Then
                                        result.ArrowheadType = AcDimArrowheadType.UserDefined
                                        result.ArrowheadName = nameArrow
                                    Else
                                        result.ArrowheadType = AcDimArrowheadType.None
                                    End If
                                End If
                            End If
                        Next i
                    End If
                End If
            End If
            ActiveDrawing.ActiveSpace.Add(result)
        Catch ex As Exception
        End Try
        Return result
    End Function

    'создание полилинии из массива точек (0-x, 1-y)
    Public Function createPolylineToArrayPoint(ByRef ArrayVertex2d As String(,), Optional ByVal boolClosed As Boolean = False, Optional boolReverse As Boolean = False) As DwgPolyline
        Dim result As DwgPolyline = Nothing
        If IsArray(ArrayVertex2d) = True Then
            If ArrayVertex2d.GetUpperBound(0) > 0 Then
                Dim rp As Topomatic.Cad.Foundation.Vector3D = Nothing
                Dim polyline As DwgPolyline = New Topomatic.Dwg.Entities.DwgPolyline()
                If boolReverse = False Then
                    For i As Integer = 0 To ArrayVertex2d.GetUpperBound(0)
                        rp.X = ArrayVertex2d(i, 0)
                        rp.Y = ArrayVertex2d(i, 1)
                        rp.Z = 0
                        Dim pos = New Topomatic.Cad.Foundation.Vector2D(rp.X, rp.Y)
                        Dim bulgeSegment As Single = 0
                        polyline.Add(New Topomatic.Cad.Foundation.BugleVector2D(pos, bulgeSegment))
                    Next i
                Else
                    For i As Integer = ArrayVertex2d.GetUpperBound(0) To 0 Step -1
                        rp.X = ArrayVertex2d(i, 0)
                        rp.Y = ArrayVertex2d(i, 1)
                        rp.Z = 0
                        Dim pos = New Topomatic.Cad.Foundation.Vector2D(rp.X, rp.Y)
                        Dim bulgeSegment As Single = 0
                        polyline.Add(New Topomatic.Cad.Foundation.BugleVector2D(pos, bulgeSegment))
                    Next i
                End If
                If polyline.Count > 1 Then
                    If IsNothing(ActiveDrawing) = False Then
                        ActiveDrawing.ActiveSpace.Add(polyline)
                        If boolClosed = True Then
                            polyline.Closed = True
                        End If
                        Return polyline
                    End If
                End If
                result = polyline
            End If
        End If
        Return result
    End Function

    'создание полилинии из полигона
    Public Function createPolylineByPolygon(ByVal polygon As NetTopologySuite.Geometries.Polygon, Optional ByRef polyline As DwgPolyline = Nothing) As DwgPolyline
        If IsNothing(polygon) = False Then
            If polygon.Coordinates.Count > 2 Then
                If IsNothing(polyline) = True Then
                    polyline = New DwgPolyline
                Else
                    polyline.Clear()
                End If
                Dim listCoordPolugon As NetTopologySuite.Geometries.Coordinate() = polygon.Coordinates
                If listCoordPolugon.Length > 2 Then
                    For i As Integer = 0 To listCoordPolugon.Length - 2
                        Dim coord As NetTopologySuite.Geometries.Coordinate = listCoordPolugon(i)
                        Dim pos As Topomatic.Cad.Foundation.Vector2D = New Topomatic.Cad.Foundation.Vector2D(coord.X, coord.Y)
                        polyline.Add(New Topomatic.Cad.Foundation.BugleVector2D(pos, 0))
                    Next i
                    Dim startVertex As NetTopologySuite.Geometries.Coordinate = listCoordPolugon(0)
                    Dim endVertex As NetTopologySuite.Geometries.Coordinate = listCoordPolugon(listCoordPolugon.Length - 1)
                    Dim distance As Double = startVertex.Distance(endVertex)
                    If distance > 0.01 Then
                        Dim pos As Topomatic.Cad.Foundation.Vector2D = New Topomatic.Cad.Foundation.Vector2D(endVertex.X, endVertex.Y)
                        polyline.Add(New Topomatic.Cad.Foundation.BugleVector2D(pos, 0))
                    End If
                    If ActiveDrawing.ActiveSpace.Entities.Contains(polyline) = False Then
                        ActiveDrawing.ActiveSpace.Entities.Add(polyline)
                    End If
                    polyline.Closed = True
                End If
            End If
        End If
        Return polyline
    End Function
    'создание полилинии из полигона
    Public Function createPolylineByLineString(ByVal polygon As NetTopologySuite.Geometries.LineString, Optional ByRef polyline As DwgPolyline = Nothing) As DwgPolyline
        If IsNothing(polygon) = False Then
            If polygon.Coordinates.Count > 2 Then
                If IsNothing(polyline) = True Then
                    polyline = New DwgPolyline
                Else
                    polyline.Clear()
                End If
                Dim listCoordPolugon As NetTopologySuite.Geometries.Coordinate() = polygon.Coordinates
                If listCoordPolugon.Length > 2 Then
                    For i As Integer = 0 To listCoordPolugon.Length - 1
                        Dim coord As NetTopologySuite.Geometries.Coordinate = listCoordPolugon(i)
                        Dim pos As Topomatic.Cad.Foundation.Vector2D = New Topomatic.Cad.Foundation.Vector2D(coord.X, coord.Y)
                        polyline.Add(New Topomatic.Cad.Foundation.BugleVector2D(pos, 0))
                    Next i
                    If ActiveDrawing.ActiveSpace.Entities.Contains(polyline) = False Then
                        ActiveDrawing.ActiveSpace.Entities.Add(polyline)
                    End If
                End If
            End If
        End If
        Return polyline
    End Function
    'обновить полилинию из полигона
    Public Function UpdatePolylineFromPolygon(ByVal polygon As NetTopologySuite.Geometries.Polygon) As Boolean
        If IsNothing(ActiveDrawing) = True Then Return False
        If IsNothing(polygon) = True Then Return False
        'обновояем внешний контур
        Dim extendPolygon As NetTopologySuite.Geometries.LineString = polygon.ExteriorRing
        Dim exterialData As CadastralTools.StructureElement = TryCast(polygon.UserData, CadastralTools.StructureElement)
        If IsNothing(exterialData) = False Then
            Dim exteralhandle As UInteger = CUInt(exterialData.IdObject)
            Dim plineCounter As DwgPolyline = Nothing
            Dim boolFindOBject As Boolean = ActiveDrawing.ActiveSpace.Entities.TryGetObject(exteralhandle, plineCounter)
            If IsNothing(plineCounter) = False Then
                ' --- Обновляем вершины ---
                Dim coords As NetTopologySuite.Geometries.Coordinate() = extendPolygon.Coordinates
                If Not (coords Is Nothing OrElse coords.Length < 4) Then
                    ' Реальных вершин — n - 1 (последняя точка = копия первой)
                    Dim vertexCount As Integer = coords.Length - 1
                    ' Удаляем все существующие вершины
                    ' --- 1. Удаляем лишние вершины (если полилиния длиннее) ---
                    ' Удаляем с конца, пока не останется ровно vertexCount вершин.
                    While plineCounter.Count > vertexCount
                        plineCounter.RemoveAt(plineCounter.Count - 1)
                    End While
                    For j As Integer = 0 To vertexCount - 1
                        Dim newBuildVertex As BugleVector2D = New BugleVector2D(New Vector2D(coords(j).X, coords(j).Y), 0)
                        If j <= plineCounter.Count - 1 Then
                            plineCounter.Item(j) = newBuildVertex
                        Else
                            plineCounter.Add(newBuildVertex)
                        End If
                    Next
                    ' Замыкаем полилинию
                    plineCounter.Closed = True
                End If

            End If
        End If
        'островки
        For i As Integer = 0 To polygon.NumInteriorRings - 1
            Dim userPolygon As NetTopologySuite.Geometries.LineString = polygon.InteriorRings(i)
            Dim userData As CadastralTools.StructureElement = TryCast(userPolygon.UserData, CadastralTools.StructureElement)
            If userData Is Nothing Then Continue For
            Dim handle As UInteger = CUInt(userData.IdObject)
            If IsNothing(handle) = False Then
                Dim plineCounter As DwgPolyline = Nothing
                Dim boolFindOBject As Boolean = ActiveDrawing.ActiveSpace.Entities.TryGetObject(handle, plineCounter)
                If IsNothing(plineCounter) = False Then
                    ' --- Обновляем вершины ---
                    Dim coords As NetTopologySuite.Geometries.Coordinate() = userPolygon.Coordinates
                    If coords Is Nothing OrElse coords.Length < 4 Then Continue For
                    ' Реальных вершин — n - 1 (последняя точка = копия первой)
                    Dim vertexCount As Integer = coords.Length - 1
                    ' --- 1. Удаляем лишние вершины (если полилиния длиннее) ---
                    ' Удаляем с конца, пока не останется ровно vertexCount вершин.
                    While plineCounter.Count > vertexCount
                        plineCounter.RemoveAt(plineCounter.Count - 1)
                    End While
                    For j As Integer = 0 To vertexCount - 1
                        Dim newBuildVertex As BugleVector2D = New BugleVector2D(New Vector2D(coords(j).X, coords(j).Y), 0)
                        If j <= plineCounter.Count - 1 Then
                            plineCounter.Item(j) = newBuildVertex
                        Else
                            plineCounter.Add(newBuildVertex)
                        End If
                    Next
                    ' Замыкаем полилинию
                    plineCounter.Closed = True
                End If

            End If
        Next
        Return True
    End Function

    '===================================================================================================
    'обновление полилинии по линии LineString
    'Аналог UpdatePolylineFromPolygon: линия не замыкается принудительно — замкнутость
    'берётся из самой геометрии (LineString.IsClosed)
    Public Function UpdatePolylineFromLineString(ByVal line As NetTopologySuite.Geometries.LineString) As Boolean
        If IsNothing(ActiveDrawing) = True Then Return False
        If IsNothing(line) = True Then Return False
        Dim userData As CadastralTools.StructureElement = TryCast(line.UserData, CadastralTools.StructureElement)
        If IsNothing(userData) = False Then
            Dim handle As String = userData.IdObject
            If IsNothing(handle) = False Then
                Dim id As UInteger = CUInt(handle)
                If id > 0 Then
                    ' --- Открываем полилинию для записи ---
                    Dim plineCounter As DwgPolyline = Nothing
                    Dim boolFindObject As Boolean = ActiveDrawing.ActiveSpace.Entities.TryGetObject(id, plineCounter)
                    If boolFindObject = True Then
                        ' --- Обновляем вершины ---
                        Dim coords As NetTopologySuite.Geometries.Coordinate() = line.Coordinates
                        UpdatePolylineVertices(plineCounter, coords, line.IsClosed)
                    End If
                End If
            End If
        End If
        Return True
    End Function
    '===================================================================================================
    'переписывает вершины существующей полилинии по массиву координат
    'Если линия замкнута, последняя координата — копия первой, реальных вершин на одну меньше
    Private Shared Function UpdatePolylineVertices(ByVal plineCounter As DwgPolyline, ByVal coords As NetTopologySuite.Geometries.Coordinate(), ByVal boolClosed As Boolean) As Boolean
        If IsNothing(plineCounter) OrElse coords Is Nothing Then Return False

        ' Линии достаточно двух точек (у кольца их минимум 4)
        Dim vertexCount As Integer = coords.Length
        If boolClosed = True Then vertexCount = coords.Length - 1
        If vertexCount < 2 Then Return False

        ' --- 1. Удаляем лишние вершины (если полилиния длиннее) ---
        ' Удаляем с конца, пока не останется ровно vertexCount вершин.
        While plineCounter.Count > vertexCount
            plineCounter.RemoveAt(plineCounter.Count - 1)
        End While
        For j As Integer = 0 To vertexCount - 1
            If j <= plineCounter.Count - 1 Then
                plineCounter.Item(j) = New BugleVector2D(New Vector2D(coords(j).X, coords(j).Y), 0)
            Else
                plineCounter.Add(New BugleVector2D(New Vector2D(coords(j).X, coords(j).Y), 0))
            End If
        Next
        ' Замыкание берём из геометрии, а не выставляем всегда True
        plineCounter.Closed = boolClosed
        Return True
    End Function



    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'Блоки
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция вставляет блок в чертеж 
    Public Function insertBlockToDrawing(ByVal insertPosition As Topomatic.Cad.Foundation.Vector3D, ByVal styleBlock As CadastralBlockStyle, ByVal listAttributes As List(Of AttributeCadastralBlock), Optional boolMAttrib As Boolean = False) As DwgInsert
        Dim result As DwgInsert = Nothing
        If IsNothing(styleBlock) = True Then Return result
        Dim nameBlock As String = styleBlock.NameBlock
        If IsNothing(nameBlock) = True Then Return result
        If nameBlock.Trim.Length = 0 Then Return result
        Dim boolHacBlk As Boolean = styleBlock.HacBlockDwg(nameBlock)
        If boolHacBlk = False Then Return result
        Dim blk As DwgBlock = ActiveDrawing.Blocks.Item(nameBlock)
        If IsNothing(blk) = False Then
            Try
                ActiveDrawing.BeginUpdate()
                Dim insertBlk As DwgInsert = New DwgInsert()
                insertBlk.Block = blk
                insertBlk.Position = insertPosition
                '=======================================================================================================
                If insertBlk.HasAttribs = True Then
                    Dim collectionAttributes As IEnumerable(Of IAttrib) = insertBlk.Attribs
                    'массив для считывания атрибутов блока
                    Dim arrayProperties As String(,) = Nothing
                    ReDim Preserve arrayProperties(20, collectionAttributes.Count - 1)
                    For i As Integer = 0 To collectionAttributes.Count - 1
                        Dim userIAttrib As IAttrib = collectionAttributes.ElementAt(i)
                        Dim userAttrib As DwgAttrib = userIAttrib
                        arrayProperties(0, i) = userAttrib.TagString
                        arrayProperties(1, i) = userAttrib.PromtString
                        arrayProperties(2, i) = userAttrib.Content
                        arrayProperties(3, i) = userAttrib.Position.X
                        arrayProperties(4, i) = userAttrib.Position.Y
                        arrayProperties(5, i) = userAttrib.Style.Name
                        arrayProperties(6, i) = userAttrib.Height
                        arrayProperties(7, i) = userAttrib.Rotation
                        arrayProperties(8, i) = userAttrib.Flags
                        arrayProperties(9, i) = userAttrib.Annotative
                        arrayProperties(10, i) = userAttrib.Oblique
                        arrayProperties(11, i) = userAttrib.Bounds.Min.X
                        arrayProperties(12, i) = userAttrib.Bounds.Min.Y
                        arrayProperties(13, i) = userAttrib.Bounds.Max.X
                        arrayProperties(14, i) = userAttrib.Bounds.Max.Y
                        arrayProperties(15, i) = userAttrib.Color.ColorIndex
                        arrayProperties(16, i) = userAttrib.Fixed
                        arrayProperties(17, i) = userAttrib.Invisible
                        arrayProperties(18, i) = userAttrib.Justify
                        arrayProperties(19, i) = userAttrib.Layer.Name
                        arrayProperties(19, i) = userAttrib.TextAlignmentPoint.X
                        arrayProperties(20, i) = userAttrib.TextAlignmentPoint.Y
                        insertBlk.RemoveAttrib(userIAttrib)
                    Next i
                    For i As Integer = 0 To arrayProperties.GetUpperBound(1)
                        Try
                            Dim nameAttr As String = arrayProperties(0, i)
                            Dim promptAttr As String = arrayProperties(1, i)
                            If nameAttr.Contains("####") Then
                                nameAttr = "ISNAME"
                            End If
                            'вставляем новый атрибут
                            If boolMAttrib = False Then
                                '==================================================================================
                                'однострочный атрибут
                                Dim userAttrib As DwgAttrib = insertBlk.AddAttrib(nameAttr, boolMAttrib)
                                userAttrib.PromtString = arrayProperties(1, i)
                                userAttrib.Content = ""
                                Dim posAttr As Vector2D = New Vector2D(Val(arrayProperties(3, i)), Val(arrayProperties(4, i)))
                                userAttrib.Position = posAttr
                                Dim styleAttr As DwgStyle = ActiveDrawing.Styles.Item(arrayProperties(5, i))
                                userAttrib.Style = styleAttr
                                userAttrib.Height = Val(arrayProperties(6, i))
                                userAttrib.Rotation = Val(arrayProperties(7, i))
                                userAttrib.Flags = Val(arrayProperties(8, i))
                                userAttrib.Annotative = CBool(arrayProperties(9, i))
                                userAttrib.Oblique = arrayProperties(10, i)
                                Dim userBound As BoundingBox2D = New BoundingBox2D
                                userBound.Min.X = arrayProperties(11, i)
                                userBound.Min.Y = arrayProperties(12, i)
                                userBound.Max.X = arrayProperties(13, i)
                                userBound.Max.Y = arrayProperties(14, i)
                                Dim userColor As CadColor = New CadColor(Val(arrayProperties(15, i)))
                                userAttrib.Color = userColor
                                userAttrib.Fixed = False
                                userAttrib.Flags = AttributeFlags.VerificationRequired
                                userAttrib.Invisible = CBool(arrayProperties(17, i))
                                userAttrib.Justify = Val(arrayProperties(18, i))
                                Dim userLayer As DwgLayer = ActiveDrawing.Layers.Item(arrayProperties(19, i))
                                userAttrib.Layer = userLayer
                                Dim posPoint As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(Val(arrayProperties(19, i)), Val(arrayProperties(20, i)), insertBlk.Position.Z)
                                userAttrib.TextAlignmentPoint = posPoint
                                userAttrib.Backward = False
                                If IsNothing(listAttributes) = False Then
                                    If listAttributes.Count > 0 Then
                                        For j As Integer = 0 To listAttributes.Count - 1
                                            Dim attrDict As AttributeCadastralBlock = listAttributes.Item(j)
                                            Dim KeyAttr As String = attrDict.TagAttribute
                                            If KeyAttr Like nameAttr Then
                                                userAttrib.Height = attrDict.HeightAttribute
                                                userAttrib.Position = New Topomatic.Cad.Foundation.Vector3D(attrDict.XAttribute, attrDict.YAttribute, 0)
                                                userAttrib.Invisible = True
                                                userAttrib.Content = attrDict.TextStringAttribute
                                            End If
                                        Next j
                                    End If
                                End If
                            Else
                                '=======================================================================================
                                'многострочный атрибут
                                Dim userAttrib As DwgMAttrib = insertBlk.AddAttrib(nameAttr, boolMAttrib)
                                userAttrib.PromtString = arrayProperties(1, i)
                                userAttrib.Content = ""
                                Dim posAttr As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(Val(arrayProperties(3, i)), Val(arrayProperties(4, i)), insertBlk.Position.Z)
                                userAttrib.Position = posAttr
                                Dim styleAttr As DwgStyle = ActiveDrawing.Styles.Item(arrayProperties(5, i))
                                userAttrib.Style = styleAttr
                                userAttrib.Height = Val(arrayProperties(6, i))
                                userAttrib.Rotation = Val(arrayProperties(7, i))
                                userAttrib.Flags = Val(arrayProperties(8, i))
                                userAttrib.Annotative = CBool(arrayProperties(9, i))
                                Dim userBound As BoundingBox2D = New BoundingBox2D
                                userBound.Min.X = arrayProperties(11, i)
                                userBound.Min.Y = arrayProperties(12, i)
                                userBound.Max.X = arrayProperties(13, i)
                                userBound.Max.Y = arrayProperties(14, i)
                                Dim userColor As CadColor = New CadColor(Val(arrayProperties(15, i)))
                                userAttrib.Color = userColor
                                userAttrib.Fixed = False
                                userAttrib.Flags = AttributeFlags.VerificationRequired
                                userAttrib.Invisible = CBool(arrayProperties(17, i))
                                userAttrib.AttachmentPoint = AttachmentPoint.BottomLeft
                                userAttrib.BackgroundColor = New CadColor(255)
                                userAttrib.BackgroundFillType = BackgroundFillType.UseFillColor
                                userAttrib.RectangleHeight = 0
                                userAttrib.RectangleHeight = 0
                                userAttrib.FillBoxScale = 1.1
                                Dim userLayer As DwgLayer = ActiveDrawing.Layers.Item(arrayProperties(19, i))
                                userAttrib.Layer = userLayer
                                Dim posPoint As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(Val(arrayProperties(19, i)), Val(arrayProperties(20, i)), insertBlk.Position.Z)
                                If IsNothing(listAttributes) = False Then
                                    If listAttributes.Count > 0 Then
                                        For j As Integer = 0 To listAttributes.Count - 1
                                            Dim attrDict As AttributeCadastralBlock = listAttributes.Item(j)
                                            Dim KeyAttr As String = attrDict.TagAttribute
                                            If KeyAttr Like nameAttr Then
                                                userAttrib.Height = attrDict.HeightAttribute
                                                userAttrib.Position = New Topomatic.Cad.Foundation.Vector3D(attrDict.XAttribute, attrDict.YAttribute, 0)
                                                userAttrib.IsVisible = attrDict.VisibleAttribute
                                                userAttrib.Content = attrDict.TextStringAttribute
                                            End If
                                        Next j
                                    End If
                                End If
                            End If
                        Catch ex As System.Exception
                        End Try
                    Next i
                End If

                ActiveDrawing.ActiveSpace.Add(insertBlk)
                ActiveDrawing.EndUpdate()
                Return insertBlk
            Catch ex1 As System.Exception
            End Try
        End If
    End Function
    Public Function insertBlockToDrawing2(ByVal insertPosition As Topomatic.Cad.Foundation.Vector3D, ByVal styleBlock As CadastralBlockStyle, ByVal listAttributes As List(Of AttributeCadastralBlock)) As DwgInsert
        Dim result As DwgInsert = Nothing
        If IsNothing(styleBlock) = True Then Return result
        Dim nameBlock As String = styleBlock.NameBlock
        If IsNothing(nameBlock) = True Then Return result
        If nameBlock.Trim.Length = 0 Then Return result
        Dim boolHacBlk As Boolean = styleBlock.HacBlockDwg(nameBlock)
        If boolHacBlk = False Then Return result
        Dim blk As DwgBlock = ActiveDrawing.Blocks.Item(nameBlock)
        If IsNothing(blk) = False Then
            Try
                ActiveDrawing.BeginUpdate()
                Dim insertBlk As DwgInsert = New DwgInsert()
                insertBlk.Block = blk
                insertBlk.Position = insertPosition
                ActiveDrawing.ActiveSpace.Entities.Add(insertBlk)
                styleBlock.setStyle(insertBlk)
                '=======================================================================================================
                If insertBlk.HasAttribs = True Then
                    Dim attribs As IEnumerable(Of IAttrib) = insertBlk.Attribs
                    'Скрытие видимости чётных блоков
                    For i As Integer = 0 To attribs.Count() - 1
                        Dim curAttrib As IAttrib = attribs.ElementAt(i)
                        Dim attrEntity As DwgAttrib = TryCast(curAttrib, DwgEntity)
                        attrEntity.Layer = insertBlk.Layer
                        If attrEntity.IsVisible = False Then
                            'Чтобы была возможность управлять видимостью атрибута, необходимо, чтобы слой вставки блока совпадал со слоем атрибута!
                            curAttrib.Flags = AttributeFlags.Invisible
                        End If
                    Next
                    'Обновление значения атрибута (Допишем порядковый номер атрибута в коллекции, для примера)
                    For i As Integer = 0 To attribs.Count() - 1
                        Dim curAttrib = attribs.ElementAt(i)
                        For j As Integer = 0 To listAttributes.Count - 1
                            Dim attrDict As AttributeCadastralBlock = listAttributes.Item(j)
                            Dim tagUserAttr As String = attrDict.TagAttribute
                            If tagUserAttr Like curAttrib.TagString Then
                                curAttrib.Content = attrDict.TextStringAttribute
                                curAttrib.Position = New Topomatic.Cad.Foundation.Vector3D(attrDict.XAttribute - insertBlk.Position.X, attrDict.YAttribute - insertBlk.Position.Y, 0)
                            End If
                        Next j
                    Next
                End If
                'styleBlock.setStyle(insertBlk)
                ActiveDrawing.EndUpdate()
                Return insertBlk
            Catch ex1 As System.Exception
            End Try
        End If
        Return result
    End Function



End Class
