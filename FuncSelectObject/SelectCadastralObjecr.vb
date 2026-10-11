Imports CadastralTools
Imports NetTopologySuite
Imports System.IO
Imports Topomatic.Alg.Offsets
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg
Imports System.Text.RegularExpressions
Imports Topomatic.Dwg.Entities
Imports System.Drawing.Text
Imports System.Drawing
Imports Topomatic.Cad.View
Imports Microsoft.Office.Interop.Excel
Imports System.Windows.Shapes
Imports Topomatic.FoundationClasses
Imports Topomatic.FoundationClasses.Undo

Public Class SelectCadastralObject
    Public Shared Function selectAllObject(ByRef userCadView As CadView) As Dictionary(Of StructureElement.nameObject, List(Of Geometries.Geometry))
        Dim result As Dictionary(Of StructureElement.nameObject, List(Of Geometries.Geometry)) = New Dictionary(Of StructureElement.nameObject, List(Of Geometries.Geometry))
        Dim roburTopoNet As FuncTopologySuite = New FuncTopologySuite()
        userCadView.SelectionSet.Clear()
        userCadView.SelectionSet.SelectAll()
        For Each selectEntity As Object In userCadView.SelectionSet
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            If TypeOf selectEntity Is DwgEntity Then
                Dim dataStructuresObject As StructureElement = Nothing
                Dim boolRecData As Boolean = FuncXRecords.getXRecords(selectEntity, dataStructuresObject)
                If boolRecData = True Then
                    Dim nameObject As StructureElement.nameObject = dataStructuresObject.NameTypeObject
                    Dim typeGeometry As StructureElement.typeGeometryObject = dataStructuresObject.TypeGeometry
                    Dim objectGeometry As Geometries.Geometry = Nothing
                    If typeGeometry = StructureElement.typeGeometryObject.polygon Then
                        If TypeOf selectEntity Is DwgPolyline Then
                            Dim polyline As DwgPolyline = selectEntity
                            If polyline.Area > 0 Then
                                Dim polygon As NetTopologySuite.Geometries.Polygon = roburTopoNet.CreatePolygonByPolyline(polyline)
                                If IsNothing(polygon) = False Then
                                    polygon.UserData = dataStructuresObject
                                    If result.ContainsKey(nameObject) = False Then
                                        result.Add(nameObject, New List(Of Geometries.Geometry) From {polygon})
                                    Else
                                        Dim listObject As List(Of Geometries.Geometry) = result.Item(nameObject)
                                        listObject.Add(polygon)
                                        result.Item(nameObject) = listObject
                                    End If
                                End If
                            End If
                        End If
                    ElseIf typeGeometry = StructureElement.typeGeometryObject.polyline Then
                        If TypeOf selectEntity Is DwgPolyline Then
                            Dim polyline As DwgPolyline = selectEntity
                            If polyline.Length > 0 Then
                                Dim line As NetTopologySuite.Geometries.LineString = roburTopoNet.CreateLineString(polyline)
                                If IsNothing(line) = False Then
                                    line.UserData = dataStructuresObject
                                    If result.ContainsKey(nameObject) = False Then
                                        result.Add(nameObject, New List(Of Geometries.Geometry) From {line})
                                    Else
                                        Dim listObject As List(Of Geometries.Geometry) = result.Item(nameObject)
                                        listObject.Add(line)
                                        result.Item(nameObject) = listObject
                                    End If
                                End If
                            End If
                        End If
                    Else
                        Dim insertPoint As Vector3D = New Vector3D()
                        If TypeOf selectEntity Is DwgLeader Then
                            Dim acLeader As DwgLeader = selectEntity
                            insertPoint = acLeader.Item(0)
                        ElseIf TypeOf selectEntity Is DwgText Then
                            Dim acText As DwgText = selectEntity
                            insertPoint = acText.Position
                        ElseIf TypeOf selectEntity Is DwgCircle Then
                            Dim acCircle As DwgCircle = selectEntity
                            insertPoint = acCircle.Center
                        ElseIf TypeOf selectEntity Is DwgInsert Then
                            Dim acBlock As DwgInsert = selectEntity
                            insertPoint = acBlock.Position
                        End If
                        Dim point As Geometries.Point = New Geometries.Point(insertPoint.X, insertPoint.Y, insertPoint.Z)
                        point.UserData = dataStructuresObject
                        If result.ContainsKey(nameObject) = False Then
                            result.Add(nameObject, New List(Of Geometries.Geometry) From {point})
                        Else
                            Dim listObject As List(Of Geometries.Geometry) = result.Item(nameObject)
                            listObject.Add(point)
                            result.Item(nameObject) = listObject
                        End If
                    End If
                End If
            End If

        Next
        Return result
    End Function

    Public Shared Function selectCadastralObjectByPolygon(ByRef userCadView As CadView, ByVal polylineFrame As DwgPolyline) As Dictionary(Of StructureElement.nameObject, List(Of Geometries.Geometry))
        Dim result As Dictionary(Of StructureElement.nameObject, List(Of Geometries.Geometry)) = New Dictionary(Of StructureElement.nameObject, List(Of Geometries.Geometry))
        Dim roburTopoNet As FuncTopologySuite = New FuncTopologySuite()
        Dim listBulgePoint As List(Of BugleVector2D) = polylineFrame.ToList
        Dim listPoint As List(Of Vector2D) = New List(Of Vector2D)
        For Each point As BugleVector2D In listBulgePoint
            listPoint.Add(New Vector2D(point.Vertex))
        Next
        listPoint.Add(listPoint.Item(0))
        userCadView.SelectionSet.Clear()
        userCadView.SelectionSet.SelectByPolygon(Hints.FrameSelectType.Contains, listPoint, Function(obj) TypeOf obj Is DwgPolyline)

        For Each selectEntity As Object In userCadView.SelectionSet
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Dim dataStructuresObject As StructureElement = Nothing
            Dim boolRecData As Boolean = FuncXRecords.getXRecords(selectEntity, dataStructuresObject)
            If boolRecData = True Then
                Dim nameObject As StructureElement.nameObject = dataStructuresObject.NameTypeObject
                Dim typeGeometry As StructureElement.typeGeometryObject = dataStructuresObject.NameTypeObject
                Dim objectGeometry As Geometries.Geometry = Nothing
                If typeGeometry = StructureElement.typeGeometryObject.polygon Then
                    If TypeOf selectEntity Is DwgPolyline Then
                        Dim polyline As DwgPolyline = selectEntity
                        If polyline.Area > 0 Then
                            Dim polygon As NetTopologySuite.Geometries.Polygon = roburTopoNet.CreatePolygonByPolyline(polyline)
                            If IsNothing(polygon) = False Then
                                polygon.UserData = dataStructuresObject
                                If result.ContainsKey(nameObject) = False Then
                                    result.Add(nameObject, New List(Of Geometries.Geometry) From {polygon})
                                Else
                                    Dim listObject As List(Of Geometries.Geometry) = result.Item(nameObject)
                                    listObject.Add(polygon)
                                    result.Item(nameObject) = listObject
                                End If
                            End If
                        End If
                    End If
                ElseIf typeGeometry = StructureElement.typeGeometryObject.polyline Then
                    If TypeOf selectEntity Is DwgPolyline Then
                        Dim polyline As DwgPolyline = selectEntity
                        If polyline.Length > 0 Then
                            Dim line As NetTopologySuite.Geometries.LineString = roburTopoNet.CreateLineString(polyline)
                            If IsNothing(line) = False Then
                                line.UserData = dataStructuresObject
                                If result.ContainsKey(nameObject) = False Then
                                    result.Add(nameObject, New List(Of Geometries.Geometry) From {line})
                                Else
                                    Dim listObject As List(Of Geometries.Geometry) = result.Item(nameObject)
                                    listObject.Add(line)
                                    result.Item(nameObject) = listObject
                                End If
                            End If
                        End If
                    End If
                Else
                    Dim insertPoint As Vector3D = New Vector3D()
                    If TypeOf selectEntity Is DwgLeader Then
                        Dim acLeader As DwgLeader = selectEntity
                        insertPoint = acLeader.Item(0)
                    ElseIf TypeOf selectEntity Is DwgText Then
                        Dim acText As DwgText = selectEntity
                        insertPoint = acText.Position
                    ElseIf TypeOf selectEntity Is DwgCircle Then
                        Dim acCircle As DwgCircle = selectEntity
                        insertPoint = acCircle.Center
                    ElseIf TypeOf selectEntity Is DwgInsert Then
                        Dim acBlock As DwgInsert = selectEntity
                        insertPoint = acBlock.Position
                    End If
                    Dim point As Geometries.Point = New Geometries.Point(insertPoint.X, insertPoint.Y, insertPoint.Z)
                    point.UserData = dataStructuresObject
                    If result.ContainsKey(nameObject) = False Then
                        result.Add(nameObject, New List(Of Geometries.Geometry) From {point})
                    Else
                        Dim listObject As List(Of Geometries.Geometry) = result.Item(nameObject)
                        listObject.Add(point)
                        result.Item(nameObject) = listObject
                    End If
                End If
            End If
        Next
        Return result
    End Function

    '==========================================================================================================================================================
    'фильтрация объектов на основе таблицы PS
    Public Shared Function selectCadastralObject(ByRef activeDocument As Topomatic.Dwg.Drawing, Optional tablePS As String = "") As Dictionary(Of String, List(Of Geometries.Geometry))
        Dim result As Dictionary(Of String, List(Of Geometries.Geometry)) = New Dictionary(Of String, List(Of Geometries.Geometry))
        If IsNothing(activeDocument) = True Then Return result
        Dim acadTopoNet As FuncTopologySuite = New FuncTopologySuite
        If IsNothing(activeDocument) = True Then Return result
        Dim userDwgObjects As DwgEntities = activeDocument.ActiveSpace.Entities
        If userDwgObjects.Count > 0 Then
            For Each dwgEntity As DwgEntity In userDwgObjects
                If TypeOf dwgEntity Is DwgPolyline Then
                    Dim polyline As DwgPolyline = dwgEntity
                    Dim dataStructuresObject As StructureElement = Nothing
                    Dim boolRecData As Boolean = FuncXRecords.getXRecords(polyline, dataStructuresObject)
                    If boolRecData = True Then
                        Dim tablePSData As String = dataStructuresObject.TablePS
                        Dim boolFindObject As Boolean = False
                        If tablePS.Trim.Length > 0 Then
                            If tablePS Like tablePSData Then
                                boolFindObject = True
                            End If
                        Else
                            boolFindObject = True
                        End If
                        If tablePSData Like "CADASTRAL_OSTROV" Then
                            boolFindObject = True
                        End If
                        If boolFindObject = True Then
                            If polyline.Area > 0 And dataStructuresObject.TypeGeometry = StructureElement.typeGeometryObject.polygon Then
                                Dim polygon As NetTopologySuite.Geometries.Polygon = acadTopoNet.CreatePolygonByPolyline(polyline)
                                If IsNothing(polygon) = False Then
                                    polygon.UserData = dataStructuresObject
                                    If result.ContainsKey(tablePSData) = False Then
                                        result.Add(tablePSData, New List(Of Geometries.Geometry) From {polygon})
                                    Else
                                        Dim listObject As List(Of Geometries.Geometry) = result.Item(tablePSData)
                                        listObject.Add(polygon)
                                        result.Item(tablePSData) = listObject
                                    End If
                                End If
                            ElseIf dataStructuresObject.TypeGeometry = StructureElement.typeGeometryObject.polyline Then
                                If polyline.Length > 0 Then
                                    Dim line As NetTopologySuite.Geometries.LineString = acadTopoNet.CreateLineString(polyline)
                                    If IsNothing(line) = False Then
                                        line.UserData = dataStructuresObject
                                        If result.ContainsKey(tablePSData) = False Then
                                            result.Add(tablePSData, New List(Of Geometries.Geometry) From {line})
                                        Else
                                            Dim listObject As List(Of Geometries.Geometry) = result.Item(tablePSData)
                                            listObject.Add(line)
                                            result.Item(tablePSData) = listObject
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        End If
        Return result
    End Function



End Class
