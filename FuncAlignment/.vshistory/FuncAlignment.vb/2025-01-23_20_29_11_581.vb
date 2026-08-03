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
Imports Topomatic.Cad.Foundation.Cogo
'Imports Topomatic.Sfc
Imports Topomatic.Cad.Foundation
Imports Topomatic.Sfc
Imports Topomatic.Sfc.Layer
Imports Topomatic.Planchet.Entities
Imports System.IO
Imports System.Text.RegularExpressions
Imports Topomatic.Visualization.Geometry
Imports Topomatic.Alg
Imports Topomatic
Imports Topomatic.ApplicationPlatform.Core
Imports Topomatic.FoundationClasses
Imports Topomatic.Alg.Road.Core
Imports Topomatic.FoundationClasses.Undo
Public Class FuncAlignment
    'функция создает полилинию из трассы и задает ее смещение и реверс
    Public Shared Function FuncOffsetAlignment(ByVal ActivDocument As Drawing, ByVal userAlignment As Alignment, ByVal offset As Double, Optional boolReverse As Boolean = False) As DwgPolyline
        FuncOffsetAlignment = Nothing
        If IsNothing(userAlignment) = True Then
            Return Nothing
        End If

        Dim acPoly3dAlign As Polyline3D = New Polyline3D
        userAlignment.Plan.CompoundLine.ToPolyLine(acPoly3dAlign)

        Dim ArrayCoord As Double(,) = Nothing
        Dim countarrayPos As Integer = 0
        If acPoly3dAlign.Count > 1 Then
            For i As Integer = 0 To acPoly3dAlign.Count - 1
                Dim pos As BugleVector3D = acPoly3dAlign.Item(i)
                ReDim Preserve ArrayCoord(3, countarrayPos)
                ArrayCoord(0, countarrayPos) = pos.Vertex.X
                ArrayCoord(1, countarrayPos) = pos.Vertex.Y
                ArrayCoord(2, countarrayPos) = pos.Bugle
                countarrayPos += 1
            Next
            If IsArray(ArrayCoord) = True Then
                If ArrayCoord.GetUpperBound(1) > 0 Then
                    Dim axisPline As DwgPolyline = RoburFunc.FuncDrawPolylineToArrayCoord(ActivDocument, ArrayCoord, False, boolReverse)
                    If IsNothing(axisPline) = False Then
                        Dim offPlineCurve2d As Polyline2DCurveOffset = New Polyline2DCurveOffset()
                        offPlineCurve2d.Offset = offset
                        Dim d As List(Of Polyline2DCurveOffset) = New List(Of Polyline2DCurveOffset)
                        d.Add(offPlineCurve2d)
                        Dim PlineCurve2d As Polyline2DCurve = New Polyline2DCurve(axisPline.ToList)
                        Dim PlineCurve2d_ As Polyline2DCurve = PlineCurve2d.Offset(d)
                        Dim newPlineCurve2D As DwgPolyline = New DwgPolyline()
                        For i As Integer = 0 To PlineCurve2d_.Count - 1
                            newPlineCurve2D.Add(PlineCurve2d_.Item(i))
                        Next i
                        ActivDocument.ActiveSpace.Add(newPlineCurve2D)
                        axisPline.Clear()
                        Return newPlineCurve2D
                    Else
                        Return Nothing
                    End If

                Else
                    Return Nothing
                End If
            Else
                Return Nothing
            End If
        Else
            Return Nothing
        End If
    End Function

    Public Shared Function FuncCreateAdditionalAxis(ByVal ActivDocument As Drawing, ByVal userAlignment As Alignment, ByVal offset As Double, Optional boolReverse As Boolean = False) As DwgPolyline
        Dim vStart As Vector2D = New Vector2D(0, offset)
        Dim vEnd As Vector2D = New Vector2D(userAlignment.Plan.CompoundLine.Length, offset)
        Dim listParam As List(Of Vector2D) = New List(Of Vector2D)
        listParam.Add(vStart)
        listParam.Add(vEnd)
        Dim additionalAxis As CompoundLine = TransitionSolver.CompoundLineDisplace(userAlignment.Plan.CompoundLine, listParam)
        Dim poly3D As Polyline3D = New Polyline3D()
        additionalAxis.ToPolyLine(poly3D)
        If poly3D.Count > 1 Then
            Dim ArrayCoord As Double(,) = Nothing
            Dim countarrayPos As Integer = 0
            For i As Integer = 0 To poly3D.Count - 1
                Dim pos As BugleVector3D = poly3D.Item(i)
                ReDim Preserve ArrayCoord(3, countarrayPos)
                ArrayCoord(0, countarrayPos) = pos.Vertex.X
                ArrayCoord(1, countarrayPos) = pos.Vertex.Y
                ArrayCoord(2, countarrayPos) = pos.Bugle
                countarrayPos += 1
            Next
            If IsArray(ArrayCoord) = True Then
                If ArrayCoord.GetUpperBound(1) > 1 Then
                    Dim axisPline As DwgPolyline = RoburFunc.FuncDrawPolylineToArrayCoord(ActivDocument, ArrayCoord, False, boolReverse)
                    Return axisPline
                End If
            End If
        Else
            Return Nothing
        End If
        Return Nothing
    End Function

    Public Shared Function FuncReturnAlignmentAndSusfaceByNameProject(ByVal ActivDocument As Drawing, ByVal nameAlignment As String, ByRef userAlignment As Alignment) As Boolean
        FuncReturnAlignmentAndSusfaceByNameProject = False
        'Получение всех моделей проекта
        Try
            Dim Project As ModelProject = ApplicationHost.Current.ActiveProject
            Dim childs As IProjectModel() = Project.Model.GetChilds()
            For Each child As IProjectModel In childs
                Dim modelUri As URI = child.Uri
                If modelUri.Extension Like ".roadx" Then
                    Dim fileNameRoad As String = Path.GetFileNameWithoutExtension(modelUri.LastPathComponent)
                    If fileNameRoad Like nameAlignment Then
                        Dim userRoadModel As RoadModel = child.Model
                        If IsNothing(userRoadModel) = False Then
                            userAlignment = userRoadModel.Alignment
                            Return True
                        End If
                    End If
                End If
            Next
        Catch ex As System.Exception
        End Try
    End Function
End Class
