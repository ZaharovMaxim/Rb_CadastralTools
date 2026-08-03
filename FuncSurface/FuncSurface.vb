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
Imports Topomatic.Dtm
Public Class FuncSurface
    Public Shared Function FuncReturnArrayNameSurfaceToProject(ByVal ActivDocument As Drawing, ByRef arrayNameSurface As String()) As Boolean
        FuncReturnArrayNameSurfaceToProject = False
        Dim Project As ModelProject = ApplicationHost.Current.ActiveProject
        Dim childs As IProjectModel() = Project.Model.GetChilds()
        Dim count As Integer = 0
        If IsArray(arrayNameSurface) = True Then
            count = arrayNameSurface.Length
        End If
        For Each child As IProjectModel In childs
            Dim modelUri As URI = child.Uri
            If modelUri.Extension Like ".sfcx" Then
                Dim fileNameModel As String = Path.GetFileNameWithoutExtension(modelUri.LastPathComponent)
                ReDim Preserve arrayNameSurface(count)
                arrayNameSurface(count) = fileNameModel
                count += 1
            End If
        Next
    End Function
    'возвращает модель поверхности
    Public Shared Function FuncReturnSurfaceToProjectByName(ByVal ActivDocument As Drawing, ByVal nameSurface As String) As Surface
        FuncReturnSurfaceToProjectByName = Nothing
        Try
            Dim Project As ModelProject = ApplicationHost.Current.ActiveProject
            Dim childs As IProjectModel() = Project.Model.GetChilds()
            For Each child As IProjectModel In childs
                Dim modelUri As URI = child.Uri
                If modelUri.Extension Like ".sfcx" Then
                    Dim fileNameModel As String = Path.GetFileNameWithoutExtension(modelUri.LastPathComponent)
                    If nameSurface.Trim Like fileNameModel.Trim Then
                        Dim userTerrainModel As TerrainModel = child.Model
                        If IsNothing(userTerrainModel) = False Then
                            Return userTerrainModel.Surface
                        End If
                    End If
                End If
            Next
        Catch ex As System.Exception
        End Try
    End Function

    Public Shared Function FuncReturnDrawingToProjectByName(ByVal ActivDocument As Drawing, ByVal nameDrawing As String) As Topomatic.Dwg.Drawing
        FuncReturnDrawingToProjectByName = Nothing
        Try
            Dim Project As ModelProject = ApplicationHost.Current.ActiveProject
            Dim childs As IProjectModel() = Project.Model.GetChilds()
            For Each child As IProjectModel In childs
                Dim modelUri As URI = child.Uri
                If modelUri.Extension Like ".dwg" Then
                    Dim fileNameModel As String = Path.GetFileNameWithoutExtension(modelUri.LastPathComponent)
                    If nameDrawing.Trim Like fileNameModel.Trim Then
                        Dim userDrawingModel As Topomatic.Dwg.Drawing = child.Model
                        If IsNothing(userDrawingModel) = False Then
                            Return userDrawingModel
                        End If
                    End If
                End If
            Next
        Catch ex As System.Exception
        End Try
    End Function
End Class
