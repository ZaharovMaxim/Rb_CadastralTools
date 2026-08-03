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
End Class
