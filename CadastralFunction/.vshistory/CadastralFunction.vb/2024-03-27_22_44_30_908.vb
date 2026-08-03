Imports System.Globalization
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities

Public Class CadastralFunction
    '==================================================================================================================================
    'функция ищет внутренние островки
    Public Shared Function FuncFindCadastralOstrov(ByVal ActivDocument As Drawing, ByVal acPoly As DwgPolyline, ByVal locadID As String, ByVal arrayData As String(,)) As Boolean
        FuncFindCadastralOstrov = False
        If IsNothing(acPoly) = False Then
            Dim acBound As BoundingBox2D = acPoly.Bounds

        End If
    End Function
End Class
