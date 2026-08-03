Imports System.Globalization
Imports System.Net
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities

Public Class CadastralFunction
    '==================================================================================================================================
    'функция ищет внутренние островки контура
    Public Shared Function FuncFindCadastralOstrov(ByVal ActivDocument As Drawing, ByVal acPoly As DwgPolyline, ByVal locadID As String, ByRef arrayData As String(,)) As Integer
        FuncFindCadastralOstrov = 0
        If IsNothing(acPoly) = False Then
            Dim countOstrov As Integer = 0

            Dim acBound As BoundingBox2D = acPoly.Bounds
            Dim minX As Double = acBound.Min.X
            Dim minY As Double = acBound.Min.Y

            Dim maxX As Double = acBound.Max.X
            Dim maxY As Double = acBound.Max.Y
            Dim listVector2d As List(Of Vector2D) = New List(Of Vector2D)
            Dim listVertPoly As List(Of BugleVector2D) = acPoly.ToList
            If listVertPoly.Count > 2 Then
                For i As Integer = 0 To listVertPoly.Count - 1
                    Dim v2d As Vector2D = New Vector2D(listVertPoly.Item(i).Vertex.X, listVertPoly.Item(i).Vertex.Y)
                    listVector2d.Add(v2d)
                Next i
            End If
            If listVector2d.Count > 2 Then
                For Each acEnt As DwgEntity In ActivDocument.ActiveSpace.Entities
                    If TypeOf (acEnt) Is DwgPolyline Then
                        Dim acPolyOstrov As DwgPolyline = acEnt
                        Dim acBountOstrov As BoundingBox2D = acPolyOstrov.Bounds
                        Dim inPoint As Vector2D = New Vector2D((acBountOstrov.Min.X + acBountOstrov.Max.X) / 2, (acBountOstrov.Min.Y + acBountOstrov.Max.Y) / 2)
                        If inPoint.X > minX And inPoint.Y > minY Then
                            If inPoint.X < maxX And inPoint.Y < maxY Then
                                'проверяем более точно
                                Dim boolCross As Boolean = Topomatic.Cad.Foundation.CadLibrary.PosInPolygon(inPoint, listVector2d, False)
                                If boolCross = True Then
                                    Dim IDKOUNTER As String = FuncXRecords.FuncReadValueXData(acPolyOstrov, "IDKOUNTER")
                                    If IDKOUNTER.Trim Like locadID Then
                                        Dim number As Integer = Val(FuncXRecords.FuncReadValueXData(acPolyOstrov, "ConditionalNumber"))
                                        Dim localIDOstrov As Integer = Val(FuncXRecords.FuncReadValueXData(acPolyOstrov, "LOCALID"))
                                        Dim listVertOstrov As List(Of BugleVector2D) = New List(Of BugleVector2D)
                                        listVertOstrov = acPolyOstrov.ToList
                                        Dim sqrtOstrov As Double = MathFunction.FuncSQRTPolylineByList2(listVertOstrov)
                                        ReDim Preserve arrayData(3, countOstrov)
                                        arrayData(0, countOstrov) = number
                                        arrayData(1, countOstrov) = acPolyOstrov.ObjectID.ToString
                                        arrayData(2, countOstrov) = sqrtOstrov
                                        arrayData(3, countOstrov) = localIDOstrov
                                        countOstrov += 1
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
                If IsArray(arrayData) = True Then
                    If arrayData.GetUpperBound(1) > 0 Then
                        Dim boolSort As Boolean = MathFunction.FuncSortStrArray(arrayData, 0)
                    End If
                End If
            End If
            Return countOstrov
        End If
    End Function

    '==================================================================================================================================
    'функция ищет контура участка
    Public Shared Function FuncFindCadastralCounter(ByVal ActivDocument As Drawing, ByVal GlobalID As String, ByRef arrayData As String(,)) As Integer
        FuncFindCadastralCounter = 0
        Dim countCounter As Integer = 0
        If IsNothing(GlobalID) = True Then Return 0
        If GlobalID.Trim.Length < 2 Then Return 0
        For Each acEnt As DwgEntity In ActivDocument.ActiveSpace.Entities
            If TypeOf (acEnt) Is DwgPolyline Then
                Dim acPolyCounter As DwgPolyline = acEnt
                Dim IDGlob As String = FuncXRecords.FuncReadValueXData(acPolyCounter, "GLOBALID")
                If IDGlob.Trim.Length > 2 Then
                    If IDGlob.Trim Like GlobalID Then
                        Dim Fullnumber As String = FuncXRecords.FuncReadValueXData(acPolyCounter, "ConditionalNumber")
                        Dim localID As String = FuncXRecords.FuncReadValueXData(acPolyCounter, "LOCALID")
                        Dim numberCounter As Integer = FuncFormatZn.FuncFindNumberCounterCadastralObject(Fullnumber)
                        Dim listVertCounter As List(Of BugleVector2D) = New List(Of BugleVector2D)
                        listVertCounter = acPolyCounter.ToList
                        Dim sqrtCounter As Double = MathFunction.FuncSQRTPolylineByList2(listVertCounter)
                        ReDim Preserve arrayData(3, countCounter)
                        arrayData(0, countCounter) = numberCounter
                        arrayData(1, countCounter) = acPolyCounter.ObjectID.ToString
                        arrayData(2, countCounter) = sqrtCounter
                        arrayData(3, countCounter) = localID
                        countCounter += 1
                    End If
                End If
            End If
        Next
        If IsArray(arrayData) = True Then
            If arrayData.GetUpperBound(1) > 0 Then
                Dim boolSort As Boolean = MathFunction.FuncSortStrArray(arrayData, 0)
            End If
        End If
        Return countCounter
    End Function

    '==================================================================================================================================
    'функция забирает в массив точки (X,Y,ID,IDRNG)
    Public Shared Function FuncFindCadastralPoints(ByVal ActivDocument As Drawing, ByVal nameBlk As String, ByRef arrayData As String(,)) As Integer
        FuncFindCadastralPoints = 0
        Dim countBlk As Integer = 0
        For Each acEnt As DwgEntity In ActivDocument.ActiveSpace.Entities
            If TypeOf (acEnt) Is DwgInsert Then
                Dim acblk As DwgInsert = acEnt
                If acblk.Block.Name Like nameBlk Then
                    ReDim Preserve arrayData(3, countBlk)
                    arrayData(0, countBlk) = acblk.Position.X
                    arrayData(1, countBlk) = acblk.Position.Y
                    arrayData(2, countBlk) = acblk.ObjectID
                    countBlk += 1
                End If
            End If
        Next
        FuncFindCadastralPoints = countBlk
    End Function

    '==================================================================================================================================
    'функция ищет точку в узле и возвращает ее id
    Public Shared Function FuncFindCadastralPoint(ByVal ActivDocument As Drawing, ByVal posVertex As Vector2D, ByVal arrayPointCollectoion As String(,), ByRef arrayData As String(,), Optional ByVal dopusk As Double = 0.05, Optional localIDCounters As String = "") As Boolean
        FuncFindCadastralPoint = False
        Dim arrayResult As String = ""
        If IsArray(arrayPointCollectoion) = True Then
            For i As Integer = 0 To arrayPointCollectoion.GetUpperBound(1)
                Dim vertBlk As Vector2D = New Vector2D(arrayPointCollectoion(0, i), arrayPointCollectoion(1, i))
                If (vertBlk - posVertex).Length <= dopusk Then
                    dopusk = (vertBlk - posVertex).Length
                    Dim BlkID As String = arrayPointCollectoion(2, i)
                    Dim idRng As String = arrayPointCollectoion(3, i)
                    If IsNothing(BlkID) = False Then
                        Dim blkFind As DwgInsert = Nothing
                        Dim boolFindObj As Boolean = ActivDocument.ActiveSpace.Entities.TryGetObject(BlkID, blkFind)
                        If boolFindObj = True Then
                            'читаем атрибут
                            Erase arrayData
                            Dim boolFindAttrib As Boolean = RoburFunc.FuncReadAttributeBlk(ActivDocument, blkFind, arrayData)
                        End If
                        If IsNothing(idRng) = True Then
                            idRng = FuncXRecords.FuncReadValueXData(blkFind, "IDRNG")
                            arrayPointCollectoion(3, i) = idRng
                        End If
                    End If
                    If localIDCounters.Trim.Length > 0 Then
                        If idRng Like localIDCounters Then
                            Return True
                        End If
                    End If
                End If
            Next i
        End If
        If IsArray(arrayData) = True Then
            Return True
        End If
    End Function
End Class
