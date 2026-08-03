Imports System.Globalization
Imports System.Net
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.Pipes.Layers.Caches.Plan.Segments.SegmentPlanCache.DrawingData

Public Class CadastralFunction
    '==================================================================================================================================
    'функция ищет внутренние островки контура
    Public Shared Function FuncFindCadastralOstrov(ByVal ActivDocument As Drawing, ByVal acPoly As DwgPolyline, ByVal localID As String, ByRef arrayData As String(,)) As Integer
        FuncFindCadastralOstrov = 0
        Try
            If IsNothing(acPoly) = False And IsNothing(ActivDocument) = False And IsNothing(localID) = False Then
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
                                        If IDKOUNTER.Trim Like localID Then
                                            Dim number As Integer = Val(FuncXRecords.FuncReadValueXData(acPolyOstrov, "ConditionalNumber"))
                                            Dim localIDOstrov As Integer = Val(FuncXRecords.FuncReadValueXData(acPolyOstrov, "LOCALID"))
                                            Dim sqrtOstrov As Double = MathFunction.GetArea(acPolyOstrov, 2)
                                            ReDim Preserve arrayData(3, countOstrov)
                                            arrayData(0, countOstrov) = number
                                            arrayData(1, countOstrov) = acPolyOstrov.ObjectID.ToString
                                            arrayData(2, countOstrov) = Math.Abs(sqrtOstrov)
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
        Catch ex As System.Exception
        End Try
    End Function

    '==================================================================================================================================
    'функция ищет контура участка
    Public Shared Function FuncFindCadastralCounter(ByVal ActivDocument As Drawing, ByVal GlobalID As String, ByRef arrayData As String(,)) As Integer
        FuncFindCadastralCounter = 0
        Dim countCounter As Integer = 0
        If IsNothing(GlobalID) = True Then Return 0
        If GlobalID.Trim.Length < 2 Then Return 0
        If IsNothing(ActivDocument) = False Then
            For Each acEnt As DwgEntity In ActivDocument.ActiveSpace.Entities
                If TypeOf (acEnt) Is DwgPolyline Then
                    Dim acPolyCounter As DwgPolyline = acEnt
                    Dim IDGlob As String = FuncXRecords.FuncReadValueXData(acPolyCounter, "GLOBALID")
                    If IDGlob.Trim.Length > 2 Then
                        If IDGlob.Trim Like GlobalID Then
                            Dim Fullnumber As String = FuncXRecords.FuncReadValueXData(acPolyCounter, "ConditionalNumber")
                            Dim localID As String = FuncXRecords.FuncReadValueXData(acPolyCounter, "LOCALID")
                            Dim numberCounter As Integer = FuncFormatZn.FuncFindNumberCounterCadastralObject(Fullnumber)
                            Dim sqrtCounter As Double = MathFunction.GetArea(acPolyCounter, 2)
                            ReDim Preserve arrayData(3, countCounter)
                            arrayData(0, countCounter) = numberCounter
                            arrayData(1, countCounter) = acPolyCounter.ObjectID.ToString
                            arrayData(2, countCounter) = Math.Abs(sqrtCounter)
                            arrayData(3, countCounter) = localID
                            countCounter += 1
                        End If
                    End If
                End If
            Next
        End If
        If IsArray(arrayData) = True Then
            If arrayData.GetUpperBound(1) > 0 Then
                Dim boolSort As Boolean = MathFunction.FuncSortStrArray(arrayData, 0)
            End If
        End If
        Return countCounter
    End Function

    '==================================================================================================================================
    'функция ищет метку участка
    Public Shared Function FuncFindCadastralLabel(ByVal ActivDocument As Drawing, ByVal localIDParcels As String) As DwgLeader
        FuncFindCadastralLabel = Nothing
        Try
            If IsNothing(ActivDocument) = False And IsNothing(localIDParcels) = False Then
                For Each acEnt As DwgEntity In ActivDocument.ActiveSpace.Entities
                    If TypeOf (acEnt) Is DwgLeader Then
                        Dim acLeaderParcel As DwgLeader = acEnt
                        Dim IDCounter As String = FuncXRecords.FuncReadValueXData(acLeaderParcel, "IDRNG")
                        If IsNothing(IDCounter) = False Then
                            If IDCounter.Trim Like localIDParcels.Trim Then
                                Return acLeaderParcel
                            End If
                        End If
                    End If
                Next
            End If
        Catch ex As System.Exception
        End Try
    End Function

    '==================================================================================================================================
    'функция обновляет метку участка
    Public Shared Function FuncUpdateCadastralLabel(ByVal ActivDocument As Drawing, ByRef leaderParcel As DwgLeader, ByVal arrayData As String(,)) As Boolean
        FuncUpdateCadastralLabel = False
        Try
            '============================================================================================================
            'обновляем метку
            If IsNothing(leaderParcel) = False Then
                Dim noteLabel As String = FuncXRecords.FuncReadValueXData(leaderParcel, "NOTE")
                If noteLabel.Trim.Length > 0 Then
                    If IsArray(arrayData) = True Then
                        For k As Integer = 0 To arrayData.GetUpperBound(1)
                            Dim nameField As String = arrayData(0, k)
                            Dim valStr As String = arrayData(1, k)
                            If noteLabel.IndexOf("<" & nameField & ">") > -1 Then
                                noteLabel = noteLabel.Replace("<" & nameField & ">", valStr)
                            End If
                        Next
                        If noteLabel.IndexOf("Short_cad_number") > -1 Then
                            Dim cad_number As String = FuncXML.FuncReturnValueByArrayProperties("Cad_number", arrayData)
                            If cad_number.Trim.Length > 0 Then
                                Dim pos As Integer = cad_number.LastIndexOf(":")
                                If cad_number.Length > pos Then
                                    Dim tempStr As String = Mid(cad_number, pos + 1)
                                    noteLabel = noteLabel.Replace("<Short_cad_number>", tempStr)
                                End If
                            End If
                        End If
                        leaderParcel.Content = noteLabel
                        Return True
                    End If
                End If
            End If
        Catch ex As System.Exception
        End Try
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

    '==================================================================================================================================
    'функция проеряет наличие в узле точки
    Public Shared Function FuncFindCadastralPointToVertex(ByVal ActivDocument As Drawing, ByVal posVertex As Vector2D, ByVal arrayPointCollectoion As String(,), Optional ByVal dopusk As Double = 0.05, Optional localIDCounters As String = "") As String
        FuncFindCadastralPointToVertex = ""
        If IsArray(arrayPointCollectoion) = True Then
            For i As Integer = 0 To arrayPointCollectoion.GetUpperBound(1)
                Dim vertBlk As Vector2D = New Vector2D(Val(arrayPointCollectoion(0, i)), Val(arrayPointCollectoion(1, i)))
                Dim BlkID As String = arrayPointCollectoion(2, i)
                If (vertBlk - posVertex).Length <= dopusk Then
                    If localIDCounters.Trim.Length > 0 Then
                        Dim blkFind As DwgInsert = Nothing
                        Dim boolFindObj As Boolean = ActivDocument.ActiveSpace.Entities.TryGetObject(BlkID, blkFind)
                        Dim idRng As String = FuncXRecords.FuncReadValueXData(blkFind, "IDRNG")
                        If idRng Like localIDCounters Then
                            Return BlkID
                        End If
                    Else
                        Return BlkID
                    End If
                End If
            Next i
        End If
    End Function

    '==============================================================================================
    'функция возвращает погрешность определения точки по категории земель
    Public Shared Function FuncReturnDeltaGeopointByCategory(ByVal Category As String) As Double
        FuncReturnDeltaGeopointByCategory = 0
        Try
            If Category.IndexOf("сельскохоз") > -1 Then
                Return 2.5
            ElseIf Category.IndexOf("населен") > -1 Then
                Return 0.1
            ElseIf Category.IndexOf("промышлен") > -1 Then
                Return 0.5
            ElseIf Category.IndexOf("охраняем") > -1 Then
                Return 5
            ElseIf Category.IndexOf("лесног") > -1 Then
                Return 5
            ElseIf Category.IndexOf("водног") > -1 Then
                Return 5
            ElseIf Category.IndexOf("запас") > -1 Then
                Return 5
            End If
        Catch
        End Try
    End Function

    '================================================================================================
    'создание таблицы координат
    Public Shared Function FuncDrawTableCoordinates(ByVal ActivDocument As Drawing, ByVal arrayData As String(,), ByVal position As Vector3D, Optional ByVal userTableStyle As DwgTableStyle = Nothing, Optional ByVal roundZN As Integer = 2) As DwgTable
        FuncDrawTableCoordinates = Nothing
        Try
            If IsNothing(userTableStyle) = True Then
                userTableStyle = ActivDocument.ActiveTableStyle
            End If
            Dim userTable As Topomatic.Dwg.Entities.DwgTable = New Topomatic.Dwg.Entities.DwgTable(userTableStyle, 4, 3, 3, 50)
            userTable.MergeCells(0, 0, 2, 0)
            userTable.MergeCells(1, 1, 2, 1)
            userTable.MergeCells(0, 1, 0, 2)
            userTable.Item(0, 0).SourceText = "Каталог координат"
            userTable.Item(0, 1).SourceText = "Обозначение характерных точек границ"
            userTable.Item(1, 1).SourceText = "Координаты, м"
            userTable.Item(1, 2).SourceText = "X"
            userTable.Item(2, 2).SourceText = "Y"
            Dim pos As Vector3D = New Vector3D(0, 0, 0)
            userTable.Position = position
            Dim countRows As Integer = 3
            If IsArray(arrayData) = True Then
                For i As Integer = 0 To arrayData.GetUpperBound(1)
                    Dim numb As String = arrayData(0, i)
                    Dim x As Double = Val(arrayData(2, i))
                    Dim y As Double = Val(arrayData(1, i))
                    Dim strX As String = FuncFormatZn.FuncFormatCoordinate(x, roundZN)
                    Dim strY As String = FuncFormatZn.FuncFormatCoordinate(y, roundZN)
                    userTable.AddRow(3)
                    userTable.Item(0, countRows).SourceText = numb
                    userTable.Item(1, countRows).SourceText = strX
                    userTable.Item(2, countRows).SourceText = strY
                    countRows += 1
                Next i
            End If
            ActivDocument.ActiveSpace.Add(userTable)
            Return userTable
        Catch ex As Exception
        End Try
    End Function


    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция вставляет блок в чертеж (height - высота, positionX, PositionY,flag,visible
    Public Shared Function FuncInsertCadastralPoint(ByVal ActivDocument As Drawing, ByVal nameBlk As String, ByVal position As Vector3D, ByVal attributeProperties As Dictionary(Of String, String(,)), Optional rotationBlk As Double = 0, Optional scaleBlk As Double = 1, Optional boolBackground As Boolean = False, Optional boolMAttrib As Boolean = False) As DwgInsert
        FuncInsertCadastralPoint = Nothing
        If IsNothing(nameBlk) = True Then Exit Function
        Dim blk As DwgBlock = ActivDocument.Blocks.Item(nameBlk)
        If IsNothing(blk) = False Then
            Try
                ActivDocument.BeginUpdate()
                Dim insertBlk As DwgInsert = New DwgInsert()
                insertBlk.Block = blk
                insertBlk.Position = position
                insertBlk.Annotative = True
                '=======================================================================================================
                If insertBlk.HasAttribs = True Then
                    Dim collectionAttributes As IEnumerable(Of IAttrib) = insertBlk.Attribs
                    'массив для считывания атрибутов блока
                    If collectionAttributes.Count > 0 And attributeProperties.Count > 0 Then
                        For i As Integer = 0 To collectionAttributes.Count - 1
                            Dim userIAttrib As IAttrib = collectionAttributes.ElementAt(i)
                            Dim userAttrib As DwgAttrib = userIAttrib
                            For j As Integer = 0 To attributeProperties.Count - 1
                                Dim attrDict As KeyValuePair(Of String, String(,)) = attributeProperties.ElementAt(j)
                                Dim KeyAttr As String = attrDict.Key
                                If KeyAttr Like userAttrib.TagString Then
                                    Dim userArrayProperties As String(,) = attrDict.Value
                                    If IsArray(userArrayProperties) = True Then
                                        Dim posAttribX As Double = 0
                                        Dim posAttribY As Double = 0
                                        Dim boolValue As Boolean = False
                                        For k As Integer = 0 To userArrayProperties.GetUpperBound(1)
                                            Dim nameF As String = userArrayProperties(0, k)
                                            If IsNothing(nameF) = False Then
                                                If nameF Like "height" Then
                                                    userAttrib.Height = Val(userArrayProperties(1, k))
                                                ElseIf nameF Like "positionX" Then
                                                    posAttribX = Val(userArrayProperties(1, k))
                                                ElseIf nameF Like "positionY" Then
                                                    posAttribY = Val(userArrayProperties(1, k))
                                                ElseIf nameF Like "flag" Then
                                                    userAttrib.Flags = Val(userArrayProperties(1, k))
                                                ElseIf nameF Like "visible" Then
                                                    userAttrib.Invisible = Val(userArrayProperties(1, k))

                                                ElseIf nameF Like "value" Then
                                                    userAttrib.Content = userArrayProperties(1, k)
                                                    boolValue = True
                                                ElseIf nameF Like "Justify" Then
                                                    userAttrib.Justify = Val(userArrayProperties(1, k))
                                                End If
                                                If posAttribX <> 0 OrElse posAttribY <> 0 Then
                                                    Dim posAttr1 As Vector3D = New Vector3D(posAttribX, posAttribY, insertBlk.Position.Z)
                                                    userAttrib.Position = posAttr1
                                                End If
                                            End If
                                        Next k
                                    End If
                                End If
                            Next j
                            If userAttrib.Invisible = False Then
                                userAttrib.Fixed = False
                                userAttrib.Flags = AttributeFlags.VerificationRequired
                            End If
                        Next i
                    End If
                End If
                ActivDocument.ActiveSpace.Add(insertBlk)
                ActivDocument.EndUpdate()
                Return insertBlk
            Catch ex1 As System.Exception
            End Try
        End If
    End Function

End Class
