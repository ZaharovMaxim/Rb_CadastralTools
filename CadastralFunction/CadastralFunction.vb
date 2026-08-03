Imports System.Globalization
Imports System.Net
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities

Public Class CadastralFunction
    '==================================================================================================================================
    'функция ищет все участки на чертеже (массив, 0-номер участка, 1-ID)
    Public Shared Function FuncFindCadastralParcels(ByVal ActivDocument As Drawing, ByVal deskObject As String, ByRef arrayData As String(,)) As Integer
        FuncFindCadastralParcels = 0
        Dim countParcels As Integer = 0
        If IsNothing(ActivDocument) = False Then
            For Each acEnt As DwgEntity In ActivDocument.ActiveSpace.Entities
                If TypeOf (acEnt) Is DwgPolyline Then
                    Dim acPolyCounter As DwgPolyline = acEnt
                    Dim strLabel As String = FuncXRecords.FuncReadValueXData(acPolyCounter, "LABEL")
                    If strLabel.Trim Like deskObject Then
                        Dim Fullnumber As String = FuncXRecords.FuncReadValueXData(acPolyCounter, "ConditionalNumber")
                        Dim numberParcels As Integer = Val(FuncFormatZn.FuncFindFullConditionalNumberCadastralObject(Fullnumber))
                        Dim numberCounter As Integer = Val(FuncFormatZn.FuncFindNumberCounterCadastralObject(Fullnumber))
                        ReDim Preserve arrayData(2, countParcels)
                        arrayData(0, countParcels) = numberParcels
                        arrayData(1, countParcels) = numberCounter
                        arrayData(2, countParcels) = acPolyCounter.ObjectID.ToString
                        countParcels += 1
                    End If
                End If
            Next
        End If
        If IsArray(arrayData) = True Then
            If arrayData.GetUpperBound(1) > 0 Then
                Dim boolSort As Boolean = MathFunction.FuncSortStrArray(arrayData, 0)
            End If
        End If
        Return countParcels
    End Function
    '===================================================================================================================================
    'функция сортирует массив с участками
    Public Shared Function FuncSortArrayParcels(ByRef GeoArray(,) As String) As Boolean
        FuncSortArrayParcels = False
        If IsArray(GeoArray) = False Then
            Return False
        End If
        If GeoArray.GetUpperBound(1) < 1 Then
            Return False
        End If
        If GeoArray.GetUpperBound(0) < 2 Then
            Return False
        End If
        Dim Temp1 As String() = Nothing
        Dim Temp2 As String() = Nothing
        Dim st1 As Integer = 0
        Dim st2 As Integer = 0
        Dim st3 As Integer = 0
        Dim st4 As Integer = 0
        Try
            For i As Integer = 0 To GeoArray.GetUpperBound(1) - 1 'берем первый элемет
                st1 = Val(GeoArray(0, i)) 'номер участка 1 строки
                st2 = Val(GeoArray(1, i)) 'номер контура 1 строки
                For j As Integer = i + 1 To GeoArray.GetUpperBound(1) 'берем второй элемент
                    'копируем вторую строку во временный массив
                    Erase Temp2
                    For k As Integer = 0 To GeoArray.GetUpperBound(0)
                        ReDim Preserve Temp2(k)
                        Temp2(k) = GeoArray(k, j)
                    Next k
                    st3 = Val(GeoArray(0, j)) 'номер участка 2 строки
                    st4 = Val(GeoArray(1, j)) 'номер контура 2 строки
                    'если номер участка 2 строки меньше чем номер участка 1 строки, меняем строки местами
                    If st3 < st1 Then
                        For k As Integer = 0 To Temp2.GetUpperBound(0)
                            GeoArray(k, j) = GeoArray(k, i)
                        Next k

                        For k As Integer = 0 To Temp2.GetUpperBound(0)
                            GeoArray(k, i) = Temp2(k)
                        Next k
                    ElseIf st3 = st1 Then
                        If st4 < st2 Then
                            For k As Integer = 0 To Temp2.GetUpperBound(0)
                                GeoArray(k, j) = GeoArray(k, i)
                            Next k

                            For k As Integer = 0 To Temp2.GetUpperBound(0)
                                GeoArray(k, i) = Temp2(k)
                            Next k
                        End If
                    End If
                Next j
            Next i
        Catch
            Return False
        End Try
        Return True
    End Function
    '==================================================================================================================================
    'функция ищет внутренние островки контура
    Public Shared Function FuncFindCadastralOstrov(ByVal ActivDocument As Drawing, ByVal acPolyParcel As DwgPolyline, ByVal localID As String, ByRef arrayData As String(,)) As Integer
        FuncFindCadastralOstrov = 0
        Try
            If IsNothing(acPolyParcel) = False And IsNothing(ActivDocument) = False And IsNothing(localID) = False Then
                Dim countOstrov As Integer = 0
                Dim acBound As BoundingBox2D = acPolyParcel.Bounds
                Dim minX As Double = acBound.Min.X
                Dim minY As Double = acBound.Min.Y

                Dim maxX As Double = acBound.Max.X
                Dim maxY As Double = acBound.Max.Y
                Dim listVector2d As List(Of Vector2D) = New List(Of Vector2D)
                Dim listVertPoly As List(Of BugleVector2D) = acPolyParcel.ToList
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
    Public Shared Function FuncFindCadastralCounter(ByVal ActivDocument As Drawing, ByVal numberParcel As String, ByRef arrayData As String(,)) As Integer
        FuncFindCadastralCounter = 0
        Dim countCounter As Integer = 0
        If IsNothing(numberParcel) = True Then Return 0
        If numberParcel.Trim.Length = 0 Then Return 0
        If IsNothing(ActivDocument) = False Then
            For Each acEnt As DwgEntity In ActivDocument.ActiveSpace.Entities
                If TypeOf (acEnt) Is DwgPolyline Then
                    Dim acPolyCounter As DwgPolyline = acEnt
                    Dim label As String = FuncXRecords.FuncReadValueXData(acPolyCounter, "LABEL")
                    If label.Trim Like "Внутренний островок" Then
                        Continue For
                    End If
                    Dim localID As String = FuncXRecords.FuncReadValueXData(acPolyCounter, "LOCALID")
                        If localID.Trim.Length > 1 Then
                            Dim Fullnumber As String = FuncXRecords.FuncReadValueXData(acPolyCounter, "ConditionalNumber")
                            Dim numberFindParcel As String = FuncFormatZn.FuncFindFullConditionalNumberCadastralObject(Fullnumber)
                            Dim numberCounter As String = FuncFormatZn.FuncFindNumberCounterCadastralObject(Fullnumber)
                            If numberFindParcel.Trim.Length > 0 Then
                                If numberFindParcel.Trim.Length Like numberParcel Then
                                    Dim sqrtCounter As Double = MathFunction.GetArea(acPolyCounter, 0)
                                    ReDim Preserve arrayData(3, countCounter)
                                    arrayData(0, countCounter) = numberCounter
                                    arrayData(1, countCounter) = acPolyCounter.ObjectID.ToString
                                    arrayData(2, countCounter) = Math.Abs(sqrtCounter)
                                    arrayData(3, countCounter) = localID
                                    countCounter += 1
                                End If
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


    '======================================================================================
    'функция возвращает еденицу измерения по коду
    Public Shared Function FuncFindCadastrUnitByCode(ByVal Unit As Integer) As String
        FuncFindCadastrUnitByCode = Unit.ToString
        Select Case Unit
            Case 3
                FuncFindCadastrUnitByCode = "мм"
            Case 4
                FuncFindCadastrUnitByCode = "см"
            Case 5
                FuncFindCadastrUnitByCode = "дм"
            Case 6
                FuncFindCadastrUnitByCode = "м"
            Case 8
                FuncFindCadastrUnitByCode = "км"
            Case 9
                FuncFindCadastrUnitByCode = "мгм"
            Case 47
                FuncFindCadastrUnitByCode = "мор.миля"
            Case 50
                FuncFindCadastrUnitByCode = "кв.мм"
            Case 51
                FuncFindCadastrUnitByCode = "кв.см"
            Case 53
                FuncFindCadastrUnitByCode = "кв.дм"
            Case 55
                FuncFindCadastrUnitByCode = "кв.м"
            Case 58
                FuncFindCadastrUnitByCode = "тыс.кв.м"
            Case 59
                FuncFindCadastrUnitByCode = "га"
            Case 61
                FuncFindCadastrUnitByCode = "кв.км"
            Case 109
                FuncFindCadastrUnitByCode = "Ар"
            Case 359
                FuncFindCadastrUnitByCode = "сут."
            Case 360
                FuncFindCadastrUnitByCode = "нед."
            Case 361
                FuncFindCadastrUnitByCode = "дек."
            Case 362
                FuncFindCadastrUnitByCode = "мес."
            Case 364
                FuncFindCadastrUnitByCode = "кв."
            Case 365
                FuncFindCadastrUnitByCode = "полугод."
            Case 356
                FuncFindCadastrUnitByCode = "час"
            Case 355
                FuncFindCadastrUnitByCode = "мин."
            Case 366
                FuncFindCadastrUnitByCode = "год"
            Case 383
                FuncFindCadastrUnitByCode = "руб."
            Case 384
                FuncFindCadastrUnitByCode = "тыс.руб."
            Case 385
                FuncFindCadastrUnitByCode = "млн.руб."
            Case 386
                FuncFindCadastrUnitByCode = "млрд.руб."
            Case 1000
                FuncFindCadastrUnitByCode = ""
            Case 1001
                FuncFindCadastrUnitByCode = ""
            Case 1002
                FuncFindCadastrUnitByCode = "руб/кв.м"
            Case 1003
                FuncFindCadastrUnitByCode = "руб/Ар"
            Case 1004
                FuncFindCadastrUnitByCode = "руб/га"
            Case 1005
                FuncFindCadastrUnitByCode = ""
            Case 1006
                FuncFindCadastrUnitByCode = "руб/кв.км"
            Case 1007
                FuncFindCadastrUnitByCode = "раб.дн."
        End Select
    End Function

    'создание полилинии из массива
    Public Shared Function FuncDrawPolylineToArrayCoord(ByVal ActivDocument As Drawing, ByRef ArrayVertex2d As String(,)) As DwgPolyline
        FuncDrawPolylineToArrayCoord = Nothing
        If IsArray(ArrayVertex2d) = True Then
            If ArrayVertex2d.GetUpperBound(1) > 0 Then
                Dim rp As Topomatic.Cad.Foundation.Vector2D = Nothing
                Dim polyline As DwgPolyline = New Topomatic.Dwg.Entities.DwgPolyline()
                For i As Integer = 0 To ArrayVertex2d.GetUpperBound(1)
                    rp.X = Val(ArrayVertex2d(3, i))
                    rp.Y = Val(ArrayVertex2d(2, i))
                    Dim pos = New Topomatic.Cad.Foundation.Vector2D(rp.X, rp.Y)
                    polyline.Add(New Topomatic.Cad.Foundation.BugleVector2D(pos, 0))
                Next i
                ActivDocument.ActiveSpace.Add(polyline)
                Return polyline
            End If
        End If
    End Function

End Class
