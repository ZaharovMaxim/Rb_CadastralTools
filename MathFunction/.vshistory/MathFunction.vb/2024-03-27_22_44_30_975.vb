Imports System.Windows.Media.Media3D
Imports System.Windows.Shapes
Imports Topomatic.Alg
Imports Topomatic.Alg.Offsets
Imports Topomatic.Cad.Foundation
Imports Topomatic.Cad.View
Imports Topomatic.Cad.View.Hints
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.Sfc

Public Class MathFunction
    '==============================================================================================================================================================
    'функция вычисляет прямоугольные координаты на основе ула и расстояния
    Public Shared Function funcCalcCoordinatesByInsPointAndAngle(ByVal InsPoint As Vector2D, ByVal Angle As Double, ByVal Lenght As Double, Optional roundZn As Integer = 3) As Vector2D
        funcCalcCoordinatesByInsPointAndAngle = New Vector2D(-1, -1)
        Try
            funcCalcCoordinatesByInsPointAndAngle.X = InsPoint.X + Lenght * Math.Cos(Angle)
            funcCalcCoordinatesByInsPointAndAngle.Y = InsPoint.Y + Lenght * Math.Sin(Angle)
            funcCalcCoordinatesByInsPointAndAngle.X = Math.Round(funcCalcCoordinatesByInsPointAndAngle.X, roundZn)
            funcCalcCoordinatesByInsPointAndAngle.Y = Math.Round(funcCalcCoordinatesByInsPointAndAngle.Y, roundZn)
        Catch ex As System.Exception
        End Try
    End Function
    '===============================================================================================================================================================
    'функфия вычисляет расстояние между двумя точками
    Public Shared Function funcCalcDistanceByToPoints2d(ByVal InsPoint1 As Vector2D, ByVal InsPoint2 As Vector2D, Optional ByVal RoundZn As Integer = 3) As Double
        funcCalcDistanceByToPoints2d = 0
        Try
            funcCalcDistanceByToPoints2d = Math.Sqrt((InsPoint1.X - InsPoint2.X) ^ 2 + (InsPoint1.Y - InsPoint2.Y) ^ 2)
            funcCalcDistanceByToPoints2d = Math.Round(funcCalcDistanceByToPoints2d, RoundZn)
        Catch ex As System.Exception
        End Try
    End Function
    '=================================================================================================================================================================
    'функция вычисляет координаты середины отрезка
    Public Shared Function funcCalcMiddleCoordByToPoints2d(ByVal InsPoint1 As Vector2D, ByVal InsPoint2 As Vector2D, Optional ByVal RoundZn As Integer = 3) As Vector2D
        funcCalcMiddleCoordByToPoints2d = New Vector2D(-1, -1)
        Try
            funcCalcMiddleCoordByToPoints2d = New Vector2D((InsPoint1.X + InsPoint2.X) / 2, (InsPoint1.Y + InsPoint2.Y) / 2)
            funcCalcMiddleCoordByToPoints2d.X = Math.Round(funcCalcMiddleCoordByToPoints2d.X, RoundZn)
            funcCalcMiddleCoordByToPoints2d.Y = Math.Round(funcCalcMiddleCoordByToPoints2d.Y, RoundZn)
        Catch ex As System.Exception
        End Try
    End Function
    '=================================================================================================================================================================
    'функция вычисляет математическое направление между 2 точками
    Public Shared Function funcCalcAngleByToPoints2d(ByVal InsPoint1 As Vector2D, ByVal InsPoint2 As Vector2D, Optional ByVal RoundZn As Integer = 6) As Double
        funcCalcAngleByToPoints2d = 0
        Try
            Dim dx As Double = InsPoint2.X - InsPoint1.X
            Dim dy As Double = InsPoint2.Y - InsPoint1.Y
            If dx = 0 And dy = 0 Then
                funcCalcAngleByToPoints2d = 0
            ElseIf dx > 0 And dy = 0 Then
                funcCalcAngleByToPoints2d = 0
            ElseIf dx = 0 And dy > 0 Then
                funcCalcAngleByToPoints2d = Math.PI / 2
            ElseIf dx < 0 And dy = 0 Then
                funcCalcAngleByToPoints2d = Math.PI
            ElseIf dx = 0 And dy < 0 Then
                funcCalcAngleByToPoints2d = (3 * Math.PI) / 2
            Else
                Dim r As Double = Math.Atan(dy / dx)
                If dx > 0 And dy > 0 Then
                    funcCalcAngleByToPoints2d = r
                ElseIf dx < 0 And dy > 0 Then
                    funcCalcAngleByToPoints2d = Math.PI / 2 + r
                ElseIf dx < 0 And dy < 0 Then
                    funcCalcAngleByToPoints2d = Math.PI + r
                ElseIf dx > 0 And dy < 0 Then
                    funcCalcAngleByToPoints2d = (3 * Math.PI) / 2 + r
                End If
            End If
            funcCalcAngleByToPoints2d = Math.Round(funcCalcAngleByToPoints2d, RoundZn)
        Catch ex As System.Exception
        End Try
    End Function
    '=================================================================================================================================================================
    'функция определяет с какой стороны находится точка (слева -1, справа 1, на линии - 0, ошибка -2)
    Public Shared Function funcLeftOrRightPointToLinearObject(ByVal acEnt As Object, ByVal InsPoint As Vector2D, Optional RoundZn As Integer = 2) As Integer
        funcLeftOrRightPointToLinearObject = -2
        Try
            Dim linear_object As ILinearObject = acEnt
            Dim polyLine As Polyline3D = New Polyline3D()
            linear_object.GetPolyline(polyLine)
            Dim station As Double = -1
            Dim offset As Double = -1
            polyLine.PosToStaOffset(InsPoint, station, offset)
            offset = Math.Round(offset, RoundZn)
            If offset = 0 Then
                Return 0
            ElseIf offset < 0 Then
                Return -1
            ElseIf offset > 0 Then
                Return 1
            End If
        Catch ex As System.Exception
        End Try
    End Function
    '==================================================================================================================================================================
    'функция сортирует двумерный массив
    '=====================================================================================
    'Функция сортирует числовой двухмерный массив по выбранному элементу (для любой размерности)
    Public Shared Function FuncSortDblArray(ByRef GeoArray(,) As Double, ByVal ind As Integer) As Boolean
        Dim Temp1 As Double() = Nothing
        Dim Temp2 As Double() = Nothing
        Dim st1 As Double = Nothing
        Dim st2 As Double = Nothing
        If IsArray(GeoArray) = False Then
            FuncSortDblArray = False
            Exit Function
        End If
        If GeoArray.GetUpperBound(1) < 1 Then
            FuncSortDblArray = False
            Exit Function
        End If
        If ind > GeoArray.GetUpperBound(1) Then
            FuncSortDblArray = False
            Exit Function
        End If
        Try
            For i As Integer = 0 To GeoArray.GetUpperBound(1) - 1 'берем первый элемет
                For j As Integer = i + 1 To GeoArray.GetUpperBound(1) 'берем второй элемент
                    'копируем из него первую строку во временный массив
                    Erase Temp2
                    For k As Integer = 0 To GeoArray.GetUpperBound(0)
                        ReDim Preserve Temp2(k)
                        Temp2(k) = GeoArray(k, j)
                    Next k
                    st1 = GeoArray(ind, i)
                    st2 = Temp2(ind)
                    If st2 < st1 Then
                        For k As Integer = 0 To Temp2.GetUpperBound(0)
                            GeoArray(k, j) = GeoArray(k, i)
                        Next k

                        For k As Integer = 0 To Temp2.GetUpperBound(0)
                            GeoArray(k, i) = Temp2(k)
                        Next k
                    End If
                Next j
            Next i
        Catch
            FuncSortDblArray = False
        End Try
        FuncSortDblArray = True
    End Function

    '=====================================================================================
    'Функция сортирует строковый двухмерный массив по выбранному элементу (для любой размерности)
    Public Shared Function FuncSortStrArray(ByRef GeoArray(,) As String, ByVal ind As Integer) As Boolean
        FuncSortStrArray = False
        If IsArray(GeoArray) = False Then
            FuncSortStrArray = False
            Exit Function
        End If
        If GeoArray.GetUpperBound(1) < 1 Then
            FuncSortStrArray = False
            Exit Function
        End If
        If ind > GeoArray.GetUpperBound(0) Then
            FuncSortStrArray = False
            Exit Function
        End If
        Dim Temp1 As String() = Nothing
        Dim Temp2 As String() = Nothing
        Dim st1 As Double = Nothing
        Dim st2 As Double = Nothing
        Dim st3 As String = Nothing
        Dim st4 As String = Nothing
        Try
            For i As Integer = 0 To GeoArray.GetUpperBound(1) - 1 'берем первый элемет
                For j As Integer = i + 1 To GeoArray.GetUpperBound(1) 'берем второй элемент
                    'копируем из него первую строку во временный массив
                    Erase Temp2
                    For k As Integer = 0 To GeoArray.GetUpperBound(0)
                        ReDim Preserve Temp2(k)
                        Temp2(k) = GeoArray(k, j)
                    Next k
                    If IsNumeric(GeoArray(ind, i)) And IsNumeric(Temp2(ind)) Then
                        st1 = Val(GeoArray(ind, i))
                        st2 = Val(Temp2(ind))
                        If st2 < st1 Then
                            For k As Integer = 0 To Temp2.GetUpperBound(0)
                                GeoArray(k, j) = GeoArray(k, i)
                            Next k

                            For k As Integer = 0 To Temp2.GetUpperBound(0)
                                GeoArray(k, i) = Temp2(k)
                            Next k
                        End If
                    Else
                        st3 = GeoArray(ind, i)
                        st4 = Temp2(ind)
                        If st4 < st3 Then
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
            FuncSortStrArray = False
        End Try
        FuncSortStrArray = True
    End Function


    '=================================================================================================================================================================
    'функция находит точку пересечения перепендикуляра проведенного к линии
    Public Shared Function funcGetNearestPointOnLine(ByRef points As Double(,), ByVal pt As Topomatic.Cad.Foundation.Vector3D) As Topomatic.Cad.Foundation.Vector3D
        funcGetNearestPointOnLine = New Topomatic.Cad.Foundation.Vector3D(0, 0, 0)
        Try
            If points.GetUpperBound(1) > 1 Then
                For i As Integer = 0 To points.GetUpperBound(1) - 1
                    Dim x1 As Double = points(0, i)
                    Dim Y1 As Double = points(1, i)
                    Dim z1 As Double = points(2, i)

                    Dim x2 As Double = points(0, i + 1)
                    Dim Y2 As Double = points(1, i + 1)
                    Dim z2 As Double = points(2, i + 1)

                    If x1 = x2 And Y1 = Y2 Then
                        Continue For
                    End If
                    Dim k As Double = ((Y2 - Y1) * (pt.X - x1) - (x2 - x1) * (pt.Y - Y1)) / ((Y2 - Y1) ^ 2 + (x2 - x1) ^ 2)
                    Dim tempX As Double = pt.X - k * (Y2 - Y1)
                    Dim tempY As Double = pt.Y + k * (x2 - x1)
                    Dim tempZ As Double = (z1 + z2) / 2
                    Dim d1 As Double = Math.Sqrt((tempX - x1) ^ 2 + (tempY - Y1) ^ 2)
                    Dim d2 As Double = Math.Sqrt((tempX - x2) ^ 2 + (tempY - Y2) ^ 2)
                    Dim d0 As Double = Math.Sqrt((x1 - x2) ^ 2 + (Y1 - Y2) ^ 2)
                    If Math.Round(d1, 4) + Math.Round(d2, 4) = Math.Round(d0, 4) Then
                        Dim newPoint As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(tempX, tempY, tempZ)
                        Return newPoint
                    End If
                Next i
            End If
        Catch ex As System.Exception
        End Try
    End Function
    '===============================================================================================
    'функция вычисляет точку пересечения двух отрезков
    Public Shared Function FuncIntersectPolyline(ByVal Pt1 As Vector2D, ByVal Pt2 As Vector2D, ByVal Pt3 As Vector2D, ByVal pt4 As Vector2D, ByRef InPoint As Vector2D) As Boolean
        FuncIntersectPolyline = False
        Dim x1 As Double = Pt1.X
        Dim y1 As Double = Pt1.Y

        Dim x2 As Double = Pt2.X
        Dim y2 As Double = Pt2.Y

        Dim x3 As Double = Pt3.X
        Dim y3 As Double = Pt3.Y

        Dim x4 As Double = pt4.X
        Dim y4 As Double = pt4.Y
        Try
            Dim d As Double = (Pt1.X - Pt2.X) * (pt4.Y - Pt3.Y) - (Pt1.Y - Pt2.Y) * (pt4.X - Pt3.X)
            Dim da As Double = (Pt1.X - Pt3.X) * (pt4.Y - Pt3.Y) - (Pt1.Y - Pt3.Y) * (pt4.X - Pt3.X)
            Dim db As Double = (Pt1.X - Pt2.X) * (Pt1.Y - Pt3.Y) - (Pt1.Y - Pt2.Y) * (Pt1.X - Pt3.X)

            Dim ta As Double = da / d
            Dim tb As Double = db / d

            If (ta >= 0 And ta <= 1 And tb >= 0 And tb <= 1) Then
                Dim dx As Double = Pt1.X + ta * (Pt2.X - Pt1.X)
                Dim dy As Double = Pt1.Y + ta * (Pt2.Y - Pt1.Y)
                InPoint = New Vector2D(dx, dy)
                FuncIntersectPolyline = True
            End If
        Catch ex As Exception
        End Try
    End Function
    '====================================================================================
    'функция проверяет лежит ли точка внутри полигона (Новый вариант)
    Public Shared Function FuncCrossPointPolyline(ByVal arrayPoint As Double(,), ByVal p As Vector2D) As Boolean
        Dim p1, p2 As Vector2D
        Dim inside As Boolean = False
        Dim poly As Double(,) = Nothing
        Dim countPoly As Integer = 0
        If IsNothing(arrayPoint) = True Then
            Return inside
        ElseIf arrayPoint.GetUpperBound(1) < 2 Then
            Return inside
        Else
            For i As Integer = 0 To arrayPoint.GetUpperBound(1)
                ReDim Preserve poly(1, i)
                poly(0, i) = arrayPoint(0, i)
                poly(1, i) = arrayPoint(1, i)
                countPoly += 1
            Next i
            If poly(0, 0) <> poly(0, poly.GetUpperBound(1)) And poly(1, 0) <> poly(1, poly.GetUpperBound(1)) Then
                ReDim Preserve poly(1, countPoly)
                poly(0, countPoly) = arrayPoint(0, 0)
                poly(1, countPoly) = arrayPoint(1, 0)
            End If
        End If
        '=========================================================================================
        Dim XOld As Double = poly(0, poly.GetUpperBound(1))
        Dim YOld As Double = poly(1, poly.GetUpperBound(1))
        Dim oldPoint As Vector2D = New Vector2D(XOld, YOld)
        For i As Integer = 0 To poly.GetUpperBound(1)
            Dim newPoint As Vector2D = New Vector2D(poly(0, i), poly(1, i))
            If newPoint.X > oldPoint.X Then
                p1 = oldPoint
                p2 = newPoint
            Else
                p1 = newPoint
                p2 = oldPoint
            End If
            If (newPoint.X < p.X) = (p.X <= oldPoint.X) AndAlso (p.Y - CLng(p1.Y)) * (p2.X - p1.X) < (p2.Y - CLng(p1.Y)) * (p.X - p1.X) Then
                inside = Not inside
            End If
            oldPoint = newPoint
        Next
        Return inside
    End Function
    '=================================================================================================
    'вычисление геометрического центра полигона
    Public Shared Function FuncFindCrossCenterPointPolyline(ByVal ArrayCoord As Double(,)) As Vector2D
        FuncFindCrossCenterPointPolyline = New Vector2D(-1, -1)
        Try
            If IsArray(ArrayCoord) = True Then
                Return New Vector2D(-1, -1)
            End If
            If ArrayCoord.GetUpperBound(1) < 2 Then
                Return New Vector2D(-1, -1)
            End If
            '2-ищем габаритный контейнер
            Dim minX As Double = 999999999999
            Dim minY As Double = 999999999999
            Dim maxX As Double = -999999999999
            Dim maxY As Double = -999999999999
            For i As Integer = 0 To ArrayCoord.GetUpperBound(1)
                Dim x As Integer = ArrayCoord(0, i)
                Dim y As Integer = ArrayCoord(0, i)
                If x > maxX Then
                    maxX = x
                End If
                If x < minX Then
                    minX = x
                End If
                If y > maxY Then
                    maxY = y
                End If
                If y < minY Then
                    minY = y
                End If
            Next
            '3-проводим горизонтальную линию по середине габаритного контейнера (ГОРИЗОНТАЛЬНУЮ ИЛИ ВЕРТИКАЛЬНУЮ)
            Dim pt1 As Vector2D
            Dim pt2 As Vector2D
            Dim dx As Double = Math.Abs(maxX - minX)
            Dim dy As Double = Math.Abs(maxY - minY)
            If dx >= dy Then
                pt1 = New Vector2D(minX, (minY + maxY) / 2)
                pt2 = New Vector2D(maxX, (minY + maxY) / 2)
            Else
                pt1 = New Vector2D((minX + maxX) / 2, minY)
                pt2 = New Vector2D((minX + maxX) / 2, maxY)
            End If
            '3-заносим в массив все пересечения фигуры с этой горизонтальной линией
            Dim CountArrayPt As Integer = 0
            Dim ArrayPt As Double(,) = Nothing
            For i As Integer = 0 To ArrayCoord.GetUpperBound(1)
                Dim pt3 As Vector2D = New Vector2D(ArrayCoord(0, i), ArrayCoord(1, i))
                Dim pt4 As Vector2D = New Vector2D(ArrayCoord(0, i + 1), ArrayCoord(1, i + 1))
                If MathFunction.funcCalcDistanceByToPoints2d(pt3, pt4) > 0 Then
                    Dim IntersectPoint As Vector2D = New Vector2D(-1, -1)
                    Dim boolIntersectLine As Boolean = MathFunction.FuncIntersectPolyline(pt1, pt2, pt3, pt4, IntersectPoint)
                    If boolIntersectLine = True Then
                        ReDim Preserve ArrayPt(1, CountArrayPt)
                        ArrayPt(0, CountArrayPt) = IntersectPoint.X
                        ArrayPt(1, CountArrayPt) = IntersectPoint.Y
                        CountArrayPt += 1
                    End If
                End If
            Next
            '4-сортируем массив (в массиве должно быть по меньшей мере 2 точки)
            If ArrayPt.GetUpperBound(1) > 0 Then
                If ArrayPt.GetUpperBound(1) > 0 Then
                    If dx >= dy Then
                        Dim boolSortArray As Boolean = MathFunction.FuncSortDblArray(ArrayPt, 0)
                    Else
                        Dim boolSortArray As Boolean = MathFunction.FuncSortDblArray(ArrayPt, 1)
                    End If
                End If
            End If
            '5-точки должна быть максимально удалена от точек пересечения
            Dim DistInsert As Double = 0
            If ArrayPt.GetUpperBound(1) > 0 Then
                For j As Integer = 0 To ArrayPt.GetUpperBound(1) - 1
                    Dim xStart As Double = ArrayPt(0, j)
                    Dim yStart As Double = ArrayPt(1, j)
                    Dim xEnd As Double = ArrayPt(0, j + 1)
                    Dim yEnd As Double = ArrayPt(1, j + 1)
                    Dim ptStart As Vector2D = New Vector2D(xStart, yStart)
                    Dim ptEnd As Vector2D = New Vector2D(xEnd, yEnd)
                    Dim TempDist As Double = MathFunction.funcCalcDistanceByToPoints2d(ptStart, ptEnd)
                    'ищем среднюю точку
                    Dim FindCenterPoint As Vector2D = New Vector2D(-1, -1)
                    If dx >= dy Then
                        FindCenterPoint = New Vector2D((xStart + xEnd) / 2, yStart)
                    Else
                        FindCenterPoint = New Vector2D(xStart, (yStart + yEnd) / 2)
                    End If
                    Dim PointCrosPoligon As Boolean = MathFunction.FuncCrossPointPolyline(ArrayCoord, FindCenterPoint)
                    If PointCrosPoligon = True Then
                        If TempDist > DistInsert Then
                            FuncFindCrossCenterPointPolyline = New Vector2D(FindCenterPoint.X, FindCenterPoint.Y)
                            DistInsert = TempDist
                        End If
                    End If
                Next j
            End If
            '6-если неудача попробуем простой вариант
            If FuncFindCrossCenterPointPolyline.X = -1 And FuncFindCrossCenterPointPolyline.Y = -1 Then
                Dim FindCenterPoint As Vector2D = New Vector2D((maxX + minX) / 2, (maxY + minY) / 2)
                Dim PointCrosPoligon As Boolean = MathFunction.FuncCrossPointPolyline(ArrayCoord, FindCenterPoint)
                If PointCrosPoligon = True Then
                    FuncFindCrossCenterPointPolyline = New Vector2D(FindCenterPoint.X, FindCenterPoint.Y)
                End If
            End If
            If FuncFindCrossCenterPointPolyline.X = -1 And FuncFindCrossCenterPointPolyline.Y = -1 Then
                FuncFindCrossCenterPointPolyline = New Vector2D(ArrayCoord(0, 0), ArrayCoord(0, 1))
            End If
        Catch ex As System.Exception
        End Try
    End Function
    '=============================================================================================
    'функция вычисляет площадь полигона
    Public Shared Function FuncSQRTPolylineByArray(ByVal ArrayCoord As Double(,), Optional RoundVertZn As Integer = 2, Optional RoundAreatZn As Integer = 0) As Double
        FuncSQRTPolylineByArray = 0
        Dim X1 As Double = 0
        Dim Y1 As Double = 0
        Dim X2 As Double = 0
        Dim Y2 As Double = 0
        Dim PSqr As Double = 0
        If IsArray(ArrayCoord) = True Then
            If ArrayCoord.GetUpperBound(1) > 1 Then
                Dim Xish As Double = Math.Round(ArrayCoord(0, 0), RoundVertZn)
                Dim Yish As Double = Math.Round(ArrayCoord(1, 0), RoundVertZn)
                Dim count As Integer = ArrayCoord.GetUpperBound(1)
                For i As Integer = 0 To ArrayCoord.GetUpperBound(1) - 1
                    X1 = Math.Round(ArrayCoord(0, i), RoundVertZn)
                    Y1 = Math.Round(ArrayCoord(1, i), RoundVertZn)
                    X2 = Math.Round(ArrayCoord(0, i + 1), RoundVertZn)
                    Y2 = Math.Round(ArrayCoord(0, i + 1), RoundVertZn)
                    PSqr = PSqr + (Y1 + Y2) * (X1 - X2)
                Next
                If Not (Xish = ArrayCoord(0, count) And Yish = ArrayCoord(1, count)) Then
                    PSqr = PSqr + (Y2 + Yish) * (X2 - Xish)
                End If
                PSqr = Math.Abs(PSqr)
                FuncSQRTPolylineByArray = Math.Round(PSqr / 2, RoundAreatZn)
            End If
        End If
    End Function
    '=============================================================================================
    'функция вычисляет площадь полигона
    Public Shared Function FuncSQRTPolylineByList(ByVal pt2dList As List(Of Vector2D), Optional RoundVertZn As Integer = 2, Optional RoundAreatZn As Integer = 0) As Double
        FuncSQRTPolylineByList = 0
        Dim X1 As Double = 0
        Dim Y1 As Double = 0
        Dim X2 As Double = 0
        Dim Y2 As Double = 0
        Dim PSqr As Double = 0
        If pt2dList.Count > 2 Then
            Dim Xish As Double = Math.Round(pt2dList.Item(0).X, RoundVertZn)
            Dim Yish As Double = Math.Round(pt2dList.Item(0).Y, RoundVertZn)
            For i As Integer = 0 To pt2dList.Count - 2
                X1 = Math.Round(pt2dList.Item(i).X, RoundVertZn)
                Y1 = Math.Round(pt2dList.Item(i).Y, RoundVertZn)
                X2 = Math.Round(pt2dList.Item(i + 1).X, RoundVertZn)
                Y2 = Math.Round(pt2dList.Item(i + 1).Y, RoundVertZn)
                PSqr = PSqr + (Y1 + Y2) * (X1 - X2)
            Next
            If Not (Xish = pt2dList.Item(pt2dList.Count - 1).X And Yish = pt2dList.Item(pt2dList.Count - 1).Y) Then
                PSqr = PSqr + (Y2 + Yish) * (X2 - Xish)
            End If
            PSqr = Math.Abs(PSqr)
            FuncSQRTPolylineByList = Math.Round(PSqr / 2, RoundAreatZn)
        End If
    End Function

    'функция проверяет есть ли в массиве нужное значение
    Public Shared Function FuncFindValueToArray(ByVal str As String, ByVal posArray As Integer, ByVal arrayData As String(,)) As Boolean
        FuncFindValueToArray = False
        If IsArray(arrayData) = True Then
            If arrayData.GetUpperBound(0) >= posArray Then
                For i As Integer = 0 To arrayData.GetUpperBound(1)
                    Dim tempZn As String = arrayData(posArray, i)
                    If IsNothing(tempZn) = False Then
                        If tempZn.Trim Like str.Trim Then
                            Return True
                        End If
                    End If
                Next i
            End If
        End If
    End Function

End Class
