
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports Topomatic
Imports Topomatic.Alg
Imports Topomatic.FoundationClasses.Lisp


Public Class CoordinateTransform
    Const ro As Double = 206264.8062 ' Число угловых секунд в радиане
    Const rad As Double = 0.0174532925199433

    '1. функция переводит географические координаты (широта, долгота, геодезическая высота) в геоцентрические (X, Y, Z)
    '============================================================================================================
    Public Shared Function funcGeocentricalCoordMax(ByVal ArrayCoord As Double(), ByVal Param As Double(), ByRef ArrayRezultCoord As Double()) As Boolean
        funcGeocentricalCoordMax = False
        Dim B As Double = ArrayCoord(0) 'широта
        Dim L As Double = ArrayCoord(1) 'долгота
        Dim H As Double = ArrayCoord(2) 'высота геодезическая

        Dim aEll As Double = Param(0) 'большая полуось
        Dim bEll As Double = Param(1) 'малая полуось

        Dim e As Double = (aEll ^ 2 - bEll ^ 2) / aEll ^ 2 'квадрат эксцентриситета эллипсоида
        Dim v As Double = aEll / (Math.Sqrt(1 - e * (Math.Sin(B * rad) ^ 2))) 'первый вертикальный радиус кривизны на широте

        Dim X As Double = (H + v) * Math.Cos(B * rad) * Math.Cos(L * rad)
        Dim Y As Double = (H + v) * Math.Cos(B * rad) * Math.Sin(L * rad)
        Dim Z As Double = ((1 - e) * v + H) * Math.Sin(B * rad)
        ReDim Preserve ArrayRezultCoord(2)
        ArrayRezultCoord(0) = X
        ArrayRezultCoord(1) = Y
        ArrayRezultCoord(2) = Z
        funcGeocentricalCoordMax = True
    End Function
    'функция переводит геоцентрические координаты (x, Y, Z) в географические  (широта, долгота, геодезическая высота)
    '==============================================================================================================================
    Public Shared Function funcGeograficCoordMax(ByVal ArrayCoord As Double(), ByVal Param As Double(), ByRef ArrayRezultCoord As Double()) As Boolean
        funcGeograficCoordMax = False
        Dim X As Double = ArrayCoord(0) 'X
        Dim Y As Double = ArrayCoord(1) 'Y
        Dim Z As Double = ArrayCoord(2) 'Z

        Dim aEll As Double = Param(0) 'большая полуось эллипсойда
        Dim bEll As Double = Param(1) 'малая полуось эллипсойда

        Dim e2 As Double = (aEll ^ 2 - bEll ^ 2) / aEll ^ 2 'квадрат эксцентриситета эллипсоида


        Dim e As Double = e2 / (1 - e2)

        Dim fp As Double = (aEll - bEll) / aEll
        Dim f As Double = 1 / fp

        Dim b As Double = aEll * (1 - fp)

        Dim p As Double = Math.Sqrt(X ^ 2 + Y ^ 2)

        Dim q As Double = Math.Atan2((Z * aEll), (p * b))

        Dim Lon As Double = Math.Atan2(Y, X)
        Dim Lat As Double = Math.Atan2((Z + e * b * (Math.Sin(q) ^ 3)), (p - e2 * aEll * (Math.Cos(q) ^ 3)))

        Dim v As Double = aEll / (Math.Sqrt(1 - e2 * (Math.Sin(Lat) ^ 2))) 'первый вертикальный радиус кривизны
        Dim h As Double = (p / Math.Cos(Lat)) - v

        ReDim Preserve ArrayRezultCoord(2)
        ArrayRezultCoord(0) = Lat * (1 / rad)
        ArrayRezultCoord(1) = Lon * (1 / rad)
        ArrayRezultCoord(2) = h
        funcGeograficCoordMax = True
    End Function
    '==========================================================================================================================
    'функция преобразует геоцентрические координаты из исходного эллипсойда в целевой (датум: по 7 параметров Гельмерта) 
    Public Shared Function VectorMatrixMultiplication(ByVal ArrayCoord As Double(), ByVal Param As Double(), ByRef ArrayRezultCoord As Double(), Optional ByVal k As Integer = 1) As Boolean
        VectorMatrixMultiplication = False
        If Param.Length >= 7 Then
            '3 параметра сдвига (k=-1 если переход обратный)
            Dim dx As Double = k * Param(0)
            Dim dy As Double = k * Param(1)
            Dim dz As Double = k * Param(2)
            'перевод угловых секуд в радианы
            Dim rx As Double = Param(3) * (Math.PI / (180 * 3600))
            Dim ry As Double = Param(4) * (Math.PI / (180 * 3600))
            Dim rz As Double = Param(5) * (Math.PI / (180 * 3600))
            rx = rx * k
            ry = ry * k
            rz = rz * k
            'масштабный коеффициент
            Dim m As Double = (1 + k * Param(6) * 10 ^ -6)

            'MatrixRotate(0, 0) = Math.Cos(ry) * Math.Cos(rz)
            'MatrixRotate(0, 1) = Math.Cos(rx) * Math.Sin(rz) + Math.Sin(rx) * Math.Sin(ry) * Math.Cos(rz)
            'MatrixRotate(0, 2) = Math.Sin(rx) * Math.Sin(rz) + Math.Cos(rx) * Math.Sin(ry) * Math.Cos(rz)

            'MatrixRotate(1, 0) = -1 * Math.Cos(ry) * Math.Sin(rz)
            'MatrixRotate(1, 1) = Math.Cos(rx) * Math.Cos(rz) - Math.Sin(rx) * Math.Sin(ry) * Math.Sin(rz)
            'MatrixRotate(1, 2) = Math.Sin(rx) * Math.Cos(rz) + Math.Cos(rx) * Math.Sin(ry) * Math.Sin(rz)

            'MatrixRotate(2, 0) = Math.Sin(ry)
            'MatrixRotate(2, 1) = -1 * Math.Sin(rx) * Math.Cos(ry)
            'MatrixRotate(2, 2) = Math.Cos(rx) * Math.Cos(ry)

            'заполняем матрицу разворота
            Dim MatrixRotate As Double(,) = Nothing
            ReDim Preserve MatrixRotate(2, 2)
            MatrixRotate(0, 0) = 1
            MatrixRotate(0, 1) = rz
            MatrixRotate(0, 2) = -1 * ry

            MatrixRotate(1, 0) = -1 * rz
            MatrixRotate(1, 1) = 1
            MatrixRotate(1, 2) = rx

            MatrixRotate(2, 0) = ry
            MatrixRotate(2, 1) = -1 * rx
            MatrixRotate(2, 2) = 1
            'умножаем матрицу разворота на вектор значений координат
            ReDim Preserve ArrayRezultCoord(2)
            For i As Integer = 0 To MatrixRotate.GetUpperBound(0)
                For j As Integer = 0 To MatrixRotate.GetUpperBound(0)
                    ArrayRezultCoord(i) += MatrixRotate(i, j) * ArrayCoord(j)
                Next j
            Next i
            'применяем масштабный коэффициент
            For i As Integer = 0 To ArrayRezultCoord.Length - 1
                ArrayRezultCoord(i) *= m
            Next i
            'добавляем смещение
            ArrayRezultCoord(0) = ArrayRezultCoord(0) + dx
            ArrayRezultCoord(1) = ArrayRezultCoord(1) + dy
            ArrayRezultCoord(2) = ArrayRezultCoord(2) + dz
        End If
        Return True
    End Function
    '===========================================================================================================================
    'функция преобразует географические координаты (широта, долгота) в топоцентрические координаты (x, y) в проекции Гауса-Крюгера (только для эллипсойде Красовского ГОСТ32453—2017). 
    Public Shared Function FuncLatLonKras_to_xyKras(ByVal ArrayCoord As Double(), ByVal Param As Double(), ByRef ArrayRezultCoord As Double()) As Boolean
        FuncLatLonKras_to_xyKras = False
        Dim B As Double = ArrayCoord(0) * rad
        Dim l As Double = (ArrayCoord(1) - Param(0)) * rad

        Dim x As Double = 6367558.4968 * B - Math.Sin(2 * B) * (16002.89 + 66.9607 * (Math.Sin(B)) ^ 2 + 0.3515 * (Math.Sin(B) ^ 4) -
        l ^ 2 * (1594561.25 + 5336.535 * (Math.Sin(B) ^ 2) + 26.79 * (Math.Sin(B) ^ 4) + 0.149 * (Math.Sin(B) ^ 6) +
        l ^ 2 * (672483.4 - 811219.9 * (Math.Sin(B) ^ 2) + 5420.0 * (Math.Sin(B) ^ 4) - 10.6 * (Math.Sin(B) ^ 6) +
        l ^ 2 * (278194 - 830174 * (Math.Sin(B) ^ 2) + 572434 * (Math.Sin(B) ^ 4) * -16010 * (Math.Sin(B) ^ 6) +
        l ^ 2 * (109500 - 574700 * (Math.Sin(B) ^ 2) + 863700 * (Math.Sin(B) ^ 4) - 398600 * (Math.Sin(B) ^ 6))))))

        Dim y As Double = l * Math.Cos(B) * (6378245 + 21346.1415 * (Math.Sin(B) ^ 2) + 107.159 * (Math.Sin(B) ^ 4) + 0.5977 * (Math.Sin(B) ^ 6 +
        l ^ 2 * (1070204.16 - 2136826.66 * (Math.Sin(B) ^ 2) + 17.98 * (Math.Sin(B) ^ 4) - 11.99 * (Math.Sin(B) ^ 6) +
        l ^ 2 * (270806 - 1523417 * (Math.Sin(B) ^ 2) + 1327645 * (Math.Sin(B) ^ 4) - 21701 * (Math.Sin(B) ^ 6) +
        l ^ 2 * (79690 - 866190 * (Math.Sin(B) ^ 2) + 1730360 * (Math.Sin(B) ^ 4) - 945460 * (Math.Sin(B) ^ 6))))))

        ReDim Preserve ArrayRezultCoord(2)
        ArrayRezultCoord(0) = x + Param(2)
        ArrayRezultCoord(1) = y + Param(3)
        ArrayRezultCoord(2) = ArrayRezultCoord(2)
        FuncLatLonKras_to_xyKras = True
    End Function
    '==============================================================================================================================
    'функция преобразует топоцентрические координаты (x, y) в проекции Гауса-Крюгера, в географические (широта, долгота) (только для эллипсойде Красовского ГОСТ32453—2017)
    Public Shared Function FuncxyKras_to_LatLonKras(ByVal ArrayCoord As Double(), ByVal Param As Double(), ByRef ArrayRezultCoord As Double()) As Boolean
        FuncxyKras_to_LatLonKras = False
        Dim x As Double = ArrayCoord(0) - Param(2)
        Dim y As Double = ArrayCoord(1) - Param(3)

        Dim b As Double = ArrayCoord(0) / 6367558.4968
        Dim B0 As Double = b + Math.Sin(2 * b) * (0.00252588685 - 0.0000149186 * (Math.Sin(b) ^ 2) + 0.00000011904 * (Math.Sin(b) ^ 4))
        Dim z0 As Double = y / (Math.Cos(B0) * 6378245)

        Dim dB As Double = -1 * z0 ^ 2 * Math.Sin(2 * B0) * (0.251684631 - 0.003369263 * (Math.Sin(B0) ^ 2) + 0.000011276 * (Math.Sin(B0) ^ 4) -
            z0 ^ 2 * (0.10500614 - 0.04559916 * (Math.Sin(B0) ^ 2) + 0.00228901 * (Math.Sin(B0) ^ 4) - 0.00002987 * (Math.Sin(B0) ^ 6) -
            z0 ^ 2 * (0.042858 - 0.025318 * (Math.Sin(B0) ^ 2) + 0.014346 * (Math.Sin(B0) ^ 4) - 0.001264 * (Math.Sin(B0) ^ 6) -
            z0 ^ 2 * (0.01672 - 0.0063 * (Math.Sin(B0) ^ 2) + 0.01188 * (Math.Sin(B0) ^ 4) - 0.00328 * (Math.Sin(B0) ^ 6)))))
        Dim Brad As Double = (B0 + dB) * (1 / rad)

        Dim l As Double = Param(0) * rad + z0 * (1 - 0.0033467108 * (Math.Sin(B0) ^ 2) - 0.0000056002 * (Math.Sin(B0) ^ 4) - 0.0000000187 * (Math.Sin(B0) ^ 6) -
        z0 ^ 2 * (0.16778975 + 0.16273586 * (Math.Sin(B0) ^ 2) - 0.0005249 * (Math.Sin(B0) ^ 4) - 0.00000846 * (Math.Sin(B0) ^ 6) -
        z0 ^ 2 * (0.0420025 + 0.1487407 * (Math.Sin(B0) ^ 2) + 0.005942 * (Math.Sin(B0) ^ 4) - 0.000015 * (Math.Sin(B0) ^ 6) -
        z0 ^ 2 * (0.01225 + 0.09477 * (Math.Sin(B0) ^ 2) + 0.03282 * (Math.Sin(B0) ^ 4) - 0.00034 * (Math.Sin(B0) ^ 6) -
        z0 ^ 2 * (0.0038 + 0.0524 * (Math.Sin(B0) ^ 2) + 0.0482 * (Math.Sin(B0) ^ 4) + 0.0032 * (Math.Sin(B0) ^ 6))))))

        Dim lrad As Double = l * (1 / rad)

        ReDim Preserve ArrayRezultCoord(2)
        ArrayRezultCoord(0) = Brad
        ArrayRezultCoord(1) = lrad
        ArrayRezultCoord(2) = ArrayRezultCoord(2)
        FuncxyKras_to_LatLonKras = True
    End Function

    '===========================================================================================================================
    'функция преобразует географические координаты (широта, долгота) в топоцентрические координаты (x, y) в проекции UTM 
    Public Shared Function FuncTraverseMercator_LatLon_to_xy(ByVal ArrayCoord As Double(), ByVal ParamTopo As Double(), ByVal ParamEll As Double(), ByRef ArrayRezultCoord As Double()) As Boolean
        FuncTraverseMercator_LatLon_to_xy = False
        'значения координат
        Dim Lat As Double = ArrayCoord(0) * rad
        Dim Lon As Double = ArrayCoord(1) * rad

        'исходные данные
        Dim Lon0 As Double = ParamTopo(0) * rad 'начальная долгота
        Dim Lat0 As Double = ParamTopo(1) * rad 'начальная широта
        Dim FN As Double = ParamTopo(2)
        Dim FE As Double = ParamTopo(3)
        Dim k0 As Double = ParamTopo(4)
        Dim a As Double = ParamEll(0) 'большая полуось
        Dim b As Double = ParamEll(1) 'малая полуось

        Dim f As Double = (a - b) / a
        Dim e As Double = Math.Sqrt(2 * f - f ^ 2)
        Dim e1 As Double = Math.Sqrt((e ^ 2 / (1 - e ^ 2)))

        Dim T As Double = Math.Tan(Lat) ^ 2
        Dim C As Double = ((e ^ 2) * Math.Cos(Lat) ^ 2) / (1 - e ^ 2)
        Dim AA As Double = (Lon - Lon0) * Math.Cos(Lat)
        Dim v As Double = a / (Math.Sqrt(1 - (e ^ 2) * (Math.Sin(Lat)) ^ 2))

        Dim a1 As Double = 1 - (e ^ 2) / 4 - 3 * (e ^ 4) / 64 - 5 * (e ^ 6) / 256
        Dim a2 As Double = 3 * (e ^ 2) / 8 + 3 * (e ^ 4) / 32 + 45 * (e ^ 6) / 1024
        Dim a3 As Double = 15 * (e ^ 4) / 256 + 45 * (e ^ 6) / 1024
        Dim a4 As Double = 35 * (e ^ 6) / 3072

        Dim M As Double = a * ((a1 * Lat - a2 * Math.Sin(2 * Lat) + a3 * Math.Sin(4 * Lat) - a4 * Math.Sin(6 * Lat)))
        Dim M0 As Double = 0

        Dim East As Double = FE + k0 * v * (AA + (1 - T + C) * (AA ^ 3) / 6 + (5 - 18 * T + (T ^ 2) + 72 * C - 58 * (e1 ^ 2)) * (AA ^ 5) / 120)
        Dim North As Double = FN + k0 * (M - M0 + v * Math.Tan(Lat) * ((AA ^ 2) / 2 + (5 - T + 9 * C + 4 * (C ^ 2)) * (AA ^ 4) / 24 + (61 - 58 * T + T ^ 2 + 600 * C - 330 * (e1 ^ 2)) * (AA ^ 6) / 720))

        ReDim Preserve ArrayRezultCoord(2)
        ArrayRezultCoord(0) = North
        ArrayRezultCoord(1) = East
        ArrayRezultCoord(2) = ArrayRezultCoord(2)
        FuncTraverseMercator_LatLon_to_xy = True
    End Function

    '===========================================================================================================================
    'функция преобразует топоцентрические координаты (x, y) в проекции UTM в географические координаты (широта, долгота)  
    Public Shared Function FuncTraverseMercator_xy_to_LatLon(ByVal ArrayCoord As Double(), ByVal ParamTopo As Double(), ByVal ParamEll As Double(), ByRef ArrayRezultCoord As Double()) As Boolean
        FuncTraverseMercator_xy_to_LatLon = False
        'значения координат
        Dim North As Double = ArrayCoord(0)
        Dim East As Double = ArrayCoord(1)

        'исходные данные
        Dim Lon0 As Double = ParamTopo(0) * rad 'начальная долгота
        Dim Lat0 As Double = ParamTopo(1) * rad 'начальная широта
        Dim FN As Double = ParamTopo(2)
        Dim FE As Double = ParamTopo(3)
        Dim k0 As Double = ParamTopo(4)
        Dim a As Double = ParamEll(0) 'большая полуось
        Dim b As Double = ParamEll(1) 'малая полуось

        Dim f As Double = (a - b) / a
        Dim e As Double = Math.Sqrt(2 * f - f ^ 2)
        Dim e1 As Double = Math.Sqrt((e ^ 2 / (1 - e ^ 2)))

        Dim M0 As Double = 0
        Dim e11 As Double = 1 - Math.Sqrt(1 - e ^ 2)
        Dim e12 As Double = 1 + Math.Sqrt(1 - e ^ 2)
        Dim e0 As Double = e11 / e12

        Dim M1 As Double = M0 + (North - FN) / k0
        Dim mu As Double = M1 / (a * (1 - (e ^ 2) / 4 - 3 * (e ^ 4) / 64 - 5 * (e ^ 6) / 256))

        Dim fi = mu + (3 * e0 / 2 - 27 * (e0 ^ 3) / 32) * Math.Sin(2 * mu) + (21 * (e0 ^ 2) / 16 - 55 * (e0 ^ 4) / 32) * Math.Sin(4 * mu) + (151 * (e0 ^ 3) / 96) * Math.Sin(6 * mu) + (1097 * (e0 ^ 4) / 512) * Math.Sin(8 * mu)

        Dim T1 As Double = Math.Tan(fi) ^ 2
        Dim C1 As Double = (e1 ^ 2) * (Math.Cos(fi) ^ 2) '2cos21
        Dim v1 As Double = a / (Math.Sqrt(1 - (e ^ 2) * (Math.Sin(fi)) ^ 2))
        Dim ro1 As Double = a * (1 - e ^ 2) / ((1 - (e ^ 2) * (Math.Sin(fi)) ^ 2)) ^ 1.5

        Dim D As Double = (East - FE) / (v1 * k0)

        Dim Lat As Double = fi - (v1 * Math.Tan(fi) / ro1) * ((D ^ 2) / 2 - (5 + 3 * T1 + 10 * C1 - 4 * C1 ^ 2 - 9 * (e1 ^ 2)) * (D ^ 4) / 24 + (61 + 90 * T1 + 298 * C1 + 45 * (T1 ^ 2) - 252 * (e1 ^ 2) - 3 * (C1 ^ 2)) * (D ^ 6) / 720)
        Dim Lon As Double = Lon0 + (D - (1 + 2 * T1 + C1) * (D ^ 3) / 6 + (5 - 2 * C1 + 28 * T1 - 3 * (C1 ^ 2) + 8 * (e1 ^ 2) + 24 * (T1 ^ 2)) * (D ^ 5) / 120) / Math.Cos(fi)

        ReDim Preserve ArrayRezultCoord(2)
        ArrayRezultCoord(0) = Lat * (1 / rad)
        ArrayRezultCoord(1) = Lon * (1 / rad)
        ArrayRezultCoord(2) = ArrayRezultCoord(2)
        FuncTraverseMercator_xy_to_LatLon = True
    End Function







    'не рабочая функция
    '===========================================================================================================================
    'функция преобразует географические координаты (широта, долгота) в топоцентрические координаты (x, y) в проекции Гауса-Крюгера (только для эллипсойде Красовского ГОСТ32453—2017). 
    Public Shared Function FuncTraverseMercaror_LatLon_to_xy(ByVal ArrayCoord As Double(), ByVal Param As Double(), ByRef ArrayRezultCoord As Double()) As Boolean
        FuncTraverseMercaror_LatLon_to_xy = False
        Dim Lat As Double = ArrayCoord(0) * rad
        Dim Lon As Double = ArrayCoord(1) * rad

        Dim Lat0 As Double = Param(1) * rad 'начальная широта
        Dim Lon0 As Double = Param(0) * rad 'начальная долгота

        Dim a As Double = 6378137 'большая полуось
        Dim b As Double = 6356752.314 'малая полуось
        Dim f As Double = (a - b) / a
        Dim e As Double = Math.Sqrt(2 * f - f ^ 2)
        Dim n As Double = f / (2 - f)
        Dim B0 As Double = (a / (1 + n)) * (1 + n ^ 2 / 4 + n ^ 4 / 64)

        Dim h1 As Double = n / 2 - (2 / 3) * n ^ 2 + (5 / 16) * n ^ 3 + (41 / 180) * n ^ 4
        Dim h2 As Double = (13 / 48) * n ^ 2 - (3 / 5) * n ^ 3 + (557 / 1440) * n ^ 4
        Dim h3 As Double = (61 / 240) * n ^ 3 - (103 / 140) * n ^ 4
        Dim h4 As Double = (49561 / 161280) * n ^ 4

        Dim M0 As Double = 0
        ' Log((1 + x) / (1 – x)) / 2
        Dim p1 As Double = e * Math.Sin(Lat)
        Dim p2 As Double = (1 + p1) / (1 - p1)
        Dim p3 As Double = (0.5 * Math.Log(p2))
        Dim p6 As Double = Math.Tan(Lat)
        Dim p7 = Math.Log(p6 + Math.Sqrt(p6 ^ 2 + 1))

        Dim Q As Double = p7 - (e * p3)
        Q = 1.279602466
        Dim be As Double = Math.Atan(Math.Sinh(Q))

        Dim p4 As Double = Math.Cos(Lat) * Math.Sin(Lon - Lon0)
        Dim p8 As Double = (1 + p4) / (1 - p4)
        Dim n0 As Double = 0.5 * Math.Log(p8)
        n0 = 1.226172663
        Dim t0 As Double = Math.Asin(Math.Sin(be) * Math.Cosh(n0))

        Dim t1 As Double = h1 * Math.Sin(2 * t0) * Math.Cosh(2 * n0)
        Dim t2 As Double = h2 * Math.Sin(4 * t0) * Math.Cosh(4 * n0)
        Dim t3 As Double = h3 * Math.Sin(6 * t0) * Math.Cosh(6 * n0)
        Dim t4 As Double = h4 * Math.Sin(8 * t0) * Math.Cosh(8 * n0)

        Dim t As Double = t0 + t1 + t2 + t3 + t4

        Dim n1 As Double = h1 * Math.Cos(2 * t0) * Math.Sinh(2 * n0)
        Dim n2 As Double = h2 * Math.Cos(4 * t0) * Math.Sinh(4 * n0)
        Dim n3 As Double = h3 * Math.Cos(6 * t0) * Math.Sinh(6 * n0)
        Dim n4 As Double = h4 * Math.Cos(8 * t0) * Math.Sinh(8 * n0)

        Dim nu As Double = n0 + n1 + n2 + n3 + n4

        'Dim kO As Double = Math.Cos(0) / Math.Sqrt(1 - e ^ 2 * Math.Sin(0))
        Dim kO As Double = 0.9997
        Dim x As Double = kO * B0 * nu
        Dim y As Double = kO * (B0 * t)

        ReDim Preserve ArrayRezultCoord(2)
        ArrayRezultCoord(0) = x
        ArrayRezultCoord(1) = y
        ArrayRezultCoord(2) = ArrayRezultCoord(2)
        FuncTraverseMercaror_LatLon_to_xy = True
    End Function

    ' Эллипсоид Красовского
    'Const aP As Double = 6378245 ' Большая полуось
    'Const alP As Double = 1 / 298.3 ' Сжатие
    'Const e2P As Double = 2 * alP - alP ^ 2 ' Квадрат эксцентриситета

    '' Эллипсоид WGS84 (GRS80, эти два эллипсоида сходны по большинству параметров)
    'Const aW As Double = 6378137 ' Большая полуось
    'Const alW As Double = 1 / 298.257223563 ' Сжатие
    'Const e2W As Double = 2 * alW - alW ^ 2 ' Квадрат эксцентриситета

    '' Вспомогательные значения для преобразования эллипсоидов
    'Const a As Double = (aP + aW) / 2
    'Const e2 As Double = (e2P + e2W) / 2
    'Const da As Double = aW - aP
    'Const de2 As Double = e2W - e2P

    '' Линейные элементы трансформирования, в метрах
    'Const dx As Double = 23.57
    'Const dy As Double = -140.95
    'Const dz As Double = -79.8

    '' Угловые элементы трансформирования, в секундах
    'Const wx As Double = 0
    'Const wy As Double = -0.35
    'Const wz As Double = -0.79

    '' Дифференциальное различие масштабов
    'Const ms As Double = -0.22


    'Public Shared Function WGS84_SK42_Lat(Bd, Ld, H) As Double
    '    WGS84_SK42_Lat = Bd - dB(Bd, Ld, H) / 3600
    'End Function

    'Public Shared Function SK42_WGS84_Lat(Bd, Ld, H) As Double
    '    SK42_WGS84_Lat = Bd + dB(Bd, Ld, H) / 3600
    'End Function

    'Public Shared Function WGS84_SK42_Long(Bd, Ld, H) As Double
    '    WGS84_SK42_Long = Ld - dL(Bd, Ld, H) / 3600
    'End Function

    'Public Shared Function SK42_WGS84_Long(Bd, Ld, H) As Double
    '    SK42_WGS84_Long = Ld + dL(Bd, Ld, H) / 3600
    'End Function

    'Public Shared Function dB(Bd, Ld, H) As Double
    '    Dim B, L, M, N As Double
    '    B = Bd * Math.PI / 180
    '    L = Ld * Math.PI / 180
    '    M = a * (1 - e2) / (1 - e2 * Math.Sin(B) ^ 2) ^ 1.5
    '    N = a * (1 - e2 * Math.Sin(B) ^ 2) ^ -0.5
    '    dB = ro / (M + H) * (N / a * e2 * Math.Sin(B) * Math.Cos(B) * da + (N ^ 2 / a ^ 2 + 1) * N * Math.Sin(B) * Math.Cos(B) * de2 / 2 - (dx * Math.Cos(L) + dy * Math.Sin(L)) * Math.Sin(B) + dz * Math.Cos(B)) - wx * Math.Sin(L) * (1 + e2 * Math.Cos(2 * B)) + wy * Math.Cos(L) * (1 + e2 * Math.Cos(2 * B)) - ro * ms * e2 * Math.Sin(B) * Math.Cos(B)
    'End Function

    'Public Shared Function dL(Bd, Ld, H) As Double
    '    Dim B, L, N As Double
    '    B = Bd * Math.PI / 180
    '    L = Ld * Math.PI / 180
    '    N = a * (1 - e2 * Math.Sin(B) ^ 2) ^ -0.5
    '    dL = ro / ((N + H) * Math.Cos(B)) * (-dx * Math.Sin(L) + dy * Math.Cos(L)) + Math.Tan(B) * (1 - e2) * (wx * Math.Cos(L) + wy * Math.Sin(L)) - wz
    'End Function

    'Public Shared Function WGS84Alt(Bd, Ld, H) As Double
    '    Dim B, L, N, dH As Double
    '    B = Bd * Math.PI / 180
    '    L = Ld * Math.PI / 180
    '    N = a * (1 - e2 * Math.Sin(B) ^ 2) ^ -0.5
    '    dH = -a / N * da + N * Math.Sin(B) ^ 2 * de2 / 2 + (dx * Math.Cos(L) + dy * Math.Sin(L)) * Math.Cos(B) + dz * Math.Sin(B) - N * e2 * Math.Sin(B) * Math.Cos(B) * (wx / ro * Math.Sin(L) - wy / ro * Math.Cos(L)) + (a ^ 2 / N + H) * ms
    '    WGS84Alt = H + dH
    'End Function




End Class
