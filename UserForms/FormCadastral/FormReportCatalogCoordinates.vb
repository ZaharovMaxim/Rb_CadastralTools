Imports Topomatic.Dwg.Layer
Imports Topomatic.Cad.View
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities

Public Class FormReportCatalogCoordinates
    Public boolRp As Boolean = False
    '=============================================================================================================================
    Public templateXML As String = ""
    Public templateDwg As String = ""
    Public ActivDocument As Topomatic.Dwg.Drawing = Nothing
    Public Sub New(ByRef cadWiever As CadView)
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        'получаем активный документ
        If IsNothing(cadWiever) Then Return
        'получаем активный слой
        Dim DrawLayer As DrawingLayer = DrawingLayer.GetDrawingLayer(cadWiever)
        If IsNothing(DrawLayer) = True Then Return
        'получаем активный документ
        ActivDocument = DrawLayer.Drawing
        If IsNothing(ActivDocument) = True Then Return
        'сброс выделения
        cadWiever.SelectionSet.Clear()
        Dim arrayPoint As String() = Nothing
        Dim boolPoint As Boolean = RoburFunc.FuncReadNameInsertBlocks(ActivDocument, arrayPoint)
        If IsArray(arrayPoint) = True Then
            ComboBox2.DataSource = arrayPoint
            For i As Integer = 0 To arrayPoint.Length - 1
                If arrayPoint(i) Like "PROJECT_POINT" Then
                    ComboBox2.Text = "PROJECT_POINT"
                    Exit For
                ElseIf arrayPoint(i) Like "TEMP_POINT" Then
                    ComboBox2.Text = "TEMP_POINT"
                    Exit For
                End If
            Next i
        End If
        '======================================================================================================================
        'загружаем блок
        Dim namePoint As String = ComboBox2.Text
        Dim blkPK1 As DwgBlock = ActivDocument.Blocks.Item(namePoint.Trim)
        'читаем атрибуты блока
        Dim arrayAttribute As String() = Nothing
        Dim boolReadBlk As Boolean = RoburFunc.FuncReadNameAttribBlk(ActivDocument, blkPK1, arrayAttribute)
        If IsArray(arrayAttribute) = True Then
            ComboBox3.DataSource = arrayAttribute
            For i As Integer = 0 To arrayAttribute.Length - 1
                Dim nameAttrib As String = arrayAttribute(i)
                If nameAttrib Like "NUM_GEOPOINT" Then
                    ComboBox3.Text = "NUM_GEOPOINT"
                    Exit For
                End If
            Next i
        Else
            ReDim arrayAttribute(0)
            arrayAttribute(0) = ""
            ComboBox3.DataSource = arrayAttribute
        End If
        Dim ArrayExl As String() = {"Таблица координат", "Документ MS Word", "Документ MS Excel", "Файл формата .csv"}
        ComboBox1.DataSource = ArrayExl
        ComboBox1.Text = "Документ MS Word"
        Dim ArrayPrefics As String() = {"00:00:0000000:0000:ЗУ", ":ЗУ", "Участок №", "Публичный сервитут №", "Сервитут №", "00:00:0000000:0000/чзу", "Часть участка №", "№", ""}
        ComboBox4.DataSource = ArrayPrefics
    End Sub
    'отмена
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        boolRp = False
        Me.Hide()
    End Sub
    'ок
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        boolRp = True
        Me.Hide()
    End Sub
    'выбран отчет в робур
    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            ComboBox1.Text = "Таблица координат"
        End If
    End Sub
    'выбран отчет в Excel
    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            ComboBox1.Text = "Документ MS Excel"
        End If
    End Sub
    'выбран отчет в Word
    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        If RadioButton3.Checked = True Then
            ComboBox1.Text = "Документ MS Word"
        End If
    End Sub
    'выбран отчет в MapInfo
    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
        If RadioButton4.Checked = True Then
            Dim arrayTablesPS As String() = Nothing
            Dim boolfindPs As Boolean = FuncXML.FuncReadPSTablesToXML(templateXML, arrayTablesPS)
            If IsArray(arrayTablesPS) = True Then
                ComboBox1.DataSource = arrayTablesPS
                For i As Integer = 0 To arrayTablesPS.Length - 1
                    If arrayTablesPS(i).Trim Like "PROJECT_PARCEL" Then
                        ComboBox1.Text = "PROJECT_PARCEL"
                    End If
                Next i
            End If
        End If
    End Sub
    'межевой план
    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton5.CheckedChanged
        If RadioButton5.Checked = True Then
            ComboBox1.Text = "Файл формата .csv"
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub FormReportCatalogCoordinates_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class