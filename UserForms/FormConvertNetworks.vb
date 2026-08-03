Imports Topomatic.Dwg

Public Class FormConvertNetworks
    Public boolShow As Boolean
    Public activDoc As Topomatic.Dwg.Drawing
    Public Sub New()

        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()

        ' Добавить код инициализации после вызова InitializeComponent().

    End Sub
    Private Sub FormConvertNetworks_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
    'отмена
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        boolShow = False
        Me.Close()
    End Sub
    'ОК
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        boolShow = True
        Me.Hide()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim userDraw As Topomatic.Dwg.Drawing = FuncSurface.FuncReturnDrawingToProjectByName(activDoc, ComboBox2.Text)
        Dim nameTypeLines As String() = {}
        Dim count As Integer = 0
        Dim strKode As String() = {"Электрокабель высогокого напряжения", "Электрокабель низкого напряжения", "Водопровод подземный", "Водопровод надземный", "Канализация ливневая", "Канализация бытовая", "Теплосеть подземная", "Теплосеть надземная", "Газопровод подземный", "Газопровод надземный", "Трубопровод спец. подземный", "Трубопровод спец. надземный", "Кабель связи в грунте", "Воздушный кабель связи", "Воздушная ЛЭП"}
        For i As Integer = 0 To userDraw.Linetypes.Count - 1
            Dim dwgLineTypeN As DwgLinetype = userDraw.Linetypes.Item(i)
            Dim nameTLine As String = dwgLineTypeN.Name
            DataGridView1.Rows.Add()
            DataGridView1.Rows(DataGridView1.Rows.Count - 2).Cells(1).Value = nameTLine
        Next
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub
End Class