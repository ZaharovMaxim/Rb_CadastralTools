

Public Class FormCreateSlopeLeader
    Public boolShow As Boolean = False
    Public Sub New()

        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()

        ' Добавить код инициализации после вызова InitializeComponent().

    End Sub

    'ОТМЕНА
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        boolShow = False
        Me.Hide()
    End Sub
    'ОК
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        boolShow = True
        Me.Hide()
    End Sub
End Class