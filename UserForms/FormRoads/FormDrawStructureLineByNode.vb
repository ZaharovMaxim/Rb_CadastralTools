Public Class FormDrawStructureLineByNode
    Public boolSelectSection As Boolean = False
    Public boolAllSelectSections As Boolean = False
    Public boolProgress As Boolean = False
    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub
    '//////////////////////////////////////////////////////////////////////////////////////////////////////
    'добавить код
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        boolSelectSection = True
        boolAllSelectSections = False
        boolProgress = False
        Me.Hide()
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'добавить все коды
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        boolSelectSection = False
        boolAllSelectSections = True
        boolProgress = False
        Me.Hide()
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'начать обработку
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        boolSelectSection = False
        boolAllSelectSections = False
        boolProgress = True
        Me.Hide()
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'отмена
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        boolSelectSection = False
        boolAllSelectSections = False
        boolProgress = False
        Me.Hide()
    End Sub

    Private Sub FormDrawStructureLineByNode_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

    End Sub
End Class