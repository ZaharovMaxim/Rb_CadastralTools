Imports System.Xml

Public Class FormImportParcels
    Public boolShow As Boolean = False
    Public templateXML As String = ""
    Public Sub New()

        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        Dim arrayNameArrib As String() = {"NUM_GEOPOINT", "NAME", "NUMBER"}
        ComboBox7.DataSource = arrayNameArrib

        Dim arrayRazd As String() = {"Пробел", ",", ";", "Табуляция"}
        ComboBox8.DataSource = arrayRazd
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'отмена
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        boolShow = False
        Me.Close()
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'ок
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        boolShow = True
        Me.Hide()
    End Sub
    'полигоны
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        Dim arrayCategory As String() = Nothing
        Dim arrayTables As String() = Nothing
        Dim arrayObject As String() = Nothing
        If CheckBox1.Checked = True Then
            If ComboBox1.Items.Count = 0 Then
                Dim boolFindCategory As Boolean = FuncXML.FuncReadCategoryToXML(templateXML, arrayCategory)
                If IsArray(arrayCategory) = True Then
                    ComboBox1.DataSource = arrayCategory
                    ComboBox1.Enabled = True
                    Dim nameCategory As String = ComboBox1.Text
                    If IsNothing(nameCategory) = False Then
                        If nameCategory.Trim.Length > 0 Then
                            Dim boolFindTables As Boolean = FuncXML.FuncReadTablesXMLByCategory(templateXML, nameCategory, arrayTables)
                            If IsArray(arrayTables) = True Then
                                ComboBox2.DataSource = arrayTables
                                ComboBox2.Enabled = True
                                Dim nameTable As String = ComboBox2.Text
                                If IsNothing(nameTable) = False Then
                                    If nameTable.Trim.Length > 0 Then
                                        Dim boolFindObject As Boolean = FuncXML.FuncReadDeskObjectByCategoryAndTablesToXML(templateXML, nameCategory, nameTable, "Полилиния", arrayObject)
                                        If IsArray(arrayObject) = True Then
                                            ComboBox3.DataSource = arrayObject
                                            ComboBox3.Enabled = True
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Else
            ComboBox1.Enabled = False
            ComboBox2.Enabled = False
            ComboBox3.Enabled = False
        End If
    End Sub
    'выбор категории
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim arrayTables As String() = Nothing
        Dim arrayObject As String() = Nothing
        Dim nameCategory As String = ComboBox1.Text
        If IsNothing(nameCategory) = False Then
            If nameCategory.Trim.Length > 0 Then
                Dim boolFindTables As Boolean = FuncXML.FuncReadTablesXMLByCategory(templateXML, nameCategory, arrayTables)
                If IsArray(arrayTables) = True Then
                    ComboBox2.DataSource = arrayTables
                    ComboBox2.Enabled = True
                    Dim nameTable As String = ComboBox2.Text
                    If IsNothing(nameTable) = False Then
                        If nameTable.Trim.Length > 0 Then
                            Dim boolFindObject As Boolean = FuncXML.FuncReadDeskObjectByCategoryAndTablesToXML(templateXML, nameCategory, nameTable, "Полилиния", arrayObject)
                            If IsArray(arrayObject) = True Then
                                ComboBox3.DataSource = arrayObject
                                ComboBox3.Enabled = True
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End Sub
    'выбор таблицы
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim arrayObject As String() = Nothing
        Dim nameCategory As String = ComboBox1.Text
        Dim nameTable As String = ComboBox2.Text
        If IsNothing(nameTable) = False Then
            If nameTable.Trim.Length > 0 Then
                Dim boolFindObject As Boolean = FuncXML.FuncReadDeskObjectByCategoryAndTablesToXML(templateXML, nameCategory, nameTable, "Полилиния", arrayObject)
                If IsArray(arrayObject) = True Then
                    ComboBox3.DataSource = arrayObject
                    ComboBox3.Enabled = True
                End If
            End If
        End If
    End Sub
    ' поворотная точка
    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        Dim arrayCategory As String() = Nothing
        Dim arrayTables As String() = Nothing
        Dim arrayObject As String() = Nothing
        If CheckBox2.Checked = True Then
            If ComboBox4.Items.Count = 0 Then
                Dim boolFindCategory As Boolean = FuncXML.FuncReadCategoryToXML(templateXML, arrayCategory)
                If IsArray(arrayCategory) = True Then
                    ComboBox4.DataSource = arrayCategory
                    ComboBox4.Enabled = True
                    Dim nameCategory As String = ComboBox4.Text
                    If IsNothing(nameCategory) = False Then
                        If nameCategory.Trim.Length > 0 Then
                            Dim boolFindTables As Boolean = FuncXML.FuncReadTablesXMLByCategory(templateXML, nameCategory, arrayTables)
                            If IsArray(arrayTables) = True Then
                                ComboBox5.DataSource = arrayTables
                                ComboBox5.Enabled = True
                                Dim nameTable As String = ComboBox5.Text
                                If IsNothing(nameTable) = False Then
                                    If nameTable.Trim.Length > 0 Then
                                        Dim boolFindObject As Boolean = FuncXML.FuncReadDeskObjectByCategoryAndTablesToXML(templateXML, nameCategory, nameTable, "Блок", arrayObject)
                                        If IsArray(arrayObject) = True Then
                                            ComboBox6.DataSource = arrayObject
                                            ComboBox6.Enabled = True
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Else
            ComboBox4.Enabled = False
            ComboBox5.Enabled = False
            ComboBox6.Enabled = False
        End If
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        Dim arrayTables As String() = Nothing
        Dim arrayObject As String() = Nothing
        Dim nameCategory As String = ComboBox4.Text
        If IsNothing(nameCategory) = False Then
            If nameCategory.Trim.Length > 0 Then
                Dim boolFindTables As Boolean = FuncXML.FuncReadTablesXMLByCategory(templateXML, nameCategory, arrayTables)
                If IsArray(arrayTables) = True Then
                    ComboBox5.DataSource = arrayTables
                    ComboBox5.Enabled = True
                    Dim nameTable As String = ComboBox5.Text
                    If IsNothing(nameTable) = False Then
                        If nameTable.Trim.Length > 0 Then
                            Dim boolFindObject As Boolean = FuncXML.FuncReadDeskObjectByCategoryAndTablesToXML(templateXML, nameCategory, nameTable, "Блок", arrayObject)
                            If IsArray(arrayObject) = True Then
                                ComboBox6.DataSource = arrayObject
                                ComboBox6.Enabled = True
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        Dim arrayObject As String() = Nothing
        Dim nameCategory As String = ComboBox4.Text
        Dim nameTable As String = ComboBox5.Text
        If IsNothing(nameTable) = False Then
            If nameTable.Trim.Length > 0 Then
                Dim boolFindObject As Boolean = FuncXML.FuncReadDeskObjectByCategoryAndTablesToXML(templateXML, nameCategory, nameTable, "Блок", arrayObject)
                If IsArray(arrayObject) = True Then
                    ComboBox6.DataSource = arrayObject
                    ComboBox6.Enabled = True
                End If
            End If
        End If
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged

    End Sub
End Class