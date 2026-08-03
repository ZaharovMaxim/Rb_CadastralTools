Imports System.Environment
Imports System.IO
Imports System.Windows.Forms
Imports Topomatic.ApplicationPlatform
Imports Topomatic.ApplicationPlatform.Core
Imports Topomatic.ApplicationPlatform.Plugins
Imports Topomatic.FoundationClasses
Public Class FormCreateCadastralPlan
    Public boolShow As Boolean = False
    Public PutchTemlate As String = ""

    Public Sub New()

        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'получем все модели
        Dim userProject As ModelProject = ApplicationHost.Current.ActiveProject
        '// Получение всех моделей проекта
        Dim childs As IProjectModel() = userProject.Model.GetChilds()
        Dim arrayNameModel As String() = Nothing
        Dim countArrayNameModel As Integer = 0
        For Each child As IProjectModel In childs
            Dim modelUri As URI = child.Uri
            If modelUri.Extension Like ".arrx" Then
                Dim pos As Integer = modelUri.AsAbsoluteUri.IndexOf("CadastralPlan")
                If pos > -1 Then
                    Dim fileNameSfcx As String = Path.GetFileNameWithoutExtension(modelUri.LastPathComponent)
                    ReDim Preserve arrayNameModel(countArrayNameModel)
                    arrayNameModel(countArrayNameModel) = fileNameSfcx
                    countArrayNameModel += 1
                End If
            End If
        Next
        If IsArray(arrayNameModel) = True Then
            ComboBox1.DataSource = arrayNameModel
        End If
        ComboBox1.DropDownStyle = ComboBoxStyle.DropDownList
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'отмена
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        boolShow = False
        Me.Hide()
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'OK
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If IsNothing(ComboBox1.Text) = False Then
            If ComboBox1.Text.Trim.Length > 0 Then
                boolShow = True
                Me.Hide()
            End If
        End If
    End Sub

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'создать новую модель
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim userProject As ModelProject = ApplicationHost.Current.ActiveProject
        If IsNothing(userProject) = False Then
            Dim folder As String = PluginCoreOps.FindModelPathId(PluginCoreOps.CreateFolder(New String() {"Модели", "LandCadastral", "CadastralPlan"}))
            Try
                Dim iUserProject As IProjectModel = ApplicationHost.Current.Plugins.Execute("mkitem", New Object() {folder, "arr"})
                ApplicationHost.Current.Plugins.Execute("activate", New Object() {iUserProject})
                If IsNothing(iUserProject) = False Then
                    Dim nameNewProject As String = Path.GetFileNameWithoutExtension(iUserProject.Uri.LastPathComponent)
                    Dim arrayModels As String() = ComboBox1.DataSource
                    Dim countArrayModels As Integer = 0
                    If IsArray(arrayModels) = True Then
                        countArrayModels = arrayModels.Length
                    End If
                    ReDim Preserve arrayModels(countArrayModels)
                    arrayModels(arrayModels.Length - 1) = nameNewProject
                    ComboBox1.DataSource = arrayModels
                    ComboBox1.Text = arrayModels(countArrayModels)
                End If
            Catch ex As System.OperationCanceledException
            End Try
        End If
    End Sub

    Private Sub FormCreateCadastralPlan_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class