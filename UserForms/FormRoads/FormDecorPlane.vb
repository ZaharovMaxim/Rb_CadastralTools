Imports System.Drawing
Imports System.IO
Imports System.Windows.Media
Imports Topomatic.ApplicationPlatform
Imports Topomatic.ApplicationPlatform.Core
Imports Topomatic.ApplicationPlatform.Plugins

Public Class FormDecorPlane
    Public boolShow As Boolean = False
    Public generalDir As String = ""
    Public dirTemplate As String() = Nothing
    Public Sub New()
        InitializeComponent()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim arrayDirSupport As String() = Nothing
        Dim boolFindDirSupport As Boolean = FuncFiles.funcReadDirectoriesSupport(arrayDirSupport)
        If IsArray(arrayDirSupport) = True Then
            For i As Integer = 0 To arrayDirSupport.Length - 1
                Dim nameFolder As String = arrayDirSupport(i)
                Dim pos As Integer = nameFolder.IndexOf("Civil3DToolsUtility")
                If pos > -1 Then
                    generalDir = Mid(nameFolder, 1, pos)
                    generalDir = generalDir & "Civil3DToolsUtility"
                    Exit For
                End If
            Next
        End If
        'шаблон оформления
        Dim directorySupport As String = generalDir & "\FileResources\Sample\"
        If Directory.Exists(directorySupport) = True Then
            Dim xmlFiles As String() = Directory.GetFiles(directorySupport, "*.xml")
            If IsArray(xmlFiles) = True Then
                For i As Integer = 0 To xmlFiles.Length - 1
                    Dim tempXml As String = xmlFiles(i)
                    ReDim Preserve dirTemplate(i)
                    dirTemplate(i) = IO.Path.GetFileNameWithoutExtension(tempXml)
                Next
            End If
        End If
        If IsArray(dirTemplate) = True Then
            ComboBox3.DataSource = dirTemplate
            ComboBox3.Tag = directorySupport
        End If
    End Sub
    Private Sub FormDecorPlane_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
    'изменение категории условных знаков
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim nameCategory As String = ComboBox2.Text
        Dim templateXml As String = ComboBox3.Tag & "\" & ComboBox3.Text & ".xml"
        If IsNothing(nameCategory) = False Then
            If nameCategory.Trim.Length > 0 Then
                If IsNothing(templateXml) = False Then
                    If File.Exists(templateXml) = True Then
                        Dim arrayTables As String() = Nothing
                        Dim boolTables As Boolean = FuncXML.FuncReadTablesXMLByCategory(templateXml, nameCategory, arrayTables)
                        If IsArray(arrayTables) = True Then
                            CheckedListBox1.DataSource = arrayTables
                        End If
                    End If
                End If
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'отмена
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        boolShow = False
        Me.Hide()
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'OK
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        boolShow = True
        Me.Hide()
    End Sub

    Private Sub ВыбратьВсеToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ВыбратьВсеToolStripMenuItem.Click
        If CheckedListBox1.Items.Count > 0 Then
            For i As Integer = 0 To CheckedListBox1.Items.Count - 1
                CheckedListBox1.SetItemChecked(i, True)
            Next
        End If
    End Sub

    Private Sub ОтменитьВыборToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОтменитьВыборToolStripMenuItem.Click
        If CheckedListBox1.Items.Count > 0 Then
            For i As Integer = 0 To CheckedListBox1.Items.Count - 1
                CheckedListBox1.SetItemChecked(i, False)
            Next
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'создать новую модель
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim userProject As ModelProject = ApplicationHost.Current.ActiveProject
        Try
            If IsNothing(userProject) = False Then
                Dim folder As String = PluginCoreOps.FindModelPathId(PluginCoreOps.CreateFolder(New String() {"Модели", "Оформление плана", "План"}))
                Try
                    Dim iUserProject As IProjectModel = ApplicationHost.Current.Plugins.Execute("mkitem", New Object() {folder, "arr"})
                    ApplicationHost.Current.Plugins.Execute("activate", New Object() {iUserProject})
                    If IsNothing(iUserProject) = False Then
                        Dim nameNewProject As String = IO.Path.GetFileNameWithoutExtension(iUserProject.Uri.LastPathComponent)
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
        Catch ex As Topomatic.ApplicationPlatform.MessageException
            MsgBox(ex.Message)
        End Try
    End Sub
End Class