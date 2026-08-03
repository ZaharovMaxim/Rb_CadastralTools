Imports System.IO
Imports System.Windows.Shapes
Imports System.Xml
Imports Topomatic.ApplicationPlatform
Imports Topomatic.ApplicationPlatform.Core
Imports Topomatic.ApplicationPlatform.Plugins
Imports Topomatic.Cad.View
Imports Topomatic.Dwg.Layer
Imports Topomatic.FoundationClasses

Public Class FormAutoCreateParcels
    Public templateXML As String = ""
    Public templateDWG As String = ""
    Public boolShow As Boolean = False
    Public arrayObjectProperties As String(,) = Nothing
    Private strCategoryTables As String = "Проект планировки и межевания территории"
    Public userCadView As CadView = Nothing
    Public arrayNameParcelsModel As String() = Nothing

    Public Sub New()

        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'проектная категория земель
        Dim arrayProjectCategory As String() = {"Земли сельскохозяйственного назначения", "Земли населенных пунктов", "Земли промышленности, энергетики, транспорта, связи, радиовещания, телевидения, информатики, земли для обеспечения космической деятельности, земли обороны, безопасности и земли иного специального назначения", "Земли особо охраняемых территорий и объектов", "Земли лесного фонда", "Земли водного фонда", "Земли запаса", "Категория не установлена"}
        ComboBox1.DataSource = arrayProjectCategory
        ComboBox1.Text = "Земли промышленности, энергетики, транспорта, связи, радиовещания, телевидения, информатики, земли для обеспечения космической деятельности, земли обороны, безопасности и земли иного специального назначения"
        'проектное разрешенное использование
        Dim arrayUtilization As String() = {"7 Транспорт", "7.1 Железнодорожный транспорт", "7.1.1 Железнодорожные пути", "7.1.2 Обслуживание железнодорожных перевозок", "7.2 Автомобильный транспорт", "7.2.1 Размещение автомобильных дорог", "7.2.2 Обслуживание перевозок пассажиров", "7.2.3 Стоянки транспорта общего пользования", "7.3 Водный транспорт", "7.4 Воздушный транспорт", "7.5	Трубопроводный транспорт", "7.6 Внеуличный транспорт", "8 Обеспечение обороны и безопасности"}
        ComboBox2.DataSource = arrayUtilization
        ComboBox2.Text = "7.2 Автомобильный транспорт"
        'проектное разрешенное использование в пределах населенного пункта
        Dim arrayUrbanUtilization As String() = {"12 Земельные участки (территории) общего пользования", "12.0.1 Улично-дорожная сеть", "12.0.2 Благоустройство территории"}
        ComboBox3.DataSource = arrayUrbanUtilization
        ComboBox3.Text = "12.0.1 Улично-дорожная сеть"
        'вид права
        Dim arrayRight As String() = {"Собственность", "Аренда", "Публичный сервитут", "Сервитут"}
        ComboBox4.DataSource = arrayRight
        ComboBox4.Text = "Собственность"
        '===============================================================================================================================
        'ищем все активные модели
        'категория объекта
        templateXML = FuncFiles.funcFindFileToDirectorySupport("Стандарт_ADSK.xml")
        templateDWG = FuncFiles.funcFindFileToDirectorySupport("Стандарт_ADSK.dwg")
        If IsNothing(templateXML) = True Then
            MsgBox("Файл с шаблонами =Стандарт_ADSK.xml= не найден. Пожалуйста, добавьте его в пути доступа к вспомогательным файлам!!!")
            Exit Sub
        ElseIf IO.File.Exists(templateXML) = False Then
            MsgBox("Файл с шаблонами =Стандарт_ADSK.xml= не найден. Пожалуйста, добавьте его в пути доступа к вспомогательным файлам!!!")
            Exit Sub
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'получем все модели
        Dim Project As ModelProject = ApplicationHost.Current.ActiveProject
        Dim childs As IProjectModel() = Project.Model.GetChilds()
        Dim arrayNameModel As String() = Nothing
        Dim arrayNameAlign As String() = Nothing

        Dim countArrayNameParcelsModel As Integer = 0
        Dim countArrayNameModel As Integer = 0
        Dim countArrayNameAlign As Integer = 0
        For Each child As IProjectModel In childs
            Dim modelUri As URI = child.Uri
            If modelUri.Extension Like ".arrx" Then
                Dim pos As Integer = modelUri.AsAbsoluteUri.IndexOf("LandCadastral")
                If pos > -1 Then
                    Dim fileNameSfcx As String = System.IO.Path.GetFileNameWithoutExtension(modelUri.LastPathComponent)
                    ReDim Preserve arrayNameModel(countArrayNameModel)
                    arrayNameModel(countArrayNameModel) = fileNameSfcx
                    countArrayNameModel += 1
                    ReDim Preserve arrayNameParcelsModel(countArrayNameParcelsModel)
                    arrayNameParcelsModel(countArrayNameParcelsModel) = fileNameSfcx
                    countArrayNameParcelsModel += 1
                End If
            ElseIf modelUri.Extension Like ".roadx" Then
                Dim fileNameSfcx As String = System.IO.Path.GetFileNameWithoutExtension(modelUri.LastPathComponent)
                ReDim Preserve arrayNameAlign(countArrayNameAlign)
                arrayNameAlign(countArrayNameAlign) = fileNameSfcx
                countArrayNameAlign += 1
            End If
        Next
        If IsArray(arrayNameModel) = True Then
            ComboBox6.DataSource = arrayNameModel
        End If
        If IsArray(arrayNameAlign) = True Then
            ComboBox9.DataSource = arrayNameAlign
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'получение земельных участков
        Dim arrayObject As String() = Nothing
        Dim nameTable As String = "Земельные участки образуемые"
        Dim boolFindObject As Boolean = FuncXML.FuncReadDeskObjectByCategoryAndTablesToXML(templateXML, strCategoryTables, nameTable, "Полилиния", arrayObject)
        If IsArray(arrayObject) = True Then
            ComboBox7.DataSource = arrayObject
        End If
        Dim deskObject As String = ComboBox7.Text
        Dim boolObjectProp As Boolean = FuncXML.FuncReadDataObjectToXMLbyDesk(templateXML, strCategoryTables, nameTable, "Полилиния", deskObject, arrayObjectProperties)
    End Sub
    Private Sub FormAutoCreateParcels_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            ComboBox5.DataSource = arrayNameParcelsModel
            ComboBox5.Enabled = True
        Else
            ComboBox5.Enabled = False
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'определить начальный номер
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If CheckBox2.Checked = True Then
            Dim strUserModel As String = ComboBox5.Text
            Dim DrawLayer As DrawingLayer = DrawingLayer.GetDrawingLayer(userCadView)
            Dim DrawShowLayers As IEnumerable(Of DrawingLayer) = DrawingLayer.GetDrawingLayers(userCadView)
            If IsNothing(strUserModel) = False Then
                If strUserModel.Trim.Length > 0 Then
                    For k1 As Integer = 0 To DrawShowLayers.Count - 1
                        Dim TempDrawLayer As DrawingLayer = DrawShowLayers.ElementAt(k1)
                        Dim nameUserFile As String = TempDrawLayer.Drawing.Filename
                        nameUserFile = IO.Path.GetFileNameWithoutExtension(nameUserFile)
                        If nameUserFile Like strUserModel Then
                            DrawLayer = TempDrawLayer
                            Exit For
                        End If
                    Next k1
                End If
            End If
            Dim arrayParcels As String(,) = Nothing
            Dim boolFindParcels As Integer = CadastralFunction.FuncFindCadastralParcels(DrawLayer.Drawing, ComboBox2.Text.Trim, arrayParcels)
            If IsArray(arrayParcels) = True Then
                If arrayParcels.GetUpperBound(1) > 0 Then
                    Dim boolSort As Boolean = MathFunction.FuncSortStrArray(arrayParcels, 0)
                End If
            End If
            Dim numParcels As Integer = arrayParcels(0, arrayParcels.GetUpperBound(1))
            If numParcels > 0 Then
                NumericUpDown1.Value = numParcels + 1
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'создать новую модель
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim userProject As ModelProject = ApplicationHost.Current.ActiveProject
        If IsNothing(userProject) = False Then
            Dim folder As String = PluginCoreOps.FindModelPathId(PluginCoreOps.CreateFolder(New String() {"Модели", "LandCadastral", "ProjectCadastralObject"}))
            Try
                Dim iUserProject As IProjectModel = ApplicationHost.Current.Plugins.Execute("mkitem", New Object() {folder, "arr"})
                ApplicationHost.Current.Plugins.Execute("activate", New Object() {iUserProject})
                If IsNothing(iUserProject) = False Then
                    Dim nameNewProject As String = IO.Path.GetFileNameWithoutExtension(iUserProject.Uri.LastPathComponent)
                    Dim arrayModels As String() = ComboBox4.DataSource
                    Dim countArrayModels As Integer = 0
                    If IsArray(arrayModels) = True Then
                        countArrayModels = arrayModels.Length
                    End If
                    ReDim Preserve arrayModels(countArrayModels)
                    arrayModels(arrayModels.Length - 1) = nameNewProject
                    ComboBox6.DataSource = arrayModels
                    ComboBox6.Text = arrayModels(countArrayModels)
                End If
            Catch ex As System.OperationCanceledException
            End Try
        End If
    End Sub
End Class