Imports System.IO
Imports System.Windows
Imports System.Windows.Forms
Imports System.Xml
Imports Topomatic.ApplicationPlatform
Imports Topomatic.ApplicationPlatform.Core
Imports Topomatic.ApplicationPlatform.Plugins
Imports Topomatic.Cad.Foundation
Imports Topomatic.Cad.View
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.Dwg.Layer
Imports Topomatic.FoundationClasses
Imports Topomatic.Sfc

Public Class FormCreateParcels
    Public templateXML As String = ""
    Public templateDWG As String = ""
    Public generalDir As String = ""
    Public nameProgram As String = FuncFiles.namePrograms
    Public boolShow As Boolean = False
    Public arrayObjectProperties As String(,) = Nothing
    Private strCategoryTables As String = "Проект планировки и межевания территории"
    Public userCadView As CadView = Nothing

    Public Sub New()
        InitializeComponent()
        'категория объекта
        '1.\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        If nameProgram.Trim.Length > 0 Then
            Dim arrayDirSupport As String() = Nothing
            Dim boolFindDirSupport As Boolean = FuncFiles.funcReadDirectoriesSupport(arrayDirSupport)
            If IsArray(arrayDirSupport) = True Then
                For i As Integer = 0 To arrayDirSupport.Length - 1
                    Dim nameFolder As String = arrayDirSupport(i)
                    Dim pos As Integer = nameFolder.IndexOf(nameProgram)
                    If pos > -1 Then
                        generalDir = Mid(nameFolder, 1, pos)
                        generalDir = generalDir & nameProgram
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
                        templateXML = xmlFiles(i)
                        Dim shortNameFile As String = IO.Path.GetFileNameWithoutExtension(templateXML)
                        Dim tempNameFileDwg As String = directorySupport & shortNameFile & ".dwg"
                        If IO.File.Exists(tempNameFileDwg) = True Then
                            templateDWG = tempNameFileDwg
                            Exit For
                        End If
                    Next
                End If
            End If
            If IsNothing(templateXML) = True Or templateXML.Trim.Length = 0 Then
                MsgBox("Файл с шаблонами =Стандарт_ADSK.xml= не найден. Пожалуйста, добавьте его в пути доступа к вспомогательным файлам!!!")
                Exit Sub
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'получем все модели
            Dim Project As ModelProject = ApplicationHost.Current.ActiveProject
            '// Получение всех моделей проекта
            Dim childs As IProjectModel() = Project.Model.GetChilds()
            Dim arrayNameModel As String() = Nothing
            Dim countArrayNameModel As Integer = 0

            For Each child As IProjectModel In childs
                Dim modelUri As URI = child.Uri
                If modelUri.Extension Like ".arrx" Then
                    Dim fileNameSfcx As String = Path.GetFileNameWithoutExtension(modelUri.LastPathComponent)
                    ReDim Preserve arrayNameModel(countArrayNameModel)
                    arrayNameModel(countArrayNameModel) = fileNameSfcx
                    countArrayNameModel += 1
                End If
            Next
            If IsArray(arrayNameModel) = True Then
                ComboBox4.DataSource = arrayNameModel
            End If

            Dim arrayTables As String() = Nothing
            Dim boolFindTables As Boolean = FuncXML.FuncReadTablesXMLByCategory(templateXML, strCategoryTables, arrayTables)
            If IsArray(arrayTables) = True Then
                ComboBox1.DataSource = arrayTables
            Else
                MsgBox("В шаблоне xml в Категории =Проект планировки территории=, таблицы не найдены. Макро прервано!!!")
                Exit Sub
            End If
            Dim arrayObject As String() = Nothing
            Dim nameTable As String = ComboBox1.Text
            Dim boolFindObject As Boolean = FuncXML.FuncReadDeskObjectByCategoryAndTablesToXML(templateXML, strCategoryTables, nameTable, "Полилиния", arrayObject)
            If IsArray(arrayObject) = True Then
                ComboBox2.DataSource = arrayObject
            End If
            Dim deskObject As String = ComboBox2.Text
            Dim boolObjectProp As Boolean = FuncXML.FuncReadDataObjectToXMLbyDesk(templateXML, strCategoryTables, nameTable, "Полилиния", deskObject, arrayObjectProperties)
            If IsArray(arrayObjectProperties) = True Then
                Dim tableXML As String = ""
                For i As Integer = 0 To arrayObjectProperties.GetUpperBound(1)
                    Dim nameFields As String = arrayObjectProperties(0, i)
                    If nameFields Like "PS" Then
                        tableXML = arrayObjectProperties(1, i)
                        Exit For
                    End If
                Next i
                If tableXML.Trim.Length > 0 Then
                    Dim dictFieldsPS As Dictionary(Of String, String(,)) = New Dictionary(Of String, String(,))
                    Dim boolFindDictionary As Boolean = FuncXRecords.FuncReadDataFieldsByTablePSToXML(templateXML, tableXML, dictFieldsPS)
                    If dictFieldsPS.Count > 0 Then
                        DataGridView1.Rows.Clear()
                        For i As Integer = 0 To dictFieldsPS.Count - 1
                            Dim userKeyPair As KeyValuePair(Of String, String(,)) = dictFieldsPS.ElementAt(i)
                            Dim nameFields As String = userKeyPair.Key
                            Dim arrayValue As String(,) = userKeyPair.Value
                            Dim deskFields As String = ""
                            Dim valueData As String = ""
                            If IsArray(arrayValue) = True Then
                                For j As Integer = 0 To arrayValue.GetUpperBound(1)
                                    Dim nameVal As String = arrayValue(0, j)
                                    Dim strVal As String = arrayValue(1, j)
                                    If strVal.Trim.Length > 0 Then
                                        If nameVal Like "Description" Then
                                            deskFields = strVal
                                        End If
                                        If nameVal Like "DefaultData" Then
                                            valueData = strVal
                                        End If
                                    End If
                                Next j
                            End If
                            If nameFields Like "LOCALID" Then Continue For
                            If nameFields Like "GLOBALID" Then Continue For
                            DataGridView1.Rows.Add()
                            If i < DataGridView1.Rows.Count Then
                                DataGridView1.Rows(i).Cells(0).Value = deskFields
                                DataGridView1.Rows(i).Cells(0).Tag = nameFields
                                If nameFields Like "LABEL" Then
                                    DataGridView1.Rows(i).Cells(1).Value = deskObject
                                ElseIf nameFields Like "ConditionalNumber" Then
                                    DataGridView1.Rows(i).Cells(1).Value = NumericUpDown1.Value
                                Else
                                    DataGridView1.Rows(i).Cells(1).Value = valueData
                                End If
                            End If
                        Next i
                        DataGridView1.Tag = tableXML
                    End If
                End If
            End If
            '===========================================================================================================================
            'префикс номера
            Dim arrayPrefParcels As String() = {"00:00:0000000:0000:ЗУ", ":0000:ЗУ", "Участок №", "", "Публ.сервитут №", "Сервитут №"}
            'ComboBox3.DataSource = arrayPrefParcels
            '===========================================================================================================================
            If CheckBox1.Checked = False Then
                NumericUpDown2.Enabled = False
            Else
                NumericUpDown2.Enabled = True
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'отмена
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        boolShow = False
        Me.Close()
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'ОК
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        boolShow = True
        Me.Hide()
    End Sub
    '==================================================================================================================
    'изменить категорию объекта
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim nameTable As String = ComboBox1.Text
        Dim arrayObject As String() = Nothing
        Dim boolFindObject As Boolean = FuncXML.FuncReadDeskObjectByCategoryAndTablesToXML(templateXML, strCategoryTables, nameTable, "Полилиния", arrayObject)
        If IsArray(arrayObject) = True Then
            ComboBox2.DataSource = arrayObject
        End If
        Dim deskObject As String = ComboBox2.Text
        Erase arrayObjectProperties
        Dim boolObjectProp As Boolean = FuncXML.FuncReadDataObjectToXMLbyDesk(templateXML, strCategoryTables, nameTable, "Полилиния", deskObject, arrayObjectProperties)
        If IsArray(arrayObjectProperties) = True Then
            Dim tableXML As String = ""
            For i As Integer = 0 To arrayObjectProperties.GetUpperBound(1)
                Dim nameFields As String = arrayObjectProperties(0, i)
                If nameFields Like "PS" Then
                    tableXML = arrayObjectProperties(1, i)
                    Exit For
                End If
            Next i
            If tableXML.Trim.Length > 0 Then
                If Not (DataGridView1.Tag Like tableXML) Then
                    Dim dictFieldsPS As Dictionary(Of String, String(,)) = New Dictionary(Of String, String(,))
                    Dim boolFindDictionary As Boolean = FuncXRecords.FuncReadDataFieldsByTablePSToXML(templateXML, tableXML, dictFieldsPS)
                    If dictFieldsPS.Count > 0 Then
                        DataGridView1.Rows.Clear()
                        For i As Integer = 0 To dictFieldsPS.Count - 1
                            Dim userKeyPair As KeyValuePair(Of String, String(,)) = dictFieldsPS.ElementAt(i)
                            Dim nameFields As String = userKeyPair.Key
                            Dim arrayValue As String(,) = userKeyPair.Value
                            Dim deskFields As String = ""
                            Dim valueData As String = ""
                            If IsArray(arrayValue) = True Then
                                For j As Integer = 0 To arrayValue.GetUpperBound(1)
                                    Dim nameVal As String = arrayValue(0, j)
                                    Dim strVal As String = arrayValue(1, j)
                                    If strVal.Trim.Length > 0 Then
                                        If nameVal Like "Description" Then
                                            deskFields = strVal
                                        End If
                                        If nameVal Like "DefaultData" Then
                                            valueData = strVal
                                        End If
                                    End If
                                Next j
                            End If
                            If nameFields Like "LOCALID" Then Continue For
                            If nameFields Like "GLOBALID" Then Continue For
                            DataGridView1.Rows.Add()
                            If i < DataGridView1.Rows.Count Then
                                DataGridView1.Rows(i).Cells(0).Value = deskFields
                                DataGridView1.Rows(i).Cells(0).Tag = nameFields
                                If nameFields Like "LABEL" And (tableXML Like "PROJECT_PARCEL" Or tableXML Like "TEMPORARY_PARCEL") Then
                                    DataGridView1.Rows(i).Cells(1).Value = deskObject
                                ElseIf nameFields Like "Name" And (tableXML Like "PROJECT_LINE" Or tableXML Like "PROJECT_ZONA") Then
                                    DataGridView1.Rows(i).Cells(1).Value = deskObject
                                ElseIf nameFields Like "Type_line" And tableXML Like "URBAN_LINE" Then
                                    DataGridView1.Rows(i).Cells(1).Value = deskObject
                                Else
                                    DataGridView1.Rows(i).Cells(1).Value = valueData
                                End If
                            End If
                        Next i
                        DataGridView1.Tag = tableXML
                    End If
                Else
                    If DataGridView1.Rows.Count > 0 Then
                        For i As Integer = 0 To DataGridView1.RowCount - 1
                            Dim nameFields As String = DataGridView1.Rows(i).Cells(0).Value
                            If nameFields Like "LABEL" And (tableXML Like "PROJECT_PARCEL" Or tableXML Like "TEMPORARY_PARCEL") Then
                                DataGridView1.Rows(i).Cells(1).Value = deskObject
                            ElseIf nameFields Like "Name" And (tableXML Like "PROJECT_LINE" Or tableXML Like "PROJECT_ZONA") Then
                                DataGridView1.Rows(i).Cells(1).Value = deskObject
                            ElseIf nameFields Like "Type_line" And tableXML Like "URBAN_LINE" Then
                                DataGridView1.Rows(i).Cells(1).Value = deskObject
                            End If
                        Next i
                    End If
                End If
            End If
        End If
    End Sub
    '==================================================================================================================
    'изменить тип объекта
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim nameTable As String = ComboBox1.Text
        Dim deskObject As String = ComboBox2.Text
        Erase arrayObjectProperties
        Dim boolObjectProp As Boolean = FuncXML.FuncReadDataObjectToXMLbyDesk(templateXML, strCategoryTables, nameTable, "Полилиния", deskObject, arrayObjectProperties)
        If IsArray(arrayObjectProperties) = True Then
            Dim tableXML As String = ""
            For i As Integer = 0 To arrayObjectProperties.GetUpperBound(1)
                Dim nameFields As String = arrayObjectProperties(0, i)
                If nameFields Like "PS" Then
                    tableXML = arrayObjectProperties(1, i)
                    Exit For
                End If
            Next i
            If tableXML.Trim.Length > 0 Then
                If Not (DataGridView1.Tag Like tableXML) Then
                    Dim dictFieldsPS As Dictionary(Of String, String(,)) = New Dictionary(Of String, String(,))
                    Dim boolFindDictionary As Boolean = FuncXRecords.FuncReadDataFieldsByTablePSToXML(templateXML, tableXML, dictFieldsPS)
                    If dictFieldsPS.Count > 0 Then
                        DataGridView1.Rows.Clear()
                        For i As Integer = 0 To dictFieldsPS.Count - 1
                            Dim userKeyPair As KeyValuePair(Of String, String(,)) = dictFieldsPS.ElementAt(i)
                            Dim nameFields As String = userKeyPair.Key
                            Dim arrayValue As String(,) = userKeyPair.Value
                            Dim deskFields As String = ""
                            Dim valueData As String = ""
                            If IsArray(arrayValue) = True Then
                                For j As Integer = 0 To arrayValue.GetUpperBound(1)
                                    Dim nameVal As String = arrayValue(0, j)
                                    Dim strVal As String = arrayValue(1, j)
                                    If strVal.Trim.Length > 0 Then
                                        If nameVal Like "Description" Then
                                            deskFields = strVal
                                        End If
                                        If nameVal Like "DefaultData" Then
                                            valueData = strVal
                                        End If
                                    End If
                                Next j
                            End If
                            If nameFields Like "LOCALID" Then Continue For
                            If nameFields Like "GLOBALID" Then Continue For
                            DataGridView1.Rows.Add()
                            If i < DataGridView1.Rows.Count Then
                                DataGridView1.Rows(i).Cells(0).Value = deskFields
                                DataGridView1.Rows(i).Cells(0).Tag = nameFields
                                If nameFields Like "LABEL" And (tableXML Like "PROJECT_PARCEL" Or tableXML Like "TEMPORARY_PARCEL") Then
                                    DataGridView1.Rows(i).Cells(1).Value = deskObject
                                ElseIf nameFields Like "Name" And (tableXML Like "PROJECT_LINE" Or tableXML Like "PROJECT_ZONA") Then
                                    DataGridView1.Rows(i).Cells(1).Value = deskObject
                                ElseIf nameFields Like "Type_line" And tableXML Like "URBAN_LINE" Then
                                    DataGridView1.Rows(i).Cells(1).Value = deskObject
                                Else
                                    DataGridView1.Rows(i).Cells(1).Value = valueData
                                End If
                            End If
                        Next i
                        DataGridView1.Tag = tableXML
                    Else
                        If DataGridView1.Rows.Count > 0 Then
                            For i As Integer = 0 To DataGridView1.RowCount - 1
                                Dim nameFields As String = DataGridView1.Rows(i).Cells(0).Value
                                If nameFields Like "LABEL" And (tableXML Like "PROJECT_PARCEL" Or tableXML Like "TEMPORARY_PARCEL") Then
                                    DataGridView1.Rows(i).Cells(1).Value = deskObject
                                ElseIf nameFields Like "Name" And (tableXML Like "PROJECT_LINE" Or tableXML Like "PROJECT_ZONA") Then
                                    DataGridView1.Rows(i).Cells(1).Value = deskObject
                                ElseIf nameFields Like "Type_line" And tableXML Like "URBAN_LINE" Then
                                    DataGridView1.Rows(i).Cells(1).Value = deskObject
                                End If
                            Next i
                        End If
                    End If
                End If
            End If
        End If
    End Sub
    '==================================================================================================================
    'участок многоконтурный
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            NumericUpDown2.Value = 1
            NumericUpDown2.Enabled = True
            If DataGridView1.Rows.Count > 0 Then
                For i As Integer = 0 To DataGridView1.RowCount - 1
                    If DataGridView1.Rows(i).Cells(0).Value Like "ConditionalNumber" Then
                        Dim numbParcel As String = NumericUpDown1.Value
                        If NumericUpDown2.Value > 0 And CheckBox1.Checked = True Then
                            numbParcel = numbParcel & "(" & NumericUpDown2.Value & ")"
                        End If
                        DataGridView1.Rows(i).Cells(1).Value = numbParcel
                    End If
                Next i
            End If
        Else
            NumericUpDown2.Value = 0
            NumericUpDown2.Enabled = False
            If DataGridView1.Rows.Count > 0 Then
                For i As Integer = 0 To DataGridView1.RowCount - 1
                    If DataGridView1.Rows(i).Cells(0).Value Like "ConditionalNumber" Then
                        Dim numbParcel As String = NumericUpDown1.Value
                        If NumericUpDown2.Value > 0 And CheckBox1.Checked = True Then
                            numbParcel = numbParcel & "(" & NumericUpDown2.Value & ")"
                        End If
                        DataGridView1.Rows(i).Cells(1).Value = numbParcel
                    End If
                Next i
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'номер участка
    Private Sub NumericUpDown1_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown1.ValueChanged
        If DataGridView1.Rows.Count > 0 Then
            For i As Integer = 0 To DataGridView1.RowCount - 1
                If DataGridView1.Rows(i).Cells(0).Value Like "ConditionalNumber" Then
                    Dim numbParcel As String = NumericUpDown1.Value
                    If NumericUpDown2.Value > 0 And CheckBox1.Checked = True Then
                        numbParcel = numbParcel & "(" & NumericUpDown2.Value & ")"
                    End If
                    DataGridView1.Rows(i).Cells(1).Value = numbParcel
                End If
            Next i
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'номер контура
    Private Sub NumericUpDown2_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown2.ValueChanged
        If DataGridView1.Rows.Count > 0 Then
            For i As Integer = 0 To DataGridView1.RowCount - 1
                If DataGridView1.Rows(i).Cells(0).Value Like "ConditionalNumber" Then
                    Dim numbParcel As String = NumericUpDown1.Value
                    If NumericUpDown2.Value > 0 And CheckBox1.Checked = True Then
                        numbParcel = numbParcel & "(" & NumericUpDown2.Value & ")"
                    End If
                    DataGridView1.Rows(i).Cells(1).Value = numbParcel
                End If
            Next i
        End If
    End Sub

    Private Sub FormCreateParcels_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'найти последний номер участка
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim numberParcel As Integer = 0
        Dim strUserModel As String = ComboBox4.Text
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
    End Sub

    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        If (e.ColumnIndex = 0 AndAlso e.Value IsNot Nothing) Then
            Dim cell As DataGridViewCell = Me.DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex)
            Dim tagStr As String = cell.Tag
            Dim strVal As String = cell.Value
            If IsNothing(tagStr) = False Then
                cell.ToolTipText = tagStr
            Else
                cell.ToolTipText = strVal
            End If
        End If
    End Sub
    'взять значения по образцу
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If IsNothing(userCadView) = False Then
            userCadView.SelectionSet.Clear()
            'выбираем объекты
            userCadView.SelectionSet.SelectObjectsAtScreen(Function(obj) TypeOf obj Is DwgObject, "Выборите объект: ")
            For Each acEnt As Object In userCadView.SelectionSet
                Dim arrayData As String(,) = Nothing
                Dim namePS As String = DataGridView1.Tag
                Dim boolFindPSData As Boolean = FuncXRecords.FuncReadXData(acEnt, namePS, arrayData)
                If IsArray(arrayData) = True Then
                    If DataGridView1.RowCount > 1 Then
                        For i As Integer = 0 To DataGridView1.RowCount - 1
                            Dim nameField As String = DataGridView1.Rows(i).Cells(0).Value
                            If IsNothing(nameField) = True Then nameField = ""
                            For j As Integer = 0 To arrayData.GetUpperBound(1)
                                Dim nameArrayField As String = arrayData(0, j)
                                If IsNothing(nameArrayField) = True Then nameArrayField = ""
                                If nameArrayField Like nameField Then
                                    Dim strVal As String = arrayData(1, j)
                                    If IsNothing(strVal) = True Then strVal = ""
                                    DataGridView1.Rows(i).Cells(1).Value = strVal
                                    Exit For
                                End If
                            Next j
                        Next i
                    End If
                End If
            Next
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'создать новую модель
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim userProject As ModelProject = ApplicationHost.Current.ActiveProject
        If IsNothing(userProject) = False Then
            If IsNothing(DataGridView1.Tag) = False Then
                Dim nameFolderProject As String = DataGridView1.Tag
                If nameFolderProject.Trim.Length > 0 Then
                    Dim folder As String = PluginCoreOps.FindModelPathId(PluginCoreOps.CreateFolder(New String() {"Модели", "LandCadastral", nameFolderProject}))
                    Try
                        Dim iUserProject As IProjectModel = ApplicationHost.Current.Plugins.Execute("mkitem", New Object() {folder, "arr"})
                        ApplicationHost.Current.Plugins.Execute("activate", New Object() {iUserProject})
                        If IsNothing(iUserProject) = False Then
                            Dim nameNewProject As String = Path.GetFileNameWithoutExtension(iUserProject.Uri.LastPathComponent)
                            Dim arrayModels As String() = ComboBox4.DataSource
                            Dim countArrayModels As Integer = 0
                            If IsArray(arrayModels) = True Then
                                countArrayModels = arrayModels.Length
                            End If
                            ReDim Preserve arrayModels(countArrayModels)
                            arrayModels(arrayModels.Length - 1) = nameNewProject
                            ComboBox4.DataSource = arrayModels
                            ComboBox4.Text = arrayModels(countArrayModels)
                        End If
                    Catch ex As System.OperationCanceledException
                    End Try
                End If
            End If
        End If
    End Sub
End Class