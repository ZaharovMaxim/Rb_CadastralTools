Imports Topomatic.ApplicationPlatform
Imports System.Windows.Forms
Imports Topomatic.ApplicationPlatform.Plugins
Imports Topomatic.Dwg.Layer
Imports Topomatic.Cad.View
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.Cad.Foundation
Imports Topomatic.Cad.View.Hints
Imports Topomatic.Alg.Prf
Imports Topomatic.Visualization
Imports Topomatic.ApplicationPlatform.Core
Imports System.IO
Imports Topomatic.FoundationClasses
Imports Topomatic.Alg
Imports Topomatic.Alg.Plan
Imports Microsoft.Office.Interop
Imports Topomatic.Acax.Export
Imports Topomatic
Imports System.Windows.Media.Media3D
Imports Path = System.IO.Path
Imports Vector3D = Topomatic.Cad.Foundation.Vector3D
Imports MS.Internal
Imports Topomatic.Visualization.Constructions
Imports Microsoft.Office.Interop.Excel
Imports NetTopologySuite.Geometries
Imports CadastralTools
Imports CadastralTools.StructureElement
Imports NetTopologySuite.LinearReferencing
Imports NetTopologySuite

Namespace RopExample1
    Partial Public Class RopExample1Module
        Inherits Topomatic.ApplicationPlatform.Plugins.PluginInitializator
        Public Overrides Sub Initialize(ByVal factory As PluginFactory)
            MyBase.Initialize(factory)
        End Sub
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'пользовательский элемент управления
        Public userControlPanelSymbolTOPO As PanelSymbolTOPO = Nothing
        Public userControlPanelPropertiesObject As PanelPropertiesObject = Nothing
        Public userControlPanelCadastralData As PanelProjectCadastral = Nothing
        Public timeModule As Date = New Date
        Public boolIns As String = "Error"

        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'установить программу
        <cmd("LoadProgram")>
        Public Sub LoadInfrastrada()
            '========================================================================================
            'загружаем пути к сборкам
            Dim myPath As String = FuncFiles.FuncGetAbsoluteFolderPathName("Выбор папки Civil3DToolsUtility")
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'прописываем пути к системным каталогами
            Dim patch1 As String = myPath & "\FileResources\Sample"
            Dim patch2 As String = myPath & "\FileResources\SHPForms"
            'Dim patch3 As String = myPath & "\TopomaticRobur\MenuTools"
            Dim bool1 As Boolean = FuncFiles.funcWriteDirectoryToDirectorySupport(patch1)
            Dim bool2 As Boolean = FuncFiles.funcWriteDirectoryToDirectorySupport(patch2)
            If bool2 And bool1 Then
                MsgBox("Пути к вспомогательным файлам поддержки успешно установлены!" & vbLf & patch1 & vbLf & patch2)
            End If
            'Dim bool3 As Boolean = FuncFiles.funcWriteDirectoryToDirectorySupport(patch3)
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'копируем шрифты
            Dim errFiles As String() = Nothing
            Dim countErrFiles As Integer = 0
            Dim userFontManager As FontManager = Topomatic.Cad.Foundation.FontManager.Current
            Dim defFileFont As String = userFontManager.DefaultFont.FilePath
            Dim folderSustemDir As String = Path.GetDirectoryName(defFileFont)
            Dim userDirFont As String = myPath & "\Fonts\"
            If System.IO.Directory.Exists(userDirFont) = True Then
                Dim fileEntries As String() = Directory.GetFiles(userDirFont)
                For Each fileName As String In fileEntries
                    Dim namefilefont As String = Path.GetFileName(fileName)
                    Dim extension As String = Path.GetExtension(fileName)
                    If extension Like ".ttf" Or extension Like ".shx" Or extension Like ".SHX" Then
                        Try
                            File.Copy(fileName, folderSustemDir & "\" & namefilefont, True)
                        Catch ex As System.UnauthorizedAccessException
                            ReDim Preserve errFiles(countErrFiles)
                            errFiles(countErrFiles) = fileName
                            countErrFiles += 1
                        Catch ex As System.ArgumentException
                            ReDim Preserve errFiles(countErrFiles)
                            errFiles(countErrFiles) = fileName
                            countErrFiles += 1
                        Catch ex As PathTooLongException
                            ReDim Preserve errFiles(countErrFiles)
                            errFiles(countErrFiles) = fileName
                            countErrFiles += 1
                        Catch ex As DirectoryNotFoundException
                            ReDim Preserve errFiles(countErrFiles)
                            errFiles(countErrFiles) = fileName
                            countErrFiles += 1
                        Catch ex As FileNotFoundException
                            ReDim Preserve errFiles(countErrFiles)
                            errFiles(countErrFiles) = fileName
                            countErrFiles += 1
                        Catch ex As IOException
                            ReDim Preserve errFiles(countErrFiles)
                            errFiles(countErrFiles) = fileName
                            countErrFiles += 1
                        Catch ex As NotSupportedException
                            ReDim Preserve errFiles(countErrFiles)
                            errFiles(countErrFiles) = fileName
                            countErrFiles += 1
                        End Try
                    End If

                Next
                If IsArray(errFiles) = True Then
                    Dim errMess As String = "Данные шрифты не были установлены в папку " & folderSustemDir
                    For i As Integer = 0 To errFiles.Length - 1
                        errMess = errMess & vbLf & errFiles(i)
                    Next i
                    MsgBox(errMess)
                End If
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'копируем формы SHP и SHX
            Dim arrayLookDirectories As String() = Nothing
            Dim boolFolders As Boolean = FuncFiles.funcReadDirectoriesSupport(arrayLookDirectories)
            If IsArray(arrayLookDirectories) = True Then
                For i As Integer = 0 To arrayLookDirectories.Length - 1
                    Erase errFiles
                    countErrFiles = 0
                    Dim tempdir As String = arrayLookDirectories(i)
                    If tempdir Like "C:\ProgramData\Topomatic\Robur road\16.0\Support" Or tempdir Like "C:\ProgramData\Topomatic\Robur road\16.0\Fonts" Then
                        If System.IO.Directory.Exists(tempdir) = True Then
                            Dim fileEntries As String() = Directory.GetFiles(patch2)
                            For Each fileName As String In fileEntries
                                Dim namefilefont As String = Path.GetFileName(fileName)
                                Dim extension As String = Path.GetExtension(fileName)
                                If extension Like ".shp" Or extension Like ".SHP" Or extension Like ".shx" Or extension Like ".SHX" Then
                                    Try
                                        File.Copy(fileName, folderSustemDir & "\" & namefilefont, True)
                                    Catch ex As System.UnauthorizedAccessException
                                        ReDim Preserve errFiles(countErrFiles)
                                        errFiles(countErrFiles) = fileName
                                        countErrFiles += 1
                                    Catch ex As System.ArgumentException
                                        ReDim Preserve errFiles(countErrFiles)
                                        errFiles(countErrFiles) = fileName
                                        countErrFiles += 1
                                    Catch ex As PathTooLongException
                                        ReDim Preserve errFiles(countErrFiles)
                                        errFiles(countErrFiles) = fileName
                                        countErrFiles += 1
                                    Catch ex As DirectoryNotFoundException
                                        ReDim Preserve errFiles(countErrFiles)
                                        errFiles(countErrFiles) = fileName
                                        countErrFiles += 1
                                    Catch ex As FileNotFoundException
                                        ReDim Preserve errFiles(countErrFiles)
                                        errFiles(countErrFiles) = fileName
                                        countErrFiles += 1
                                    Catch ex As IOException
                                        ReDim Preserve errFiles(countErrFiles)
                                        errFiles(countErrFiles) = fileName
                                        countErrFiles += 1
                                    Catch ex As NotSupportedException
                                        ReDim Preserve errFiles(countErrFiles)
                                        errFiles(countErrFiles) = fileName
                                        countErrFiles += 1
                                    End Try
                                End If
                            Next
                        End If
                    End If
                    If IsArray(errFiles) = True Then
                        Dim errMess As String = "Данные Формы не были установлены в папку " & tempdir
                        For j As Integer = 0 To errFiles.Length - 1
                            errMess = errMess & vbLf & errFiles(j)
                        Next j
                        MsgBox(errMess)
                    End If
                Next i
            End If
        End Sub

        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'создание панели условных знаков
        <cmd("ProjectCadastralSumbols")>
        Private Function CreatePanelSymbolTOPO() As System.Windows.Forms.Control
            If Not (ApplicationHost.Current.ActiveProject.ActiveDocument.UID = Consts.PlanWindow) Then
                MsgBox("Для активации палитры условных знаков необходимо переключиться в окно ПЛАН!!!")
                Exit Function
            End If
            Dim cadWiever As CadView = Me.CadView
            If IsNothing(cadWiever) Then
                Exit Function
            End If
            'получаем активный слой
            Dim DrawLayer As DrawingLayer = DrawingLayer.GetDrawingLayer(cadWiever)
            If IsNothing(DrawLayer) = True Then
                Exit Function
            End If
            Dim arrayDrawLayer As String() = Nothing
            Dim DrawShowLayers As IEnumerable(Of DrawingLayer) = DrawingLayer.GetDrawingLayers(cadWiever)
            If DrawShowLayers.Count > 0 Then
                For k1 As Integer = 0 To DrawShowLayers.Count - 1
                    Dim DrawLayerTemp As DrawingLayer = DrawShowLayers.ElementAt(k1)
                    Dim nameUserFile As String = DrawLayerTemp.Drawing.Filename
                    nameUserFile = System.IO.Path.GetFileNameWithoutExtension(nameUserFile)
                    ReDim Preserve arrayDrawLayer(k1)
                    arrayDrawLayer(k1) = nameUserFile
                Next k1
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            userControlPanelSymbolTOPO = New PanelSymbolTOPO()
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            If IsArray(arrayDrawLayer) = True Then
                userControlPanelSymbolTOPO.ComboBox1.DataSource = arrayDrawLayer
                userControlPanelSymbolTOPO.userCadView = cadWiever
            End If
            Dim M As Double = cadWiever.AnnotationScale
            userControlPanelSymbolTOPO.Label2.Text = "М 1:" & Math.Round(M * 1000, 0)
            userControlPanelSymbolTOPO.Label2.Tag = M
            Return userControlPanelSymbolTOPO
        End Function
        'обновление панели условных знаков
        <cmd("RefreshSymbolTOPO")>
        Private Function RefreshSymbolTOPO() As System.Windows.Forms.Control
            Dim cadWiever As CadView = Me.CadView
            If IsNothing(cadWiever) Then
                Exit Function
            End If
            If ApplicationHost.Current.ActiveProject.ActiveDocument.UID = Consts.PlanWindow Then
                Dim M As Double = cadWiever.AnnotationScale
                userControlPanelSymbolTOPO.Label2.Text = "М 1:" & Math.Round(M * 1000, 0)
                userControlPanelSymbolTOPO.Label2.Tag = M
            End If
            userControlPanelSymbolTOPO.userCadView = cadWiever
        End Function
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'панель пользовательских свойств
        <cmd("CreatePropertiesObject")>
        Private Function CreatePropertiesObject() As System.Windows.Forms.Control
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Dim cadWiever As CadView = Me.CadView
            If IsNothing(cadWiever) Then
                Exit Function
            End If
            If ApplicationHost.Current.ActiveProject.ActiveDocument.UID = Consts.PlanWindow Then
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                userControlPanelPropertiesObject = New PanelPropertiesObject()
                userControlPanelPropertiesObject.userCadView = cadWiever
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                Return userControlPanelPropertiesObject
            End If
        End Function
        'обновление панели пользовательских свойств
        <cmd("RefreshPropertiesObject")>
        Private Function RefreshPropertiesObject() As System.Windows.Forms.Control
            Dim cadWiever As CadView = Me.CadView
            If IsNothing(cadWiever) Then
                Exit Function
            End If
            If Not (ApplicationHost.Current.ActiveProject.ActiveDocument.UID = Consts.PlanWindow) Then
                Exit Function
            End If
            userControlPanelPropertiesObject.userCadView = cadWiever
            If cadWiever.SelectionSet.Count > 0 Then
                For Each acEnt As Object In cadWiever.SelectionSet
                    If TypeOf acEnt Is DwgEntity Then
                        Dim acObj As DwgEntity = acEnt
                        If Not (acObj.ObjectID Like userControlPanelPropertiesObject.ComboBox_DataTable.Tag) Then
                            userControlPanelPropertiesObject.BeginPropertiesGridUpdate()
                            Try
                                userControlPanelPropertiesObject.selectDwgObject = Nothing
                                userControlPanelPropertiesObject.dataStructureElement = Nothing
                                userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows.Clear()
                                userControlPanelPropertiesObject.ComboBox_DataTable.Tag = ""
                                userControlPanelPropertiesObject.ComboBox_DataTable.DataSource = Nothing
                                userControlPanelPropertiesObject.SetGlobalPropertiesMode()

                                Dim arrayTablesPS As String() = Nothing
                                Dim boolfindTables As Boolean = FuncXRecords.FuncReadTablesPS(acObj, arrayTablesPS)
                                If IsArray(arrayTablesPS) = True Then
                                    userControlPanelPropertiesObject.ComboBox_DataTable.DataSource = arrayTablesPS
                                    userControlPanelPropertiesObject.ComboBox_DataTable.Tag = acObj.ObjectID

                                    Dim projectStructuresIndex As Integer = -1
                                    For comboIndex As Integer = 0 To userControlPanelPropertiesObject.ComboBox_DataTable.Items.Count - 1
                                        Dim comboItem As Object = userControlPanelPropertiesObject.ComboBox_DataTable.Items(comboIndex)
                                        If comboItem IsNot Nothing AndAlso String.Equals(comboItem.ToString(), StructureElement.tableXRecords.PROJECT_STRUCTURES.ToString(), StringComparison.OrdinalIgnoreCase) Then
                                            projectStructuresIndex = comboIndex
                                            Exit For
                                        End If
                                    Next comboIndex

                                    If projectStructuresIndex > -1 Then
                                        userControlPanelPropertiesObject.ComboBox_DataTable.SelectedIndex = projectStructuresIndex
                                        Dim dataElement As StructureElement = New StructureElement
                                        Dim boolFindProperties As Boolean = FuncXRecords.getXRecords(acObj, dataElement)
                                        If boolFindProperties = True Then
                                            userControlPanelPropertiesObject.selectDwgObject = acObj
                                            userControlPanelPropertiesObject.dataStructureElement = dataElement
                                            userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows.Add(11)
                                            userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows(0).Cells(0).Value = "LABEL"
                                            userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows(0).Cells(1).Value = dataElement.Label
                                            userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows(1).Cells(0).Value = "NameClassObject"
                                            userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows(1).Cells(1).Value = dataElement.NameClassObject.ToString
                                            userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows(2).Cells(0).Value = "TypeGeometry"
                                            userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows(2).Cells(1).Value = dataElement.TypeGeometry.ToString
                                            userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows(3).Cells(0).Value = "NameTypeObject"
                                            userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows(3).Cells(1).Value = dataElement.NameTypeObject
                                            userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows(4).Cells(0).Value = "TableObject"
                                            userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows(4).Cells(1).Value = dataElement.TableObject
                                            userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows(5).Cells(0).Value = "Description"
                                            userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows(5).Cells(1).Value = dataElement.Description
                                            userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows(5).Cells(0).Value = "TablePS"
                                            userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows(5).Cells(1).Value = dataElement.TablePS
                                            userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows(6).Cells(0).Value = "KeyParameter"
                                            userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows(6).Cells(1).Value = dataElement.KeyParameter
                                            userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows(7).Cells(0).Value = "IdStructure"
                                            userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows(7).Cells(1).Value = dataElement.IdStructure
                                            userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows(8).Cells(0).Value = "IdElement"
                                            userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows(8).Cells(1).Value = dataElement.IdElement
                                            userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows(9).Cells(0).Value = "Note"
                                            userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows(9).Cells(1).Value = dataElement.Note
                                            userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows(10).Cells(0).Value = "IdObject"
                                            userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows(10).Cells(1).Value = acObj.ObjectID.ToString
                                        End If
                                    End If
                                End If
                            Finally
                                userControlPanelPropertiesObject.EndPropertiesGridUpdate()
                            End Try
                        End If
                    End If
                    Exit For
                Next
            Else
                userControlPanelPropertiesObject.BeginPropertiesGridUpdate()
                Try
                    userControlPanelPropertiesObject.selectDwgObject = Nothing
                    userControlPanelPropertiesObject.dataStructureElement = Nothing
                    userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows.Clear()
                    userControlPanelPropertiesObject.ComboBox_DataTable.Tag = ""
                    userControlPanelPropertiesObject.ComboBox_DataTable.DataSource = Nothing
                    userControlPanelPropertiesObject.SetGlobalPropertiesMode()
                Finally
                    userControlPanelPropertiesObject.EndPropertiesGridUpdate()
                End Try
            End If
            Return userControlPanelPropertiesObject

        End Function
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'панель редактирования кадастровых данных
        <cmd("CreatePropertiesCadastralData")>
        Private Function CreatePropertiesCadastralData() As System.Windows.Forms.Control
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Dim userCadView As CadView = Me.CadView
            If IsNothing(userCadView) Then
                Exit Function
            End If
            If ApplicationHost.Current.ActiveProject.ActiveDocument.UID = Consts.PlanWindow Then
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                userControlPanelCadastralData = New PanelProjectCadastral()
                userControlPanelCadastralData.cadViewPanel = userCadView
                'получаем активный слой
                Dim DrawLayer As DrawingLayer = DrawingLayer.GetDrawingLayer(userCadView)
                If IsNothing(DrawLayer) = True Then
                    MsgBox("Для запуска программы, пожалуйста сделайте активируйте модель!")
                    Return Nothing
                End If
                'получаем активный документ
                Dim ActivDocument As Topomatic.Dwg.Drawing = DrawLayer.Drawing
                If IsNothing(ActivDocument) = True Then
                    Return Nothing
                End If
                userControlPanelCadastralData.activDocumentPanel = ActivDocument
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                Dim namePrograms As String = FuncFiles.namePrograms
                Dim generalDirectoryPrograms As String = ""
                Dim putchTemplatesXML As String = ""
                Dim putchTemplateDwg As String = ""
                Dim boolFindSupportFile As Boolean = FuncFiles.FuncReadProperties(namePrograms, generalDirectoryPrograms, putchTemplatesXML, putchTemplateDwg)
                If generalDirectoryPrograms.Trim.Length = 0 Then
                    Return Nothing
                End If
                userControlPanelCadastralData.generalDirectoryPrograms = generalDirectoryPrograms
                userControlPanelCadastralData.putchTemplatesXML = putchTemplatesXML
                userControlPanelCadastralData.putchTemplateDwg = putchTemplateDwg
                Return userControlPanelCadastralData
            End If
        End Function
        <cmd("RefreshPanelCadastralData")>
        Private Function RefreshPanelCadastralData() As System.Windows.Forms.Control
            Try
                Dim userCadView As CadView = userControlPanelCadastralData.cadViewPanel
                If userCadView IsNot CadView Then
                    userControlPanelCadastralData.cadViewPanel = Me.CadView
                End If

                If IsNothing(ApplicationHost.Current.ActiveProject.ActiveDocument) = True Then
                    Exit Function
                End If
            Catch ex As System.Exception
            End Try
        End Function
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'редактировать XML
        <cmd("EditXMLSymbols")>
        Public Sub EditXMLSymbols()
            Dim cadWiever As CadView = Me.CadView
            If IsNothing(cadWiever) Then Return
            'получаем активный слой
            Dim DrawLayer As DrawingLayer = DrawingLayer.GetDrawingLayer(cadWiever)
            If IsNothing(DrawLayer) = True Then Return
            'получаем активный документ
            Dim ActivDocument As Topomatic.Dwg.Drawing = DrawLayer.Drawing
            If IsNothing(ActivDocument) = True Then Return
            'сброс выделения
            cadWiever.SelectionSet.Clear()
            Dim formEditXML As FormEditXMLSymbol = New FormEditXMLSymbol
            formEditXML.userCadView = cadWiever
Line1:
            Try
                formEditXML.ShowDialog()
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'выбор объекта
                If formEditXML.boolSelectObject = True Then
                    'сброс выделения
                    cadWiever.SelectionSet.Clear()
                    cadWiever.SelectionSet.SelectOneObjectAtScreen(Function(obj) TypeOf obj Is Object, "Выборите объект: ")
                    If cadWiever.SelectionSet.Count > 0 Then
                        Dim userNode As TreeNode = formEditXML.TreeView1.SelectedNode
                        Dim boolWriteData As Boolean = False
                        For Each row As DataGridViewRow In formEditXML.DataGridView1.Rows
                            If row.Cells(0).Value Like "Description" Then
                                If row.Cells(1).Value Like userNode.Text Then
                                    boolWriteData = True
                                    formEditXML.boolIzmDataGrid = False
                                    Exit For
                                End If
                            End If
                        Next
                        Dim arrayData As String(,) = userNode.Tag
                        Dim countArrayData As Integer = 1
                        Dim number As Integer = userNode.Parent.Nodes.Count + 1
                        Dim nameObject As String = ""
                        Dim nameLineType As String = ""
                        Dim nameLayer As String = ""
                        Dim userColor As Integer = 7
                        Dim userpattern As String = ""
                        Dim lineWidthUser As Double = 20
                        Dim lineScaleUser As Double = 1
                        Dim obj_code As String = ""
                        For Each acEnt As Object In cadWiever.SelectionSet
                            If TypeOf acEnt Is DwgEntity Then
                                Dim userCadEntity As DwgEntity = acEnt
                                nameLayer = userCadEntity.Layer.Name
                                nameLineType = userCadEntity.Linetype.Name
                                Dim lineTypeUser As DwgLinetype = userCadEntity.Linetype
                                userpattern = StyleCadastralObject.FuncReturnPatternLineTypeCode(lineTypeUser)
                                lineWidthUser = userCadEntity.Lineweight
                                lineScaleUser = userCadEntity.LinetypeScale
                                userColor = userCadEntity.Color.ColorIndex
                            ElseIf TypeOf acEnt Is Topomatic.Sfc.Layer.Wrappers.SurfaceStructureLineWrapper Then
                                Dim userStructureLine As Topomatic.Sfc.Layer.Wrappers.SurfaceStructureLineWrapper = acEnt
                                'nameLineType = userStructureLine.LinearSemantic.
                                'Dim lineTypeUser As DwgLinetype = userCadEntity.Linetype
                                'userpattern = FuncStyles.FuncReturnPatternLineTypeCode(lineTypeUser)
                                'lineWidthUser = userStructureLine.Lineweight
                                'lineScaleUser = userCadEntity.LinetypeScale
                                'userColor = userCadEntity.Color.ColorIndex
                                obj_code = userStructureLine.LinearObjectCode.Key
                            End If
                        Next
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'записываем данные
                        For Each row As DataGridViewRow In formEditXML.DataGridView1.Rows
                            If row.Cells(0).Value Like "Number" Then
                                If number > 0 Then
                                    row.Cells(1).Value = number
                                    row.Cells(1).Style.ForeColor = System.Drawing.Color.Green
                                End If
                            ElseIf row.Cells(0).Value Like "Name" Then
                                If nameObject.Trim.Length = 0 Then
                                    row.Cells(1).Value = nameLineType
                                    row.Cells(1).Style.ForeColor = System.Drawing.Color.Green
                                Else
                                    row.Cells(1).Value = nameObject
                                    row.Cells(1).Style.ForeColor = System.Drawing.Color.Green
                                End If
                            ElseIf row.Cells(0).Value Like "LineTypeCode" Then
                                If userpattern.Trim.Length > 0 Then
                                    row.Cells(1).Value = userpattern
                                    row.Cells(1).Style.ForeColor = System.Drawing.Color.Green
                                End If
                            ElseIf row.Cells(0).Value Like "Layer" Then
                                If nameLayer.Trim.Length > 0 Then
                                    row.Cells(1).Value = nameLayer
                                    row.Cells(1).Style.ForeColor = System.Drawing.Color.Green
                                End If
                            ElseIf row.Cells(0).Value Like "Color" Then
                                If userColor > 0 Then
                                    row.Cells(1).Value = userColor
                                    row.Cells(1).Style.ForeColor = System.Drawing.Color.Green
                                End If
                            ElseIf row.Cells(0).Value Like "LineWidth" Then
                                If lineWidthUser >= 0 Then
                                    row.Cells(1).Value = lineWidthUser
                                    row.Cells(1).Style.ForeColor = System.Drawing.Color.Green
                                End If
                            ElseIf row.Cells(0).Value Like "LineScale" Then
                                If lineScaleUser > 0 Then
                                    row.Cells(1).Value = lineScaleUser
                                    row.Cells(1).Style.ForeColor = System.Drawing.Color.Green
                                End If
                            ElseIf row.Cells(0).Value Like "Object_code" Then
                                If obj_code.Trim.Length > 0 Then
                                    row.Cells(1).Value = obj_code
                                    row.Cells(1).Style.ForeColor = System.Drawing.Color.Green
                                End If
                            End If
                        Next
                        '=====================================================================================================
                        'заносим данные в массив
                        For Each row As DataGridViewRow In formEditXML.DataGridView1.Rows
                            Dim namefield As String = row.Cells(0).Value
                            Dim value As String = row.Cells(1).Value
                            If IsNothing(value) = True Then value = ""
                            If IsNothing(namefield) = False Then
                                If namefield.Trim.Length > 0 Then

                                    If arrayData.GetUpperBound(1) <= countArrayData Then
                                        arrayData(0, countArrayData) = namefield
                                        arrayData(1, countArrayData) = value

                                        countArrayData += 1
                                    End If

                                End If
                            End If
                        Next
                        userNode.Tag = arrayData
                    End If
                    formEditXML.boolSelectObject = False
                    formEditXML.boolIzmDataGrid = True
                    GoTo Line1
                End If
            Catch ex As System.StackOverflowException
            End Try
        End Sub
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'Проект полосы отвода
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '0. Исходные данные
        '==============================================================================================================================
        'Импорт КПТ с оформлением
        <cmd("RosreestrImportKPT")>
        Public Sub RosreestrImportKPT()
            'получаем активное окно
            Dim userCadView As CadView = Me.CadView
            If IsNothing(userCadView) Then Return
            'получаем активный слой
            Dim DrawLayer As DrawingLayer = DrawingLayer.GetDrawingLayer(userCadView)
            If IsNothing(DrawLayer) = True Then Return
            'получаем активный документ
            Dim ActivDocument As Topomatic.Dwg.Drawing = DrawLayer.Drawing
            If IsNothing(ActivDocument) = True Then Return
            Dim rng As Random = New Random()
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'читаем масштаб аннотаций (текущий)
            Dim MDraw As Integer = userCadView.AnnotationScale
            'передаем масштаб чертежа в форму
            DrawLayer = DrawingLayer.GetDrawingLayer(userCadView)
            If IsNothing(DrawLayer) = True Then Return
            'получаем активный документ
            ActivDocument = DrawLayer.Drawing
            If IsNothing(ActivDocument) = True Then Return
            '1.\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Dim namePrograms As String = FuncFiles.namePrograms
            Dim generalDir As String = ""
            Dim templateXML As String = ""
            Dim templateDWG As String = ""
            Dim categoryZn As String = "Земельный кадастр"
            Dim boolFindFileTemplate As Boolean = FuncFiles.FuncReadProperties(namePrograms, generalDir, templateXML, templateDWG)
            Dim FormCadastralImport As CadastralTools.FormImportKPT = New CadastralTools.FormImportKPT
            ActivDocument.BeginUpdate()
            Try
                'передаем масштаб чертежа в форму
                Dim nameKvartal As String = ""
                Dim tempPath As String = Path.GetTempPath()
                tempPath = tempPath & "rosreestr"
                FormCadastralImport.initDirectory = tempPath
                'запускаем форму
                FormCadastralImport.ShowDialog()
                If FormCadastralImport.boolShow = False Then
                    Exit Sub
                End If
                If FormCadastralImport.PutchTemlate.Trim.Length > 0 Then
                    templateXML = FormCadastralImport.PutchTemlate
                End If
                If IsNothing(templateXML) = True Then
                    MsgBox("Шаблон оформления не найден!!! Макро прервано.")
                    Exit Sub
                End If
                If File.Exists(templateXML) = False Then
                    MsgBox("Шаблон оформления не найден!!! Макро прервано.")
                    Exit Sub
                End If

                Dim dictCadastralNumber As List(Of String) = New List(Of String) 'список кадастровый номеров для проверки дубликатов
                If FormCadastralImport.boolImportXML = True Then
                    If Directory.Exists(FormCadastralImport.initDirectory) = True Then
                        Try
                            tempPath = Path.Combine(FormCadastralImport.initDirectory, Guid.NewGuid().ToString())
                            Directory.CreateDirectory(tempPath)
                        Catch ex As Exception
                            MsgBox("Не удалось создать папку для выходных данных.")
                            Exit Sub
                        End Try
                    Else
                        Exit Sub
                    End If
                    If FormCadastralImport.boolDubleObject = True Then
                        dictCadastralNumber = CadastralFunction.getNumberCadastralObject(ActivDocument)
                    End If
                    Try
                        If Directory.Exists(tempPath) = False Then
                            Directory.CreateDirectory(tempPath)
                        End If
                        If Directory.Exists(tempPath) = True Then
                            CadastralTools.ImportKPT.RosreestrImportKPT(tempPath, nameKvartal)
                        End If
                    Catch ex As System.Exception
                    End Try
                Else
                    tempPath = FormCadastralImport.initDirectory
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                Dim files As String() = Directory.GetFiles(tempPath)
                If files.Length = 0 Then
                    MsgBox("Не удалось корректно прочитать XML. Макро прервано.")
                    Exit Sub
                End If
                Dim cadastralToolsNetTopoSuit As CadastralTools.FuncNetTopologySuite = New CadastralTools.FuncNetTopologySuite
                Dim roburNetTopoSuit As FuncTopologySuite = New FuncTopologySuite
                Dim drawObject As CreateDWGObject = New CreateDWGObject
                'библиотека стилей отображения

                Dim dictonaryStyle As Dictionary(Of String, StyleCadastralObject) = New Dictionary(Of String, StyleCadastralObject)
                '===============================================================================================================================================================
                'выбираем внешнюю границу для отсечки лишних объектов
                Dim boundPolygon As NetTopologySuite.Geometries.Polygon = Nothing
                If FormCadastralImport.boolSelectBound = True Then
                    'сброс выделения
                    userCadView.SelectionSet.Clear()
                    userCadView.SelectionSet.SelectOneObjectAtScreen(Function(obj) TypeOf obj Is DwgPolyline, "Выберите внешнюю границу: ")
                    For Each acEnt As Object In userCadView.SelectionSet
                        If TypeOf acEnt Is DwgPolyline Then
                            Dim boundPoly As DwgPolyline = acEnt
                            If boundPoly.Area > 0 Then
                                boundPolygon = roburNetTopoSuit.CreatePolygonByPolyline(boundPoly)
                            End If
                        End If
                    Next
                    If boundPolygon.Area < 1 Then
                        MsgBox("Выберите другую внешнюю границу.")
                        Exit Sub
                    End If
                End If
                'читаем земельные участки
                For Each file As String In files
                    Try
                        Dim listDictionaryData As List(Of Dictionary(Of String, Object)) = CadastralTools.JsonValidator.CreateObjectsFromJsonFile(file)
                        If listDictionaryData.Count > 0 Then
                            Dim dataCadastralObject As CadastralTools.StructureElement = New CadastralTools.StructureElement
                            Dim dictLandProperties As Dictionary(Of String, String) = New Dictionary(Of String, String)
                            dataCadastralObject.Label = "Земельный кадастр"
                            Dim idStructure As String = ""
                            Dim tableZn As String = ""
                            Dim PSTable As String = ""
                            Dim typeZn As String = "Полилиния"
                            Dim Label As String = ""
                            Dim idName As Integer = 0
                            Dim countOstrov As Integer = 1
                            Dim listPolygonCoordinates As List(Of String(,)) = New List(Of String(,))
                            For Each cadastralObject As Dictionary(Of String, Object) In listDictionaryData
                                If IsNothing(cadastralObject) = True Then Continue For
                                If cadastralObject.Count > 0 Then
                                    If cadastralObject.ContainsKey("PS") = True Then
                                        PSTable = cadastralObject.Item("PS")
                                    End If
                                    If cadastralObject.ContainsKey("LABEL") = True Then
                                        Label = cadastralObject.Item("LABEL")
                                    End If
                                    If cadastralObject.ContainsKey("NameTypeObject") = True Then
                                        idName = Val(cadastralObject.Item("NameTypeObject"))
                                    End If
                                    'читаем свойства объекта
                                    Dim boolReadDict As Boolean = CadastralTools.DataFormatting.ConvertToDictionaryOfString(cadastralObject, dictLandProperties)

                                    If cadastralObject.ContainsKey("ListCoord") Then
                                        Dim dataCoordParcel As Dictionary(Of String, Object) = cadastralObject("ListCoord")

                                        For Each kvp As KeyValuePair(Of String, Object) In dataCoordParcel
                                            Dim dataCoordCounter As List(Of Object) = TryCast(kvp.Value, List(Of Object))
                                            If dataCoordCounter Is Nothing OrElse dataCoordCounter.Count = 0 Then Continue For

                                            For Each coordItem As Object In dataCoordCounter
                                                Dim listCoordinates As List(Of Object) = TryCast(coordItem, List(Of Object))
                                                If listCoordinates Is Nothing OrElse listCoordinates.Count < 3 Then Continue For

                                                ' Извлекаем списки координат (индексы 1 и 2)
                                                Dim elementCoordX As List(Of Object) = TryCast(listCoordinates(1), List(Of Object))
                                                Dim elementCoordY As List(Of Object) = TryCast(listCoordinates(2), List(Of Object))

                                                If elementCoordX Is Nothing OrElse elementCoordY Is Nothing Then Continue For
                                                If elementCoordX.Count = 0 OrElse elementCoordX.Count <> elementCoordY.Count Then Continue For

                                                ' Создаём массив координат [Точки x Y]
                                                Dim pointCount As Integer = elementCoordX.Count
                                                Dim arrayCoord(pointCount - 1, 1) As String  ' [индекс точки, 0=X, 1=Y]

                                                For k As Integer = 0 To pointCount - 1
                                                    arrayCoord(k, 0) = elementCoordX(k).ToString()
                                                    arrayCoord(k, 1) = elementCoordY(k).ToString()
                                                Next k

                                                If IsNothing(boundPolygon) = False Then
                                                    If boundPolygon.Area > 1 Then
                                                        If Not (PSTable Like "LAND_CADASTRAL_DIVISION") Then
                                                            If arrayCoord.GetUpperBound(1) = 0 Then
                                                                'это точка
                                                                Dim point As NetTopologySuite.Geometries.Point = New NetTopologySuite.Geometries.Point(arrayCoord(0, 0), arrayCoord(1, 0))
                                                                Dim boolIntersect As Boolean = cadastralToolsNetTopoSuit.IsPointInsidePolygon(point, boundPolygon)
                                                                If boolIntersect = False Then
                                                                    Continue For
                                                                End If
                                                            Else
                                                                Dim lineStr As NetTopologySuite.Geometries.LineString = cadastralToolsNetTopoSuit.CreateLineString(arrayCoord)
                                                                Dim boolIntersect As Boolean = cadastralToolsNetTopoSuit.DoesLineIntersectPolygon(lineStr, boundPolygon)
                                                                If boolIntersect = False Then
                                                                    Continue For
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                                listPolygonCoordinates.Add(arrayCoord)
                                            Next coordItem
                                        Next kvp
                                    End If
                                End If
                            Next
                            'преобразуйм в единую строку
                            Dim keyParam As String = CadastralTools.JsonValidator.convertDictionaryToJson(dictLandProperties)
                            '==========================================================================================================
                            'включена опция проверки дубликатов
                            Dim cadNumber As String = ""
                            Dim cadObjectNumber As Object = CadastralTools.JsonValidator.GetValueFromJson(keyParam, "Cad_number")
                            If TypeOf cadObjectNumber Is String Then
                                cadNumber = cadObjectNumber
                            Else
                                cadObjectNumber = CadastralTools.JsonValidator.GetValueFromJson(keyParam, "Reg_numb_border")
                                If TypeOf cadObjectNumber Is String Then
                                    cadNumber = cadObjectNumber
                                End If
                            End If
                            If dictCadastralNumber.Count > 0 Then
                                If dictCadastralNumber.Contains(cadNumber) = True Then
                                    Continue For
                                End If
                            End If
                            '===========================================================================================================
                            'читаем координаты объекта

                            '==========================================================================================================
                            'ФОРМИРУЕМ СЕМАНТИКУ
                            dataCadastralObject.NameTypeObject = CadastralTools.StructureElement.GetEnumByValue(GetType(nameObject), idName)
                            dataCadastralObject.Description = CadastralTools.StructureElement.GetEnumDescriptionByValue(GetType(nameObject), idName)
                            dataCadastralObject.NameClassObject = CadastralTools.StructureElement.classObject.LAND_CADASTRAL
                            dataCadastralObject.TypeGeometry = CadastralTools.StructureElement.typeGeometryObject.polygon
                            dataCadastralObject.KeyParameter = keyParam
                            dataCadastralObject.IdElement = Guid.NewGuid.ToString
                            dataCadastralObject.TablePS = PSTable
                            Select Case PSTable
                                Case "LAND_CADASTRAL_DIVISION"
                                    tableZn = "ЕГРН-Кадастровое деление"
                                Case "LAND_CADASTRAL_PARCEL"
                                    tableZn = "ЕГРН-Кадастровые участки"
                                Case "LAND_CADASTRAL_PART"
                                    tableZn = "ЕГРН-Кадастровые участки"
                                Case "LAND_CADASTRAL_BUILDING"
                                    tableZn = "ЕГРН-Здания (строения)"
                                Case "LAND_CADASTRAL_CONSTRUCTION"
                                    dataCadastralObject.TypeGeometry = CadastralTools.StructureElement.typeGeometryObject.line
                                    tableZn = "ЕГРН-Сооружения"
                                Case "LAND_CADASTRAL_UNFINISHED_CONSTRUCTION"
                                    dataCadastralObject.TypeGeometry = CadastralTools.StructureElement.typeGeometryObject.line
                                    tableZn = "ЕГРН-Объекты незавершенного строительства"
                                Case "LAND_CADASTRAL_ZONE"
                                    tableZn = "ЕГРН-Территориальные зоны"
                                Case "LAND_CADASTRAL_BOUNDARY"
                                    tableZn = "ЕГРН-Границы территорий"
                                    If dataCadastralObject.NameTypeObject = nameObject.WaterBoundary Then
                                        dataCadastralObject.TypeGeometry = CadastralTools.StructureElement.typeGeometryObject.line
                                    End If
                                Case "LAND_CADASTRAL_REDLINE"
                                    tableZn = "ЕГРН-Границы территорий"
                                    dataCadastralObject.TypeGeometry = CadastralTools.StructureElement.typeGeometryObject.line
                                Case "LAND_CADASTRAL_COASTLINE"
                                    tableZn = "ЕГРН-Границы территорий"
                                    dataCadastralObject.TypeGeometry = CadastralTools.StructureElement.typeGeometryObject.line
                            End Select
                            '=========================================================================================================================
                            'назначаем стиль
                            Dim newLineStyle As CadastralPolylineStyle = New CadastralPolylineStyle
                            newLineStyle.TemplateXML = templateXML
                            newLineStyle.ActiveProjectDrawing = ActivDocument
                            newLineStyle.NameStyle = dataCadastralObject.Description
                            If dictonaryStyle.ContainsKey(dataCadastralObject.Description) = False Then
                                Dim arrayProperties As String(,) = {}
                                newLineStyle.setObjectStyle(templateXML, categoryZn, tableZn, dataCadastralObject.TypeGeometry, dataCadastralObject.Description)
                                dictonaryStyle.Add(dataCadastralObject.Description, newLineStyle)
                            Else
                                newLineStyle = dictonaryStyle.Item(dataCadastralObject.Description)
                            End If
                            '=============================================================================================================================
                            'рисуем контур
                            If listPolygonCoordinates.Count > 0 Then
                                Dim listPolygon As Dictionary(Of Object, NetTopologySuite.Geometries.Polygon) = New Dictionary(Of Object, NetTopologySuite.Geometries.Polygon)
                                For i As Integer = 0 To listPolygonCoordinates.Count - 1
                                    Dim arrayCoord As String(,) = listPolygonCoordinates.ElementAt(i)
                                    If arrayCoord.Length > 2 Then
                                        Dim acPolyline As DwgPolyline = drawObject.createPolylineToArrayPoint(arrayCoord)
                                        If IsNothing(acPolyline) = False Then
                                            If dataCadastralObject.TypeGeometry = typeGeometryObject.polygon Then
                                                acPolyline.Closed = True
                                                Dim sourcePolygon As NetTopologySuite.Geometries.Polygon = roburNetTopoSuit.CreatePolygonByPolyline(acPolyline)
                                                listPolygon.Add(acPolyline, sourcePolygon)
                                                If listPolygon.Count > 1 Then
                                                    Dim boolFindOstrov As Boolean = False
                                                    For j As Integer = 0 To listPolygon.Count - 2
                                                        Dim destPolygon As NetTopologySuite.Geometries.Polygon = listPolygon.ElementAt(j).Value
                                                        If destPolygon.Area > sourcePolygon.Area Then
                                                            If cadastralToolsNetTopoSuit.IsPolygonInside(destPolygon, sourcePolygon) = True Then
                                                                boolFindOstrov = True
                                                                Exit For
                                                            ElseIf cadastralToolsNetTopoSuit.IsPolygonInside(sourcePolygon, destPolygon) = True Then
                                                                boolFindOstrov = True
                                                                Exit For
                                                            End If
                                                        End If
                                                    Next j
                                                    If boolFindOstrov = True Then
                                                        Dim objOstrov As CadastralTools.StructureElement = New CadastralTools.StructureElement
                                                        objOstrov.Label = dataCadastralObject.Label
                                                        objOstrov.NameClassObject = CadastralTools.StructureElement.classObject.LAND_CADASTRAL
                                                        dataCadastralObject.TypeGeometry = CadastralTools.StructureElement.typeGeometryObject.polygon
                                                        objOstrov.NameTypeObject = dataCadastralObject.NameTypeObject
                                                        objOstrov.Description = "Внутренний островок"
                                                        objOstrov.IdStructure = dataCadastralObject.IdElement
                                                        objOstrov.IdElement = Guid.NewGuid.ToString
                                                        objOstrov.TablePS = "CADASTRAL_OSTROV"
                                                        Dim dictOstrov As Dictionary(Of String, String) = New Dictionary(Of String, String)
                                                        dictOstrov.Add("ConditionalNumber", countOstrov)
                                                        dictOstrov.Add("ParentNumber", cadNumber)
                                                        Dim keyParamOstrov As String = CadastralTools.JsonValidator.convertDictionaryToJson(dictOstrov)
                                                        objOstrov.KeyParameter = keyParamOstrov
                                                        Dim boolWriteDataOstriv As Boolean = FuncXRecords.setXRecords(acPolyline, CadastralTools.StructureElement.tableXRecords.PROJECT_STRUCTURES, objOstrov)
                                                        ActivDocument.ActiveSpace.Entities.Add(acPolyline)
                                                        newLineStyle.setPolylineStyle(acPolyline)
                                                        countOstrov += 1
                                                        Continue For
                                                    End If
                                                End If
                                            End If
                                            ActivDocument.ActiveSpace.Entities.Add(acPolyline)
                                            newLineStyle.setPolylineStyle(acPolyline)
                                            Dim boolWriteData As Boolean = FuncXRecords.setXRecords(acPolyline, CadastralTools.StructureElement.tableXRecords.PROJECT_STRUCTURES, dataCadastralObject)
                                        End If
                                    Else
                                        Dim acCircle As DwgCircle = New DwgCircle
                                        acCircle.Center = New Vector3D(arrayCoord(0, 0), arrayCoord(0, 1), 0)
                                        acCircle.Radius = 0.5
                                        newLineStyle.setlineStyle(acCircle)
                                        dataCadastralObject.TypeGeometry = typeGeometryObject.block
                                        ActivDocument.ActiveSpace.Entities.Add(acCircle)
                                        Dim boolWriteData As Boolean = FuncXRecords.setXRecords(acCircle, CadastralTools.StructureElement.tableXRecords.PROJECT_STRUCTURES, dataCadastralObject)
                                        Continue For
                                    End If
                                Next i
                            End If
                        End If
                    Catch ex As System.Exception
                    End Try
                Next

                If FormCadastralImport.boolImportXML = True Then
                    If nameKvartal.Trim.Length > 0 Then
                        Dim boolKVName As Boolean = CadastralTools.SupportFiles.RenameDirectory(tempPath, nameKvartal)
                    End If
                End If
            Finally
                ActivDocument.EndUpdate()
            End Try
        End Sub

        '==============================================================================================================================
        'Импорт из текстового файла
        <cmd("ImportCoordinates")>
        Public Sub ImportCoordinates()
            'получаем активное окно
            Dim userCadView As CadView = Me.CadView
            If IsNothing(userCadView) Then Return
            'получаем активный слой
            Dim DrawLayer As DrawingLayer = DrawingLayer.GetDrawingLayer(userCadView)
            If IsNothing(DrawLayer) = True Then Return
            'получаем активный документ
            Dim ActivDocument As Topomatic.Dwg.Drawing = DrawLayer.Drawing
            If IsNothing(ActivDocument) = True Then Return
            Dim rng As Random = New Random()
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'читаем масштаб аннотаций (текущий)
            Dim MDraw As Integer = userCadView.AnnotationScale
            Dim templateXML As String = FuncFiles.funcFindFileToDirectorySupport("Стандарт_ADSK.xml")
            Dim templateDwg As String = FuncFiles.funcFindFileToDirectorySupport("Стандарт_ADSK.dwg")
            If IsNothing(templateXML) = True Or templateXML.Trim.Length = 0 Then
                MsgBox("Файл с шаблонами =Стандарт_ADSK.xml= не найден. Пожалуйста, добавьте его в пути доступа к вспомогательным файлам!!!")
                Exit Sub
            End If
            '=========================================================================================
            Dim formImport As FormImportParcels = New FormImportParcels()
            formImport.ShowDialog()
            If formImport.boolShow = False Then
                Exit Sub
            End If
            'полигоны
            Dim categoryParcel As String = formImport.ComboBox1.Text
            Dim tableParcel As String = formImport.ComboBox2.Text
            Dim deskParcel As String = formImport.ComboBox3.Text
            Dim arrayPropertiesParcels As String(,) = Nothing
            If formImport.CheckBox1.Checked = True Then
                Dim boolPropParcels As Boolean = FuncXML.FuncReadDataObjectToXMLbyDesk(templateXML, categoryParcel, tableParcel, "Полилиния", deskParcel, arrayPropertiesParcels)
            End If
            Dim parcelsLayer As DwgLayer = ActivDocument.ActiveLayer
            Dim parcelsColor As CadColor = New CadColor(7)
            Dim parcelsTypeLine As DwgLinetype = ActivDocument.ActiveLinetype
            Dim parcelsScaleTypeLine As Double = 1
            Dim parcelsWightLine As Double = 20
            Dim parcelsGWidth As Double = 0
            Dim parcelsPSTable As String = ""
            If IsArray(arrayPropertiesParcels) = True Then
                'Dim boolCreateStyle As Boolean = StyleCadastralObject.createStyleLinearObjec(ActivDocument, templateXML, arrayPropertiesParcels, parcelsLayer, parcelsTypeLine, parcelsColor, parcelsScaleTypeLine, parcelsWightLine, parcelsPSTable, parcelsGWidth)
            End If
            'точки
            Dim categoryPoint As String = formImport.ComboBox4.Text
            Dim tablePoint As String = formImport.ComboBox5.Text
            Dim deskPoint As String = formImport.ComboBox6.Text
            Dim arrayPropertiesPoints As String(,) = Nothing
            If formImport.CheckBox2.Checked = True Then
                Dim boolPropObject As Boolean = FuncXML.FuncReadDataObjectToXMLbyDesk(templateXML, categoryPoint, tablePoint, "Блок", deskPoint, arrayPropertiesPoints)
            End If
            Dim pointLayer As DwgLayer = ActivDocument.ActiveLayer
            Dim pointColor As CadColor = New CadColor(7)
            Dim pointTypeLine As DwgLinetype = ActivDocument.ActiveLinetype
            Dim pointScaleLineType As Double = 1
            Dim pointWightLine As Double = 20
            Dim pointScale As Double = 1
            Dim pointAngle As Double = 0
            Dim pointPSTable As String = ""
            Dim pointAttributeHeight As Double = 1.5
            Dim pointName As String = "PROJECT_POINT"
            Dim pointNote As String = ""
            If IsArray(arrayPropertiesPoints) = True Then
                'FuncDecorPlane.FuncCreateStylePointObject(ActivDocument, templateXML, arrayPropertiesPoints, pointName, pointLayer, pointTypeLine, pointColor, pointScaleLineType, pointWightLine, pointScale, pointAngle, pointAttributeHeight, pointPSTable, pointNote)
            End If
            '========================================================================
            'загружаем блок
            Dim blkPoint As DwgBlock = RoburFunc.FuncReadBlockByName(ActivDocument, pointName)
            If IsNothing(blkPoint) = True And formImport.CheckBox2.Checked = True Then
                blkPoint = RoburFunc.FuncInsertBlockByDWGFile(ActivDocument, templateDwg, pointName)
            End If
            'имя атрибута
            Dim attributeTag As String = formImport.ComboBox7.Text
            Dim posAttribX As Double = 1.5
            Dim posAttribY As Double = -1 * (pointAttributeHeight / 2)
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'тело программы
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'выбор файла
            Dim LineErrNumber As Integer = 0
            Dim filterFile As String = "txt files (*.txt)|*.txt"
            Dim FileRez As String = FuncFiles.FuncGetAbsoluteFilePathName("Открыть файл txt", filterFile)
            If FileRez.Trim.Length = 0 Then Exit Sub
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'читаем файл xml
            If File.Exists(FileRez) = True Then
                Dim arrayXYZ As String(,) = Nothing
                Dim countArrayXYZ As Integer = 0
                Dim startRead As Integer = formImport.NumericUpDown7.Value
                Dim r As String = formImport.ComboBox8.Text
                If r Like "Пробел" Then
                    r = ""
                ElseIf r Like "Табуляция" Then
                    r = Chr(9)
                End If
                Dim countLine As Integer = 1
                Dim countCounters As Integer = 0
                Using reader As StreamReader = New StreamReader(FileRez, True)
                    Do Until reader.EndOfStream
                        Dim userTxtLine As String = reader.ReadLine() 'считываем строку
                        If countLine >= startRead Then
                            If IsNothing(userTxtLine) = False Then
                                If userTxtLine.Trim.Length > 0 Then
                                    userTxtLine = userTxtLine.Trim
                                    Dim tempArray As String() = userTxtLine.Split(r)
                                    If IsArray(tempArray) = True Then
                                        Dim numCounter As Integer = 0
                                        Dim numPoint As Integer = 0
                                        Dim X As Double = 0
                                        Dim Y As Double = 0
                                        Dim h As Double = 0
                                        Dim desk As String = ""
                                        If formImport.NumericUpDown1.Value > 0 Then

                                        End If


                                        If tempArray.Length > 2 Then
                                            ReDim Preserve arrayXYZ(3, countArrayXYZ)
                                            arrayXYZ(1, countArrayXYZ) = tempArray(0)
                                            arrayXYZ(2, countArrayXYZ) = tempArray(1)
                                            arrayXYZ(3, countArrayXYZ) = tempArray(2)
                                            countArrayXYZ += 1
                                            Continue Do
                                        End If
                                    Else
                                        countCounters += 1
                                    End If
                                Else
                                    countCounters += 1
                                End If
                            Else
                                countCounters += 1
                            End If
                        End If
                        countLine += 1
                    Loop
                    If IsArray(arrayXYZ) = True Then
                        If arrayXYZ.GetUpperBound(1) > 0 Then
                            Dim acPline As DwgPolyline = CadastralFunction.FuncDrawPolylineToArrayCoord(ActivDocument, arrayXYZ)
                            For i As Integer = 0 To arrayXYZ.GetUpperBound(1)
                                'вставляем блок
                                Dim posInsertBlk As Vector2D = New Vector2D(Val(arrayXYZ(3, i)), Val(arrayXYZ(2, i)))
                                'назначаем номер точке
                                Dim oboznPoint As String = arrayXYZ(1, i)
                                'формируем артрибут
                                Dim dictionaryAttr As Dictionary(Of String, String(,)) = New Dictionary(Of String, String(,))
                                Dim arrayAttrb As String(,) = Nothing
                                ReDim arrayAttrb(1, 3)
                                arrayAttrb(0, 0) = "value"
                                arrayAttrb(1, 0) = oboznPoint
                                arrayAttrb(0, 1) = "height"
                                'arrayAttrb(1, 1) = attribHeight
                                arrayAttrb(0, 2) = "positionX"
                                arrayAttrb(1, 2) = posAttribX
                                arrayAttrb(0, 3) = "positionY"
                                arrayAttrb(1, 3) = posAttribY
                                dictionaryAttr.Add(attributeTag, arrayAttrb)
                                'Dim insertBlk As DwgInsert = CadastralFunction.FuncInsertCadastralPoint(ActivDocument, namePoint, posInsertBlk, dictionaryAttr, anglePoint)
                            Next i
                            Erase arrayXYZ
                            countArrayXYZ = 0
                        End If
                    End If
                End Using
            End If
        End Sub
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '1. создание объектов
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '================================================================================================================================
        'создать объект вручную
        <cmd("CreateCadastralObject")>
        Public Sub CreateCadastralObject()
            'получаем активное окно
            Dim userCadView As CadView = Me.CadView
            If IsNothing(userCadView) Then Return
            'получаем активный слой
            Dim DrawLayer As DrawingLayer = DrawingLayer.GetDrawingLayer(userCadView)
            If IsNothing(DrawLayer) = True Then Return
            'получаем активный документ
            Dim ActivDocument As Topomatic.Dwg.Drawing = DrawLayer.Drawing
            If IsNothing(ActivDocument) = True Then Return
            Dim rng As Random = New Random()
            Dim globalID As String = ""
            '1.\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'вызываем диалоговое окно с первоначальными настройками
            Dim filej As String = "k:\Civil3DToolsUtility"
            userCadView.SelectionSet.Clear()
            Dim formCreateParcels As CadastralTools.FormCreateCadastralObject = New FormCreateCadastralObject(filej)
            formCreateParcels.ShowDialog()
            If formCreateParcels.boolShow = False Then
                Exit Sub
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'если участок многоконтурный формируем глобальный ID

            'читаем префикс участка
            Dim numberParcel As Integer = formCreateParcels.numberParcel
            Dim numberCounter As Integer = formCreateParcels.numberCounter
            Dim fileTemlateXML1 As String = formCreateParcels.fileTemlateXML
            Dim dictionaryPropertiesObject As Dictionary(Of String, String) = formCreateParcels.resultDictionary
            Dim arrayPropertyObject As String(,) = formCreateParcels.arrayPropertiesZn
            Dim arrayPropertiesLeader As String(,) = formCreateParcels.arrayPropertiesLeader
            Dim nameObject As CadastralTools.StructureElement.nameObject = formCreateParcels.nameTypeObject
            Dim typeGeometryObject As StructureElement.typeGeometryObject = formCreateParcels.typeGeometry
            Dim deskObject As String = formCreateParcels.nameObject
            Dim deskLeader As String = formCreateParcels.deskLeader
            Dim nameTableObject As String = formCreateParcels.nameTableZn
            If dictionaryPropertiesObject.ContainsKey("GLOBALID") = True Then
                globalID = dictionaryPropertiesObject.Item("GLOBALID")
            End If
            If dictionaryPropertiesObject.Count > 0 Then
                If arrayPropertyObject.Length > 0 Then
                    '=========================================================================================================================
                    'назначаем стиль объекту
                    Dim styleObject As CadastralPolylineStyle = New CadastralPolylineStyle()
                    styleObject.TemplateXML = fileTemlateXML1
                    styleObject.ActiveProjectDrawing = ActivDocument
                    Dim arrayProperties As String(,) = {}
                    styleObject.setObjectStyle(formCreateParcels.fileTemlateXML, formCreateParcels.nameCategory, formCreateParcels.nameTableZn, typeGeometryObject, formCreateParcels.nameObject)

                    Dim formatLeader As String = CadastralTools.MathFunction.findValueToArray2d(arrayPropertiesLeader, "NOTE")
                    Dim globalClassTopology As CadastralTools.FuncNetTopologySuite = New CadastralTools.FuncNetTopologySuite()
                    Dim localClassTopology As FuncTopologySuite = New FuncTopologySuite()
                    Dim classDrawObject As CreateDWGObject = New CreateDWGObject()
                    classDrawObject.ActiveDrawing = ActivDocument
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'выбираем объекты
                    If typeGeometryObject = typeGeometryObject.polygon Or typeGeometryObject = typeGeometryObject.polyline Then
                        userCadView.SelectionSet.SelectObjectsAtScreen(Function(obj) TypeOf obj Is DwgPolyline, "Выборите полилинии для преобразования: ")
                        For Each acEnt As Object In userCadView.SelectionSet
                            If TypeOf (acEnt) Is DwgPolyline Then
                                Dim plineParcels As DwgPolyline = acEnt
                                Dim fileObjectName As String = plineParcels.Drawing.Filename
                                If Not (fileObjectName Like ActivDocument.Filename) Then
                                    plineParcels = plineParcels.Clone()
                                    If IsNothing(ActivDocument) = False Then
                                        ActivDocument.ActiveSpace.Add(plineParcels)
                                    End If
                                End If
                                'точка вставки мультивыноски
                                Dim innerPoint As NetTopologySuite.Geometries.Point = Nothing
                                'обрабатываем полигон
                                If typeGeometryObject = typeGeometryObject.polygon Then
                                    If plineParcels.Count > 2 Then
                                        Dim polygon As NetTopologySuite.Geometries.Polygon = localClassTopology.CreatePolygonByPolyline(plineParcels)
                                        If IsNothing(polygon) = False Then
                                            polygon = globalClassTopology.RemoveDuplicateVertices(polygon, 0.01)
                                            polygon = globalClassTopology.ReversePolygonPreserveOrientation(polygon)
                                            plineParcels = classDrawObject.createPolylineByPolygon(polygon, plineParcels)
                                        End If
                                        'ищем внутреннюю точку
                                        innerPoint = globalClassTopology.GetInteriorPoint(polygon)
                                    End If
                                    If dictionaryPropertiesObject.ContainsKey("Area") = True Then
                                        dictionaryPropertiesObject.Item("Area") = Math.Round(plineParcels.Area, 0)
                                    End If
                                    'обрабатываем не замкнутую полилинию
                                ElseIf typeGeometryObject = typeGeometryObject.polyline Then
                                    If plineParcels IsNot Nothing AndAlso plineParcels.Count > 1 Then
                                        Dim cadastralLine As NetTopologySuite.Geometries.LineString = localClassTopology.CreateLineString(plineParcels)
                                        If cadastralLine IsNot Nothing AndAlso Not cadastralLine.IsEmpty AndAlso cadastralLine.Length > 0 Then
                                            ' Создаем полилинию (если нужно)
                                            plineParcels = classDrawObject.createPolylineByLineString(cadastralLine, plineParcels)
                                            ' Ищем внутреннюю точку на ГРАНИЦЕ (используем cadastralLine, а не plineParcels)
                                            Dim locationIndexedLine As New LocationIndexedLine(cadastralLine)
                                            ' Вычисляем половину длины от cadastralLine
                                            Dim totalLength As Double = cadastralLine.Length / 2.0
                                            ' Создаем "карту" для преобразования расстояния в LinearLocation
                                            Dim lengthLocationMap As New LengthLocationMap(cadastralLine)
                                            ' Преобразуем длину в LinearLocation
                                            Dim targetLocation As LinearLocation = lengthLocationMap.GetLocation(totalLength)
                                            ' Извлекаем координату
                                            Dim midCoordinate As Coordinate = locationIndexedLine.ExtractPoint(targetLocation)
                                            innerPoint = New NetTopologySuite.Geometries.Point(midCoordinate)
                                        End If
                                    End If
                                    If dictionaryPropertiesObject.ContainsKey("Lenght") = True Then
                                        dictionaryPropertiesObject.Item("Lenght") = Math.Round(plineParcels.Length, 2)
                                    End If
                                End If
                                'обновляем свойства
                                Dim conditionalParcel As String = numberParcel
                                If numberCounter > 0 Then
                                    conditionalParcel = numberParcel & "(" & numberCounter & ")"
                                    numberCounter += 1
                                Else
                                    numberParcel += 1
                                End If
                                'обновляем условный номер и локальный ID
                                If dictionaryPropertiesObject.ContainsKey("ConditionalNumber") = True Then
                                    dictionaryPropertiesObject.Item("ConditionalNumber") = conditionalParcel
                                End If
                                If dictionaryPropertiesObject.ContainsKey("LOCALID") = True Then
                                    dictionaryPropertiesObject.Item("LOCALID") = Guid.NewGuid.ToString
                                End If
                                'формируем свойства объекта
                                Dim keyParamObject As String = CadastralTools.JsonValidator.convertDictionaryToJson(dictionaryPropertiesObject)
                                Dim dataCadastralObject As CadastralTools.StructureElement = New CadastralTools.StructureElement
                                dataCadastralObject.Label = formCreateParcels.nameCategory
                                dataCadastralObject.NameClassObject = CadastralTools.StructureElement.classObject.TERRITORIAL_PLANNING
                                dataCadastralObject.TypeGeometry = formCreateParcels.typeGeometry
                                dataCadastralObject.NameTypeObject = nameObject
                                dataCadastralObject.TableObject = nameTableObject
                                dataCadastralObject.Description = deskObject
                                dataCadastralObject.KeyParameter = keyParamObject
                                dataCadastralObject.TablePS = formCreateParcels.PS
                                dataCadastralObject.IdStructure = globalID
                                dataCadastralObject.IdElement = Guid.NewGuid.ToString
                                dataCadastralObject.IdObject = plineParcels.ObjectID
                                Dim boolWriteDataOstriv As Boolean = FuncXRecords.setXRecords(plineParcels, CadastralTools.StructureElement.tableXRecords.PROJECT_STRUCTURES, dataCadastralObject)
                                styleObject.setPolylineStyle(plineParcels)
                                '============================================================================================================
                                'рисуем выноску
                                '=========================================================================================================================
                                'назначаем стиль выноске
                                If arrayPropertiesLeader.Length > 0 Then
                                    Dim styleLeaderObject As CadastralLeaderStyle = New CadastralLeaderStyle()
                                    styleLeaderObject.ActiveProjectDrawing = ActivDocument
                                    styleLeaderObject.TemplateXML = fileTemlateXML1
                                    Dim tempLayerLeader As String = CadastralTools.MathFunction.findValueToArray2d(arrayPropertiesLeader, "Layer")
                                    If Mid(tempLayerLeader.Trim, 1, 1) Like "_" Then
                                        Dim boolCreateLeaderStyle As Boolean = styleLeaderObject.createStyleLeader(arrayPropertiesLeader, styleObject.Layer.Name)
                                    Else
                                        Dim boolCreateLeaderStyle As Boolean = styleLeaderObject.createStyleLeader(arrayPropertiesLeader)
                                    End If
                                    '===============================================================
                                    'вторая точка
                                    Dim pointBound As BoundingBox2D = plineParcels.Bounds
                                    Dim maskLabel As String = ""
                                    Dim strNoteLeader As String = CadastralTools.MathFunction.findValueToArray2d(arrayPropertiesLeader, "Note")
                                    If strNoteLeader.Trim.Length > 0 Then
                                        conditionalParcel = CadastralTools.DataFormatting.getTextLeader(strNoteLeader, dictionaryPropertiesObject, maskLabel)
                                        If conditionalParcel.Trim.Length > 0 Then
                                            Dim leaderObject As DwgLeader = classDrawObject.createLeader(conditionalParcel, New Vector2D(innerPoint.X, innerPoint.Y), pointBound.Max)
                                            If IsNothing(leaderObject) = False Then
                                                styleLeaderObject.setLeaderStyle(leaderObject)
                                                Dim createXdata As StructureElement = CadastralTools.CadastralLeader.getLeaderXData(dataCadastralObject, deskLeader, maskLabel, fileTemlateXML1)
                                                Dim boolWriteDataLeader As Boolean = FuncXRecords.setXRecords(leaderObject, CadastralTools.StructureElement.tableXRecords.PROJECT_STRUCTURES, createXdata)
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    ElseIf typeGeometryObject = typeGeometryObject.block Then
                        'загружаем в чертеж блок



                        'CadCursors.GetPoint(userCadView,, "Укажите точку вставки блока: ")
                        'userCadView.SelectionSet.SelectObjectsAtScreen(Function(obj) TypeOf obj Is DwgInsert, "Выборите блоки для преобразования: ")
                    End If
                End If
            End If
        End Sub
        '================================================================================================================================
        'полуавтоматическое создание участков
        <cmd("AutoCreateCadastralObject")>
        Public Sub AutoCreateCadastralObject()
            'получаем активное окно
            Dim userCadView As CadView = Me.CadView
            If IsNothing(userCadView) Then Return
            'получаем активный слой
            Dim DrawLayer As DrawingLayer = DrawingLayer.GetDrawingLayer(userCadView)
            If IsNothing(DrawLayer) = True Then Return
            'получаем активный документ
            Dim ActivDocument As Topomatic.Dwg.Drawing = DrawLayer.Drawing
            If IsNothing(ActivDocument) = True Then Return
            Dim rng As Random = New Random()
            '1.\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '1.\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Dim generalDir As String = ""
            Dim templateXML As String = ""
            Dim templateDWG As String = ""
            Dim arrayDirSupport As String() = Nothing
            Dim boolFindDirSupport As Boolean = FuncFiles.funcReadDirectoriesSupport(arrayDirSupport)
            If IsArray(arrayDirSupport) = True Then
                For i As Integer = 0 To arrayDirSupport.Length - 1
                    Dim nameFolder As String = arrayDirSupport(i)
                    Dim pos As Integer = nameFolder.IndexOf("InfrastradaToolsUtility")
                    If pos > -1 Then
                        generalDir = Mid(nameFolder, 1, pos)
                        generalDir = generalDir & "InfrastradaToolsUtility"
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
                        Dim shortNameFile As String = System.IO.Path.GetFileNameWithoutExtension(templateXML)
                        Dim tempNameFileDwg As String = directorySupport & shortNameFile & ".dwg"
                        If System.IO.File.Exists(tempNameFileDwg) = True Then
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
            'сброс выделения
            userCadView.SelectionSet.Clear()
            userCadView.SelectionSet.SelectOneObjectAtScreen(Function(obj) TypeOf obj Is DwgPolyline, "Выборите полилинию по которой будут создаваться новые объекты: ")
            Dim parentObject As DwgPolyline = Nothing
            Dim l1 As List(Of Vector2D) = New List(Of Vector2D)
            For Each acEnt As Object In userCadView.SelectionSet
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                If TypeOf acEnt Is DwgPolyline Then
                    parentObject = acEnt
                    If parentObject.Closed = False Then
                        If parentObject.Count > 2 Then
                            parentObject.Closed = True
                        Else
                            MsgBox("Выбранная полилиния не замкнута. Макро прервано!!!")
                            Continue For
                        End If
                    End If

                    Dim p3d1 As Polyline3D = New Polyline3D()
                    parentObject.GetPolyline(p3d1)
                    p3d1.ConvertToPosArray(l1, ValueConverter.Eps)
                    Exit For
                End If
            Next
            If IsNothing(parentObject) = True Then Exit Sub
            '===================================================================================================================
            userCadView.SelectionSet.Clear()
            Dim formAutoCreateParcels As FormAutoCreateParcels = New FormAutoCreateParcels
            formAutoCreateParcels.ShowDialog()
            If formAutoCreateParcels.boolShow = False Then
                Exit Sub
            End If
            Dim arrayUserParcels As DwgPolyline() = Nothing
            Dim countArrayUserParcels As Integer = 0
            If formAutoCreateParcels.CheckBox2.Checked = False Then
                userCadView.SelectionSet.SelectObjectsAtScreen(Function(obj) TypeOf obj Is DwgPolyline, "Выборите кадастровые участки: ")
                For Each acEnt As Object In userCadView.SelectionSet
                    If TypeOf (acEnt) Is DwgPolyline Then
                        Dim plineCadastralParcel As DwgPolyline = acEnt
                        Dim boolIntersect As Boolean = plineCadastralParcel.IntersectWith(parentObject.Bounds, 1)
                        If boolIntersect = True Then
                            Dim p3d2 As Polyline3D = New Polyline3D()
                            plineCadastralParcel.GetPolyline(p3d2)
                            Dim l2 As List(Of Vector2D) = New List(Of Vector2D)
                            p3d2.ConvertToPosArray(l2, ValueConverter.Eps)
                            Dim op = New PolygonOperation()
                            Dim l3 As List(Of List(Of Vector2D)) = New List(Of List(Of Vector2D))
                            l3 = op.Intersection(l1, l2)
                            If l3.Count > 0 Then
                                For i As Integer = 0 To l3.Count - 1
                                    Dim l4 As List(Of Vector2D) = l3.Item(i)
                                    Dim polyline As DwgPolyline = New Topomatic.Dwg.Entities.DwgPolyline()
                                    For j As Integer = 0 To l4.Count - 1
                                        Dim pos = l4.Item(j)
                                        polyline.Add(New Topomatic.Cad.Foundation.BugleVector2D(pos, 0))
                                    Next j
                                    ActivDocument.ActiveSpace.Add(polyline)
                                    ReDim Preserve arrayUserParcels(countArrayUserParcels)
                                    arrayUserParcels(countArrayUserParcels) = polyline
                                    countArrayUserParcels += 1
                                Next i
                            End If
                        End If
                    End If
                Next
            Else
                userCadView.SelectionSet.SelectAll()
            End If

        End Sub
        '================================================================================================================================
        'создание островка
        <cmd("CreateOstrovObject")>
        Public Sub CreateOstrovObject()
            'получаем активное окно
            Dim userCadView As CadView = Me.CadView
            If IsNothing(userCadView) Then Return
            'получаем активный слой
            Dim DrawLayer As DrawingLayer = DrawingLayer.GetDrawingLayer(userCadView)
            If IsNothing(DrawLayer) = True Then Return
            'получаем активный документ
            Dim ActivDocument As Topomatic.Dwg.Drawing = DrawLayer.Drawing
            If IsNothing(ActivDocument) = True Then Return
            Dim rng As Random = New Random()
            '1.\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '1.\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Dim generalDir As String = ""
            Dim templateXML As String = ""
            Dim templateDWG As String = ""
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
            Dim dictionaryTemplate As Dictionary(Of String, String) = CadastralTools.TemplateXML.readTemplatesXML(generalDir)
            If dictionaryTemplate.Count > 0 Then
                templateXML = dictionaryTemplate.ElementAt(0).Value
            End If
            If IsNothing(templateXML) = True Or templateXML.Trim.Length = 0 Then
                MsgBox("Файл с шаблонами =Стандарт_ADSK.xml= не найден. Пожалуйста, добавьте его в пути доступа к вспомогательным файлам!!!")
                Exit Sub
            End If
            'сброс выделения
            userCadView.SelectionSet.Clear()
            userCadView.SelectionSet.SelectOneObjectAtScreen(Function(obj) TypeOf obj Is DwgPolyline, "Выборите объект для которого требуется создать островок: ")
            Dim dataStructuresObject As StructureElement = Nothing
            Dim parentObject As DwgPolyline = Nothing
            For Each acEnt As Object In userCadView.SelectionSet
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                Dim boolRecData As Boolean = FuncXRecords.getXRecords(acEnt, dataStructuresObject)
                If boolRecData = True Then
                    parentObject = acEnt
                    Exit For
                End If
            Next
            'выбираем объекты внутри выбранного контура
            Dim dictionaryCadastralObject As Dictionary(Of StructureElement.nameObject, List(Of NetTopologySuite.Geometries.Geometry)) = SelectCadastralObject.selectAllObject(userCadView)
            Dim cadastralNetTopoSuite As CadastralTools.FuncNetTopologySuite = New FuncNetTopologySuite
            Dim roburNetTopoSuite As FuncTopologySuite = New FuncTopologySuite()
            Dim classDrawObject As CreateDWGObject = New CreateDWGObject()
            classDrawObject.ActiveDrawing = ActivDocument
            Dim lineStyle As CadastralPolylineStyle = New CadastralPolylineStyle(parentObject)
            If IsNothing(dataStructuresObject) = False And IsNothing(parentObject) = False Then
                Dim parentID As String = dataStructuresObject.IdElement
                Dim keyParam As String = dataStructuresObject.KeyParameter
                Dim parentPolygon As NetTopologySuite.Geometries.Polygon = roburNetTopoSuite.CreatePolygonByPolyline(parentObject)
                Dim arrayPropertiesCadastralOstrov As String(,) = {}
                Dim boolFindPropertiesOstrov As Boolean = CadastralTools.TemplateXML.FuncReadTablePSToXML(templateXML, "CADASTRAL_OSTROV", arrayPropertiesCadastralOstrov)
                Dim numberOstrov As Integer = CadastralTools.CadastralOstrov.getLastNumber(parentPolygon, dictionaryCadastralObject)
                numberOstrov += 1
                Dim listPolygon As List(Of NetTopologySuite.Geometries.Geometry) = Nothing
                Dim listOstrov As List(Of NetTopologySuite.Geometries.Geometry) = Nothing
                If dictionaryCadastralObject.ContainsKey(dataStructuresObject.NameTypeObject) = True Then
                    listPolygon = dictionaryCadastralObject.Item(dataStructuresObject.NameTypeObject)
                End If
                If dictionaryCadastralObject.ContainsKey(StructureElement.nameObject.ProjectOstrov) = True Then
                    listOstrov = dictionaryCadastralObject.Item(StructureElement.nameObject.ProjectOstrov)
                End If
                parentPolygon = CadastralTools.CadastralPolygon.getCounter(parentPolygon, listPolygon, listOstrov)
                Dim areaParcel As Double = parentPolygon.Area
                If arrayPropertiesCadastralOstrov.Length > 0 Then
                    'список полигонов островков для их упорядочивания
                    Dim listPolygonRing As List(Of NetTopologySuite.Geometries.Polygon) = New List(Of NetTopologySuite.Geometries.Polygon)
                    userCadView.SelectionSet.Clear()
                    userCadView.SelectionSet.SelectObjectsAtScreen(Function(obj) TypeOf obj Is DwgPolyline, "Выборите полилинии для преобразования их в островки: ")
                    For Each acEnt As Object In userCadView.SelectionSet
                        If TypeOf (acEnt) Is DwgPolyline Then
                            Dim plineOstrov As DwgPolyline = acEnt
                            If plineOstrov.Count > 2 Then
                                If plineOstrov.Area > 0 Then
                                    If Not (plineOstrov.Drawing.Layers.Drawing.Filename = parentObject.Drawing.Layers.Drawing.Filename) Then
                                        plineOstrov = plineOstrov.Clone
                                        ActivDocument.ActiveSpace.Add(plineOstrov)
                                    ElseIf plineOstrov.HasExtensionDictionary = True Then
                                        plineOstrov = plineOstrov.Clone
                                        ActivDocument.ActiveSpace.Add(plineOstrov)
                                    End If
                                End If
                                Dim ostrovPolygon As NetTopologySuite.Geometries.Polygon = roburNetTopoSuite.CreatePolygonByPolyline(plineOstrov)
                                ostrovPolygon = cadastralNetTopoSuite.RemoveDuplicateVertices(ostrovPolygon, 0.01)
                                ostrovPolygon = cadastralNetTopoSuite.ReversePolygonPreserveOrientation(ostrovPolygon)
                                If ostrovPolygon.Coordinates.Count > 2 Then
                                    If ostrovPolygon.Area > 0 Then
                                        ostrovPolygon.UserData = plineOstrov.ObjectID
                                        listPolygonRing.Add(ostrovPolygon)
                                    End If
                                End If
                            End If
                        End If
                    Next
                    If listPolygonRing.Count > 2 Then
                        'упорядочиваем полигоны
                        listPolygonRing = cadastralNetTopoSuite.SortByNearestNeighbor(listPolygonRing)
                    End If

                    For Each userPolygon As NetTopologySuite.Geometries.Polygon In listPolygonRing
                        Dim idPolyline As UInteger = CUInt(userPolygon.UserData)
                        Dim polylineRing As DwgPolyline = New DwgPolyline()
                        Dim boolFindPolyline As Boolean = ActivDocument.ActiveSpace.Entities.TryGetObject(idPolyline, polylineRing)
                        If boolFindPolyline = True Then
                            polylineRing = classDrawObject.createPolylineByPolygon(userPolygon, polylineRing)
                            'формируем свойства объекта
                            Dim dataCadastralObject As CadastralTools.StructureElement = CadastralTools.CadastralOstrov.getProjectOstrovXData(dataStructuresObject, numberOstrov, arrayPropertiesCadastralOstrov)
                            dataCadastralObject.IdObject = polylineRing.ObjectID
                            Dim boolWriteDataOstriv As Boolean = FuncXRecords.setXRecords(polylineRing, CadastralTools.StructureElement.tableXRecords.PROJECT_STRUCTURES, dataCadastralObject)
                            lineStyle.setPolylineStyle(polylineRing)
                            parentPolygon = cadastralNetTopoSuite.AddPolygonAsHole(parentPolygon, userPolygon)
                            numberOstrov += 1
                        End If
                    Next
                    Dim rezArea As Double = Math.Round(parentPolygon.Area, 0)
                    keyParam = CadastralTools.JsonValidator.SetJsonFieldValue(keyParam, "Area", rezArea)
                    dataStructuresObject.KeyParameter = keyParam
                    Dim boolNewData As Boolean = FuncXRecords.setXRecords(parentObject, tableXRecords.PROJECT_STRUCTURES, dataStructuresObject)
                End If
            End If
        End Sub
        '================================================================================================================================
        'добавить контур
        <cmd("AddCadastralObjectCounter")>
        Public Sub CreateCounterObject()
            'получаем активное окно
            Dim userCadView As CadView = Me.CadView
            If IsNothing(userCadView) Then Return
            'получаем активный слой
            Dim DrawLayer As DrawingLayer = DrawingLayer.GetDrawingLayer(userCadView)
            If IsNothing(DrawLayer) = True Then Return
            'получаем активный документ
            Dim ActivDocument As Topomatic.Dwg.Drawing = DrawLayer.Drawing
            If IsNothing(ActivDocument) = True Then Return
            Dim rng As Random = New Random()
            '1.\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Dim generalDir As String = ""
            Dim templateXML As String = ""
            Dim templateDWG As String = ""
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
            Dim dictionaryTemplate As Dictionary(Of String, String) = CadastralTools.TemplateXML.readTemplatesXML(generalDir)
            If dictionaryTemplate.Count > 0 Then
                templateXML = dictionaryTemplate.ElementAt(0).Value
            End If
            If IsNothing(templateXML) = True Or templateXML.Trim.Length = 0 Then
                MsgBox("Файл с шаблонами =Стандарт_ADSK.xml= не найден. Пожалуйста, добавьте его в пути доступа к вспомогательным файлам!!!")
                Exit Sub
            End If
            'сброс выделения
            userCadView.SelectionSet.Clear()
            userCadView.SelectionSet.SelectOneObjectAtScreen(Function(obj) TypeOf obj Is DwgPolyline, "Выборите объект для которого требуется создать контур: ")
            Dim dataParentParcel As StructureElement = Nothing
            Dim parentPolyline As DwgPolyline = Nothing
            For Each acEnt As Object In userCadView.SelectionSet
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                Dim boolRecData As Boolean = FuncXRecords.getXRecords(acEnt, dataParentParcel)
                If boolRecData = True Then
                    parentPolyline = acEnt
                    Exit For
                Else
                    MsgBox("Не верный тип объекта.")
                    Exit Sub
                End If
            Next
            'выбираем объекты внутри выбранного контура
            Dim dictionaryCadastralObject As Dictionary(Of StructureElement.nameObject, List(Of NetTopologySuite.Geometries.Geometry)) = SelectCadastralObject.selectAllObject(userCadView)
            Dim cadastralNetTopoSuite As CadastralTools.FuncNetTopologySuite = New FuncNetTopologySuite
            Dim roburNetTopoSuite As FuncTopologySuite = New FuncTopologySuite()
            Dim classDrawObject As CreateDWGObject = New CreateDWGObject()
            classDrawObject.ActiveDrawing = ActivDocument
            Dim lineStyle As CadastralPolylineStyle = New CadastralPolylineStyle(parentPolyline)
            Dim numberParcel As Integer = 1
            Dim counterParcel As Integer = 1
            If IsNothing(dataParentParcel) = False And IsNothing(parentPolyline) = False Then
                Dim parentID As String = dataParentParcel.IdElement
                Dim keyParam As String = dataParentParcel.KeyParameter
                Dim dictionaryDataParent As Dictionary(Of String, Object) = CadastralTools.JsonValidator.DeserializeJsonToDictionary(keyParam)
                If dictionaryDataParent.Count > 0 Then
                    If dictionaryDataParent.ContainsKey("ConditionalNumber") = True Then
                        Dim conditionalParentParcel As String = dictionaryDataParent.Item("ConditionalNumber")
                        numberParcel = CadastralTools.DataFormatting.GetNumberBeforeParenthesis(conditionalParentParcel)
                        If numberParcel = 0 Then numberParcel = 1
                        counterParcel = CadastralTools.DataFormatting.GetNumberInParenthesis(conditionalParentParcel)
                        If counterParcel = 0 Then
                            counterParcel = 1
                            conditionalParentParcel = numberParcel & "(" & counterParcel & ")"
                            keyParam = CadastralTools.JsonValidator.SetJsonFieldValue(keyParam, "ConditionalNumber", conditionalParentParcel)
                            dataParentParcel.KeyParameter = keyParam
                            If dataParentParcel.IdStructure.Length < 2 Then
                                dataParentParcel.IdStructure = Guid.NewGuid.ToString
                            End If
                            Dim boolRecData As Boolean = FuncXRecords.setXRecords(parentPolyline, tableXRecords.PROJECT_STRUCTURES, dataParentParcel)
                        End If
                    End If
                End If
                '====================================================================================================================================
                'выьираем контура для выбранного участка
                userCadView.SelectionSet.Clear()
                userCadView.SelectionSet.SelectObjectsAtScreen(Function(obj) TypeOf obj Is DwgPolyline, "Выборите полилинии для преобразования их в островки: ")
                Dim listCounter As List(Of NetTopologySuite.Geometries.Geometry) = New List(Of Geometry)
                For Each acEnt As Object In userCadView.SelectionSet
                    If TypeOf (acEnt) Is DwgPolyline Then
                        Dim plineCounter As DwgPolyline = acEnt
                        If Not (plineCounter.Drawing.Layers.Drawing.Filename = parentPolyline.Drawing.Layers.Drawing.Filename) Then
                            plineCounter = plineCounter.Clone
                            ActivDocument.ActiveSpace.Add(plineCounter)
                        End If
                        If dataParentParcel.TypeGeometry = typeGeometryObject.polygon Then
                            If plineCounter.Count > 2 Then
                                If plineCounter.Area > 0 Then
                                    Dim counterPolygon As NetTopologySuite.Geometries.Polygon = roburNetTopoSuite.CreatePolygonByPolyline(plineCounter)
                                    counterPolygon = cadastralNetTopoSuite.RemoveDuplicateVertices(counterPolygon, 0.01)
                                    counterPolygon = cadastralNetTopoSuite.ReversePolygonPreserveOrientation(counterPolygon)
                                    If counterPolygon.Coordinates.Count > 2 Then
                                        If counterPolygon.Area > 0 Then
                                            counterPolygon.UserData = plineCounter.ObjectID
                                            listCounter.Add(counterPolygon)
                                        End If
                                    End If
                                End If
                            End If
                        ElseIf dataParentParcel.TypeGeometry = typeGeometryObject.line Then
                            Dim cadastralLine As NetTopologySuite.Geometries.LineString = roburNetTopoSuite.CreateLineString(plineCounter)
                            If cadastralLine IsNot Nothing AndAlso Not cadastralLine.IsEmpty AndAlso cadastralLine.Length > 0 Then
                                cadastralLine = cadastralNetTopoSuite.RemoveDuplicateVertices(cadastralLine, 0.01)
                                If cadastralLine.Coordinates.Count > 2 Then
                                    If cadastralLine.Area > 0 Then
                                        cadastralLine.UserData = plineCounter.ObjectID
                                        listCounter.Add(cadastralLine)
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
                If listCounter.Count > 2 Then
                    'упорядочиваем полигоны
                    listCounter = cadastralNetTopoSuite.SortByNearestNeighbor(listCounter)
                End If
                '======================================================================================================================
                'создаем контура
                For Each topoObject As NetTopologySuite.Geometries.Geometry In listCounter
                    Dim idPolyline As UInteger = CUInt(topoObject.UserData)
                    Dim polylineCounter As DwgPolyline = New DwgPolyline()
                    Dim boolFindPolyline As Boolean = ActivDocument.ActiveSpace.Entities.TryGetObject(idPolyline, polylineCounter)
                    If boolFindPolyline = True Then
                        If TypeOf topoObject Is NetTopologySuite.Geometries.Polygon Then
                            polylineCounter = classDrawObject.createPolylineByPolygon(topoObject, polylineCounter)
                        ElseIf TypeOf topoObject Is LineString Then
                            polylineCounter = classDrawObject.createPolylineByLineString(topoObject, polylineCounter)
                        End If
                        Dim newDataCounter As StructureElement = New StructureElement
                        newDataCounter = dataParentParcel
                        Dim keyParamCounter As String = newDataCounter.KeyParameter
                        counterParcel += 1
                        keyParamCounter = CadastralTools.JsonValidator.SetJsonFieldValue(keyParam, "ConditionalNumber", numberParcel & "(" & counterParcel & ")")
                        newDataCounter.KeyParameter = keyParamCounter
                        newDataCounter.IdElement = Guid.NewGuid.ToString
                        Dim boolWriteDataOstriv As Boolean = FuncXRecords.setXRecords(polylineCounter, CadastralTools.StructureElement.tableXRecords.PROJECT_STRUCTURES, newDataCounter)
                        lineStyle.setPolylineStyle(polylineCounter)
                    End If
                Next
            End If
        End Sub
        '================================================================================================================================
        'заменить контур
        <cmd("ReplaceCadastralObject")>
        Public Sub ReplaceCadastralObject()
            'получаем активное окно
            Dim userCadView As CadView = Me.CadView
            If IsNothing(userCadView) Then Return
            'получаем активный слой
            Dim DrawLayer As DrawingLayer = DrawingLayer.GetDrawingLayer(userCadView)
            If IsNothing(DrawLayer) = True Then Return
            'получаем активный документ
            Dim ActivDocument As Topomatic.Dwg.Drawing = DrawLayer.Drawing
            If IsNothing(ActivDocument) = True Then Return
            Dim rng As Random = New Random()
            '1.\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Dim generalDir As String = ""
            Dim templateXML As String = ""
            Dim templateDWG As String = ""
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
            'сброс выделения
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            userCadView.SelectionSet.Clear()
            userCadView.SelectionSet.SelectOneObjectAtScreen(Function(obj) TypeOf obj Is Object, "Выборите исходный полигон который требуется заменить: ")
            Dim parentParcel As DwgPolyline = Nothing
            Dim dataParent As StructureElement = Nothing
            Dim stylePolyline As CadastralPolylineStyle = Nothing
            For Each acEnt As Object In userCadView.SelectionSet
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                If TypeOf acEnt Is DwgPolyline Then
                    parentParcel = acEnt
                    Dim boolReadDataParcel As Boolean = FuncXRecords.getXRecords(parentParcel, dataParent)
                    If boolReadDataParcel = True Then
                        stylePolyline = New CadastralPolylineStyle(parentParcel)
                        stylePolyline.ActiveProjectDrawing = ActivDocument
                        Exit For
                    End If
                End If
            Next
            If IsNothing(dataParent) = False Then
                userCadView.SelectionSet.Clear()
                userCadView.SelectionSet.SelectObjectsAtScreen(Function(obj) TypeOf obj Is Object, "Выборите новый полигон: ")
                For Each acEnt As Object In userCadView.SelectionSet
                    If TypeOf (acEnt) Is DwgPolyline Then
                        Dim plineCounter As DwgPolyline = acEnt
                        stylePolyline.setPolylineStyle(plineCounter)
                        Dim boolXRec As Boolean = FuncXRecords.setXRecords(plineCounter, CadastralTools.StructureElement.tableXRecords.PROJECT_STRUCTURES, dataParent)
                        If boolXRec = True Then
                            ActivDocument.ActiveSpace.Entities.Remove(parentParcel)
                        End If

                    End If
                Next
            End If

        End Sub
        '===============================================================================================================================
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '2 Нумеровать полигон
        <cmd("AutoNumberCadastralObject")>
        Public Sub AutoNumberCadastralObject()
            'получаем активное окно
            Dim userCadView As CadView = Me.CadView
            If IsNothing(userCadView) Then Return
            'получаем активный слой
            Dim DrawLayer As DrawingLayer = DrawingLayer.GetDrawingLayer(userCadView)
            If IsNothing(DrawLayer) = True Then Return
            'получаем активный документ
            Dim ActivDocument As Topomatic.Dwg.Drawing = DrawLayer.Drawing
            If IsNothing(ActivDocument) = True Then Return
            'сброс выделения
            userCadView.SelectionSet.Clear()
            '1.\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'находим файл с шаблоном
            '1.\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'делаем настройки для диалогового окна УСТАНОВКИ
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'читаем масштаб аннотаций (текущий)
            Dim MAnno As Integer = userCadView.AnnotationScale * 1000
            Dim generalDir As String = FuncFiles.funcFindDirectorySupport("Civil3DToolsUtility")
            Dim formNumbersPoint As CadastralTools.FormCreatePoint = New CadastralTools.FormCreatePoint(MAnno, generalDir)
            'вызываем диалоговое окно УСТАНОВКИ
            formNumbersPoint.ShowDialog()
            If formNumbersPoint.boolShow = False Then
                Exit Sub
            End If
            'путь к шаблону
            Dim templateDWG As String = formNumbersPoint.PutchTemlateDwg
            Dim templateXML As String = formNumbersPoint.PutchTemlateXML
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'настройки
            Dim idStyle As ObjectId = Nothing 'текстовые стиль
            'префикс номера
            Dim prefNumber As String = formNumbersPoint.prefNumber
            'метод определения координат
            Dim MetodCoord As String = formNumbersPoint.metodCoord
            'способ закрепления
            Dim Zakrep As String = formNumbersPoint.zakrPoint
            'погрешность определения точки
            Dim Delta As String = formNumbersPoint.icnarisityPoint
            'поиск дубликатов
            Dim boolDoublePoint As Boolean = formNumbersPoint.boolDublePoint
            Dim tolerance As Double = formNumbersPoint.tolerance
            'округрение вершин
            Dim boolRoundVertex As Boolean = formNumbersPoint.boolRoundVertex
            'по часовой стрелке
            Dim boolClock As Boolean = formNumbersPoint.boolClockVertex
            'северо-западный угол
            Dim boolSVAngle As Boolean = formNumbersPoint.boolSEVertex
            '========================================================================================
            Dim M As Double = formNumbersPoint.scalePoint * 1000 'масштаб
            Dim HText As Double = (M * formNumbersPoint.hText) / 1000 ' высота подписи
            Dim DText As Double = (M * formNumbersPoint.offsetText) / 1000 'расстояние смещения
            '==================================================================================================
            'размер блока
            Dim nameBlk As String = ""
            Dim HPoint As Double = 1 'размер точки
            Dim ScalePoint As Vector3D = New Vector3D((HPoint * M) / 1000, (HPoint * M) / 1000, (HPoint * M) / 1000)
            '===========================================================================================
            'характеристика атрибута
            Dim nameAttributeNumber As String = formNumbersPoint.nameAttribyteNumber
            Dim attributeHeight As Double = 2
            Dim arrayAttributes As String(,) = formNumbersPoint.arrayAttribytes
            '============================================================================================
            Dim categoryZN As String = formNumbersPoint.categoryZN
            Dim tableZn As String = formNumbersPoint.tableZN
            Dim typeObject As CadastralTools.StructureElement.typeGeometryObject = StructureElement.typeGeometryObject.block
            Dim descrittionBlockStyle As String = formNumbersPoint.deskBlockStyle
            Dim descriptionObjectStyle As String = formNumbersPoint.deskObjectStyle
            '===============================================================================================
            Dim ArrayPoligon As String(,) = Nothing 'массив для хранения всех участков
            Dim CountPline As Integer = 0 'счетчик участков
            Dim ArrayTablesXML As String(,) = Nothing 'массив с таблицами для фильтрации
            Dim arrayPropertiesBlock As String(,) = {}
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'открываем транзакцию
            Dim styleBlockReference As CadastralBlockStyle = New CadastralBlockStyle()
            Dim acadNetTopogy As FuncTopologySuite = New FuncTopologySuite()
            Dim cadastralNetTopologySuite As CadastralTools.FuncNetTopologySuite = New CadastralTools.FuncNetTopologySuite
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'блокируем чертеж
            ActivDocument.BeginUpdate()
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'копируем блок
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            styleBlockReference.ActiveProjectDrawing = ActivDocument
            styleBlockReference.TemplateDwg = templateDWG
            styleBlockReference.TemplateXML = templateXML
            Dim boolReadPropertiesBlock As Boolean = CadastralTools.TemplateXML.FuncReadPropertiesObjectToXML(templateXML, categoryZN, tableZn, typeObject, descrittionBlockStyle, arrayPropertiesBlock)
            If arrayPropertiesBlock.Length > 0 Then
                styleBlockReference.createBlockStyle(arrayPropertiesBlock)
                attributeHeight = Val(CadastralTools.MathFunction.findValueToArray2d(arrayPropertiesBlock, "AttributeHeight"))
            End If
            nameBlk = styleBlockReference.NameBlock
            Dim findBlk As Boolean = styleBlockReference.HacBlockDwg(nameBlk)
            If findBlk = False Then
                ActivDocument.EndUpdate()
                MsgBox("Имя блока не найдено в чертеже. Загрузка из шаблона также не удалась. Попробуйте добавить его в ручную.")
                Exit Sub
            End If
            'находим наибольший номер
            Dim listBlockReference As Dictionary(Of String, NetTopologySuite.Geometries.Point) = acadNetTopogy.getPointByListBlockReference(ActivDocument, nameBlk, nameAttributeNumber)
            Dim LargNumber As Integer = 1
            If listBlockReference.Count > 0 Then
                LargNumber = CadastralTools.CadastralPoint.getLargNumber(listBlockReference, nameAttributeNumber)
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'начальный номер точки
            Dim rezNumberPoint As GetPointResult = CadCursors.GetIntegerWithDefault(userCadView, LargNumber, vbLf & "Введите начальный номер точки: ")
            Try
                'выбираем примитивы
                Dim CountParcels As Integer = 0
                Dim dictCadastralParcels As New Dictionary(Of Integer, String(,))
                userCadView.SelectionSet.SelectObjectsAtScreen(Function(obj) TypeOf obj Is DwgPolyline, "Выборите объекты для нумерации вершин: ")
                For Each acEnt As Object In userCadView.SelectionSet
                    If TypeOf (acEnt) Is DwgPolyline Then
                        Dim plineParcels As DwgPolyline = acEnt
                        Dim fileObjectName As String = plineParcels.Drawing.Filename
                        If fileObjectName Like ActivDocument.Filename Then
                            Dim deskObject As String = ""
                            Dim dataPolyline As StructureElement = Nothing
                            Dim boolReadData As Boolean = FuncXRecords.getXRecords(plineParcels, dataPolyline)
                            If boolReadData = True Then
                                Dim PSObject As String = dataPolyline.TablePS
                                If PSObject Like "CADASTRAL_OSTROV" Or PSObject Like "CADASTRAL_OBJECT" Then
                                    Continue For
                                End If
                                deskObject = dataPolyline.Description
                                If deskObject Like descriptionObjectStyle Then
                                    Dim categoryEarth As String = ""
                                    Dim ConditionNumber As String = CadastralTools.JsonValidator.GetValueFromJson(dataPolyline.KeyParameter, "ConditionalNumber")
                                    If Delta <= 0 Then
                                        categoryEarth = CadastralTools.JsonValidator.GetValueFromJson(dataPolyline.KeyParameter, "Category")
                                    End If
                                    If IsNothing(ConditionNumber) = False Then
                                        If ConditionNumber.Trim.Length > 0 Then
                                            Dim numberParcel As Integer = CadastralTools.DataFormatting.GetNumberBeforeParenthesis(ConditionNumber)
                                            '=====================================================================================================
                                            If dictCadastralParcels.ContainsKey(numberParcel) = False Then
                                                ReDim Preserve ArrayPoligon(3, 0)
                                                ArrayPoligon(0, 0) = numberParcel 'номер участка
                                                ArrayPoligon(1, 0) = plineParcels.ObjectID 'хангл полилинии
                                                ArrayPoligon(2, 0) = dataPolyline.TablePS
                                                If categoryEarth.Trim.Length > 0 Then
                                                    Delta = CadastralTools.CadastralInfo.getDeltaPositionPointByCategoryEarth(categoryEarth)
                                                End If
                                                ArrayPoligon(3, 0) = Delta
                                                dictCadastralParcels.Add(numberParcel, ArrayPoligon)
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
                '==================================================================================================
                If dictCadastralParcels.Count = 0 Then
                    ActivDocument.EndUpdate()
                    MsgBox("Массив с участками пуст!!!")
                    Exit Sub
                Else
                    If dictCadastralParcels.Count > 1 Then
                        dictCadastralParcels = dictCadastralParcels.OrderBy(Function(kv) kv.Key).ToDictionary(Function(kv) kv.Key, Function(kv) kv.Value)
                    End If
                End If
                '====================================================================================================
                'выбираем полигоны
                Dim dictCadastralObject As New Dictionary(Of String, List(Of Geometries.Geometry))
                '====================================================================================================
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'массив с участками
                Dim classDrawObject As CreateDWGObject = New CreateDWGObject()
                classDrawObject.ActiveDrawing = ActivDocument
                For i As Integer = 0 To dictCadastralParcels.Count - 1
                    ArrayPoligon = dictCadastralParcels.ElementAt(i).Value
                    'получаем доступ к полилинии для извлечения координат
                    Dim numberParcel As Integer = Val(ArrayPoligon(0, 0))
                    Dim handlePLine As String = ArrayPoligon(1, 0)
                    Dim tablePS As String = ArrayPoligon(2, 0)
                    Delta = ArrayPoligon(3, 0)
                    If dictCadastralObject.ContainsKey(tablePS) = False Then
                        dictCadastralObject = SelectCadastralObject.selectCadastralObject(ActivDocument, tablePS)
                    End If
                    'составляем участок
                    If dictCadastralObject.ContainsKey(tablePS) = True Then
                        Dim listParcels As List(Of NetTopologySuite.Geometries.Geometry) = dictCadastralObject.Item(tablePS)
                        Dim listOstrov As New List(Of NetTopologySuite.Geometries.Geometry)
                        If dictCadastralObject.ContainsKey("CADASTRAL_OSTROV") = True Then
                            listOstrov = dictCadastralObject.Item("CADASTRAL_OSTROV")
                        End If
                        Dim dictCounters As Dictionary(Of Integer, NetTopologySuite.Geometries.Geometry) = CadastralTools.CadastralPolygon.getCountersByNumber(numberParcel, listParcels)
                        If dictCounters.Count > 0 Then
                            For j As Integer = 0 To dictCounters.Count - 1
                                Dim tempCounter As NetTopologySuite.Geometries.Geometry = dictCounters.ElementAt(j).Value
                                Dim triplets As New List(Of (VertexIndex As Integer, Prev As Coordinate, Current As Coordinate, [Next] As Coordinate))
                                Dim dataCounter As StructureElement = tempCounter.UserData
                                If TypeOf tempCounter Is NetTopologySuite.Geometries.Polygon Then
                                    Dim counter As NetTopologySuite.Geometries.Polygon = tempCounter
                                    dataCounter = counter.UserData
                                    counter = CadastralTools.FuncNetTopologySuite.RemoveDuplicateVertices(counter)
                                    If boolSVAngle = True Then
                                        counter = CadastralTools.FuncNetTopologySuite.NorthWestVertexFirst(counter)
                                    End If
                                    If boolClock = True Then
                                        counter = CadastralTools.FuncNetTopologySuite.ReversePolygonPreserveOrientation(counter)
                                    End If
                                    If boolRoundVertex = True Then
                                        counter = CadastralTools.FuncNetTopologySuite.RoundPolygonVertices(counter, tolerance)
                                    End If
                                    Dim boolUpdateVertexPolyline As Boolean = classDrawObject.UpdatePolylineFromPolygon(counter)
                                    triplets = CadastralTools.CadastralPoint.GetPolygonVertexTriplets(counter)
                                ElseIf TypeOf tempCounter Is NetTopologySuite.Geometries.LineString Then
                                    Dim counter As NetTopologySuite.Geometries.LineString = tempCounter
                                    dataCounter = counter.UserData
                                    counter = CadastralTools.FuncNetTopologySuite.RemoveDuplicateVertices(counter)
                                    If boolRoundVertex = True Then
                                        counter = CadastralTools.FuncNetTopologySuite.RoundLineStringVertices(counter, tolerance)
                                    End If
                                    Dim boolUpdateVertexPolyline As Boolean = classDrawObject.UpdatePolylineFromLineString(counter)
                                    triplets = CadastralTools.CadastralPoint.GetLineStringVertexTriplets(counter)
                                End If

                                If triplets.Count > 0 Then
                                    For Each t In triplets
                                        Dim insAttr As Double() = {}
                                        Dim Vert0 As NetTopologySuite.Geometries.Point = New NetTopologySuite.Geometries.Point(t.Prev.X, t.Prev.Y)
                                        Dim Vert1 As NetTopologySuite.Geometries.Point = New NetTopologySuite.Geometries.Point(t.Current.X, t.Current.Y)
                                        Dim Vert2 As NetTopologySuite.Geometries.Point = New NetTopologySuite.Geometries.Point(t.Next.X, t.Next.Y)
                                        'ищем уже существующую точку
                                        If boolDoublePoint = True Then
                                            Dim boolFindPoint As Boolean = CadastralTools.CadastralPoint.HasPoint(Vert1, listBlockReference, "", tolerance)
                                            If boolFindPoint = True Then
                                                Continue For
                                            End If
                                        End If
                                        Dim boolInsAttr As Boolean = CadastralTools.CadastralPoint.calculatePositionAttributeBlock(Vert0, Vert1, Vert2, insAttr, DText, True)
                                        'заполняем атрибуты
                                        Dim listAttributes As List(Of AttributeCadastralBlock) = New List(Of AttributeCadastralBlock)
                                        For k As Integer = 1 To arrayAttributes.GetUpperBound(0)
                                            Dim attribyteNumberPoint As AttributeCadastralBlock = New AttributeCadastralBlock()
                                            'номер
                                            attribyteNumberPoint.TagAttribute = arrayAttributes(k, 0)
                                            attribyteNumberPoint.DefinitionAttribute = arrayAttributes(k, 1)
                                            If attribyteNumberPoint.TagAttribute Like nameAttributeNumber Then
                                                attribyteNumberPoint.TextStringAttribute = prefNumber & LargNumber
                                                attribyteNumberPoint.HeightAttribute = HText
                                                attribyteNumberPoint.XAttribute = insAttr(0)
                                                attribyteNumberPoint.YAttribute = insAttr(1)
                                                attribyteNumberPoint.HorizontalAlignment = insAttr(2)
                                                attribyteNumberPoint.VerticalAlignment = insAttr(3)
                                            ElseIf attribyteNumberPoint.TagAttribute Like "DELTA" Then
                                                attribyteNumberPoint.TextStringAttribute = Delta
                                            ElseIf attribyteNumberPoint.TagAttribute Like "FASTDESK" Then
                                                attribyteNumberPoint.TextStringAttribute = Zakrep
                                            ElseIf attribyteNumberPoint.TagAttribute Like "METHODCOORDINATES" Then
                                                attribyteNumberPoint.TextStringAttribute = MetodCoord
                                            End If
                                            listAttributes.Add(attribyteNumberPoint)
                                        Next k
                                        'вставляем блок
                                        Dim insertPointBlk As Vector3D = New Vector3D(t.Current.X, t.Current.Y, 0)
                                        Dim insertBlk As DwgInsert = classDrawObject.insertBlockToDrawing2(insertPointBlk, styleBlockReference, listAttributes)
                                        If IsNothing(insertBlk) = False Then
                                            Dim boolAttrib As Boolean = AttributeCadastralBlock.OverwriteAttributeBlk(insertBlk, listAttributes)
                                            Dim dataPoint As CadastralTools.StructureElement = CadastralTools.CadastralPoint.createDataPoint(categoryZN, tableZn, descrittionBlockStyle, dataCounter.IdElement, LargNumber, prefNumber, Delta, Zakrep, MetodCoord, insertBlk.Handle.ToString)
                                            Dim dictData As Dictionary(Of String, Object) = CadastralTools.JsonValidator.DeserializeJsonToDictionary(dataPoint.KeyParameter)
                                            Dim recData As Boolean = FuncXRecords.setXRecords(insertBlk, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataPoint)
                                        End If
                                        Vert1.UserData = LargNumber
                                        listBlockReference.Add(insertBlk.Handle.ToString, Vert1)
                                        LargNumber += 1
                                    Next
                                End If
                            Next j
                        End If
                    End If
                Next i
            Catch ex As System.Exception
                MsgBox("Ошибка!!! " & ex.ToString)
            End Try
            ActivDocument.EndUpdate()
        End Sub
        '==============================================================================================================================
        'Перенумеровать полигон
        <cmd("ReNumberPolygon")>
        Public Sub ReNumberPolygon()
            'получаем активное окно
            Dim cadWiever As CadView = Me.CadView
            If IsNothing(cadWiever) Then Return
            'получаем активный слой
            Dim DrawLayer As DrawingLayer = DrawingLayer.GetDrawingLayer(cadWiever)
            If IsNothing(DrawLayer) = True Then Return
            'получаем активный документ
            Dim ActivDocument As Topomatic.Dwg.Drawing = DrawLayer.Drawing
            If IsNothing(ActivDocument) = True Then Return
            'сброс выделения
            cadWiever.SelectionSet.Clear()
            '1.\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'находим файл с шаблоном
            '1.\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Dim generalDir As String = ""
            Dim templateXML As String = ""
            Dim templateDWG As String = ""
            Dim arrayDirSupport As String() = Nothing
            Dim boolFindDirSupport As Boolean = FuncFiles.funcReadDirectoriesSupport(arrayDirSupport)
            If IsArray(arrayDirSupport) = True Then
                For i As Integer = 0 To arrayDirSupport.Length - 1
                    Dim nameFolder As String = arrayDirSupport(i)
                    Dim pos As Integer = nameFolder.IndexOf("InfrastradaToolsUtility")
                    If pos > -1 Then
                        generalDir = Mid(nameFolder, 1, pos)
                        generalDir = generalDir & "InfrastradaToolsUtility"
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
                        Dim shortNameFile As String = System.IO.Path.GetFileNameWithoutExtension(templateXML)
                        Dim tempNameFileDwg As String = directorySupport & shortNameFile & ".dwg"
                        If System.IO.File.Exists(tempNameFileDwg) = True Then
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
            'заносим шаблоны в массив
            Dim arrayTemlates As String() = {templateXML, templateDWG}
            'форма инициализируем
            Dim formInsertPoint As FormInsertPointPPT = New FormInsertPointPPT(cadWiever)
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'запускаем форму
            formInsertPoint.ShowDialog()
            If formInsertPoint.boolShow = False Then
                Exit Sub
            End If
            'масштаб чертежа
            Dim M As Double = cadWiever.AnnotationScale
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'смотрим характеристики блока
            Dim categoryPoint As String = formInsertPoint.ComboBox4.Text.Trim
            Dim tablePoint As String = formInsertPoint.ComboBox1.Text.Trim
            Dim deskPoint As String = formInsertPoint.ComboBox5.Text.Trim
            Dim arrayPropertiesObject As String(,) = Nothing
            Dim boolPropObject As Boolean = FuncXML.FuncReadDataObjectToXMLbyDesk(templateXML, categoryPoint, tablePoint, "Блок", deskPoint, arrayPropertiesObject)
            '========================================================================
            'начальные установки
            'блок
            Dim namePoint As String = formInsertPoint.ComboBox2.Text
            Dim blkPoint As DwgBlock = RoburFunc.FuncReadBlockByName(ActivDocument, namePoint)
            If IsNothing(blkPoint) = True Then
                blkPoint = RoburFunc.FuncInsertBlockByDWGFile(ActivDocument, templateDWG, namePoint)
            End If
            If IsNothing(blkPoint) = True Then
                MsgBox("Блок не найден!!! Макро прервано!!!")
                Exit Sub
            End If
            'слой
            Dim nameLayerPoint As String = formInsertPoint.ComboBox3.Text
            Dim layerPoint As DwgLayer = RoburFunc.FuncFindLayerDwg(ActivDocument, nameLayerPoint)
            If IsNothing(layerPoint) = True Then
                'layerPoint = FuncDecorPlane.FuncCreateLayerByXML(ActivDocument, templateXML, nameLayerPoint)
            End If
            If IsNothing(layerPoint) = True Then
                layerPoint = ActivDocument.ActiveLayer
            End If
            'тип линии
            Dim lineTypePoint As DwgLinetype = ActivDocument.Linetypes.ByLayer
            'цвет
            Dim indStrColor As String = formInsertPoint.TextBox2.Text
            If indStrColor Like "ПоСлою" Then
                indStrColor = "256"
            ElseIf indStrColor Like "ПоБлоку" Then
                indStrColor = "0"
            End If
            Dim indColor As Integer = Val(indColor)
            If indColor < 0 Then
                indColor = 0
            ElseIf indColor > 256 Then
                indColor = 256
            End If
            Dim colorPoint As CadColor = New CadColor(indColor)
            Dim scaleTypeLinePoint As Double = 1
            Dim widthLinePoint As Double = 20
            Dim scalePoint As Double = formInsertPoint.NumericUpDown1.Value
            If scalePoint <= 0 Then
                scalePoint = 1
            End If
            'поворот блока
            Dim anglePoint As Double = 0
            'размер атрибута
            Dim attribHeight As Double = formInsertPoint.NumericUpDown2.Value
            If attribHeight <= 0 Then
                attribHeight = 1.5
            End If
            'имя атрибута
            Dim attributeTag As String = formInsertPoint.ComboBox6.Text
            If IsNothing(attributeTag) = True Then
                MsgBox("Не верно задано имя атрибута!!!")
                Exit Sub
            End If
            'таблица PS
            Dim PSTable As String = formInsertPoint.ComboBox7.Text
            'примечание
            Dim note As String = FuncXML.FuncReturnValueByArrayProperties("NOTE", arrayPropertiesObject)
            'префикс номера
            Dim prefNumber As String = formInsertPoint.ComboBox8.Text
            'номер
            Dim startNumber As Integer = formInsertPoint.NumericUpDown4.Value
            'положение атрибута
            Dim posAttribX As Double = 1.5
            Dim posAttribY As Double = -1 * (attribHeight / 2)
            'забираем точки в массив
            Dim arrayPointCollections As String(,) = Nothing
            Dim numberPointCollections As Integer = CadastralFunction.FuncFindCadastralPoints(ActivDocument, namePoint, arrayPointCollections)
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Dim arrayParcels As String() = Nothing
            Dim countArrayParcels As Integer = 0
            Dim [userSelect] As String = "Все"
            CadCursors.GetUserSelect(cadWiever, [userSelect], Nothing, "Перенумеровать объекты:", "Все", "Выбрать")
            Select Case [userSelect]
                Case "Все"
                    cadWiever.SelectionSet.Clear()
                    cadWiever.SelectionSet.SelectOneObjectAtScreen(Function(obj) TypeOf obj Is DwgPolyline, "Выберите полилинию для образца: ")
                    Dim typeParcels As String = ""
                    For Each acEnt As Object In cadWiever.SelectionSet
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        If TypeOf acEnt Is DwgPolyline Then
                            typeParcels = FuncXRecords.FuncReadValueXData(acEnt, "LABEL")
                            If IsNothing(typeParcels) = False Then
                                Exit For
                            End If
                        End If
                    Next
                    If typeParcels.Trim.Length > 0 Then
                        For Each acEnt As DwgEntity In ActivDocument.ActiveSpace.Entities
                            Try
                                If TypeOf acEnt Is DwgPolyline Then
                                    Dim typeParcelUser As String = FuncXRecords.FuncReadValueXData(acEnt, "LABEL")
                                    If typeParcelUser.Trim Like typeParcels.Trim Then
                                        ReDim Preserve arrayParcels(countArrayParcels)
                                        arrayParcels(countArrayParcels) = acEnt.ObjectID
                                        countArrayParcels += 1
                                    End If
                                End If
                            Catch ex As System.Exception
                            End Try
                        Next
                    End If
                Case "Выбрать"
                    cadWiever.SelectionSet.Clear()
                    cadWiever.SelectionSet.SelectOneObjectAtScreen(Function(obj) TypeOf obj Is DwgPolyline, "Выборите полилинии которые необходимо пронумеровать: ")
                    For Each acEnt As Object In cadWiever.SelectionSet
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        If TypeOf acEnt Is DwgPolyline Then
                            Dim localIDParcel As String = FuncXRecords.FuncReadValueXData(acEnt, "LOCALID")
                            If localIDParcel.Trim.Length > 0 Then
                                ReDim Preserve arrayParcels(countArrayParcels)
                                arrayParcels(countArrayParcels) = acEnt.ObjectID
                                countArrayParcels += 1
                            End If
                        End If
                    Next
            End Select
            If IsArray(arrayParcels) = True Then
                '=======================================================================================================
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                Dim arrayObject As String(,) = Nothing '0-Номер участка,1-handle,2-GLobalID,2-LocalID,3-ha
                Dim countArrayObject As Integer = 0
                Dim labelParcel As String = "Земельный участок"
                For i As Integer = 0 To arrayParcels.Length - 1
                    Dim hgCounters As String = arrayParcels(i)
                    Dim acPolyCounters As DwgObject = Nothing
                    Dim boolFindObject As Boolean = ActivDocument.ActiveSpace.Entities.TryGetObject(hgCounters, acPolyCounters)
                    If acPolyCounters.HasExtensionDictionary = True Then
                        Dim LABEL As String = FuncXRecords.FuncReadValueXData(acPolyCounters, "LABEL")
                        If LABEL Like "Внутренний островок" Then
                            Continue For
                        Else
                            Dim conditionalParcel As String = FuncXRecords.FuncReadValueXData(acPolyCounters, "ConditionalNumber")
                            Dim pos1 As Integer = conditionalParcel.IndexOf("(")
                            Dim numberParcel As String = FuncFormatZn.FuncFindFullConditionalNumberCadastralObject(conditionalParcel)
                            Dim globalID As String = FuncXRecords.FuncReadValueXData(acPolyCounters, "GLOBALID")
                            Dim localID As String = FuncXRecords.FuncReadValueXData(acPolyCounters, "LOCALID")
                            Dim boolFindGLOBALID As Boolean = False
                            If globalID.Trim.Length > 2 Then
                                boolFindGLOBALID = MathFunction.FuncFindValueToArray(globalID, 2, arrayObject)
                            End If
                            If boolFindGLOBALID = False Then
                                ReDim Preserve arrayObject(4, countArrayObject)
                                arrayObject(0, countArrayObject) = Val(numberParcel)
                                arrayObject(1, countArrayObject) = acPolyCounters.ObjectID
                                arrayObject(2, countArrayObject) = globalID
                                arrayObject(3, countArrayObject) = localID
                                arrayObject(4, countArrayObject) = LABEL
                                countArrayObject += 1
                            End If
                        End If
                    End If
                Next i
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'участки найдены начинаем нумеровать
                If IsArray(arrayObject) = True Then
                    If arrayObject.GetUpperBound(1) > 0 Then
                        Dim boolSortArray As Boolean = MathFunction.FuncSortStrArray(arrayObject, 0)
                    End If
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'обрабатываем участки
                    Dim arrayParcel As String(,) = Nothing
                    Dim countArrayParcel As Integer = 0
                    '===============================================================================================================
                    For i As Integer = 0 To arrayObject.GetUpperBound(1)
                        'сбрасываем номер, еси требуется
                        If formInsertPoint.CheckBox5.Checked = False Then
                            startNumber = formInsertPoint.NumericUpDown4.Value
                        End If
                        'берем объект
                        Dim hgFirstCounters As String = arrayObject(1, i)
                        'ищем контура
                        Dim globalID As String = arrayObject(2, i) 'глобальный ID
                        Dim localIDParcel As String = arrayObject(3, i) 'глобальный ID
                        If globalID.Trim.Length > 2 Then
                            'ищем входящие контура
                            Dim arrayDataCounters As String(,) = Nothing
                            Dim countCounters As Integer = CadastralFunction.FuncFindCadastralCounter(ActivDocument, globalID, arrayDataCounters)
                            If IsArray(arrayDataCounters) = True Then
                                For j As Integer = 0 To arrayDataCounters.GetUpperBound(1)
                                    Dim hgCounters As String = arrayDataCounters(1, j)
                                    Dim localIDCounter As String = arrayDataCounters(3, j) 'локальный ID
                                    If IsNothing(hgCounters) = False Then
                                        If hgCounters.Trim.Length > 0 Then
                                            Dim acPolyCounters As DwgObject = Nothing
                                            Dim boolFindObject As Boolean = ActivDocument.ActiveSpace.Entities.TryGetObject(hgCounters, acPolyCounters)
                                            If boolFindObject = True Then
                                                If TypeOf acPolyCounters Is DwgPolyline Then
                                                    Dim acPoly As DwgPolyline = acPolyCounters
                                                    'удаляем совпадающие вершины
                                                    Dim boolDelDubleVertex As Boolean = MathFunction.FuncDeleteDubleVertexPline(acPoly, 0.01)
                                                    'округляем вершины до см
                                                    If formInsertPoint.CheckBox1.Checked = True Then
                                                        Dim boolroundVertex As Boolean = MathFunction.FuncRoundVertexPline(acPoly, 2)
                                                    End If
                                                    'по часовой стрелке, делаем порядок следования вершин
                                                    If formInsertPoint.CheckBox4.Checked = True Then
                                                        Dim areaPline As Double = MathFunction.GetArea(acPoly, 2)
                                                        If areaPline > 0 Then
                                                            Dim boolClockReverseVertex As Boolean = MathFunction.FuncReversVertexPline(acPoly)
                                                        End If
                                                    End If
                                                    'северо-западный угол
                                                    If formInsertPoint.CheckBox3.Checked = True Then
                                                        Dim boolNEAngle As Boolean = MathFunction.FuncNorthWestAnglePline(acPoly)
                                                    End If
                                                    'категория земель
                                                    Dim category As String = FuncXRecords.FuncReadValueXData(acPoly, "Category")
                                                    Dim delta_geopoint As Double = CadastralFunction.FuncReturnDeltaGeopointByCategory(category)
                                                    'нумеруем полигон
                                                    Dim countVert As Integer = acPoly.Count
                                                    For k As Integer = 0 To countVert - 1
                                                        Dim x As Double = acPoly.ElementAt(k).Vertex.X
                                                        Dim y As Double = acPoly.ElementAt(k).Vertex.Y
                                                        Dim posInsertBlk As Vector2D = New Vector2D(x, y)
                                                        'проверяем на наличие блока
                                                        If IsArray(arrayPointCollections) = True Then
                                                            Dim idPoint As String = CadastralFunction.FuncFindCadastralPointToVertex(ActivDocument, posInsertBlk, arrayPointCollections, 0.01)
                                                            If idPoint.Trim.Length > 0 Then
                                                                'делаем замену номера
                                                                Dim blkPointUser As DwgInsert = Nothing
                                                                Dim boolFindPoint As Boolean = ActivDocument.ActiveSpace.Entities.TryGetObject(idPoint, blkPointUser)
                                                                If boolFindPoint = True Then
                                                                    Dim idRng As String = FuncXRecords.FuncReadValueXData(blkPointUser, "IDRNG")
                                                                    If idRng.Trim Like localIDCounter Then
                                                                        'назначаем номер точке
                                                                        Dim oboznPoint As String = prefNumber & startNumber
                                                                        'формируем артрибут
                                                                        Dim arrayData As String(,) = Nothing
                                                                        ReDim arrayData(1, 0)
                                                                        arrayData(0, 0) = attributeTag
                                                                        arrayData(1, 0) = oboznPoint
                                                                        Dim boolreplaseBlk As Boolean = RoburFunc.FuncOverwriteAttributeBlk(ActivDocument, blkPointUser, arrayData)
                                                                        Dim arrayRec As String(,) = Nothing
                                                                        ReDim Preserve arrayRec(1, 0)
                                                                        arrayRec(0, 0) = "Num_geopoint"
                                                                        arrayRec(1, 0) = oboznPoint
                                                                        If IsArray(arrayRec) = True Then
                                                                            Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXData(blkPointUser, PSTable, arrayRec, templateXML)
                                                                        End If
                                                                        startNumber += 1
                                                                    End If
                                                                End If
                                                            Else
                                                                'назначаем номер точке
                                                                Dim oboznPoint As String = prefNumber & startNumber
                                                                'формируем артрибут
                                                                Dim dictionaryAttr As Dictionary(Of String, String(,)) = New Dictionary(Of String, String(,))
                                                                Dim arrayAttrb As String(,) = Nothing
                                                                ReDim arrayAttrb(1, 3)
                                                                arrayAttrb(0, 0) = "value"
                                                                arrayAttrb(1, 0) = oboznPoint
                                                                arrayAttrb(0, 1) = "height"
                                                                arrayAttrb(1, 1) = attribHeight
                                                                arrayAttrb(0, 2) = "positionX"
                                                                arrayAttrb(1, 2) = posAttribX
                                                                arrayAttrb(0, 3) = "positionY"
                                                                arrayAttrb(1, 3) = posAttribY
                                                                dictionaryAttr.Add(attributeTag, arrayAttrb)
                                                                Dim insertBlk As DwgInsert = CadastralFunction.FuncInsertCadastralPoint(ActivDocument, namePoint, posInsertBlk, dictionaryAttr, anglePoint)
                                                                If IsNothing(insertBlk) = False Then
                                                                    insertBlk.Layer = layerPoint
                                                                    insertBlk.Color = colorPoint
                                                                    insertBlk.Linetype = lineTypePoint
                                                                    insertBlk.Lineweight = widthLinePoint
                                                                    insertBlk.Scale = New Vector3D(scalePoint, scalePoint, scalePoint)
                                                                    Dim arrayRec As String(,) = Nothing
                                                                    ReDim Preserve arrayRec(1, 5)
                                                                    arrayRec(0, 0) = "LABEL"
                                                                    arrayRec(1, 0) = deskPoint
                                                                    arrayRec(0, 1) = "Num_geopoint"
                                                                    arrayRec(1, 1) = oboznPoint
                                                                    arrayRec(0, 2) = "Delta_geopoint"
                                                                    arrayRec(1, 2) = delta_geopoint
                                                                    arrayRec(0, 3) = "GeopointZacrep"
                                                                    arrayRec(1, 3) = formInsertPoint.ComboBox10.Text
                                                                    arrayRec(0, 4) = "Geopoint_opred"
                                                                    arrayRec(1, 4) = formInsertPoint.ComboBox9.Text
                                                                    arrayRec(0, 5) = "IDRNG"
                                                                    arrayRec(1, 5) = localIDCounter
                                                                    If IsArray(arrayRec) = True Then
                                                                        Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXData(insertBlk, PSTable, arrayRec, templateXML)
                                                                    End If
                                                                    Dim countColl As Integer = arrayPointCollections.GetUpperBound(1) + 1
                                                                    ReDim Preserve arrayPointCollections(3, countColl)
                                                                    arrayPointCollections(0, countColl) = x
                                                                    arrayPointCollections(1, countColl) = y
                                                                    arrayPointCollections(2, countColl) = insertBlk.ObjectID
                                                                    arrayPointCollections(3, countColl) = localIDCounter
                                                                    startNumber += 1
                                                                End If
                                                            End If
                                                        End If
                                                    Next k
                                                    'ищем внутренние островки
                                                    Dim arrayDataOstrov As String(,) = Nothing
                                                    Dim countOstrov As Integer = CadastralFunction.FuncFindCadastralOstrov(ActivDocument, acPoly, localIDCounter, arrayDataOstrov)
                                                    If IsArray(arrayDataOstrov) = True Then
                                                        For k As Integer = 0 To arrayDataOstrov.GetUpperBound(1)
                                                            Dim hgOstrov As String = arrayDataOstrov(1, k)
                                                            Dim localIDOstrov As String = arrayDataOstrov(3, k)
                                                            If IsNothing(hgOstrov) = False Then
                                                                Dim objOstrov As DwgObject = Nothing
                                                                Dim findOstrov As Boolean = ActivDocument.ActiveSpace.Entities.TryGetObject(hgOstrov, objOstrov)
                                                                If findOstrov = True Then
                                                                    If TypeOf objOstrov Is DwgPolyline Then
                                                                        Dim acPolyOstrov As DwgPolyline = objOstrov
                                                                        'удаляем совпадающие вершины
                                                                        Dim boolDelDubleVertexOstrov As Boolean = MathFunction.FuncDeleteDubleVertexPline(acPolyOstrov, 0.01)
                                                                        'округляем вершины до см
                                                                        If formInsertPoint.CheckBox1.Checked = True Then
                                                                            Dim boolroundVertex As Boolean = MathFunction.FuncRoundVertexPline(acPolyOstrov, 2)
                                                                        End If
                                                                        'по часовой стрелке, делаем порядок следования вершин
                                                                        If formInsertPoint.CheckBox4.Checked = True Then
                                                                            Dim areaPline As Double = MathFunction.GetArea(acPolyOstrov, 2)
                                                                            If areaPline > 0 Then
                                                                                Dim boolClockReverseVertex As Boolean = MathFunction.FuncReversVertexPline(acPolyOstrov)
                                                                            End If
                                                                        End If
                                                                        'северо-западный угол
                                                                        If formInsertPoint.CheckBox3.Checked = True Then
                                                                            Dim boolNEAngle As Boolean = MathFunction.FuncNorthWestAnglePline(acPolyOstrov)
                                                                        End If
                                                                        For k1 As Integer = 0 To acPolyOstrov.Count - 1
                                                                            Dim x As Double = acPolyOstrov.ElementAt(k1).Vertex.X
                                                                            Dim y As Double = acPolyOstrov.ElementAt(k1).Vertex.Y
                                                                            Dim posInsertBlk As Vector2D = New Vector2D(x, y)
                                                                            'проверяем на наличие блока
                                                                            If IsArray(arrayPointCollections) = True Then
                                                                                Dim idPoint As String = CadastralFunction.FuncFindCadastralPointToVertex(ActivDocument, posInsertBlk, arrayPointCollections, 0.01)
                                                                                If idPoint.Trim.Length > 0 Then
                                                                                    'делаем замену номера
                                                                                    Dim blkPointUser As DwgInsert = Nothing
                                                                                    Dim boolFindPoint As Boolean = ActivDocument.ActiveSpace.Entities.TryGetObject(idPoint, blkPointUser)
                                                                                    If boolFindPoint = True Then
                                                                                        Dim idRng As String = FuncXRecords.FuncReadValueXData(blkPointUser, "IDRNG")
                                                                                        If idRng.Trim Like localIDOstrov Then
                                                                                            'назначаем номер точке
                                                                                            Dim oboznPoint As String = prefNumber & startNumber
                                                                                            'формируем артрибут
                                                                                            Dim arrayData As String(,) = Nothing
                                                                                            ReDim arrayData(1, 0)
                                                                                            arrayData(0, 0) = attributeTag
                                                                                            arrayData(1, 0) = oboznPoint
                                                                                            Dim boolreplaseBlk As Boolean = RoburFunc.FuncOverwriteAttributeBlk(ActivDocument, blkPointUser, arrayData)
                                                                                            Dim arrayRec As String(,) = Nothing
                                                                                            ReDim Preserve arrayRec(1, 0)
                                                                                            arrayRec(0, 0) = "Num_geopoint"
                                                                                            arrayRec(1, 0) = oboznPoint
                                                                                            If IsArray(arrayRec) = True Then
                                                                                                Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXData(blkPointUser, PSTable, arrayRec, templateXML)
                                                                                            End If
                                                                                            startNumber += 1
                                                                                        End If
                                                                                    End If
                                                                                Else
                                                                                    'назначаем номер точке
                                                                                    Dim oboznPoint As String = prefNumber & startNumber
                                                                                    'формируем артрибут
                                                                                    Dim dictionaryAttr As Dictionary(Of String, String(,)) = New Dictionary(Of String, String(,))
                                                                                    Dim arrayAttrb As String(,) = Nothing
                                                                                    ReDim arrayAttrb(1, 3)
                                                                                    arrayAttrb(0, 0) = "value"
                                                                                    arrayAttrb(1, 0) = oboznPoint
                                                                                    arrayAttrb(0, 1) = "height"
                                                                                    arrayAttrb(1, 1) = attribHeight
                                                                                    arrayAttrb(0, 2) = "positionX"
                                                                                    arrayAttrb(1, 2) = posAttribX
                                                                                    arrayAttrb(0, 3) = "positionY"
                                                                                    arrayAttrb(1, 3) = posAttribY
                                                                                    dictionaryAttr.Add(attributeTag, arrayAttrb)
                                                                                    Dim insertBlk As DwgInsert = CadastralFunction.FuncInsertCadastralPoint(ActivDocument, namePoint, posInsertBlk, dictionaryAttr, anglePoint)
                                                                                    If IsNothing(insertBlk) = False Then
                                                                                        insertBlk.Layer = layerPoint
                                                                                        insertBlk.Color = colorPoint
                                                                                        insertBlk.Linetype = lineTypePoint
                                                                                        insertBlk.Lineweight = widthLinePoint
                                                                                        insertBlk.Scale = New Vector3D(scalePoint, scalePoint, scalePoint)
                                                                                        Dim arrayRec As String(,) = Nothing
                                                                                        ReDim Preserve arrayRec(1, 5)
                                                                                        arrayRec(0, 0) = "LABEL"
                                                                                        arrayRec(1, 0) = deskPoint
                                                                                        arrayRec(0, 1) = "Num_geopoint"
                                                                                        arrayRec(1, 1) = oboznPoint
                                                                                        arrayRec(0, 2) = "Delta_geopoint"
                                                                                        arrayRec(1, 2) = delta_geopoint
                                                                                        arrayRec(0, 3) = "GeopointZacrep"
                                                                                        arrayRec(1, 3) = formInsertPoint.ComboBox10.Text
                                                                                        arrayRec(0, 4) = "Geopoint_opred"
                                                                                        arrayRec(1, 4) = formInsertPoint.ComboBox9.Text
                                                                                        arrayRec(0, 5) = "IDRNG"
                                                                                        arrayRec(1, 5) = localIDCounter
                                                                                        If IsArray(arrayRec) = True Then
                                                                                            Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXData(insertBlk, PSTable, arrayRec, templateXML)
                                                                                        End If
                                                                                        Dim countColl As Integer = arrayPointCollections.GetUpperBound(1) + 1
                                                                                        ReDim Preserve arrayPointCollections(3, countColl)
                                                                                        arrayPointCollections(0, countColl) = x
                                                                                        arrayPointCollections(1, countColl) = y
                                                                                        arrayPointCollections(2, countColl) = insertBlk.ObjectID
                                                                                        arrayPointCollections(3, countColl) = localIDCounter
                                                                                        startNumber += 1
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                        Next k1
                                                                    End If
                                                                End If
                                                            End If
                                                        Next k
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                Next j
                            End If
                        Else
                            'обособленный участок
                            If formInsertPoint.CheckBox5.Checked = False Then
                                startNumber = formInsertPoint.NumericUpDown4.Value
                            End If
                            Dim acPolyCounters As DwgObject = Nothing
                            Dim boolFindObject As Boolean = ActivDocument.ActiveSpace.Entities.TryGetObject(hgFirstCounters, acPolyCounters)
                            If boolFindObject = True Then
                                If TypeOf acPolyCounters Is DwgPolyline Then
                                    Dim acPoly As DwgPolyline = acPolyCounters
                                    'удаляем совпадающие вершины
                                    Dim boolDelDubleVertex As Boolean = MathFunction.FuncDeleteDubleVertexPline(acPoly, 0.01)
                                    'округляем вершины до см
                                    If formInsertPoint.CheckBox1.Checked = True Then
                                        Dim boolroundVertex As Boolean = MathFunction.FuncRoundVertexPline(acPoly, 2)
                                    End If
                                    'по часовой стрелке, делаем порядок следования вершин
                                    If formInsertPoint.CheckBox4.Checked = True Then
                                        Dim areaPline As Double = MathFunction.GetArea(acPoly, 2)
                                        If areaPline > 0 Then
                                            Dim boolClockReverseVertex As Boolean = MathFunction.FuncReversVertexPline(acPoly)
                                        End If
                                    End If
                                    'северо-западный угол
                                    If formInsertPoint.CheckBox3.Checked = True Then
                                        Dim boolNEAngle As Boolean = MathFunction.FuncNorthWestAnglePline(acPoly)
                                    End If
                                    'категория земель
                                    Dim category As String = FuncXRecords.FuncReadValueXData(acPoly, "Category")
                                    Dim delta_geopoint As Double = CadastralFunction.FuncReturnDeltaGeopointByCategory(category)
                                    'нумеруем полигон
                                    Dim countVert As Integer = acPoly.Count
                                    For k As Integer = 0 To countVert - 1
                                        Dim x As Double = acPoly.ElementAt(k).Vertex.X
                                        Dim y As Double = acPoly.ElementAt(k).Vertex.Y
                                        Dim posInsertBlk As Vector2D = New Vector2D(x, y)
                                        If IsArray(arrayPointCollections) = True Then
                                            Dim idPoint As String = CadastralFunction.FuncFindCadastralPointToVertex(ActivDocument, posInsertBlk, arrayPointCollections, 0.01)
                                            If idPoint.Trim.Length > 0 Then
                                                'делаем замену номера
                                                Dim blkPointUser As DwgInsert = Nothing
                                                Dim boolFindPoint As Boolean = ActivDocument.ActiveSpace.Entities.TryGetObject(idPoint, blkPointUser)
                                                If boolFindPoint = True Then
                                                    Dim idRng As String = FuncXRecords.FuncReadValueXData(blkPointUser, "IDRNG")
                                                    If idRng.Trim Like localIDParcel Then
                                                        'назначаем номер точке
                                                        Dim oboznPoint As String = prefNumber & startNumber
                                                        'формируем артрибут
                                                        Dim arrayData As String(,) = Nothing
                                                        ReDim arrayData(1, 0)
                                                        arrayData(0, 0) = attributeTag
                                                        arrayData(1, 0) = oboznPoint
                                                        Dim boolreplaseBlk As Boolean = RoburFunc.FuncOverwriteAttributeBlk(ActivDocument, blkPointUser, arrayData)
                                                        Dim arrayRec As String(,) = Nothing
                                                        ReDim Preserve arrayRec(1, 0)
                                                        arrayRec(0, 0) = "Num_geopoint"
                                                        arrayRec(1, 0) = oboznPoint
                                                        If IsArray(arrayRec) = True Then
                                                            Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXData(blkPointUser, PSTable, arrayRec, templateXML)
                                                        End If
                                                        startNumber += 1
                                                    End If
                                                End If
                                            Else
                                                'назначаем номер точке
                                                Dim oboznPoint As String = prefNumber & startNumber
                                                'формируем артрибут
                                                Dim dictionaryAttr As Dictionary(Of String, String(,)) = New Dictionary(Of String, String(,))
                                                Dim arrayAttrb As String(,) = Nothing
                                                ReDim arrayAttrb(1, 3)
                                                arrayAttrb(0, 0) = "value"
                                                arrayAttrb(1, 0) = oboznPoint
                                                arrayAttrb(0, 1) = "height"
                                                arrayAttrb(1, 1) = attribHeight
                                                arrayAttrb(0, 2) = "positionX"
                                                arrayAttrb(1, 2) = posAttribX
                                                arrayAttrb(0, 3) = "positionY"
                                                arrayAttrb(1, 3) = posAttribY
                                                dictionaryAttr.Add(attributeTag, arrayAttrb)
                                                Dim insertBlk As DwgInsert = CadastralFunction.FuncInsertCadastralPoint(ActivDocument, namePoint, posInsertBlk, dictionaryAttr, anglePoint)
                                                If IsNothing(insertBlk) = False Then
                                                    insertBlk.Layer = layerPoint
                                                    insertBlk.Color = colorPoint
                                                    insertBlk.Linetype = lineTypePoint
                                                    insertBlk.Lineweight = widthLinePoint
                                                    insertBlk.Scale = New Vector3D(scalePoint, scalePoint, scalePoint)
                                                    Dim arrayRec As String(,) = Nothing
                                                    ReDim Preserve arrayRec(1, 5)
                                                    arrayRec(0, 0) = "LABEL"
                                                    arrayRec(1, 0) = deskPoint
                                                    arrayRec(0, 1) = "Num_geopoint"
                                                    arrayRec(1, 1) = oboznPoint
                                                    arrayRec(0, 2) = "Delta_geopoint"
                                                    arrayRec(1, 2) = delta_geopoint
                                                    arrayRec(0, 3) = "GeopointZacrep"
                                                    arrayRec(1, 3) = formInsertPoint.ComboBox10.Text
                                                    arrayRec(0, 4) = "Geopoint_opred"
                                                    arrayRec(1, 4) = formInsertPoint.ComboBox9.Text
                                                    arrayRec(0, 5) = "IDRNG"
                                                    arrayRec(1, 5) = localIDParcel
                                                    If IsArray(arrayRec) = True Then
                                                        Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXData(insertBlk, PSTable, arrayRec, templateXML)
                                                    End If
                                                    Dim countColl As Integer = arrayPointCollections.GetUpperBound(1) + 1
                                                    ReDim Preserve arrayPointCollections(3, countColl)
                                                    arrayPointCollections(0, countColl) = x
                                                    arrayPointCollections(1, countColl) = y
                                                    arrayPointCollections(2, countColl) = insertBlk.ObjectID
                                                    arrayPointCollections(3, countColl) = localIDParcel
                                                    startNumber += 1
                                                End If

                                            End If
                                        End If
                                    Next k
                                    'ищем внутренние островки
                                    Dim arrayDataOstrov As String(,) = Nothing
                                    Dim localID As String = FuncXRecords.FuncReadValueXData(acPoly, "LOCALID")
                                    Dim countOstrov As Integer = CadastralFunction.FuncFindCadastralOstrov(ActivDocument, acPoly, localID, arrayDataOstrov)
                                    If IsArray(arrayDataOstrov) = True Then
                                        For k As Integer = 0 To arrayDataOstrov.GetUpperBound(1)
                                            Dim hgOstrov As String = arrayDataOstrov(1, k)
                                            Dim localIDOstrov As String = arrayDataOstrov(3, k)
                                            If IsNothing(hgOstrov) = False Then
                                                Dim objOstrov As DwgObject = Nothing
                                                Dim findOstrov As Boolean = ActivDocument.ActiveSpace.Entities.TryGetObject(hgOstrov, objOstrov)
                                                If findOstrov = True Then
                                                    If TypeOf objOstrov Is DwgPolyline Then
                                                        Dim acPolyOstrov As DwgPolyline = objOstrov
                                                        'удаляем совпадающие вершины
                                                        Dim boolDelDubleVertex1 As Boolean = MathFunction.FuncDeleteDubleVertexPline(acPolyOstrov, 0.01)
                                                        'округляем вершины до см
                                                        If formInsertPoint.CheckBox1.Checked = True Then
                                                            Dim boolroundVertex As Boolean = MathFunction.FuncRoundVertexPline(acPolyOstrov, 2)
                                                        End If
                                                        'по часовой стрелке, делаем порядок следования вершин
                                                        If formInsertPoint.CheckBox4.Checked = True Then
                                                            Dim areaPline As Double = MathFunction.GetArea(acPolyOstrov, 2)
                                                            If areaPline > 0 Then
                                                                Dim boolClockReverseVertex As Boolean = MathFunction.FuncReversVertexPline(acPolyOstrov)
                                                            End If
                                                        End If
                                                        'северо-западный угол
                                                        If formInsertPoint.CheckBox3.Checked = True Then
                                                            Dim boolNEAngle As Boolean = MathFunction.FuncNorthWestAnglePline(acPolyOstrov)
                                                        End If

                                                        For k1 As Integer = 0 To acPolyOstrov.Count - 1
                                                            Dim x As Double = acPolyOstrov.ElementAt(k1).Vertex.X
                                                            Dim y As Double = acPolyOstrov.ElementAt(k1).Vertex.Y
                                                            Dim posInsertBlk As Vector2D = New Vector2D(x, y)
                                                            If IsArray(arrayPointCollections) = True Then
                                                                Dim idPoint As String = CadastralFunction.FuncFindCadastralPointToVertex(ActivDocument, posInsertBlk, arrayPointCollections, 0.01)
                                                                If idPoint.Trim.Length > 0 Then
                                                                    'делаем замену номера
                                                                    Dim blkPointUser As DwgInsert = Nothing
                                                                    Dim boolFindPoint As Boolean = ActivDocument.ActiveSpace.Entities.TryGetObject(idPoint, blkPointUser)
                                                                    If boolFindPoint = True Then
                                                                        Dim idRng As String = FuncXRecords.FuncReadValueXData(blkPointUser, "IDRNG")
                                                                        If idRng.Trim Like localIDOstrov Then
                                                                            'назначаем номер точке
                                                                            Dim oboznPoint As String = prefNumber & startNumber
                                                                            'формируем артрибут
                                                                            Dim arrayData As String(,) = Nothing
                                                                            ReDim arrayData(1, 0)
                                                                            arrayData(0, 0) = attributeTag
                                                                            arrayData(1, 0) = oboznPoint
                                                                            Dim boolreplaseBlk As Boolean = RoburFunc.FuncOverwriteAttributeBlk(ActivDocument, blkPointUser, arrayData)
                                                                            Dim arrayRec As String(,) = Nothing
                                                                            ReDim Preserve arrayRec(1, 0)
                                                                            arrayRec(0, 1) = "Num_geopoint"
                                                                            arrayRec(1, 1) = oboznPoint
                                                                            If IsArray(arrayRec) = True Then
                                                                                Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXData(blkPointUser, PSTable, arrayRec, templateXML)
                                                                            End If
                                                                            startNumber += 1
                                                                        End If
                                                                    End If
                                                                Else
                                                                    'назначаем номер точке
                                                                    Dim oboznPoint As String = prefNumber & startNumber
                                                                    'формируем артрибут
                                                                    Dim dictionaryAttr As Dictionary(Of String, String(,)) = New Dictionary(Of String, String(,))
                                                                    Dim arrayAttrb As String(,) = Nothing
                                                                    ReDim arrayAttrb(1, 3)
                                                                    arrayAttrb(0, 0) = "value"
                                                                    arrayAttrb(1, 0) = oboznPoint
                                                                    arrayAttrb(0, 1) = "height"
                                                                    arrayAttrb(1, 1) = attribHeight
                                                                    arrayAttrb(0, 2) = "positionX"
                                                                    arrayAttrb(1, 2) = posAttribX
                                                                    arrayAttrb(0, 3) = "positionY"
                                                                    arrayAttrb(1, 3) = posAttribY
                                                                    dictionaryAttr.Add(attributeTag, arrayAttrb)
                                                                    Dim insertBlk As DwgInsert = CadastralFunction.FuncInsertCadastralPoint(ActivDocument, namePoint, posInsertBlk, dictionaryAttr, anglePoint)
                                                                    If IsNothing(insertBlk) = False Then
                                                                        insertBlk.Layer = layerPoint
                                                                        insertBlk.Color = colorPoint
                                                                        insertBlk.Linetype = lineTypePoint
                                                                        insertBlk.Lineweight = widthLinePoint
                                                                        insertBlk.Scale = New Vector3D(scalePoint, scalePoint, scalePoint)
                                                                        Dim arrayRec As String(,) = Nothing
                                                                        ReDim Preserve arrayRec(1, 5)
                                                                        arrayRec(0, 0) = "LABEL"
                                                                        arrayRec(1, 0) = deskPoint
                                                                        arrayRec(0, 1) = "Num_geopoint"
                                                                        arrayRec(1, 1) = oboznPoint
                                                                        arrayRec(0, 2) = "Delta_geopoint"
                                                                        arrayRec(1, 2) = delta_geopoint
                                                                        arrayRec(0, 3) = "GeopointZacrep"
                                                                        arrayRec(1, 3) = formInsertPoint.ComboBox10.Text
                                                                        arrayRec(0, 4) = "Geopoint_opred"
                                                                        arrayRec(1, 4) = formInsertPoint.ComboBox9.Text
                                                                        arrayRec(0, 5) = "IDRNG"
                                                                        arrayRec(1, 5) = localID
                                                                        If IsArray(arrayRec) = True Then
                                                                            Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXData(insertBlk, PSTable, arrayRec, templateXML)
                                                                        End If
                                                                        Dim countColl As Integer = arrayPointCollections.GetUpperBound(1) + 1
                                                                        ReDim Preserve arrayPointCollections(3, countColl)
                                                                        arrayPointCollections(0, countColl) = x
                                                                        arrayPointCollections(1, countColl) = y
                                                                        arrayPointCollections(2, countColl) = insertBlk.ObjectID
                                                                        arrayPointCollections(3, countColl) = localID
                                                                        startNumber += 1
                                                                    End If
                                                                End If
                                                            End If
                                                        Next k1
                                                    End If
                                                End If
                                            End If
                                        Next k
                                    End If
                                End If
                            End If
                        End If
                    Next i
                End If
            End If
        End Sub
        '================================================================================================================================
        'удалить точки
        <cmd("DeleteCadastralPoints")>
        Public Sub DeleteCadastralPoints()
            'получаем активное окно
            Dim userCadView As CadView = Me.CadView
            If IsNothing(userCadView) Then Return
            'получаем активный слой
            Dim DrawLayer As DrawingLayer = DrawingLayer.GetDrawingLayer(userCadView)
            If IsNothing(DrawLayer) = True Then Return
            'получаем активный документ
            Dim ActivDocument As Topomatic.Dwg.Drawing = DrawLayer.Drawing
            If IsNothing(ActivDocument) = True Then Return
            Dim userDrawing As Dwg.Drawing = DrawLayer.Drawing
            'сброс выделения
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            userCadView.SelectionSet.Clear()
            Dim [userSelect] As String = "Все"
            CadCursors.GetUserSelect(userCadView, [userSelect], Nothing, "Способ удаления точек:", "Все точки", "Участок")

            Select Case [userSelect]
                Case "Все точки"
                    userCadView.SelectionSet.Clear()
                    userCadView.SelectionSet.SelectOneObjectAtScreen(Function(obj) TypeOf obj Is Object, "Укажите точку в качестве образца: ")
                    Dim pointName As String = ""
                    For Each acEnt As Object In userCadView.SelectionSet
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        If TypeOf acEnt Is DwgInsert Then
                            Dim insBlk As DwgInsert = acEnt
                            pointName = insBlk.Block.Name
                            If IsNothing(pointName) = False Then
                                Exit For
                            End If
                        End If
                    Next

                    userCadView.SelectionSet.Clear()
                    Dim arrayPoints As DwgInsert() = Nothing
                    Dim countArrayPoints As Integer = 0
                    For Each acEnt As DwgEntity In ActivDocument.ActiveSpace.Entities
                        Try
                            If TypeOf acEnt Is DwgInsert Then
                                Dim userBlk As DwgInsert = acEnt
                                If userBlk.Block.Name.Trim Like pointName.Trim Then
                                    ReDim Preserve arrayPoints(countArrayPoints)
                                    arrayPoints(countArrayPoints) = userBlk
                                    countArrayPoints += 1
                                End If
                            End If
                        Catch ex As System.Exception
                        End Try
                    Next
                    If IsArray(arrayPoints) = True Then
                        userDrawing.BeginUpdate()
                        Try
                            For i As Integer = 0 To arrayPoints.Length - 1
                                Dim userBlk As DwgInsert = arrayPoints(i)
                                ActivDocument.ActiveSpace.Entities.Remove(userBlk)
                            Next
                        Catch ex As Exception
                        Finally
                            userDrawing.EndUpdate()
                            userCadView.SelectionSet.Clear()
                        End Try
                    End If
                Case "Участок"

                    userCadView.SelectionSet.Clear()
                    userCadView.SelectionSet.SelectOneObjectAtScreen(Function(obj) TypeOf obj Is Object, "Выборите Участок: ")
                    Dim localIDParcel As String = ""
                    Dim parcels As DwgPolyline = Nothing
                    For Each acEnt As Object In userCadView.SelectionSet
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        If TypeOf acEnt Is DwgPolyline Then
                            If IsNothing(acEnt) = False Then
                                localIDParcel = FuncXRecords.FuncReadValueXData(acEnt, "LOCALID")
                                If localIDParcel.Trim.Length > 0 Then
                                    parcels = acEnt
                                End If
                            End If
                        End If
                    Next

                    If IsNothing(parcels) = False Then
                        userCadView.SelectionSet.Clear()
                        For i As Integer = 0 To parcels.Count - 1
                            Dim buildVertex As BugleVector2D = parcels.Item(i)
                            userCadView.SelectionSet.SelectAtPoint(New Vector2D(buildVertex.Vertex.X, buildVertex.Vertex.Y), Function(obj) TypeOf obj Is DwgInsert)
                            For Each acBlkInsert As DwgInsert In userCadView.SelectionSet
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                Dim localIDPoint = FuncXRecords.FuncReadValueXData(acBlkInsert, "IDRNG")
                                If localIDPoint.Trim Like localIDParcel Then
                                    ActivDocument.ActiveSpace.Entities.Remove(acBlkInsert)
                                End If
                            Next
                        Next
                    End If
            End Select
        End Sub
        '================================================================================================================================
        'найти точку
        <cmd("FindCadastralPoints")>
        Public Sub FindCadastralPoints()
            'получаем активное окно
            Dim userCadView As CadView = Me.CadView
            If IsNothing(userCadView) Then Return
            'получаем активный слой
            Dim DrawLayer As DrawingLayer = DrawingLayer.GetDrawingLayer(userCadView)
            If IsNothing(DrawLayer) = True Then Return
            'получаем активный документ
            Dim ActivDocument As Topomatic.Dwg.Drawing = DrawLayer.Drawing
            If IsNothing(ActivDocument) = True Then Return

            'сброс выделения
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            userCadView.SelectionSet.Clear()
            Dim numberPoint As String = ""
            CadCursors.GetString(userCadView, numberPoint, "Введите номер точки для поиска:")
            If numberPoint.Trim.Length = 0 Then
                Exit Sub
            End If
            Dim arrayBlk As String() = Nothing
            Dim countArrayBlk As Integer = 0
            For Each acEnt As DwgEntity In ActivDocument.ActiveSpace.Entities
                Try
                    If TypeOf acEnt Is DwgInsert Then
                        Dim userBlk As DwgInsert = acEnt
                        Dim attrs As IEnumerable(Of IAttrib) = userBlk.Attribs
                        If attrs.Count > 0 Then
                            For Each attr As IAttrib In attrs
                                Dim userStr As String = attr.Content
                                If userStr.Trim Like numberPoint.Trim Then
                                    ReDim Preserve arrayBlk(countArrayBlk)
                                    arrayBlk(countArrayBlk) = userBlk.ObjectID
                                    countArrayBlk += 1
                                End If
                            Next
                        End If
                    End If
                Catch ex As System.Exception
                End Try
            Next
            If IsArray(arrayBlk) = True Then
                For i As Integer = 0 To arrayBlk.Length - 1
                    Dim insBlk As DwgInsert = Nothing
                    Dim hg As UInteger = arrayBlk(i)
                    Dim boolShow As Boolean = ActivDocument.ActiveSpace.Entities.TryGetObject(hg, insBlk)
                    If IsNothing(insBlk) = False Then
                        Dim boundblk As BoundingBox2D = insBlk.Bounds
                        userCadView.ZoomBound(boundblk, True)
                        Dim [userSelect] As String = "Все"
                        CadCursors.GetUserSelect(userCadView, [userSelect], Nothing, "Выход?", "Да", "Нет")
                        Select Case [userSelect]
                            Case "Да"
                                Exit For
                        End Select
                    End If
                Next i
            End If
        End Sub

        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '3. Добавить метки участкам
        '================================================================================================================================
        'создать метку
        <cmd("CreateLabelsParcels")>
        Public Sub CreateLabelsParcels()
            'получаем активное окно
            Dim userCadView As CadView = Me.CadView
            If IsNothing(userCadView) Then Return
            'получаем активный слой
            Dim DrawLayer As DrawingLayer = DrawingLayer.GetDrawingLayer(userCadView)
            If IsNothing(DrawLayer) = True Then Return
            'получаем активный документ
            Dim ActivDocument As Topomatic.Dwg.Drawing = DrawLayer.Drawing
            If IsNothing(ActivDocument) = True Then Return
            Dim rng As Random = New Random()
            '=============================================================================================================================
            '1.\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Dim generalDir As String = ""
            Dim templateXML As String = ""
            Dim templateDWG As String = ""
            Dim arrayDirSupport As String() = Nothing
            Dim boolFindDirSupport As Boolean = FuncFiles.funcReadDirectoriesSupport(arrayDirSupport)
            If IsArray(arrayDirSupport) = True Then
                For i As Integer = 0 To arrayDirSupport.Length - 1
                    Dim nameFolder As String = arrayDirSupport(i)
                    Dim pos As Integer = nameFolder.IndexOf("InfrastradaToolsUtility")
                    If pos > -1 Then
                        generalDir = Mid(nameFolder, 1, pos)
                        generalDir = generalDir & "InfrastradaToolsUtility"
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
                        Dim shortNameFile As String = System.IO.Path.GetFileNameWithoutExtension(templateXML)
                        Dim tempNameFileDwg As String = directorySupport & shortNameFile & ".dwg"
                        If System.IO.File.Exists(tempNameFileDwg) = True Then
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
            '================================================================================================================================
            'заносим шаблоны в массив
            Dim arrayTemlates As String() = {templateXML, templateDWG}
            'форма инициализируем
            Dim formCreateLabel As FormCreateLabelParcels = New FormCreateLabelParcels(userCadView)
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'запускаем форму
            formCreateLabel.ShowDialog()
            If formCreateLabel.boolShow = False Then
                Exit Sub
            End If
            'масштаб чертежа
            Dim M As Double = userCadView.AnnotationScale
            Dim nameCategory As String = formCreateLabel.ComboBox4.Text
            Dim nameTable As String = formCreateLabel.ComboBox1.Text
            Dim deskMLeader As String = formCreateLabel.ComboBox5.Text
            '====================================================================================
            'имя создаваемого объекта
            Dim nameCreateObject As String = ""
            '====================================================================================
            'слой
            Dim objectLayer As DwgLayer = ActivDocument.ActiveLayer
            Dim textLayer As String = formCreateLabel.ComboBox3.Text
            '====================================================================================
            Dim objectUserTypeLine As DwgLinetype = ActivDocument.Linetypes.ByLayer
            Dim objectColor As CadColor = New CadColor(7)
            Dim indColor As Integer = Val(formCreateLabel.TextBox2.Text)
            Dim scaleUserTypeLine As Double = 1
            Dim widthUserTypeLine As Double = 20
            Dim TextStyleUser As String = formCreateLabel.ComboBox7.Text
            Dim hText As Double = formCreateLabel.NumericUpDown2.Value
            Dim nameArrowUser As String = formCreateLabel.ComboBox6.Text
            Dim nameArrow As String = ""
            If nameArrowUser Like "нет" Then
                nameArrow = ""
            ElseIf nameArrowUser Like "Стрелка" Then
                nameArrow = "_Closed"
            ElseIf nameArrowUser Like "Точка" Then
                nameArrow = "_Dot"
            ElseIf nameArrowUser Like "Малая точка" Then
                nameArrow = "_DotSmall"
            Else
                nameArrow = nameArrowUser
            End If

            Dim arrowSize As Double = formCreateLabel.NumericUpDown1.Value
            Dim tablePs As String = "CADASTRAL_OBJECT"
            Dim note As String = ""
            Dim numerator As String = ""
            Dim denumerator As String = ""
            'Dim noteLabel As String = "numerator=&lt;ConditionalNumber&gt;; denominator=Площадь &lt;Area&gt; кв.м."
            Dim arrayPropertyMLeader As String(,) = Nothing
            Dim boolPropObject As Boolean = FuncXML.FuncReadDataObjectToXMLbyDesk(templateXML, nameCategory, nameTable, "Мультивыноска", deskMLeader, arrayPropertyMLeader)
            If IsArray(arrayPropertyMLeader) = True Then
                For i As Integer = 0 To arrayPropertyMLeader.GetUpperBound(1)
                    Dim nameField As String = arrayPropertyMLeader(0, i)
                    If nameField Like "Layer" Then
                        If Not (formCreateLabel.ComboBox3.Text.Trim Like arrayPropertyMLeader(1, i)) Then
                            arrayPropertyMLeader(1, i) = textLayer
                        End If
                    ElseIf nameField Like "Color" Then
                        If Not (formCreateLabel.TextBox2.Text.Trim Like arrayPropertyMLeader(1, i)) Then
                            arrayPropertyMLeader(1, i) = indColor
                        End If
                    ElseIf nameField Like "ArrowSize" Then
                        arrayPropertyMLeader(1, i) = formCreateLabel.NumericUpDown1.Value
                    ElseIf nameField Like "TextHeight" Then
                        arrayPropertyMLeader(1, i) = formCreateLabel.NumericUpDown2.Value
                    ElseIf nameField Like "TextStyle" Then
                        arrayPropertyMLeader(1, i) = TextStyleUser
                    End If
                Next i
                'FuncDecorPlane.FuncCreateStyleLeaderObject(ActivDocument, templateXML, arrayPropertyMLeader, nameCreateObject, objectLayer, objectUserTypeLine, objectColor, scaleUserTypeLine, widthUserTypeLine, TextStyleUser, hText, nameArrow, arrowSize, tablePs, note)
                If note.Trim.Length > 0 Then
                    Dim ArrayNote As String() = note.Split(";")
                    If IsArray(ArrayNote) = True Then
                        For j As Integer = 0 To ArrayNote.Length - 1
                            Dim str As String = ArrayNote(j)
                            Dim pos As Integer = str.IndexOf("=")
                            If pos > 0 Then
                                Dim NameField As String = Mid(str, 1, pos)
                                Dim valStr As String = Mid(str, pos + 2)
                                If j < ArrayNote.Length - 1 Then
                                    For k As Integer = j + 1 To ArrayNote.Length - 1
                                        Dim TempStr As String = ArrayNote(k)
                                        If TempStr.IndexOf("=") = -1 Then
                                            valStr = valStr & ";" & TempStr
                                            j = k
                                        End If
                                    Next k
                                End If
                                If NameField.Trim Like "numerator" Then
                                    numerator = valStr
                                ElseIf NameField.Trim Like "denominator" Then
                                    denumerator = valStr
                                End If
                            End If
                        Next j
                    End If
                End If
            End If
            Dim noteLabel As String = ""
            If formCreateLabel.ComboBox2.Text.Trim Like "00:00:0000000:0000:ЗУ" Then
                noteLabel = "<Cad_number>:ЗУ" & numerator
            ElseIf formCreateLabel.ComboBox2.Text.Trim Like ":0000:ЗУ" Then
                noteLabel = ":<Short_cad_number>:ЗУ" & numerator
            ElseIf formCreateLabel.ComboBox2.Text.Trim Like "00:00:0000000:0000/чзу" Then
                noteLabel = "<Cad_number>/чзу" & numerator
            ElseIf formCreateLabel.ComboBox2.Text.Trim Like ":0000/чзу" Then
                noteLabel = ":<Short_cad_number>/чзу" & numerator
            Else
                noteLabel = formCreateLabel.ComboBox2.Text & numerator
            End If

            If denumerator.Trim.Length > 0 Then
                noteLabel = noteLabel & " \P " & denumerator
            End If

            '===========================================================================================================
            'выбор объектов
            Dim arrayObject As Object() = Nothing
            Dim countArrayobject As Integer = 0
            If formCreateLabel.RadioButton1.Checked = True Then
                userCadView.SelectionSet.SelectObjectsAtScreen(Function(obj) TypeOf obj Is Object, "Выборите Участки: ")
                '=======================================================================================================
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                For Each acEnt As Object In userCadView.SelectionSet
                    If TypeOf acEnt Is DwgPolyline Then
                        If acEnt.HasExtensionDictionary = True Then
                            ReDim Preserve arrayObject(countArrayobject)
                            arrayObject(countArrayobject) = acEnt
                            countArrayobject += 1
                        End If
                    End If
                Next
            Else
                For Each acEnt As DwgEntity In ActivDocument.ActiveSpace.Entities
                    If TypeOf acEnt Is DwgPolyline Then
                        If acEnt.HasExtensionDictionary = True Then
                            ReDim Preserve arrayObject(countArrayobject)
                            arrayObject(countArrayobject) = acEnt
                            countArrayobject += 1
                        End If
                    End If
                Next
            End If

            '=======================================================================================================
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            If IsArray(arrayObject) = True Then
                For i As Integer = 0 To arrayObject.Length - 1
                    Dim acEnt As DwgEntity = arrayObject(i)
                    If acEnt.HasExtensionDictionary = True Then
                        Dim arrayTables As String() = Nothing
                        Dim boolTables As Boolean = FuncXRecords.FuncReadTablesPS(acEnt, arrayTables)
                        If IsArray(arrayTables) = True Then
                            For j As Integer = 0 To arrayTables.Count - 1
                                Dim nameTablePS As String = arrayTables(j)
                                If nameTablePS.Trim Like "CADASTRAL_OSTROV" Or nameTablePS.Trim Like "IS_SYSTEM" Then
                                    Continue For
                                End If

                                Dim arrayData As String(,) = Nothing
                                Dim boolReadData As Boolean = FuncXRecords.FuncReadXData(acEnt, nameTablePS, arrayData)
                                If IsArray(arrayData) = True Then
                                    Dim localID As String = FuncXRecords.FuncReadValueXData(acEnt, "LOCALID")
                                    For k As Integer = 0 To arrayData.GetUpperBound(1)
                                        Dim nameField As String = arrayData(0, k)
                                        Dim valStr As String = arrayData(1, k)
                                        If IsNothing(valStr) = True Then valStr = ""
                                        If nameField Like "Area" Then
                                            If TypeOf acEnt Is DwgPolyline Then
                                                Dim polyCounter As DwgPolyline = acEnt
                                                Dim sqrtPoly As Double = MathFunction.GetArea(polyCounter, 2)
                                                sqrtPoly = Math.Abs(sqrtPoly)
                                                If IsNothing(localID) = False Then
                                                    Dim arrayostrov As String(,) = Nothing
                                                    Dim boolOstrov As Boolean = CadastralFunction.FuncFindCadastralOstrov(ActivDocument, polyCounter, localID, arrayostrov)
                                                    If IsArray(arrayostrov) = True Then
                                                        For k1 As Integer = 0 To arrayostrov.GetUpperBound(1)
                                                            Dim sqrtOstrov As Double = Val(arrayostrov(2, k1))
                                                            sqrtPoly -= sqrtOstrov
                                                        Next k1
                                                    End If
                                                    numerator = numerator.Replace("<Area>", Math.Round(sqrtPoly, 0))
                                                    denumerator = denumerator.Replace("<Area>", Math.Round(sqrtPoly, 0))
                                                End If
                                            End If
                                        ElseIf nameField Like "Lenght" Then
                                            If TypeOf acEnt Is DwgPolyline Then
                                                Dim polyCounter As DwgPolyline = acEnt
                                                Dim l As Double = Math.Round(polyCounter.Length, 2)
                                                numerator = numerator.Replace("<Lenght>", l)
                                                denumerator = denumerator.Replace("<Lenght>", l)
                                            End If
                                        ElseIf nameField Like "ConditionalNumber" Then
                                            Dim prefParcel As String = formCreateLabel.ComboBox2.Text.Trim
                                            Dim cad_number As String = FuncXRecords.FuncReadValueXData(acEnt, "Cad_number")
                                            Dim conditionalNumber As String = FuncXRecords.FuncReadValueXData(acEnt, "ConditionalNumber")
                                            If prefParcel.IndexOf("00:00:0000000:0000:ЗУ") > -1 Then
                                                Dim pos As Integer = cad_number.LastIndexOf(":")
                                                If pos = cad_number.Length Then
                                                    conditionalNumber = cad_number & "ЗУ" & conditionalNumber
                                                Else
                                                    conditionalNumber = cad_number & ":ЗУ" & conditionalNumber
                                                End If
                                            ElseIf prefParcel.IndexOf(":0000:ЗУ") > -1 Then
                                                Dim pos As Integer = cad_number.LastIndexOf(":")
                                                conditionalNumber = ":" & Mid(cad_number, pos + 1) & ":ЗУ" & conditionalNumber
                                            ElseIf prefParcel.IndexOf("00:00:0000000:0000/чзу") > -1 Then
                                                conditionalNumber = cad_number & "/чзу" & conditionalNumber
                                            ElseIf prefParcel.IndexOf(":0000/чзу") > -1 Then
                                                Dim pos As Integer = cad_number.LastIndexOf(":")
                                                conditionalNumber = ":" & Mid(cad_number, pos + 1) & "/чзу" & conditionalNumber
                                            Else
                                                conditionalNumber = prefParcel & conditionalNumber
                                            End If
                                            numerator = numerator.Replace("<ConditionalNumber>", conditionalNumber)
                                            denumerator = denumerator.Replace("<ConditionalNumber>", conditionalNumber)
                                        Else
                                            Dim pos As Integer = numerator.IndexOf(nameField)
                                            If pos > 0 Then
                                                Dim prefSymb As String = Mid(numerator, pos - 1, 1)
                                                Dim postSymb As String = Mid(numerator, pos + 1, 1)
                                                If prefSymb Like "<" And postSymb Like ">" Then
                                                    numerator = numerator.Replace("<" & nameField & ">", valStr)
                                                End If
                                            End If
                                            pos = denumerator.IndexOf(nameField)
                                            If pos > 0 Then
                                                Dim prefSymb As String = Mid(denumerator, pos - 1, 1)
                                                Dim postSymb As String = Mid(denumerator, pos + 1, 1)
                                                If prefSymb Like "<" And postSymb Like ">" Then
                                                    denumerator = denumerator.Replace("<" & nameField & ">", valStr)
                                                End If
                                            End If
                                        End If
                                    Next k
                                    '============================================================================================================
                                    'строка подписи
                                    Dim textLeder As String = ""
                                    If numerator.Trim.Length > 0 Then
                                        textLeder = numerator
                                    End If
                                    If denumerator.Trim.Length > 0 Then
                                        textLeder = textLeder & "\P" & denumerator
                                    End If

                                    '============================================================================================================
                                    'вставляем выноску в участок
                                    If formCreateLabel.RadioButton5.Checked = True Then
                                        'ищем мультивыноску
                                        Dim acLeaderParcel As DwgLeader = CadastralFunction.FuncFindCadastralLabel(ActivDocument, localID)
                                        If IsNothing(acLeaderParcel) = False Then
                                            acLeaderParcel.Content = textLeder
                                        End If
                                    Else
                                        If TypeOf acEnt Is DwgPolyline Then
                                            Dim polyCounter As DwgPolyline = acEnt
                                            Dim centerPolygon As Vector2D = New Vector2D(-1, -1)
                                            If polyCounter.Closed = True Then
                                                Dim listVertex As List(Of BugleVector2D) = polyCounter.ToList
                                                Dim arrayVertex As Double(,) = Nothing
                                                For k As Integer = 0 To listVertex.Count - 1
                                                    ReDim Preserve arrayVertex(1, k)
                                                    arrayVertex(0, k) = listVertex.Item(k).Vertex.X
                                                    arrayVertex(1, k) = listVertex.Item(k).Vertex.Y
                                                Next k
                                                centerPolygon = MathFunction.FuncFindCrossCenterPointPolyline(arrayVertex)
                                            End If
                                            Dim pointInsText As Vector2D = polyCounter.Bounds.Max
                                            Dim userMLeader As DwgLeader = Nothing
                                            Dim styleTextUser As DwgStyle = StyleCadastralObject.FuncReturnTextStyle(ActivDocument, TextStyleUser)
                                            If IsNothing(styleTextUser) = True Then
                                                styleTextUser = ActivDocument.Styles.Standard
                                            End If
                                            If formCreateLabel.RadioButton3.Checked = True Then
                                                userMLeader = RoburFunc.FuncDrawLeader(ActivDocument, textLeder, centerPolygon, centerPolygon, M, hText, styleTextUser, nameArrow, arrowSize)
                                                userMLeader.HorizontalAlignment = Topomatic.Dwg.HorizontalAlignment.TextMiddle
                                            Else
                                                userMLeader = RoburFunc.FuncDrawLeader(ActivDocument, textLeder, centerPolygon, pointInsText, M, hText, styleTextUser, nameArrow, arrowSize)
                                                userMLeader.HorizontalAlignment = Topomatic.Dwg.HorizontalAlignment.FirstLineUnderline
                                            End If

                                            If IsNothing(userMLeader) = False Then
                                                userMLeader.Layer = objectLayer
                                                userMLeader.Color = objectColor
                                                userMLeader.Linetype = objectUserTypeLine
                                                userMLeader.Lineweight = widthUserTypeLine
                                                userMLeader.LinetypeScale = scaleUserTypeLine
                                                userMLeader.Annotative = True
                                                userMLeader.AutoMirror = True

                                                If formCreateLabel.CheckBox1.Checked = True Then
                                                    userMLeader.BackgroundFillType = BackgroundFillType.UseWindowColor
                                                End If

                                                Dim arrayRec As String(,) = Nothing
                                                ReDim Preserve arrayRec(1, 3)
                                                arrayRec(0, 0) = "LABEL"
                                                arrayRec(1, 0) = deskMLeader
                                                arrayRec(0, 1) = "Number"
                                                arrayRec(1, 1) = "1"
                                                arrayRec(0, 2) = "IDRNG"
                                                arrayRec(1, 2) = localID
                                                arrayRec(0, 3) = "NOTE"
                                                arrayRec(1, 3) = noteLabel
                                                If IsArray(arrayRec) = True Then
                                                    Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXData(userMLeader, tablePs, arrayRec, templateXML)
                                                End If
                                                Exit For
                                            End If
                                        End If
                                    End If
                                End If
                            Next j
                        End If
                    End If
                Next i
            End If
        End Sub
        '=================================================================================================================================
        'удалить метки
        <cmd("DeleteLabelsParcels")>
        Public Sub DeleteLabelsParcels()
            'получаем активное окно
            Dim userCadView As CadView = Me.CadView
            If IsNothing(userCadView) Then Return
            'получаем активный слой
            Dim DrawLayer As DrawingLayer = DrawingLayer.GetDrawingLayer(userCadView)
            If IsNothing(DrawLayer) = True Then Return
            'получаем активный документ
            Dim ActivDocument As Topomatic.Dwg.Drawing = DrawLayer.Drawing
            If IsNothing(ActivDocument) = True Then Return

            Dim userDrawing As Dwg.Drawing = DrawLayer.Drawing

            'сброс выделения
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            userCadView.SelectionSet.Clear()
            userCadView.SelectionSet.SelectOneObjectAtScreen(Function(obj) TypeOf obj Is Object, "Выборите метку для образца: ")

            Dim labelLeader As String = ""
            For Each acEnt As Object In userCadView.SelectionSet
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                If TypeOf acEnt Is DwgLeader Then
                    labelLeader = FuncXRecords.FuncReadValueXData(acEnt, "LABEL")
                    If IsNothing(labelLeader) = False Then
                        Exit For
                    End If
                End If
            Next

            Dim arrayObject As DwgLeader() = Nothing
            Dim countArrayObject As Integer = 0

            For Each acEnt As DwgEntity In ActivDocument.ActiveSpace.Entities
                Try
                    If TypeOf acEnt Is DwgLeader Then
                        Dim labelLeaderSelect As String = FuncXRecords.FuncReadValueXData(acEnt, "LABEL")
                        If IsNothing(labelLeaderSelect) = False Then
                            If labelLeaderSelect.Trim Like labelLeader.Trim Then
                                Dim userLeader As DwgLeader = acEnt
                                ReDim Preserve arrayObject(countArrayObject)
                                arrayObject(countArrayObject) = userLeader
                                countArrayObject += 1
                                'userLeader.Content = ""
                                'userLeader.Clear()
                            End If
                        End If
                    End If
                Catch ex As System.Exception
                End Try
            Next

            userDrawing.BeginUpdate()
            Try
                If IsArray(arrayObject) = True Then
                    For i As Integer = 0 To arrayObject.Length - 1
                        Dim userLeader As DwgLeader = arrayObject(i)
                        ActivDocument.ActiveSpace.Entities.Remove(userLeader)
                    Next
                End If
            Catch ex As Exception
            Finally
                userDrawing.EndUpdate()
                userCadView.SelectionSet.Clear()
            End Try

        End Sub
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '4.Отчеты
        '==================================================================================================================================
        'экспорт данных
        <cmd("ExportCadastralData")>
        Public Sub ExportCadastralData()
            'получаем активное окно
            Dim userCadView As CadView = Me.CadView
            If IsNothing(userCadView) Then Return
            'получаем активный слой
            Dim DrawLayer As DrawingLayer = DrawingLayer.GetDrawingLayer(userCadView)
            If IsNothing(DrawLayer) = True Then Return
            'получаем активный документ
            Dim ActivDocument As Topomatic.Dwg.Drawing = DrawLayer.Drawing
            If IsNothing(ActivDocument) = True Then Return

            'сброс выделения
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            userCadView.SelectionSet.Clear()
            userCadView.SelectionSet.SelectObjectsAtScreen(Function(obj) TypeOf obj Is Object, "Выборите объекты: ")
            If userCadView.SelectionSet.Count > 0 Then
                Dim strFilter As String = "AutoCAD files (*.dwg)|*.dwg|All files (*.*)|*.*"
                Dim strFileName As String = FuncFiles.FuncSaveFilePathName("Экспорт данных", strFilter)
                If strFileName.Trim.Length > 0 Then
                    Dim drawing As Dwg.Drawing = New Dwg.Drawing()
                    Dim blockEnt As DwgBlock = drawing.ActiveSpace
                    For Each acEnt As Object In userCadView.SelectionSet
                        If TypeOf acEnt Is DwgEntity Then
                            Dim acUserObject As DwgEntity = acEnt
                            Dim destAcObject As DwgEntity = acUserObject.Clone
                            blockEnt.Add(destAcObject)
                        End If
                    Next
                    Dim provider As DrawingExportProvider = DrawingExportProvider.GetProvider("DWG")
                    provider.SaveToFile(strFileName, drawing)
                End If
            End If
        End Sub
        '==================================================================================================================================
        'Отчет по координатам
        <cmd("ReportCoordinatesPoints")>
        Public Sub ReportCoordinatesPoints()
            'получаем активное окно
            Dim cadWiever As CadView = Me.CadView
            If IsNothing(cadWiever) Then Return
            'получаем активный слой
            Dim DrawLayer As DrawingLayer = DrawingLayer.GetDrawingLayer(cadWiever)
            If IsNothing(DrawLayer) = True Then Return
            'получаем активный документ
            Dim ActivDocument As Topomatic.Dwg.Drawing = DrawLayer.Drawing
            If IsNothing(ActivDocument) = True Then Return

            Dim generalDir As String = ""
            Dim templateXML As String = ""
            Dim templateDWG As String = ""
            Dim arrayDirSupport As String() = Nothing
            Dim boolFindDirSupport As Boolean = FuncFiles.funcReadDirectoriesSupport(arrayDirSupport)
            If IsArray(arrayDirSupport) = True Then
                For i As Integer = 0 To arrayDirSupport.Length - 1
                    Dim nameFolder As String = arrayDirSupport(i)
                    Dim pos As Integer = nameFolder.IndexOf("InfrastradaToolsUtility")
                    If pos > -1 Then
                        generalDir = Mid(nameFolder, 1, pos)
                        generalDir = generalDir & "InfrastradaToolsUtility"
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
                        Dim shortNameFile As String = System.IO.Path.GetFileNameWithoutExtension(templateXML)
                        Dim tempNameFileDwg As String = directorySupport & shortNameFile & ".dwg"
                        If System.IO.File.Exists(tempNameFileDwg) = True Then
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

            'сброс выделения
            cadWiever.SelectionSet.Clear()
            '1.\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Dim formReport As FormReportCatalogCoordinates = New FormReportCatalogCoordinates(cadWiever)
            formReport.templateDwg = templateDWG
            formReport.templateXML = templateXML

            formReport.ShowDialog()
            If formReport.boolRp = False Then
                Exit Sub
            End If

            Dim namePoint As String = formReport.ComboBox2.Text
            If IsNothing(namePoint) = True Then
                MsgBox("Имя блока выбрано не корректно! Макро прервано!!!")
                Exit Sub
            End If
            If namePoint.Trim.Length = 0 Then
                MsgBox("Имя блока выбрано не корректно! Макро прервано!!!")
                Exit Sub
            End If
            Dim BlkPoint As DwgBlock = ActivDocument.Blocks.Item(namePoint)
            If IsNothing(BlkPoint) = True Then
                MsgBox("Имя блока выбрано не корректно! Макро прервано!!!")
                Exit Sub
            End If

            Dim nameAttrib As String = formReport.ComboBox3.Text
            If IsNothing(nameAttrib) = True Then
                MsgBox("Имя атрибута блока выбрано не корректно! Макро прервано!!!")
                Exit Sub
            End If
            If nameAttrib.Trim.Length = 0 Then
                MsgBox("Имя атрибута блока выбрано не корректно! Макро прервано!!!")
                Exit Sub
            End If

            Dim parcelDictionary As Dictionary(Of String, String(,)) = New Dictionary(Of String, String(,))
            Dim arrayReportParcels As String(,) = Nothing
            Dim countArrayRaportParcels As Integer = 0

            Dim roundCoord As Integer = formReport.NumericUpDown1.Value
            Dim roundArea As Integer = formReport.NumericUpDown2.Value
            Dim dopusk As Double = formReport.NumericUpDown3.Value

            Dim arrayPointCollections As String(,) = Nothing
            Dim numberPointCollections As Integer = CadastralFunction.FuncFindCadastralPoints(ActivDocument, namePoint, arrayPointCollections)
            'габаритный контейнер для mapinfo
            Dim minX As Double = 900000000
            Dim minY As Double = 900000000
            Dim maxX As Double = -1
            Dim maxY As Double = -1
            '2.\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'выбираем объекты
            CadView.SelectionSet.SelectObjectsAtScreen(Function(obj) TypeOf obj Is Object, "Выборите Участки: ")
            '=======================================================================================================
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Dim arrayObject As String(,) = Nothing '0-Номер участка,1-handle,2-GLobalID,2-LocalID,3-ha
            Dim countArrayObject As Integer = 0
            '=======================================================================================================
            'находим все участки в выбранной коллекции
            For Each acEnt As Object In cadWiever.SelectionSet
                If TypeOf acEnt Is DwgPolyline Then
                    Dim acPoly As DwgPolyline = acEnt
                    If acPoly.HasExtensionDictionary = True Then
                        Dim LABEL As String = FuncXRecords.FuncReadValueXData(acPoly, "LABEL")
                        If LABEL Like "Внутренний островок" Then
                            Continue For
                        Else
                            Dim globalID As String = FuncXRecords.FuncReadValueXData(acPoly, "GLOBALID")
                            If globalID.Trim.Length > 2 Then
                                Dim boolArrayGlobalid As Boolean = False
                                If IsArray(arrayObject) = True Then
                                    For i As Integer = 0 To arrayObject.GetUpperBound(1)
                                        Dim tempGlobalID As String = arrayObject(2, i)
                                        If tempGlobalID Like globalID Then
                                            boolArrayGlobalid = True
                                            Exit For
                                        End If
                                    Next i
                                End If
                                If boolArrayGlobalid = True Then
                                    Continue For
                                End If
                            End If
                            Dim conditionalParcel As String = FuncXRecords.FuncReadValueXData(acPoly, "ConditionalNumber")
                            Dim pos1 As Integer = conditionalParcel.IndexOf("(")
                            Dim numberParcel As String = FuncFormatZn.FuncFindFullConditionalNumberCadastralObject(conditionalParcel)
                            Dim localID As String = FuncXRecords.FuncReadValueXData(acPoly, "LOCALID")
                            ReDim Preserve arrayObject(4, countArrayObject)
                            arrayObject(0, countArrayObject) = Val(numberParcel)
                            arrayObject(1, countArrayObject) = acPoly.ObjectID
                            arrayObject(2, countArrayObject) = globalID
                            arrayObject(3, countArrayObject) = localID
                            arrayObject(4, countArrayObject) = LABEL
                            countArrayObject += 1
                        End If
                    End If
                End If
            Next
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'участки найдены начинаем доставать координаты
            If IsArray(arrayObject) = True Then
                If arrayObject.GetUpperBound(1) > 0 Then
                    Dim boolSortArray As Boolean = MathFunction.FuncSortStrArray(arrayObject, 0)
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'обрабатываем участки
                Dim posStartParcel As Integer = 0
                For i As Integer = 0 To arrayObject.GetUpperBound(1)
                    Dim numberParcel As String = arrayObject(0, i)
                    Dim hgFirstCounters As String = arrayObject(1, i)
                    Dim globalID As String = arrayObject(2, i)
                    Dim localID As String = arrayObject(3, i)
                    Dim sqrtParcels As Double = 0
                    Dim arrayCoordinates As String(,) = Nothing
                    Dim countArrayCoordinates As Integer = 0
                    '===========================================================================================
                    If IsNothing(hgFirstCounters) = False Then
                        If hgFirstCounters.Trim.Length > 0 Then
                            Dim acPolyParcels As DwgObject = Nothing
                            Dim boolFindObject As Boolean = ActivDocument.ActiveSpace.Entities.TryGetObject(hgFirstCounters, acPolyParcels)
                            If boolFindObject = True Then
                                If TypeOf acPolyParcels Is DwgPolyline Then
                                    Dim acParentPoly As DwgPolyline = acPolyParcels
                                    'ищем входящие контура
                                    Dim arrayDataCounters As String(,) = Nothing
                                    Dim countCounters As Integer = 1
                                    Dim numParcels As Integer = Val(arrayObject(0, i))
                                    If globalID.Trim.Length > 0 Then
                                        countCounters = CadastralFunction.FuncFindCadastralCounter(ActivDocument, numParcels, arrayDataCounters)
                                    End If
                                    If IsArray(arrayDataCounters) = False Then
                                        ReDim Preserve arrayDataCounters(3, 0)
                                        arrayDataCounters(0, 0) = "0"
                                        arrayDataCounters(1, 0) = hgFirstCounters
                                        Dim sqrtPoly As Double = Math.Abs(MathFunction.GetArea(acParentPoly, roundCoord))
                                        sqrtPoly = Math.Round(sqrtPoly, roundArea)
                                        arrayDataCounters(2, 0) = Math.Abs(sqrtPoly)
                                        arrayDataCounters(3, 0) = localID
                                    End If
                                    '==========================================================================================================
                                    Dim startPosRec As Integer = 0
                                    For j As Integer = 0 To arrayDataCounters.GetUpperBound(1)
                                        Dim numberCounters As String = arrayDataCounters(0, j)
                                        Dim hgCounters As String = arrayDataCounters(1, j)
                                        Dim sqrtPolyCounters As Double = Val(arrayDataCounters(2, j))
                                        Dim localIDCounter As String = arrayDataCounters(3, j)
                                        countCounters = 1
                                        'запоминаем начало участка
                                        startPosRec = countArrayCoordinates

                                        '0-номер контура,1-обозначение участка,2-площадь или длина,3-тип объекта, 4-количество объектов,5-количество вершин,6-ObjectID объекта
                                        Dim acPolyCountersObj As DwgObject = Nothing
                                        Dim boolFindCounters As Boolean = ActivDocument.ActiveSpace.Entities.TryGetObject(hgCounters, acPolyCountersObj)
                                        If boolFindCounters = True Then
                                            If TypeOf acPolyCountersObj Is DwgPolyline Then
                                                'доступ к контуру
                                                Dim acPolyCounters As DwgPolyline = acPolyCountersObj
                                                'габаритный контейнер (для mapInfo)
                                                Dim boundPoly As BoundingBox2D = acPolyCounters.Bounds
                                                If minX > boundPoly.Min.X Then
                                                    minX = boundPoly.Min.X
                                                End If
                                                If minY > boundPoly.Min.Y Then
                                                    minY = boundPoly.Min.Y
                                                End If
                                                If maxX < boundPoly.Max.X Then
                                                    maxX = boundPoly.Max.X
                                                End If
                                                If maxY < boundPoly.Max.Y Then
                                                    maxY = boundPoly.Max.Y
                                                End If
                                                '===============================================================================================
                                                'ищем островки
                                                Dim arrayDataOstrov As String(,) = Nothing
                                                Dim boolFindOstrov As Boolean = CadastralFunction.FuncFindCadastralOstrov(ActivDocument, acPolyCounters, localIDCounter, arrayDataOstrov)
                                                If IsArray(arrayDataOstrov) = True Then
                                                    For k As Integer = 0 To arrayDataOstrov.GetUpperBound(1)
                                                        Dim sqrtOstrov As Double = Val(arrayDataOstrov(2, k))
                                                        sqrtPolyCounters -= sqrtOstrov
                                                        countCounters += 1
                                                    Next k
                                                End If
                                                '==============================================================================================
                                                ReDim Preserve arrayCoordinates(6, countArrayCoordinates)
                                                arrayCoordinates(0, countArrayCoordinates) = numberCounters
                                                arrayCoordinates(1, countArrayCoordinates) = numberParcel
                                                arrayCoordinates(2, countArrayCoordinates) = sqrtPolyCounters
                                                arrayCoordinates(3, countArrayCoordinates) = ""
                                                arrayCoordinates(4, countArrayCoordinates) = countCounters
                                                arrayCoordinates(5, countArrayCoordinates) = ""
                                                arrayCoordinates(6, countArrayCoordinates) = hgCounters

                                                If acPolyCounters.Closed = True Then
                                                    arrayCoordinates(3, countArrayCoordinates) = "REGION"
                                                    arrayCoordinates(5, countArrayCoordinates) = acPolyCounters.Count + 1
                                                Else
                                                    arrayCoordinates(3, countArrayCoordinates) = "PLINE "
                                                    arrayCoordinates(5, countArrayCoordinates) = acPolyCounters.Count
                                                End If
                                                countArrayCoordinates += 1
                                                Dim startPos As Integer = countArrayCoordinates
                                                'запасываем координаты
                                                For k As Integer = 0 To acPolyCounters.Count - 1
                                                    Dim x As Double = acPolyCounters.ElementAt(k).Vertex.X
                                                    Dim y As Double = acPolyCounters.ElementAt(k).Vertex.Y
                                                    '0-Номер точки,1-x,2-y,3-z,4-дельта,5-способ хакрепления,6-способ определения
                                                    ReDim Preserve arrayCoordinates(6, countArrayCoordinates)
                                                    arrayCoordinates(0, countArrayCoordinates) = ""
                                                    arrayCoordinates(1, countArrayCoordinates) = ""
                                                    arrayCoordinates(2, countArrayCoordinates) = ""
                                                    arrayCoordinates(3, countArrayCoordinates) = ""
                                                    arrayCoordinates(4, countArrayCoordinates) = ""
                                                    arrayCoordinates(5, countArrayCoordinates) = ""
                                                    arrayCoordinates(6, countArrayCoordinates) = ""
                                                    arrayCoordinates(1, countArrayCoordinates) = Math.Round(x, roundCoord)
                                                    arrayCoordinates(2, countArrayCoordinates) = Math.Round(y, roundCoord)
                                                    '=========================================================================================
                                                    'ищем номер точки
                                                    If IsArray(arrayPointCollections) = True Then
                                                        Dim arrayDataPoint As String(,) = Nothing
                                                        Dim boolFindNumber As Boolean = CadastralFunction.FuncFindCadastralPoint(ActivDocument, New Vector2D(x, y), arrayPointCollections, arrayDataPoint, dopusk, localIDCounter)
                                                        If IsArray(arrayDataPoint) = True Then
                                                            Dim num_point As String = FuncXML.FuncReturnValueByArrayProperties(nameAttrib, arrayDataPoint, 2)
                                                            arrayCoordinates(0, countArrayCoordinates) = num_point
                                                            If formReport.RadioButton5.Checked = True Then
                                                                Dim d_point As String = FuncXML.FuncReturnValueByArrayProperties("Delta_geopoint", arrayDataPoint, 2)
                                                                Dim zakrep_point As String = FuncXML.FuncReturnValueByArrayProperties("GeopointZacrep", arrayDataPoint, 2)
                                                                Dim opred_point As String = FuncXML.FuncReturnValueByArrayProperties("Geopoint_opred", arrayDataPoint, 2)
                                                                arrayCoordinates(4, countArrayCoordinates) = d_point
                                                                arrayCoordinates(5, countArrayCoordinates) = zakrep_point
                                                                arrayCoordinates(6, countArrayCoordinates) = opred_point
                                                            End If
                                                        End If
                                                    End If
                                                    countArrayCoordinates += 1
                                                Next k
                                                'записываем первую точку полигона повторно
                                                If acPolyCounters.Closed = True Then
                                                    ReDim Preserve arrayCoordinates(6, countArrayCoordinates)
                                                    arrayCoordinates(0, countArrayCoordinates) = arrayCoordinates(0, startPos)
                                                    arrayCoordinates(1, countArrayCoordinates) = arrayCoordinates(1, startPos)
                                                    arrayCoordinates(2, countArrayCoordinates) = arrayCoordinates(2, startPos)
                                                    arrayCoordinates(3, countArrayCoordinates) = arrayCoordinates(3, startPos)
                                                    arrayCoordinates(4, countArrayCoordinates) = arrayCoordinates(4, startPos)
                                                    arrayCoordinates(5, countArrayCoordinates) = arrayCoordinates(5, startPos)
                                                    arrayCoordinates(6, countArrayCoordinates) = arrayCoordinates(6, startPos)
                                                    countArrayCoordinates += 1
                                                End If

                                                If IsArray(arrayDataOstrov) = True Then
                                                    For k As Integer = 0 To arrayDataOstrov.GetUpperBound(1)
                                                        Dim hgOstrov As String = arrayDataOstrov(1, k)
                                                        Dim localIDOstrov As String = arrayDataOstrov(3, k)
                                                        If IsNothing(hgOstrov) = False Then
                                                            Dim objOstrov As DwgObject = Nothing
                                                            Dim findOstrov As Boolean = ActivDocument.ActiveSpace.Entities.TryGetObject(hgOstrov, objOstrov)
                                                            If findOstrov = True Then
                                                                If TypeOf objOstrov Is DwgPolyline Then
                                                                    Dim acPolyOstrov As DwgPolyline = objOstrov
                                                                    ReDim Preserve arrayCoordinates(6, countArrayCoordinates)
                                                                    arrayCoordinates(0, countArrayCoordinates) = ""
                                                                    arrayCoordinates(1, countArrayCoordinates) = ""
                                                                    arrayCoordinates(2, countArrayCoordinates) = ""
                                                                    arrayCoordinates(3, countArrayCoordinates) = ""
                                                                    arrayCoordinates(4, countArrayCoordinates) = ""
                                                                    arrayCoordinates(5, countArrayCoordinates) = ""
                                                                    arrayCoordinates(6, countArrayCoordinates) = ""
                                                                    arrayCoordinates(5, countArrayCoordinates) = acPolyOstrov.Count + 1
                                                                    arrayCoordinates(3, countArrayCoordinates) = "Островок"
                                                                    countArrayCoordinates += 1
                                                                    startPos = countArrayCoordinates
                                                                    For k1 As Integer = 0 To acPolyOstrov.Count - 1
                                                                        Dim x As Double = acPolyOstrov.ElementAt(k1).Vertex.X
                                                                        Dim y As Double = acPolyOstrov.ElementAt(k1).Vertex.Y
                                                                        ReDim Preserve arrayCoordinates(6, countArrayCoordinates)
                                                                        arrayCoordinates(0, countArrayCoordinates) = ""
                                                                        arrayCoordinates(1, countArrayCoordinates) = ""
                                                                        arrayCoordinates(2, countArrayCoordinates) = ""
                                                                        arrayCoordinates(3, countArrayCoordinates) = ""
                                                                        arrayCoordinates(4, countArrayCoordinates) = ""
                                                                        arrayCoordinates(5, countArrayCoordinates) = ""
                                                                        arrayCoordinates(6, countArrayCoordinates) = ""
                                                                        arrayCoordinates(1, countArrayCoordinates) = Math.Round(x, roundCoord)
                                                                        arrayCoordinates(2, countArrayCoordinates) = Math.Round(y, roundCoord)
                                                                        '=========================================================================================
                                                                        'ищем номер точки
                                                                        If IsArray(arrayPointCollections) = True Then
                                                                            Dim arrayDataPoint As String(,) = Nothing
                                                                            Dim boolFindNumber As Boolean = CadastralFunction.FuncFindCadastralPoint(ActivDocument, New Vector2D(x, y), arrayPointCollections, arrayDataPoint, dopusk, localIDOstrov)
                                                                            If IsArray(arrayDataPoint) = True Then
                                                                                Dim num_point As String = FuncXML.FuncReturnValueByArrayProperties(nameAttrib, arrayDataPoint, 2)
                                                                                arrayCoordinates(0, countArrayCoordinates) = num_point
                                                                                If formReport.RadioButton5.Checked = True Then
                                                                                    Dim d_point As String = FuncXML.FuncReturnValueByArrayProperties("Delta_geopoint", arrayDataPoint, 2)
                                                                                    Dim zakrep_point As String = FuncXML.FuncReturnValueByArrayProperties("GeopointZacrep", arrayDataPoint, 2)
                                                                                    Dim opred_point As String = FuncXML.FuncReturnValueByArrayProperties("Geopoint_opred", arrayDataPoint, 2)
                                                                                    arrayCoordinates(4, countArrayCoordinates) = d_point
                                                                                    arrayCoordinates(5, countArrayCoordinates) = zakrep_point
                                                                                    arrayCoordinates(6, countArrayCoordinates) = opred_point
                                                                                End If
                                                                            End If
                                                                        End If
                                                                        countArrayCoordinates += 1
                                                                    Next k1
                                                                    'записываем первую точку полигона повторно
                                                                    ReDim Preserve arrayCoordinates(6, countArrayCoordinates)
                                                                    arrayCoordinates(0, countArrayCoordinates) = arrayCoordinates(0, startPos)
                                                                    arrayCoordinates(1, countArrayCoordinates) = arrayCoordinates(1, startPos)
                                                                    arrayCoordinates(2, countArrayCoordinates) = arrayCoordinates(2, startPos)
                                                                    arrayCoordinates(3, countArrayCoordinates) = arrayCoordinates(3, startPos)
                                                                    arrayCoordinates(4, countArrayCoordinates) = arrayCoordinates(4, startPos)
                                                                    arrayCoordinates(5, countArrayCoordinates) = arrayCoordinates(5, startPos)
                                                                    arrayCoordinates(6, countArrayCoordinates) = arrayCoordinates(6, startPos)
                                                                    countArrayCoordinates += 1
                                                                    'вычисляем площадь островка
                                                                End If
                                                            End If
                                                        End If
                                                    Next k
                                                End If
                                            End If
                                        End If
                                    Next j
                                    If parcelDictionary.ContainsKey(numberParcel) = False Then
                                        parcelDictionary.Add(numberParcel, arrayCoordinates)
                                    Else
                                        numberParcel = numberParcel & "_" & i
                                        parcelDictionary.Add(numberParcel, arrayCoordinates)
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next i
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'выводим отчет
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'MS Word
            If formReport.RadioButton3.Checked = True Then
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'Word документ
                Dim WApp As Word.Application = Nothing
                Dim WDoc As Word.Document = Nothing
                Dim boolWDVisibl As Boolean = False
                If formReport.RadioButton3.Checked = True Then
                    WApp = CreateObject("Word.Application")
                    WDoc = WApp.Documents.Add()
                End If
                WApp.Visible = True
                '================================================================================
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'записываем данные
                If parcelDictionary.Count > 0 Then
                    For i As Integer = 0 To parcelDictionary.Count - 1
                        Dim dictParcel As KeyValuePair(Of String, String(,)) = parcelDictionary.ElementAt(i)
                        Dim numberParcel As String = dictParcel.Key
                        Dim arrayCoord As String(,) = dictParcel.Value
                        Dim text1 As String = "Каталог координат поворотных точек границы: " & numberParcel
                        Dim prefNumber As String = formReport.ComboBox4.Text
                        If prefNumber.IndexOf("00:00:0000000:0000") > -1 Then
                            Dim hgParcel As String = ""
                            If IsArray(arrayCoord) = True Then
                                hgParcel = arrayCoord(6, 0)
                                Dim acEct As DwgEntity = Nothing
                                Dim boolEnt As Boolean = ActivDocument.ActiveSpace.Entities.TryGetObject(hgParcel, acEct)
                                Dim cad_number As String = FuncXRecords.FuncReadValueXData(acEct, "Cad_number")
                                prefNumber.Replace("00:00:0000000:0000", cad_number)
                                text1 = "Каталог координат поворотных точек границы: " & prefNumber & arrayCoord(0, 1)
                            End If
                        Else
                            text1 = "Каталог координат поворотных точек границы: " & prefNumber & arrayCoord(0, 1)
                        End If

                        Dim SetPos As Integer = WDoc.Content.End - 1
                        WDoc.Range(SetPos).Paragraphs.Add()
                        SetPos = WDoc.Content.End - 1
                        WDoc.Range(SetPos).Font.Name = "Times New Roman"
                        WDoc.Range(SetPos).Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft

                        Dim l1 As Integer = text1.Length
                        WDoc.Range(SetPos).Text = text1
                        WDoc.Range(SetPos + l1).Paragraphs.Add()
                        SetPos = SetPos + l1 + 1

                        Dim range As Word.Range = WDoc.Range(SetPos)
                        WDoc.Range(SetPos).Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                        Dim WDocTbl As Word.Table = WDoc.Tables.Add(range, 3, 3, Word.WdDefaultTableBehavior.wdWord9TableBehavior, Word.WdAutoFitBehavior.wdAutoFitContent)

                        WDocTbl.Style = "Сетка таблицы"
                        WDocTbl.ApplyStyleRowBands = True
                        WDocTbl.Cell(1, 1).Merge(MergeTo:=WDocTbl.Cell(2, 1))
                        WDocTbl.Cell(1, 1).Range.Text = "Обозначение" & vbLf & "характерных" & vbLf & "точек границ"
                        WDocTbl.Cell(1, 1).Range.Bold = 1
                        WDocTbl.Cell(1, 1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                        WDocTbl.Cell(1, 1).Range.Font.Name = "Times New Roman"
                        WDocTbl.Cell(1, 1).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter

                        WDocTbl.Cell(1, 2).Merge(MergeTo:=WDocTbl.Cell(1, 3))
                        WDocTbl.Cell(1, 2).Range.Text = "Координаты, м"
                        WDocTbl.Cell(1, 2).Range.Bold = 1
                        WDocTbl.Cell(1, 2).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                        WDocTbl.Cell(1, 2).Range.Font.Name = "Times New Roman"
                        WDocTbl.Cell(1, 2).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter

                        WDocTbl.Cell(2, 2).Range.Text = "X"
                        WDocTbl.Cell(2, 2).Range.Bold = 1
                        WDocTbl.Cell(2, 2).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                        WDocTbl.Cell(2, 2).Range.Font.Name = "Times New Roman"
                        WDocTbl.Cell(2, 2).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter

                        WDocTbl.Cell(2, 3).Range.Text = "Y"
                        WDocTbl.Cell(2, 3).Range.Bold = 1
                        WDocTbl.Cell(2, 3).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                        WDocTbl.Cell(2, 3).Range.Font.Name = "Times New Roman"
                        WDocTbl.Cell(2, 3).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter

                        Dim count As Integer = 3
                        Dim areaParcel As Double = 0
                        For j As Integer = 0 To arrayCoord.GetUpperBound(1)
                            Dim typeObject As String = arrayCoord(3, j)
                            If typeObject.Trim Like "PLINE" Or typeObject.Trim Like "REGION" Then
                                Dim numbCounter As String = arrayCoord(0, j)
                                Dim areaCounter As Double = Val(arrayCoord(2, j))
                                areaParcel += areaCounter
                                If Not (numbCounter Like "0") Then
                                    WDocTbl.Rows.Add()
                                    WDocTbl.Cell(count, 1).Merge(MergeTo:=WDocTbl.Cell(count, 3))
                                    WDocTbl.Cell(count, 1).Range.Text = "(" & arrayCoord(0, j) & ")"
                                    WDocTbl.Cell(count, 1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                                    WDocTbl.Cell(count, 1).Range.Font.Name = "Times New Roman"
                                    WDocTbl.Cell(count, 1).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                    count += 1
                                End If
                            ElseIf typeObject.Trim Like "Островок" Then
                                WDocTbl.Rows.Add()
                                WDocTbl.Cell(count, 1).Merge(MergeTo:=WDocTbl.Cell(count, 3))
                                WDocTbl.Cell(count, 1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                                WDocTbl.Cell(count, 1).Range.Font.Name = "Times New Roman"
                                WDocTbl.Cell(count, 1).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                count += 1
                            Else
                                WDocTbl.Rows.Add()
                                Dim num As String = arrayCoord(0, j)
                                If IsNothing(num) = False Then
                                    WDocTbl.Cell(count, 1).Range.Text = num
                                    WDocTbl.Cell(count, 1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                                    WDocTbl.Cell(count, 1).Range.Font.Name = "Times New Roman"
                                    WDocTbl.Cell(count, 1).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                End If
                                Dim X As String = arrayCoord(2, j)
                                If IsNothing(X) = False Then
                                    Dim xi As Double = Val(X)
                                    WDocTbl.Cell(count, 2).Range.Text = xi
                                    WDocTbl.Cell(count, 2).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                                    WDocTbl.Cell(count, 2).Range.Font.Name = "Times New Roman"
                                    WDocTbl.Cell(count, 2).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                End If
                                Dim Y As String = arrayCoord(1, j)
                                If IsNothing(Y) = False Then
                                    Dim yi As Double = Val(Y)
                                    WDocTbl.Cell(count, 3).Range.Text = yi
                                    WDocTbl.Cell(count, 3).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                                    WDocTbl.Cell(count, 3).Range.Font.Name = "Times New Roman"
                                    WDocTbl.Cell(count, 3).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                End If
                                count += 1
                            End If
                        Next j

                        WDocTbl.Rows.Add()
                        WDocTbl.Cell(count, 1).Range.Text = "Площадь: "
                        WDocTbl.Cell(count, 1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                        WDocTbl.Cell(count, 1).Range.Font.Name = "Times New Roman"
                        WDocTbl.Cell(count, 1).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter

                        WDocTbl.Cell(count, 2).Range.Text = Math.Round(areaParcel, 0)
                        WDocTbl.Cell(count, 2).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                        WDocTbl.Cell(count, 2).Range.Font.Name = "Times New Roman"
                        WDocTbl.Cell(count, 2).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter

                        WDocTbl.Cell(count, 3).Range.Text = "кв.м."
                        WDocTbl.Cell(count, 3).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                        WDocTbl.Cell(count, 3).Range.Font.Name = "Times New Roman"
                        WDocTbl.Cell(count, 3).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    Next i
                End If
                WDoc = Nothing
                WApp.Quit()
                WApp = Nothing
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'MS Excel
            ElseIf formReport.RadioButton2.Checked = True Then
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'запускаем Excel
                Dim xlApp As Excel.Application = Nothing
                Dim xlBook As Excel.Workbook = Nothing
                Dim xlSheets As Excel.Worksheet = Nothing
                Dim StartRows As Integer = 2
                If formReport.RadioButton2.Checked = True Then
                    xlApp = CreateObject("Excel.Application")
                    xlBook = xlApp.Workbooks.Add()
                End If
                xlApp.Visible = True
                '================================================================================
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'записываем данные
                If parcelDictionary.Count > 0 Then
                    Dim areaParcel As Double = 0
                    For i As Integer = 0 To parcelDictionary.Count - 1
                        Dim dictParcel As KeyValuePair(Of String, String(,)) = parcelDictionary.ElementAt(i)
                        Dim numberParcel As String = dictParcel.Key
                        Dim text1 As String = "Условный номер участка: " & numberParcel

                        Dim cell1 As Excel.Range 'первая ячейка в строке
                        Dim cell2 As Excel.Range 'последняя ячейка в строке(6)
                        Dim CellRng As Excel.Range 'массив ячеек
                        xlSheets = xlBook.ActiveSheet
                        If StartRows = 2 Then
                            xlSheets.Columns.Range("A:A").ColumnWidth = 16
                            xlSheets.Columns.Range("B:B").ColumnWidth = 13
                            xlSheets.Columns.Range("C:C").ColumnWidth = 13
                        End If

                        cell1 = xlSheets.Cells(StartRows, 1) 'ссылка на первую ячейку
                        cell2 = xlSheets.Cells(StartRows, 3) 'ссылка на 6 ячейку в строке
                        CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                        CellRng.Merge(Type.Missing)
                        cell1.Font.Name = "Times New Roman"
                        cell1.Font.Bold = 1
                        cell1.Font.Size = 12
                        cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                        cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                        CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                        cell1.Value = text1
                        '===========================================================================
                        cell1 = xlSheets.Cells(StartRows + 1, 1) 'ссылка на первую ячейку
                        cell2 = xlSheets.Cells(StartRows + 2, 1) 'ссылка на 6 ячейку в строке
                        CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                        CellRng.Merge(Type.Missing)
                        cell1.Font.Name = "Times New Roman"
                        cell1.Font.Bold = 1
                        cell1.Font.Size = 12
                        cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                        cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                        CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                        cell1.Value = "Обозначение" & vbLf & "характерных" & vbLf & "точек границ"
                        CellRng.WrapText = True
                        '==============================================================================
                        cell1 = xlSheets.Cells(StartRows + 1, 2) 'ссылка на первую ячейку
                        cell2 = xlSheets.Cells(StartRows + 1, 3) 'ссылка на 6 ячейку в строке
                        CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                        CellRng.Merge(Type.Missing)
                        cell1.Font.Name = "Times New Roman"
                        cell1.Font.Bold = 1
                        cell1.Font.Size = 12
                        cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                        cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                        CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                        cell1.Value = "Координаты, м"
                        '==============================================================================
                        cell1 = xlSheets.Cells(StartRows + 2, 2) 'ссылка на первую ячейку
                        cell1.Font.Name = "Times New Roman"
                        cell1.Font.Bold = 1
                        cell1.Font.Size = 12
                        cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                        cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                        cell1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                        cell1.Value = "X"
                        '==============================================================================
                        cell1 = xlSheets.Cells(StartRows + 2, 3) 'ссылка на первую ячейку
                        cell1.Font.Name = "Times New Roman"
                        cell1.Font.Bold = 1
                        cell1.Font.Size = 12
                        cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                        cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                        cell1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                        cell1.Value = "Y"
                        '====================================================================================
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'записываем координаты
                        StartRows = StartRows + 3
                        Dim arrayCoord As String(,) = dictParcel.Value
                        For j As Integer = 0 To arrayCoord.GetUpperBound(1)
                            Dim typeObject As String = arrayCoord(3, j)
                            If typeObject.Trim Like "PLINE" Or typeObject.Trim Like "REGION" Then
                                Dim numbCounter As String = arrayCoord(0, j)
                                Dim areaCounter As Double = Val(arrayCoord(2, j))
                                areaParcel += areaCounter
                                If Not (numbCounter Like "0") Then
                                    cell1 = xlSheets.Cells(StartRows, 1) 'ссылка на первую ячейку
                                    cell2 = xlSheets.Cells(StartRows, 3) 'ссылка на 3 ячейку в строке
                                    CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                                    CellRng.Merge(Type.Missing)
                                    cell1.Font.Name = "Times New Roman"
                                    cell1.Font.Size = 12
                                    cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                    cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                    CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                    cell1.Value = "(" & numbCounter & ")"
                                    StartRows += 1
                                End If
                            ElseIf typeObject.Trim Like "Островок" Then
                                cell1 = xlSheets.Cells(StartRows, 1) 'ссылка на первую ячейку
                                cell2 = xlSheets.Cells(StartRows, 3) 'ссылка на 3 ячейку в строке
                                CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                                CellRng.Merge(Type.Missing)
                                cell1.Font.Name = "Times New Roman"
                                cell1.Font.Size = 12
                                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                StartRows += 1
                            Else
                                'Координаты
                                cell1 = xlSheets.Cells(StartRows, 1) 'ссылка на первую ячейку
                                cell2 = xlSheets.Cells(StartRows, 3) 'ссылка на 3 ячейку в строке
                                CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                                CellRng.Font.Name = "Times New Roman"
                                CellRng.Font.Size = 12
                                CellRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                CellRng.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                'номер точки
                                xlSheets.Cells(StartRows, 1).Value = arrayCoord(0, j)
                                Dim X As String = arrayCoord(2, j)
                                If IsNothing(X) = False Then
                                    Dim xi As Double = Val(X)
                                    xi = Math.Round(xi, 2)
                                    cell1 = xlSheets.Cells(StartRows, 3)
                                    cell1.Value = FuncFormatZn.FuncFormatCoordinate(xi, roundCoord)
                                End If
                                Dim Y As String = arrayCoord(1, j)
                                If IsNothing(Y) = False Then
                                    Dim yi As Double = Val(Y)
                                    yi = Math.Round(yi, 2)
                                    cell1 = xlSheets.Cells(StartRows, 2)
                                    cell1.Value = FuncFormatZn.FuncFormatCoordinate(yi, roundCoord)
                                End If
                                StartRows += 1
                            End If
                        Next j
                        cell1 = xlSheets.Cells(StartRows, 1) 'ссылка на первую ячейку
                        cell2 = xlSheets.Cells(StartRows, 3) 'ссылка на 3 ячейку в строке
                        CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                        CellRng.Font.Name = "Times New Roman"
                        CellRng.Font.Size = 12
                        CellRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                        CellRng.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                        CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                        'номер точки
                        xlSheets.Cells(StartRows, 1).Value = "Площадь: "
                        cell1 = xlSheets.Cells(StartRows, 2)
                        cell1.Value = Math.Round(areaParcel, 0)
                        cell1 = xlSheets.Cells(StartRows, 3)
                        cell1.Value = "кв.м."
                    Next i
                End If
                xlBook = Nothing
                xlApp.Quit()
                xlApp = Nothing
            ElseIf formReport.RadioButton4.Checked = True Then
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'mid mif
                Dim ParcelMif1 As StreamWriter = Nothing
                Dim ParcelMid1 As StreamWriter = Nothing
                Dim TableMif As String = formReport.ComboBox1.Text
                '========================================================================
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                Dim strFilter As String = "mif files (*.mif)|*.mif|All files (*.*)|*.*"
                Dim nameFile As String = FuncFiles.FuncSaveFilePathName("Сохранить файл", strFilter)
                If nameFile.Trim.Length > 0 Then
                    'открываем 2 текстовые файла на запись
                    Dim ParcelFile1 As String = nameFile
                    Dim ParcelFile2 As String = nameFile & ".mid"
                    ParcelMif1 = New StreamWriter(ParcelFile1, False, System.Text.Encoding.Default)
                    ParcelMid1 = New StreamWriter(ParcelFile2, False, System.Text.Encoding.Default)

                    'заполняем заголовок таблицы Участки.mif
                    ParcelMif1.WriteLine("Version 300")
                    ParcelMif1.WriteLine("Delimiter " & """" & "," & """")
                    ParcelMif1.WriteLine("CoordSys NonEarth Units " & """" & "m" & """" & " Bounds(" & Math.Round(minX, 0) - 10 & ", " & Math.Round(minY, 0) - 10 & ")(" & Math.Round(maxX, 0) + 10 & ", " & Math.Round(maxY, 0) + 10 & ")")
                    'подсчитываем число полей
                    Dim dictionaryPSField As Dictionary(Of String, String(,)) = New Dictionary(Of String, String(,))
                    Dim boolXml As Boolean = FuncXRecords.FuncReadDataFieldsByTablePSToXML(templateXML, TableMif, dictionaryPSField)
                    Dim countColumns As Integer = 0
                    Dim arrayField As String(,) = Nothing
                    If dictionaryPSField.Count > 0 Then
                        For i As Integer = 0 To dictionaryPSField.Count - 1
                            Dim kf As KeyValuePair(Of String, String(,)) = dictionaryPSField.ElementAt(i)
                            Dim fieldName As String = kf.Key
                            If fieldName Like "LOCALID" Then
                                Continue For
                            ElseIf fieldName Like "GLOBALID" Then
                                Continue For
                            Else
                                Dim arrayF As String(,) = kf.Value
                                If IsArray(arrayF) = True Then
                                    ReDim Preserve arrayField(2, countColumns)
                                    arrayField(1, countColumns) = fieldName
                                    For j As Integer = 0 To arrayF.GetUpperBound(1)
                                        If arrayF(0, j).Trim Like "TypeField" Then
                                            arrayField(2, countColumns) = arrayF(1, j).Trim
                                        ElseIf arrayF(0, j).Trim Like "DisplayOrder" Then
                                            arrayField(0, countColumns) = arrayF(1, j).Trim
                                        End If
                                    Next j
                                    countColumns += 1
                                End If
                            End If
                        Next
                        If IsArray(arrayField) = True Then
                            If arrayField.GetUpperBound(1) > 0 Then
                                Dim boolSort As Boolean = MathFunction.FuncSortStrArray(arrayField, 0)
                            End If
                            ParcelMif1.WriteLine("COLUMNS " & arrayField.GetUpperBound(1) + 1)
                            For i As Integer = 0 To arrayField.GetUpperBound(1)
                                Dim typeField As String = "Char(1000)"
                                If Val(arrayField(2, i)) = 0 Then
                                    typeField = "Integer"
                                ElseIf Val(arrayField(2, i)) = 1 Then
                                    typeField = "Decimal"
                                End If
                                ParcelMif1.WriteLine("     " & arrayField(1, i) & " " & typeField)
                            Next
                            ParcelMif1.WriteLine("DATA")
                        Else
                            ParcelMif1.WriteLine("COLUMNS 1")
                            ParcelMif1.WriteLine("     blank char(1)")
                            ParcelMif1.WriteLine("DATA")
                        End If
                    Else
                        ParcelMif1.WriteLine("COLUMNS 1")
                        ParcelMif1.WriteLine("     blank char(1)")
                        ParcelMif1.WriteLine("DATA")
                    End If
                    ParcelMif1.WriteLine(" ")
                    If parcelDictionary.Count > 0 Then
                        For i As Integer = 0 To parcelDictionary.Count - 1
                            Dim dictParcel As KeyValuePair(Of String, String(,)) = parcelDictionary.ElementAt(i)
                            Dim numberParcel As String = dictParcel.Key
                            Dim arrayCoord As String(,) = dictParcel.Value
                            If IsArray(arrayCoord) = True Then
                                Try
                                    '==================================================================================
                                    Dim typeObject As String = ""
                                    For j1 As Integer = 0 To arrayCoord.GetUpperBound(1)
                                        typeObject = arrayCoord(3, j1)
                                        If typeObject.Trim Like "PLINE" Then
                                            ParcelMif1.WriteLine("PLINE" & " " & Val(arrayCoord(4, j1)))
                                        ElseIf typeObject.Trim Like "REGION" Then
                                            ParcelMif1.WriteLine("REGION " & Val(arrayCoord(4, j1)))
                                            ParcelMif1.WriteLine(Val("  " & Val(arrayCoord(5, j1))))
                                            'записываем mid файл
                                            Dim StrMid As String = ""
                                            Dim hgObject As String = arrayCoord(6, j1)
                                            Dim acEnt As DwgEntity = Nothing
                                            Dim boolFindobject As Boolean = ActivDocument.ActiveSpace.Entities.TryGetObject(hgObject, acEnt)
                                            If boolFindobject = True Then
                                                Dim arrayData As String(,) = Nothing
                                                Dim boolReadData As Boolean = FuncXRecords.FuncReadXData(acEnt, TableMif, arrayData)
                                                If IsArray(arrayData) = True Then
                                                    For j As Integer = 0 To arrayField.GetUpperBound(1)
                                                        Dim nameField1 As String = arrayField(1, j).Trim
                                                        For k As Integer = 0 To arrayData.GetUpperBound(1)
                                                            Dim nameField2 As String = arrayData(0, k).Trim
                                                            If nameField1 Like nameField2 Then
                                                                If StrMid.Trim.Length = 0 Then
                                                                    StrMid = Chr(34) & arrayData(1, k) & Chr(34)
                                                                Else
                                                                    StrMid = StrMid & "," & Chr(34) & arrayData(1, k) & Chr(34)
                                                                End If
                                                            End If
                                                        Next k
                                                    Next j
                                                End If
                                                If StrMid.Trim.Length > 0 Then
                                                    ParcelMid1.WriteLine(StrMid)
                                                    StrMid = ""
                                                End If
                                            End If

                                        ElseIf typeObject.Trim Like "Островок" Then
                                            ParcelMif1.WriteLine(Val("  " & Val(arrayCoord(5, j1))))
                                        Else
                                            Dim numPoint As String = arrayCoord(0, j1)
                                            Dim x As Double = Val(arrayCoord(1, j1))
                                            Dim y As Double = Val(arrayCoord(2, j1))
                                            ParcelMif1.WriteLine(Math.Round(x, roundCoord) & " " & Math.Round(y, roundCoord))
                                        End If
                                    Next j1
                                    If arrayCoord(3, 0).Trim Like "REGION" Then
                                        ParcelMif1.WriteLine("Pen(12, 2, 16711680)")
                                        ParcelMif1.WriteLine("BRUSH(2,16777215)")
                                    ElseIf arrayCoord(3, 0).Trim Like "PLINE" Then
                                        ParcelMif1.WriteLine("Pen(12, 2, 16711680)")
                                    Else
                                        ParcelMif1.WriteLine("SYMBOL (35,16777215,10)")
                                    End If

                                Catch ex As Exception
                                End Try
                            End If
                        Next i
                    End If
                    'закрываем файлы
                    ParcelMid1.Close()
                    ParcelMif1.Close()
                End If
            ElseIf formReport.RadioButton1.Checked = True Then
                'таблица Робур
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                If parcelDictionary.Count > 0 Then
                    Dim posTable As Vector3D = New Vector3D(0, 0, 0)
                    Dim areaParcel As Double = 0
                    Dim sizeTable As Double = 0
                    For i As Integer = 0 To parcelDictionary.Count - 1
                        If i = 0 Then
                            CadCursors.GetPoint(cadWiever, posTable, "Укажите точку вставки таблицы: ")
                        End If
                        Dim dictParcel As KeyValuePair(Of String, String(,)) = parcelDictionary.ElementAt(i)
                        Dim numberParcel As String = dictParcel.Key
                        Dim arrayCoord As String(,) = dictParcel.Value
                        If IsArray(arrayCoord) = True Then
                            Try
                                '==================================================================================
                                Dim userTableStyle As DwgTableStyle = ActivDocument.ActiveTableStyle
                                Dim userTable As Topomatic.Dwg.Entities.DwgTable = New Topomatic.Dwg.Entities.DwgTable(userTableStyle, 4, 3, 3, 50)
                                userTable.MergeCells(0, 0, 2, 0)
                                userTable.MergeCells(1, 1, 2, 1)
                                userTable.MergeCells(0, 1, 0, 2)
                                userTable.Item(0, 0).SourceText = "Каталог координат "
                                userTable.Item(1, 0).SourceText = "Обозначение характерных точек границ"
                                userTable.Item(1, 1).SourceText = "Координаты, м"
                                userTable.Item(2, 1).SourceText = "X"
                                userTable.Item(2, 2).SourceText = "Y"
                                userTable.Position = posTable
                                Dim countRows As Integer = 3
                                If IsArray(arrayCoord) = True Then
                                    For j As Integer = 0 To arrayCoord.GetUpperBound(1)
                                        Dim typeObject As String = arrayCoord(3, j)
                                        If typeObject.Trim Like "PLINE" Or typeObject.Trim Like "REGION" Then
                                            Dim numbCounter As String = arrayCoord(0, j)
                                            Dim areaCounter As Double = Val(arrayCoord(2, j))
                                            areaParcel += areaCounter
                                            If Not (numbCounter Like "0") Then
                                                userTable.AddRow(1)
                                                Dim countRow As Integer = userTable.RowsCount
                                                userTable.MergeCells(0, countRows, 2, countRows)
                                                userTable.Item(countRows, 0).SourceText = "(" & numbCounter & ")"
                                                countRows += 1
                                            End If
                                        ElseIf typeObject.Trim Like "Островок" Then
                                            userTable.AddRow(1)
                                            userTable.MergeCells(0, countRows, 2, countRows)
                                            countRows += 1
                                        Else
                                            Dim numb As String = arrayCoord(0, j)
                                            Dim x As Double = Val(arrayCoord(2, j))
                                            Dim y As Double = Val(arrayCoord(1, j))
                                            Dim strX As String = FuncFormatZn.FuncFormatCoordinate(x, roundCoord)
                                            Dim strY As String = FuncFormatZn.FuncFormatCoordinate(y, roundCoord)
                                            userTable.AddRow(1)
                                            userTable.Item(countRows, 0).SourceText = numb
                                            userTable.Item(countRows, 1).SourceText = strX
                                            userTable.Item(countRows, 2).SourceText = strY
                                            countRows += 1
                                        End If
                                    Next j
                                    userTable.AddRow(1)
                                    userTable.Item(countRows, 0).SourceText = "Площадь:"
                                    userTable.Item(countRows, 1).SourceText = Math.Round(areaParcel, 0)
                                    userTable.Item(countRows, 2).SourceText = "кв.м."
                                End If
                                ActivDocument.ActiveSpace.Add(userTable)
                                Dim userBound As BoundingBox2D = userTable.Bounds
                                Dim offset As Double = (userBound.Max.X - userBound.Min.X) + 5
                                posTable = New Vector3D(posTable.X + offset, posTable.Y, 0)
                            Catch ex As Exception
                            End Try
                        End If
                    Next i
                End If
            End If
        End Sub
        '===================================================================================================================================
        'Сведения о наличии зон с особыми условиями использования территории
        <cmd("ReportCadastralZones")>
        Public Sub ReportCadastralZones()
            'получаем активное окно
            Dim userCadView As CadView = Me.CadView
            If IsNothing(userCadView) Then Return
            'получаем активный слой
            Dim DrawLayer As DrawingLayer = DrawingLayer.GetDrawingLayer(userCadView)
            If IsNothing(DrawLayer) = True Then Return
            'получаем активный документ
            Dim ActivDocument As Topomatic.Dwg.Drawing = DrawLayer.Drawing
            If IsNothing(ActivDocument) = True Then Return
            '================================================================================================================
            'получаем все трассы
            Dim arrayAlign As Alignment() = Nothing
            Dim countAlign As Integer = 0
            Dim iUserProject As IProjectModel() = PluginCoreOps.FindModels(New String() {"road"})
            If IsArray(iUserProject) = True Then
                For i As Integer = 0 To iUserProject.Length - 1
                    Dim userProject As IProjectModel = iUserProject(i)
                    Try
                        Dim am As IAlignmentContainer = userProject.LockReadContainer(Of IAlignmentContainer)()
                        If IsNothing(am) = False Then
                            Dim align As Topomatic.Alg.Alignment = am.Alignment
                            ReDim Preserve arrayAlign(countAlign)
                            arrayAlign(countAlign) = align
                            countAlign += 1
                        End If
                    Catch ex As Exception
                    End Try
                Next
            End If
            'массив для записи
            '0-StartPK, 1-off, 2-EndPK, 3-off, 4-Реесстровый номер, 5-Тип зоны, 6-Описание
            Dim arrayZone As String(,) = Nothing
            Dim countArrayZone As Integer = 0
            '0-StartPK, 1-Тип сооружения, 2-Характеристика, 3-Примечание
            Dim arrayBuilds As String(,) = Nothing
            Dim countArrayBuids As Integer = 0
            '0-StartPK, 1-Кадастровый номер, 2-Площадь, 4-Разрешенное использование, 5-Вид права, 6-Примечание
            Dim arrayParcels As String(,) = Nothing
            Dim countArrayParcels As Integer = 0
            '0-StartPK, 1-EndPK, 2-Наименование, 3-Характеристика, 4-Протяженность
            Dim arrayConstr As String(,) = Nothing
            Dim countArrayConstr As Integer = 0

            'сброс выделения
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            userCadView.SelectionSet.Clear()
            userCadView.SelectionSet.SelectOneObjectAtScreen(Function(obj) TypeOf obj Is Object, "Выберите границу полосы отвода: ")
            Dim p3d1 As Polyline3D = New Polyline3D()
            Dim boundPoly As BoundingBox2D = Nothing
            Dim l1 As List(Of Vector2D) = New List(Of Vector2D)
            If userCadView.SelectionSet.Count > 0 Then
                For Each acEnt As Object In userCadView.SelectionSet
                    If TypeOf acEnt Is DwgPolyline Then
                        Dim acPoly As DwgPolyline = acEnt
                        boundPoly = acPoly.Bounds
                        acPoly.GetPolyline(p3d1)
                        p3d1.ConvertToPosArray(l1, ValueConverter.Eps)
                    End If
                Next
                '=============================================================================================================
                'выбираем зоны
                userCadView.SelectionSet.Clear()
                userCadView.SelectionSet.SelectObjectsAtScreen(Function(obj) TypeOf obj Is Object, "Выборите объекты: ")
                For Each acZoneEnt As Object In userCadView.SelectionSet
                    Try
                        If TypeOf acZoneEnt Is DwgPolyline Then
                            Dim plineCadastralZone As DwgPolyline = acZoneEnt
                            'ищем пересечения
                            Dim boolIntersect As Boolean = plineCadastralZone.IntersectWith(boundPoly, 1)
                            If boolIntersect = True Then 'пересечение найдено
                                'читаем расширенные данные
                                Dim arrayTbl As String() = Nothing
                                Dim boolFindPs As Boolean = FuncXRecords.FuncReadTablesPS(plineCadastralZone, arrayTbl)
                                If IsArray(arrayTbl) = True Then 'массив с таблицами найден
                                    For i As Integer = 0 To arrayTbl.Length - 1
                                        If arrayTbl(i) Like "LAND_CADASTRAL_ZONE" Or arrayTbl(i) Like "PROJECT_ZONA" Then
                                            Dim p3d2 As Polyline3D = New Polyline3D()
                                            plineCadastralZone.GetPolyline(p3d2)
                                            Dim l2 As List(Of Vector2D) = New List(Of Vector2D)
                                            p3d2.ConvertToPosArray(l2, ValueConverter.Eps)
                                            Dim op = New PolygonOperation()
                                            Dim l3 As List(Of List(Of Vector2D)) = New List(Of List(Of Vector2D))
                                            l3 = op.Intersection(l1, l2)
                                            'зона пересекает полосу отвода
                                            If l3.Count > 0 Then
                                                'читаем характеристики зоны
                                                Dim arrayData As String(,) = Nothing
                                                ReDim Preserve arrayZone(6, countArrayZone)
                                                arrayZone(0, countArrayZone) = "" 'начальное значение пикета начала
                                                arrayZone(2, countArrayZone) = 0 'начальное значение конечного пикета
                                                If arrayTbl(i) Like "LAND_CADASTRAL_ZONE" Then
                                                    Dim boolReadData As Boolean = FuncXRecords.FuncReadXData(plineCadastralZone, "LAND_CADASTRAL_ZONE", arrayData)
                                                    'массив с характеристиками найден
                                                    If IsArray(arrayData) = True Then
                                                        For j As Integer = 0 To arrayData.GetUpperBound(1)
                                                            If arrayData(0, j) Like "Reg_numb_border" Then
                                                                arrayZone(4, countArrayZone) = arrayData(1, j)
                                                            ElseIf arrayData(0, j) Like "Type_zone" Then
                                                                arrayZone(5, countArrayZone) = arrayData(1, j)
                                                            ElseIf arrayData(0, j) Like "ContentRestrictions" Then
                                                                arrayZone(6, countArrayZone) = arrayData(1, j)
                                                            End If
                                                        Next j
                                                    End If
                                                Else
                                                    Dim boolReadData As Boolean = FuncXRecords.FuncReadXData(plineCadastralZone, "PROJECT_ZONA", arrayData)
                                                    'массив с характеристиками найден
                                                    If IsArray(arrayData) = True Then
                                                        For j As Integer = 0 To arrayData.GetUpperBound(1)
                                                            If arrayData(0, j) Like "Name" Then
                                                                arrayZone(5, countArrayZone) = arrayData(1, j)
                                                            ElseIf arrayData(0, j) Like "KeyParameter" Then
                                                                arrayZone(6, countArrayZone) = arrayData(1, j)
                                                            End If
                                                        Next j
                                                    End If
                                                End If
                                                'находим пикетаж по каждому пересечению
                                                For j As Integer = 0 To l3.Count - 1
                                                    Dim l4 As List(Of Vector2D) = l3.Item(j)
                                                    Dim PKStartG As Double = 99999999999
                                                    Dim PKEndG As Double = -9999999999
                                                    'перебираем отдельные точки и находим пикетаж
                                                    For k As Integer = 0 To l4.Count - 1
                                                        Dim pos As Vector2D = l4.Item(k)
                                                        For k2 As Integer = 0 To arrayAlign.Count - 1
                                                            Dim pkStart As Double = 0
                                                            Dim pkOff As Double = 0
                                                            Dim tempAlign As Topomatic.Alg.Alignment = arrayAlign(k2)
                                                            Dim boolIns As Boolean = tempAlign.Plan.CompoundLine.PosToStaOffset(pos, pkStart, pkOff)
                                                            If boolIns = True Then
                                                                If pkStart < PKStartG Then
                                                                    PKStartG = pkStart
                                                                End If
                                                                If pkStart > PKEndG Then
                                                                    PKEndG = pkStart
                                                                End If
                                                            End If
                                                        Next k2
                                                    Next k
                                                    If arrayZone(0, countArrayZone).Trim.Length = 0 Then
                                                        arrayZone(0, countArrayZone) = "ПК" & FuncFormatZn.FuncFormatPK(PKStartG, 2) & " - ПК" & FuncFormatZn.FuncFormatPK(PKEndG, 2)
                                                    Else
                                                        arrayZone(0, countArrayZone) = vbLf & "ПК" & FuncFormatZn.FuncFormatPK(PKStartG, 2) & " - ПК" & FuncFormatZn.FuncFormatPK(PKEndG, 2)
                                                    End If
                                                Next j
                                                countArrayZone += 1
                                            End If
                                        ElseIf arrayTbl(i) Like "LAND_CADASTRAL_BUILDING" Or arrayTbl(i) Like "PROJECT_BUILDING" Then
                                            Dim p3d2 As Polyline3D = New Polyline3D()
                                            plineCadastralZone.GetPolyline(p3d2)
                                            Dim l2 As List(Of Vector2D) = New List(Of Vector2D)
                                            p3d2.ConvertToPosArray(l2, ValueConverter.Eps)
                                            Dim op = New PolygonOperation()
                                            Dim l3 As List(Of List(Of Vector2D)) = New List(Of List(Of Vector2D))
                                            l3 = op.Intersection(l1, l2)
                                            'зона пересекает полосу отвода
                                            If l3.Count > 0 Then
                                                'читаем характеристики зоны
                                                Dim arrayData As String(,) = Nothing
                                                ReDim Preserve arrayZone(6, countArrayZone)
                                                arrayZone(0, countArrayZone) = "" 'начальное значение пикета начала
                                                arrayZone(2, countArrayZone) = 0 'начальное значение конечного пикета
                                                If arrayTbl(i) Like "LAND_CADASTRAL_BUILDING" Then
                                                    Dim boolReadData As Boolean = FuncXRecords.FuncReadXData(plineCadastralZone, "LAND_CADASTRAL_BUILDING", arrayData)
                                                    'массив с характеристиками найден
                                                    If IsArray(arrayData) = True Then
                                                        For j As Integer = 0 To arrayData.GetUpperBound(1)
                                                            If arrayData(0, j) Like "Reg_numb_border" Then
                                                                arrayZone(4, countArrayZone) = arrayData(1, j)
                                                            ElseIf arrayData(0, j) Like "Type_zone" Then
                                                                arrayZone(5, countArrayZone) = arrayData(1, j)
                                                            ElseIf arrayData(0, j) Like "ContentRestrictions" Then
                                                                arrayZone(6, countArrayZone) = arrayData(1, j)
                                                            End If
                                                        Next j
                                                    End If
                                                Else
                                                    Dim boolReadData As Boolean = FuncXRecords.FuncReadXData(plineCadastralZone, "PROJECT_ZONA", arrayData)
                                                    'массив с характеристиками найден
                                                    If IsArray(arrayData) = True Then
                                                        For j As Integer = 0 To arrayData.GetUpperBound(1)
                                                            If arrayData(0, j) Like "Name" Then
                                                                arrayZone(5, countArrayZone) = arrayData(1, j)
                                                            ElseIf arrayData(0, j) Like "KeyParameter" Then
                                                                arrayZone(6, countArrayZone) = arrayData(1, j)
                                                            End If
                                                        Next j
                                                    End If
                                                End If
                                                'находим пикетаж по каждому пересечению
                                                For j As Integer = 0 To l3.Count - 1
                                                    Dim l4 As List(Of Vector2D) = l3.Item(j)
                                                    Dim PKStartG As Double = 99999999999
                                                    Dim PKEndG As Double = -9999999999
                                                    'перебираем отдельные точки и находим пикетаж
                                                    For k As Integer = 0 To l4.Count - 1
                                                        Dim pos As Vector2D = l4.Item(k)
                                                        For k2 As Integer = 0 To arrayAlign.Count - 1
                                                            Dim pkStart As Double = 0
                                                            Dim pkOff As Double = 0
                                                            Dim tempAlign As Topomatic.Alg.Alignment = arrayAlign(k2)
                                                            Dim boolIns As Boolean = tempAlign.Plan.CompoundLine.PosToStaOffset(pos, pkStart, pkOff)
                                                            If boolIns = True Then
                                                                If pkStart < PKStartG Then
                                                                    PKStartG = pkStart
                                                                End If
                                                                If pkStart > PKEndG Then
                                                                    PKEndG = pkStart
                                                                End If
                                                            End If
                                                        Next k2
                                                    Next k
                                                    If arrayZone(0, countArrayZone).Trim.Length = 0 Then
                                                        arrayZone(0, countArrayZone) = "ПК" & FuncFormatZn.FuncFormatPK(PKStartG, 2) & " - ПК" & FuncFormatZn.FuncFormatPK(PKEndG, 2)
                                                    Else
                                                        arrayZone(0, countArrayZone) = vbLf & "ПК" & FuncFormatZn.FuncFormatPK(PKStartG, 2) & " - ПК" & FuncFormatZn.FuncFormatPK(PKEndG, 2)
                                                    End If
                                                Next j
                                                countArrayZone += 1
                                            End If
                                        End If


                                    Next i
                                End If
                            End If
                        End If
                    Catch ex As System.InvalidOperationException
                    End Try
                Next
            End If

            If IsArray(arrayZone) = True Then
                'записываем таблицу Word
                Dim WApp As Word.Application = Nothing
                Dim WDoc As Word.Document = Nothing
                Dim boolWDVisibl As Boolean = False
                WApp = CreateObject("Word.Application")
                WDoc = WApp.Documents.Add()
                WApp.Visible = True

                Dim SetPos As Integer = WDoc.Content.End - 1
                WDoc.Range(SetPos).Paragraphs.Add()
                SetPos = WDoc.Content.End - 1
                WDoc.Range(SetPos).Font.Name = "Times New Roman"
                WDoc.Range(SetPos).Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                Dim text1 As String = "Таблица №1. Сведения о наличии зон с особыми условиями использования территории"
                Dim ltext As Integer = text1.Length
                WDoc.Range(SetPos).Text = text1
                WDoc.Range(SetPos + ltext).Paragraphs.Add()
                SetPos = SetPos + ltext + 1

                Dim range As Word.Range = WDoc.Range(SetPos)
                WDoc.Range(SetPos).Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                Dim WDocTbl As Word.Table = WDoc.Tables.Add(range, 2, 4, Word.WdDefaultTableBehavior.wdWord9TableBehavior, Word.WdAutoFitBehavior.wdAutoFitContent)

                WDocTbl.Style = "Сетка таблицы"
                WDocTbl.ApplyStyleRowBands = True
                WDocTbl.Cell(1, 1).Range.Text = "№ п/п"
                WDocTbl.Cell(1, 1).Range.Bold = 1
                WDocTbl.Cell(1, 1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                WDocTbl.Cell(1, 1).Range.Font.Name = "Times New Roman"
                WDocTbl.Cell(1, 1).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter

                WDocTbl.Cell(1, 2).Range.Text = "Местоположение, ПК"
                WDocTbl.Cell(1, 2).Range.Bold = 1
                WDocTbl.Cell(1, 2).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                WDocTbl.Cell(1, 2).Range.Font.Name = "Times New Roman"
                WDocTbl.Cell(1, 2).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter

                WDocTbl.Cell(1, 3).Range.Text = "Наименование зоны"
                WDocTbl.Cell(1, 3).Range.Bold = 1
                WDocTbl.Cell(1, 3).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                WDocTbl.Cell(1, 3).Range.Font.Name = "Times New Roman"
                WDocTbl.Cell(1, 3).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter

                WDocTbl.Cell(1, 4).Range.Text = "Реестровый номер"
                WDocTbl.Cell(1, 4).Range.Bold = 1
                WDocTbl.Cell(1, 4).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                WDocTbl.Cell(1, 4).Range.Font.Name = "Times New Roman"
                WDocTbl.Cell(1, 4).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter

                Dim count As Integer = 2
                Dim areaParcel As Double = 0
                For j As Integer = 0 To arrayZone.GetUpperBound(1)
                    WDocTbl.Rows.Add()
                    WDocTbl.Cell(count, 1).Range.Text = j + 1
                    WDocTbl.Cell(count, 1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                    WDocTbl.Cell(count, 1).Range.Font.Name = "Times New Roman"
                    WDocTbl.Cell(count, 1).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter

                    WDocTbl.Cell(count, 2).Range.Text = arrayZone(0, j)
                    WDocTbl.Cell(count, 2).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                    WDocTbl.Cell(count, 2).Range.Font.Name = "Times New Roman"
                    WDocTbl.Cell(count, 2).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter

                    WDocTbl.Cell(count, 3).Range.Text = arrayZone(5, j)
                    WDocTbl.Cell(count, 3).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                    WDocTbl.Cell(count, 3).Range.Font.Name = "Times New Roman"
                    WDocTbl.Cell(count, 3).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter

                    WDocTbl.Cell(count, 4).Range.Text = arrayZone(4, j)
                    WDocTbl.Cell(count, 4).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                    WDocTbl.Cell(count, 4).Range.Font.Name = "Times New Roman"
                    WDocTbl.Cell(count, 4).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    count += 1
                Next j
                WDoc = Nothing
                WApp = Nothing
            End If
        End Sub

        'Сведения о наличии зон с особыми условиями использования территории
        <cmd("UpdateProgram")>
        Public Sub UpdateProgram()
            'получаем активное окно
            Dim userCadView As CadView = Me.CadView
            If IsNothing(userCadView) Then Return
            'получаем активный слой
            Dim DrawLayer As DrawingLayer = DrawingLayer.GetDrawingLayer(userCadView)
            If IsNothing(DrawLayer) = True Then Return
            'получаем активный документ
            Dim ActivDocument As Topomatic.Dwg.Drawing = DrawLayer.Drawing
            If IsNothing(ActivDocument) = True Then Return

            ' Dim tester As New GoogleDriveTester()
            ' Dim folderUrl As String = "https://drive.google.com/drive/folders/1OWYT_urmd3QMrX0ZZPz5r3Uj3NxvXZrV?usp=drive_link"
            'tester.TestFolderAccess(folderUrl)

            Dim formDownload As CadastralTools.FormUpdateProgram = New FormUpdateProgram
            formDownload.ShowDialog()



        End Sub


    End Class
End Namespace
