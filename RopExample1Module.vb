Imports Topomatic.ApplicationPlatform
Imports System.Windows.Forms
Imports Topomatic.ApplicationPlatform.Plugins
Imports Topomatic.Dwg.Layer
Imports Topomatic.Cad.View
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.Cad.Foundation
Imports Topomatic.Cad.View.Hints
Imports Topomatic.Sfc
Imports Topomatic.Alg.Prf
Imports Topomatic.Visualization.Runtime
Imports Topomatic.Visualization
Imports Topomatic.Alg.Road.Urb.UserStrips
Imports Topomatic.ApplicationPlatform.Core
Imports Topomatic.Planchet.Entities
Imports System.IO
Imports Topomatic.Crs.Templates
Imports Topomatic.Alg.Crs
Imports Topomatic.FoundationClasses
Imports Topomatic.Alg.Road.Urb.Border
Imports Topomatic.Alg
Imports Topomatic.Smt
Imports Topomatic.Alg.Plan
Imports Topomatic.Alg.Road
Imports Topomatic.Alg.Road.Urb.Descent
Imports System.Windows.Documents
Imports System.Environment
Imports System.Text
Imports Microsoft.Office.Interop
Imports System.Xml
Imports Topomatic.Alg.Runtime.Tools.StraighteningSimplePlanSolver
Imports Topomatic.Sfc.Style
Imports System.Net.Security
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.ToolTip
Imports Topomatic.Pipes.Layers.Caches.Plan.Segments.SegmentPlanCache.DrawingData
Imports Topomatic.Planchet
Imports System.Drawing
Imports Topomatic.Pipes.Runtime.Plt.PlanCross.Fields
Imports Topomatic.Pipes.Layers.Profile.StaticNodeDrawers.NodeExtendedHeaderDrawer
Imports Topomatic.Alg.Runtime.ServiceClasses
Imports Topomatic.Controls.Dialogs
Imports Topomatic.Acax.Export
Imports Topomatic
Imports Topomatic.Pipes.Runtime.Plt.Profile.DescrEnumConverters
Imports Topomatic.FoundationClasses.Vcs
Imports System.Windows.Shapes
Imports System.Windows.Media.Media3D
Imports Path = System.IO.Path
Imports Vector3D = Topomatic.Cad.Foundation.Vector3D
Imports System.Media
Imports System.ComponentModel
Imports MS.Internal
Imports System.Net.Mime.MediaTypeNames
Imports Topomatic.Cad.Foundation.Brep
Imports Topomatic.Cad.View.Tools
Imports Topomatic.Alg.Road.Urb.Descent.DescentStrip
Imports Topomatic.Visualization.Constructions
Imports Topomatic.Alg.Runtime.Tools.ProfileVisibleCalculator
Imports System.Windows.Forms.Control
Imports Topomatic.ApplicationPlatform.Core.ProjectStateConfiguration
Imports Topomatic.Arrangements
Imports Topomatic.FoundationClasses.Undo
Imports Topomatic.Alg.Road.Core
Imports Topomatic.Alg.Layers.Wrappers
Imports System.Net
Imports Newtonsoft.Json
Imports System.Runtime.InteropServices.ComTypes
Imports Topomatic.FoundationClasses.Lisp.LMath
Imports System.Windows.Media.Animation
Imports Microsoft.Office.Interop.Excel
Imports System.Windows.Controls
Imports Topomatic.Sfc.Layer
Imports Topomatic.Visualization.Geometry
Imports Topomatic.Visualization.LinearSolidBuilder
Imports Topomatic.Cad.View.Design
Imports System.IO.Packaging
Imports Topomatic.Alg.Bridges
Imports System.Security.Cryptography
Imports NetTopologySuite.Geometries
Imports NetTopologySuite.Algorithm.Hull
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
            If IO.Directory.Exists(userDirFont) = True Then
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
                        If IO.Directory.Exists(tempdir) = True Then
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
        <cmd("CreateSymbolTOPO")>
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
                    nameUserFile = IO.Path.GetFileNameWithoutExtension(nameUserFile)
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
                    If TypeOf acEnt Is DwgObject Then
                        Dim acObj As DwgObject = acEnt
                        If Not (acObj.ObjectID Like userControlPanelPropertiesObject.ComboBox1.Tag) Then
                            Dim arrayTablesPS As String() = Nothing
                            Dim boolfindTables As Boolean = FuncXRecords.FuncReadTablesPS(acObj, arrayTablesPS)
                            If IsArray(arrayTablesPS) = True Then
                                For i As Integer = 0 To arrayTablesPS.Length - 1
                                    Dim nameTablePs As String = arrayTablesPS(i)
                                    If IsNothing(nameTablePs) = False Then
                                        If nameTablePs.Trim.Length > 0 Then
                                            Dim arrayData As String(,) = Nothing
                                            Dim boolFindProperties As Boolean = FuncXRecords.FuncReadXData(acObj, nameTablePs, arrayData)
                                            If IsArray(arrayData) = True Then
                                                userControlPanelPropertiesObject.DataGridView1.Rows.Clear()
                                                Dim countRow As Integer = 0
                                                For j As Integer = 0 To arrayData.GetUpperBound(1)
                                                    Dim nameF As String = arrayData(0, j)
                                                    Dim valF As String = arrayData(1, j)
                                                    If IsNothing(nameF) = False Then
                                                        If nameF.Trim.Length > 0 Then
                                                            If IsNothing(valF) = True Then valF = ""
                                                            userControlPanelPropertiesObject.DataGridView1.Rows.Add()
                                                            userControlPanelPropertiesObject.DataGridView1.Rows(countRow).Cells(0).Value = nameF
                                                            userControlPanelPropertiesObject.DataGridView1.Rows(countRow).Cells(1).Value = valF
                                                            countRow += 1
                                                        End If
                                                    End If
                                                Next j
                                            End If
                                        End If
                                    End If
                                Next i
                            End If
                            If IsArray(arrayTablesPS) = True Then
                                userControlPanelPropertiesObject.ComboBox1.DataSource = arrayTablesPS
                                userControlPanelPropertiesObject.ComboBox1.Tag = acObj.ObjectID
                                Exit For
                            End If
                        End If
                    End If
                    Exit For
                Next
            Else
                userControlPanelPropertiesObject.DataGridView1.Rows.Clear()
                userControlPanelPropertiesObject.ComboBox1.Tag = ""
                Dim arrayTemp As String() = Nothing
                userControlPanelPropertiesObject.ComboBox1.DataSource = arrayTemp
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
                                userpattern = FuncStyles.FuncReturnPatternLineTypeCode(lineTypeUser)
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
            Dim FormKPTImport As FormRosreestrKPTImport = New FormRosreestrKPTImport
            'запускаем форму
            FormKPTImport.ShowDialog()
            If FormKPTImport.boolShow = False Then
                Exit Sub
            End If
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
            Dim boolFindFileTemplate As Boolean = FuncFiles.FuncReadProperties(namePrograms, generalDir, templateXML, templateDWG)
            '=========================================================================================================================
            Dim activUserModel As String = FormKPTImport.ComboBox1.Text
            Dim userProject As ModelProject = ApplicationHost.Current.ActiveProject
            '// Получение всех моделей проекта
            Dim childs As IProjectModel() = userProject.Model.GetChilds()
            Dim userIProjectModel As IProjectModel = Nothing
            For Each userIProjectModel In childs
                Dim modelUri As URI = userIProjectModel.Uri
                If modelUri.Extension Like ".arrx" Then
                    Dim fileNameSfcx As String = Path.GetFileNameWithoutExtension(modelUri.LastPathComponent)
                    If fileNameSfcx.Trim Like activUserModel.Trim Then
                        ApplicationHost.Current.Plugins.Execute("activate", New Object() {userIProjectModel})
                        Exit For
                    End If
                End If
            Next
            Dim DrawShowLayers As IEnumerable(Of DrawingLayer) = DrawingLayer.GetDrawingLayers(userCadView)
            If IsNothing(activUserModel) = False Then
                If activUserModel.Trim.Length > 0 Then
                    For k1 As Integer = 0 To DrawShowLayers.Count - 1
                        Dim TempDrawLayer As DrawingLayer = DrawShowLayers.ElementAt(k1)
                        Dim nameUserFile As String = TempDrawLayer.Drawing.Filename
                        nameUserFile = IO.Path.GetFileNameWithoutExtension(nameUserFile)
                        If nameUserFile Like activUserModel Then
                            DrawLayer = TempDrawLayer
                            ActivDocument = DrawLayer.Drawing
                            Exit For
                        End If
                    Next k1
                End If
            End If
            userIProjectModel.LockWrite()
            ActivDocument.BeginUpdate()
            Try




            Finally
                ActivDocument.EndUpdate()
                userIProjectModel.UnlockWrite()
            End Try
        End Sub
        '==============================================================================================================================
        'Фильтрация КПТ
        <cmd("FiltreKPT")>
        Public Sub FiltreKPT()
            'получаем активное окно
            Dim userCadView As CadView = Me.CadView
            If IsNothing(userCadView) Then Return
            'получаем активный слой
            Dim DrawLayer As DrawingLayer = DrawingLayer.GetDrawingLayer(userCadView)
            If IsNothing(DrawLayer) = True Then Return
            'получаем активный документ
            Dim ActivDocument As Topomatic.Dwg.Drawing = DrawLayer.Drawing
            If IsNothing(ActivDocument) = True Then Return
            Dim userProject As ModelProject = ApplicationHost.Current.ActiveProject
            'передаем масштаб чертежа в форму
            Dim FormKPTImport As FormCreateCadastralPlan = New FormCreateCadastralPlan
            'запускаем форму
            FormKPTImport.ShowDialog()
            If FormKPTImport.boolShow = False Then
                Exit Sub
            End If
            Dim activUserModel As String = FormKPTImport.ComboBox1.Text
            '1.\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Dim namePrograms As String = FuncFiles.namePrograms
            Dim generalDir As String = ""
            Dim templateXML As String = ""
            Dim templateDWG As String = ""
            Dim boolFindFileTemplate As Boolean = FuncFiles.FuncReadProperties(namePrograms, generalDir, templateXML, templateDWG)
            '// Получение всех моделей проекта
            Dim childs As IProjectModel() = userProject.Model.GetChilds()
            Dim userIProjectModel As IProjectModel = Nothing
            For Each userIProjectModel In childs
                Dim modelUri As URI = userIProjectModel.Uri
                If modelUri.Extension Like ".arrx" Then
                    Dim fileNameSfcx As String = Path.GetFileNameWithoutExtension(modelUri.LastPathComponent)
                    If fileNameSfcx.Trim Like activUserModel.Trim Then
                        ApplicationHost.Current.Plugins.Execute("activate", New Object() {userIProjectModel})
                        Exit For
                    End If
                End If
            Next
            Dim DrawShowLayers As IEnumerable(Of DrawingLayer) = DrawingLayer.GetDrawingLayers(userCadView)
            If IsNothing(activUserModel) = False Then
                If activUserModel.Trim.Length > 0 Then
                    For k1 As Integer = 0 To DrawShowLayers.Count - 1
                        Dim TempDrawLayer As DrawingLayer = DrawShowLayers.ElementAt(k1)
                        Dim nameUserFile As String = TempDrawLayer.Drawing.Filename
                        nameUserFile = IO.Path.GetFileNameWithoutExtension(nameUserFile)
                        If nameUserFile Like activUserModel Then
                            DrawLayer = TempDrawLayer
                            ActivDocument = DrawLayer.Drawing
                            Exit For
                        End If
                    Next k1
                End If
            End If

            Dim acPlineBoundares As DwgPolyline = Nothing
            userCadView.SelectionSet.Clear()
            userCadView.SelectionSet.SelectObjectsAtScreen(Function(obj) TypeOf obj Is Object, "Выборите внешнюю границу: ")
            For Each acZoneEnt As Object In userCadView.SelectionSet
                If TypeOf acZoneEnt Is DwgPolyline Then
                    acPlineBoundares = acZoneEnt
                    Exit For
                End If
            Next
            Dim classTopo As FuncTopologySuite = New FuncTopologySuite
            Dim polygon1 = classTopo.CreatePolygon(acPlineBoundares)
            '// копируем объекты
            For Each userIProjectModel In childs
                Dim modelUri As URI = userIProjectModel.Uri
                If modelUri.Extension Like ".arrx" Then
                    Dim pos As Integer = modelUri.AsAbsoluteUri.IndexOf("LandInventory")
                    If pos > -1 Then
                        Dim cadastralIDrawing As IDrawingContainer = userIProjectModel.LockRead()
                        Dim cadastralDrawing As Dwg.Drawing = cadastralIDrawing.Drawing
                        For Each acEnt As DwgEntity In cadastralDrawing.ActiveSpace.Entities
                            If TypeOf (acEnt) Is DwgPolyline Then
                                Dim poly As DwgPolyline = acEnt
                                If poly.Closed = True Then
                                    Dim polygon2 = classTopo.CreatePolygon(poly)
                                    Dim boolI As Boolean = classTopo.DoPolygonsIntersect(polygon1, polygon2)
                                    If boolI = True Then
                                        Dim copyObject As DwgEntity = acEnt.Clone()
                                        If ActivDocument.ActiveSpace.Entities.Contains(copyObject) = False Then
                                            ActivDocument.ActiveSpace.Entities.Add(copyObject)
                                        End If
                                    End If
                                Else
                                    Dim polygon2 = classTopo.CreateLineString(poly)
                                    Dim boolI As Boolean = classTopo.DoesLineIntersectPolygon(polygon2, polygon1)
                                    If boolI = True Then
                                        Dim copyObject As DwgEntity = acEnt.Clone()
                                        If ActivDocument.ActiveSpace.Entities.Contains(copyObject) = False Then
                                            ActivDocument.ActiveSpace.Entities.Add(copyObject)
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If
            Next
        End Sub
        '==============================================================================================================================
        'Импорт КПТ
        <cmd("RosreestrImportKPT2")>
        Public Sub RosreestrImportKPT2()
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
            Dim FormKPTImport As FormRosreestrKPTImport = New FormRosreestrKPTImport
            'запускаем форму
            FormKPTImport.ShowDialog()
            If FormKPTImport.boolShow = False Then
                Exit Sub
            End If
            '1.\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Dim namePrograms As String = FuncFiles.namePrograms
            Dim generalDir As String = ""
            Dim templateXML As String = ""
            Dim templateDWG As String = ""
            Dim boolFindFileTemplate As Boolean = FuncFiles.FuncReadProperties(namePrograms, generalDir, templateXML, templateDWG)
            '=========================================================================================================================
            Dim activUserModel As String = FormKPTImport.ComboBox1.Text
            Dim userProject As ModelProject = ApplicationHost.Current.ActiveProject
            '// Получение всех моделей проекта
            Dim childs As IProjectModel() = userProject.Model.GetChilds()
            For Each child As IProjectModel In childs
                Dim modelUri As URI = child.Uri
                If modelUri.Extension Like ".arrx" Then
                    Dim fileNameSfcx As String = Path.GetFileNameWithoutExtension(modelUri.LastPathComponent)
                    If fileNameSfcx.Trim Like activUserModel.Trim Then
                        ApplicationHost.Current.Plugins.Execute("activate", New Object() {child})
                        Exit For
                    End If
                End If
            Next
            Dim DrawShowLayers As IEnumerable(Of DrawingLayer) = DrawingLayer.GetDrawingLayers(userCadView)
            If IsNothing(activUserModel) = False Then
                If activUserModel.Trim.Length > 0 Then
                    For k1 As Integer = 0 To DrawShowLayers.Count - 1
                        Dim TempDrawLayer As DrawingLayer = DrawShowLayers.ElementAt(k1)
                        Dim nameUserFile As String = TempDrawLayer.Drawing.Filename
                        nameUserFile = IO.Path.GetFileNameWithoutExtension(nameUserFile)
                        If nameUserFile Like activUserModel Then
                            DrawLayer = TempDrawLayer
                            ActivDocument = DrawLayer.Drawing
                            Exit For
                        End If
                    Next k1
                End If
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '============================================================================================
            'начальные данные для считывания файла xml
            '============================================================================================
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Dim ObjectCollections As String(,) = Nothing
            Dim CountObjectCollection As Integer = 0
            Dim boolKvartal As Boolean = False 'обрабатывается квартал
            Dim boolspatial_data As Boolean = False
            Dim boolBuild As Boolean = False 'обрабатывается здание
            Dim boolParcel As Boolean = False 'обрабатывается участок
            Dim boolPart As Boolean = False 'обрабатываеся часть участка, здания, сооружения
            Dim boolCounter As Boolean = False 'обрабатывается контур участка
            Dim boolconstruction As Boolean = False 'обрабатывается сооружение
            Dim boolunderconstruction As Boolean = False 'обрабатывается объект незавершенного строительства
            Dim boolboundary As Boolean = False 'обрабатывается граница
            Dim boolcityboundary As Boolean = False 'обрабатывается граница города
            Dim boolzone As Boolean = False 'обрабатывается зона
            Dim boolcoastline As Boolean = False 'береговая полоса
            Dim boolNastr As Boolean = False 'флаг Говорит о том что нужно ли проверять настройки (слой, цвет и т.п.)
            '============================================================================================
            Dim boolNumberSubParcel As Boolean = False
            Dim boolSpatialElements As Boolean = False
            Dim boolCode As Boolean = False
            Dim boolType As Boolean = False
            Dim boolval As Boolean = False
            Dim boolcomdata As Boolean = False
            Dim boolunit As Boolean = False
            'начальное значение номера точки
            Dim Numb As Integer = 1 'номер
            Dim NXYZ As String(,) = Nothing 'массив для хранения координат объектов
            Dim boolX As Boolean = False
            Dim X As String = ""
            Dim boolY As Boolean = False
            Dim Y As String = ""
            Dim boolOrgNumb As Boolean = False
            Dim OrgNumb As String = "" 'номер
            Dim boolnum_geopoint As Boolean = False
            Dim num_geopoint As String = ""
            Dim boolGeopoint_opred As Boolean = False
            Dim Geopoint_opred As String = ""
            Dim boolDelta As Boolean = False
            Dim Delta As String = "0" 'погрешность
            Dim boolgeopoint_zacrep As Boolean = False
            Dim geopoint_zacrep As String = "-" 'закрепление
            Dim boolR As Boolean = False
            Dim r As String = "" 'радиус окружности
            Dim boolGeopointValue As Boolean = False
            Dim GeopointValue As String = ""
            Dim ArrayCounterParcel As String(,) = Nothing
            Dim CountArrayCounterParcel As Integer = 0
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'тип объекта
            '0-кадастровый квартал
            Dim boolarea_quarter As Boolean = False
            Dim Area_quarter As String = ""
            Dim CadastralBlock As String = "-"
            Dim boolKvartalNumber As Boolean = False
            '1. Общие данные для участков, ОКС и сооружений
            Dim boolCadNumber As Boolean = False 'кадастровый номер объекта
            Dim Cad_Number As String = "" 'полный кадастровый номер объекта
            Dim Cad_Number_Note As String = "" 'для единого землепользования
            Dim Cad_Number_Short As String = "" 'краткий кадастровый номер
            Dim CodTypeObject As String = "" 'код объекта (Земельный участок, сооружение, и т.п.)
            Dim ValTypeObject As String = "" 'значение объекта
            'общие для участков и зданий (участки-разрешенное использование, знания - тип здания (жилое, нежилое)
            Dim boolpermitted_use As Boolean = False 'разрешенное использование
            '1. Земельные участки
            Dim StateSubParcel As String = "" 'тип подучастка 01-ранее учтенный 06-учтенный ....
            Dim NumParcel As String = "" 'краткий кадастровый номер
            Dim boolNumberContour As Boolean = False
            Dim NumberContour As String = "" 'номер контура
            Dim NumberSubParcel As String = "" 'номер части
            Dim boolcommon_land_cad_number As Boolean = False
            Dim common_land_cad_number As String = ""
            'тип участка
            Dim boolSubTypeParcel As Boolean = False
            Dim SubTypeParcel As String = "" 'тип подучастка землепользование, многоконтурный участок и т.п
            'вид участка
            Dim StateParcel As String = "06" '06-учтенный
            Dim boolStateParcel As Boolean = False
            'категория земель
            Dim boolCategory As Boolean = False
            Dim Category As String = ""
            Dim CodCategory As String = ""
            'разрешенное использование 
            Dim boolRazrIsp As Boolean = False
            Dim byRazrIspByDoc As Boolean = False
            Dim RazrIsp As String = ""
            Dim RazrIspByDoc As String = ""
            'адрес
            '===============================================================
            Dim boolAdresParcel As Boolean = False
            Dim AddresParcel As String = ""
            Dim boolokato As Boolean = False
            Dim okato As String = ""
            Dim boolkladr As Boolean = False
            Dim kladr As String = ""
            Dim boolpostcode As Boolean = False
            Dim postcode As String = ""
            Dim boolregion As Boolean = False
            Dim region As String = ""
            Dim boolDistrict As Boolean = False
            Dim booltypeDistrict As Boolean = False
            Dim typeDistrict As String = ""
            Dim boolnameDistrict As Boolean = False
            Dim nameDistrict As String = ""
            Dim boolLocality As Boolean = False
            Dim booltypeLocality As Boolean = False
            Dim typeLocality As String = ""
            Dim boolnameLocality As Boolean = False
            Dim nameLocalyty As String = ""
            Dim boolStreet As Boolean = False
            Dim booltypeStreet As Boolean = False
            Dim type_street As String = ""
            Dim boolnameStreet As Boolean = False
            Dim name_street As String = ""
            Dim boolLevel1 As Boolean = False
            Dim booltypeLevel1 As Boolean = False
            Dim typeLevel1 As String = ""
            Dim boolnameLevel1 As Boolean = False
            Dim nameLevel1 As String = ""
            Dim boolLevel2 As Boolean = False
            Dim booltypeLevel2 As Boolean = False
            Dim typeLevel2 As String = ""
            Dim boolnameLevel2 As Boolean = False
            Dim nameLevel2 As String = ""
            Dim boolLevel3 As Boolean = False
            Dim booltypeLevel3 As Boolean = False
            Dim typeLevel3 As String = ""
            Dim boolnameLevel3 As Boolean = False
            Dim nameLevel3 As String = ""
            Dim readable_address As String = ""
            Dim boolreadable_address As Boolean = False
            '====================================================================
            'здания
            Dim boolpurpose As Boolean = False 'тип здания жилое нежилое
            Dim purpose_build As String = ""
            Dim boolposition_description As Boolean = False 'описание местоположения
            Dim pos_description_build As String = ""
            Dim codeBuild As String = "" 'назначение код здания
            Dim permitted_use As String = ""
            Dim boolname As Boolean = False
            Dim Assignatuin As String = ""
            Dim MaterialWall As String = ""
            Dim CountFloat As String = ""
            '============================================================================
            'сооружения
            Dim purpose_construction As String = ""
            Dim boolparams As Boolean = False 'параметры сооружения
            Dim boolbase_parameter As Boolean = False
            Dim booldepth As Boolean = False 'глубина залегания
            Dim baseParameters As String = ""
            Dim depth As String = ""
            Dim boolextension As Boolean = False 'длина
            Dim extension As String = ""
            Dim boolvolume As Boolean = False
            Dim volume As String = ""
            Dim boolbuilt_up_area As Boolean = False
            Dim boolareaConstr As Boolean = False
            '============================================================================
            'границы и зоны
            Dim booltype_boundary As Boolean = False
            Dim type_boundary As String = ""
            Dim boolregistration_date As Boolean = False
            Dim registr_data As String = ""
            Dim boolRegNumbZone As Boolean = False 'номер территориальной зоны
            Dim RegNumbZone As String = "" 'номер зоны
            Dim boolType_zone As Boolean = False
            Dim cod_boundary As String = ""
            Dim type_zona As String = ""
            Dim code_zona As String = ""
            Dim boolNumberZona As Boolean = False
            Dim boolIndexZona As Boolean = False
            Dim NumberZona As String = ""
            Dim IndexZona As String = ""
            '====================================================================================
            'береговые линии
            Dim boolwater_object_type As Boolean = False
            Dim boolwater_object_name As Boolean = False
            Dim water_object_name As String = ""
            Dim codwater_object_type As String = ""
            Dim typewater_object_type As String = ""
            '====================================================================================
            'собственник
            Dim boolRightType As Boolean = False
            Dim RightType As String = ""
            Dim boolOwner As Boolean = False
            Dim boolOwnerName As Boolean = False
            Dim ownerName As String = ""
            Dim boolresident As Boolean = False
            Dim encumbranceName As String = ""
            Dim boolrecord_info As Boolean = False
            Dim boolRightName As Boolean = False

            '===============================================================================
            'реквизиты выписки
            Dim booldate_formation As Boolean = False
            Dim dataVP As String = "-"
            Dim boolregistration_number As Boolean = False
            Dim NumberVP As String = ""
            Dim boolgroup_top_requisites As Boolean = False
            '========================================================
            'площадь участка общая
            Dim boolGlobalArea As Boolean = False
            Dim GlobalArea As String = ""
            Dim deltaArea As String = ""
            Dim booldeltaArea As Boolean = False
            'кадастровая стоимость
            Dim boolCostParcel As Boolean = False
            Dim CostParcel As String = ""
            'Счетчик вершин контура (для массива координат)
            Dim countcoord As Integer = -1
            Dim TypeLineSubParcel As String = ""
            Dim LocalID As Integer = 0
            Dim GlobalID As Integer = 0
            '=======================================================================================================
            'переменные для оформления контуров
            Dim LayerObject As String = "0"
            Dim LineTypeObj As String = "Continuous"
            Dim IndexColorObj As Integer = 256
            Dim WidthrObj As Integer = 20
            Dim ScLine As Double = 1
            Dim GlobalContLine As Double = 0
            Dim NamePSSetTable As String = "OBJECT"
            Dim IDPSTAble As ObjectId = Nothing
            Dim OldCodObject As String = ""
            '========================================================================================================
            'массив с кадастровыми номерами всех кадастровых объектов
            Dim ArrayOldCadastrNumbers As String() = Nothing
            Dim CountArrayOldCadastrNumbers As Integer = 0
            Dim ArrayOldCadastrRegNumbers As String() = Nothing
            Dim CountArrayOldRegNumbers As Integer = 0
            Dim ArrayOldCadastrKVNumbers As String() = Nothing
            Dim CountArrayOldKVNumbers As Integer = 0
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'тело программы
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'выбор файла
            Dim LineErrNumber As Integer = 0
            Dim filterFile As String = "xml files (*.xml)|*.xml"
            Dim FileRez As String = FuncFiles.FuncGetAbsoluteFilePathName("Открыть файл xml", filterFile)
            If FileRez.Trim.Length = 0 Then Exit Sub
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'читаем файл xml
            If File.Exists(FileRez) = True Then
                Dim reader As XmlTextReader = New XmlTextReader(FileRez)
                Dim countFile As Integer = 0
                While reader.Read()
                    Dim NameBlock As String = reader.Name 'читаем ветку
                    LineErrNumber = reader.LineNumber
                    Select Case reader.NodeType
                        Case XmlNodeType.Element
                            'блок общих данных
                            '=====================================================================
                            If NameBlock Like "common_data" Then
                                boolcomdata = True
                            ElseIf NameBlock Like "type" Then
                                boolType = True
                            ElseIf NameBlock Like "code" Then
                                boolCode = True
                            ElseIf NameBlock Like "value" Then
                                boolval = True
                            ElseIf NameBlock Like "unit" Then
                                boolunit = True
                            ElseIf NameBlock Like "name" Then
                                boolname = True
                                '============================================================================
                                'кадастровый квартал
                            ElseIf NameBlock Like "area_quarter" Then
                                boolarea_quarter = True 'площадь квартала
                            ElseIf NameBlock Like "cadastral_number" Then
                                boolKvartalNumber = True 'кадастровый номер квартала (КПТ)
                            ElseIf NameBlock Like "quarter_cad_number" Then
                                boolKvartalNumber = True 'кадастровый номер квартала (выписка)
                                '============================================================================
                                'часть участка и контур участка
                            ElseIf NameBlock Like "object_part" Then
                                boolPart = True 'часть кадастрового участка
                                NumberSubParcel = ""
                            ElseIf NameBlock Like "object_parts" Then
                                boolNastr = False 'сбрасываем настройки
                            ElseIf NameBlock Like "contours" Then
                                'CountCounter = 0
                            ElseIf NameBlock Like "contour" Then
                                'контур объекта
                                boolCounter = True
                                NumberContour = ""
                            ElseIf NameBlock Like "State" Then
                                boolStateParcel = True
                                '=========================================================
                                'адрес
                            ElseIf NameBlock Like "address" Then
                                boolAdresParcel = True
                            ElseIf NameBlock Like "okato" Then
                                boolokato = True
                            ElseIf NameBlock Like "kladr" Then
                                boolkladr = True
                            ElseIf NameBlock Like "postal_code" Then
                                boolpostcode = True
                            ElseIf NameBlock Like "region" Then
                                boolregion = True
                            ElseIf NameBlock Like "type_district" Then
                                booltypeDistrict = True
                            ElseIf NameBlock Like "name_district" Then
                                boolnameDistrict = True
                            ElseIf NameBlock Like "type_locality" Then
                                booltypeLocality = True
                            ElseIf NameBlock Like "name_locality" Then
                                boolnameLocality = True
                            ElseIf NameBlock Like "type_street" Then
                                booltypeStreet = True
                            ElseIf NameBlock Like "name_street" Then
                                boolnameStreet = True
                            ElseIf NameBlock Like "type_level1" Then
                                booltypeLevel1 = True
                            ElseIf NameBlock Like "name_level1" Then
                                boolnameLevel1 = True
                            ElseIf NameBlock Like "type_level2" Then
                                booltypeLevel2 = True
                            ElseIf NameBlock Like "name_level2" Then
                                boolnameLevel2 = True
                            ElseIf NameBlock Like "type_level3" Then
                                booltypeLevel3 = True
                            ElseIf NameBlock Like "name_level3" Then
                                boolnameLevel3 = True
                            ElseIf NameBlock Like "readable_address" Then
                                boolreadable_address = True 'неструктурированный адрес
                            ElseIf NameBlock Like "position_description" Then
                                boolposition_description = True 'описание местоположения
                                '===============================================================================
                                'кадастровые номера и обозначения зон
                            ElseIf NameBlock Like "cad_number" Then
                                boolCadNumber = True 'кадастровый номер участка здания стоения ОНС
                            ElseIf NameBlock Like "part_number" Then
                                boolNumberSubParcel = True 'номер части участка
                            ElseIf NameBlock Like "number_pp" Then
                                boolNumberContour = True 'номер контура
                                LocalID = rng.Next
                                If GlobalID < 1 Then
                                    GlobalID = rng.Next
                                End If
                            ElseIf NameBlock Like "reg_numb_border" Then
                                boolRegNumbZone = True 'обозначение границы и ЗОУИТ
                            ElseIf NameBlock Like "common_land_cad_number" Then
                                boolcommon_land_cad_number = True
                                '=================================================================================
                                'площади
                            ElseIf NameBlock Like "area" And boolbase_parameter = False Then
                                boolGlobalArea = True 'площадь объекта
                            ElseIf NameBlock Like "inaccuracy" Then
                                booldeltaArea = True 'погрешность определения площади
                                '====================================================================================
                                'кадастровая стоимость
                            ElseIf NameBlock Like "cost" Then
                                boolCostParcel = True
                                '====================================================================================
                                'тип участка
                            ElseIf NameBlock Like "subtype" Then
                                boolSubTypeParcel = True 'землепользование, обособленный участок и т.п.
                            ElseIf NameBlock Like "category" Then
                                boolCategory = True 'категория земель
                            ElseIf NameBlock Like "permitted_use" Then
                                boolpermitted_use = True
                            ElseIf NameBlock Like "by_document" Then
                                byRazrIspByDoc = True 'разрешенное использование по документу
                                '=======================================
                                'здания, сооружения (характеристики)
                            ElseIf NameBlock Like "purpose" Then
                                boolpurpose = True 'жилое нежилое многоквартирный дом
                            ElseIf NameBlock Like "permitted_use" Then
                                boolpermitted_use = True 'разрешенное использование, для здания - жилое, нежилое
                                '==============================================
                                'характеристики строения
                            ElseIf NameBlock Like "params" Then
                                boolparams = True 'блок параметров
                            ElseIf NameBlock Like "base_parameter" Then
                                boolbase_parameter = True 'блок с базовыми параметрами
                            ElseIf NameBlock Like "occurence_depth" Then
                                booldepth = True 'глубина
                            ElseIf NameBlock Like "depth" Then
                                booldepth = True 'тоже глубина
                            ElseIf NameBlock Like "extension" Then
                                boolextension = True 'протяженность
                            ElseIf NameBlock Like "volume" Then
                                boolvolume = True 'объем
                            ElseIf NameBlock Like "area" And boolbase_parameter = True Then
                                boolareaConstr = True 'площадь
                            ElseIf NameBlock Like "built_up_area" And boolbase_parameter = True Then
                                boolbuilt_up_area = True 'площадь
                                '=======================================================
                                'территориальные зоны границы
                            ElseIf NameBlock Like "record_info" Then
                                boolrecord_info = True
                            ElseIf NameBlock Like "type_boundary" Then
                                booltype_boundary = True 'тип границы
                            ElseIf NameBlock Like "type_zone" Then
                                boolType_zone = True 'описание зоны
                            ElseIf NameBlock Like "registration_date" Then
                                boolregistration_date = True 'дата регистрации зоны
                            ElseIf NameBlock Like "reg_numb_border" Then
                                boolRegNumbZone = True 'номер зоны
                            ElseIf NameBlock Like "number" And boolzone = True Then
                                boolNumberZona = True 'номер зоны
                            ElseIf NameBlock Like "index" And boolzone = True Then
                                boolIndexZona = True 'номер зоны
                                '=======================================================
                                'береговые полосы
                            ElseIf NameBlock Like "water_object_type" Then
                                boolwater_object_type = True
                            ElseIf NameBlock Like "water_object_name" Then
                                boolwater_object_name = True
                                '=====================================================================
                                'сведения о правах
                            ElseIf NameBlock Like "right_type" Then
                                boolRightType = True 'тип собственности
                            ElseIf NameBlock Like "right_holder" Then
                                boolOwner = True 'тип собственности
                            ElseIf NameBlock Like "name" And boolOwner = True Then
                                boolOwnerName = True 'тип собственности
                            ElseIf NameBlock Like "resident" Then
                                boolresident = True 'арендатор
                                '=====================================================================
                                'реквизиты выписки
                            ElseIf NameBlock Like "group_top_requisites" Then
                                boolgroup_top_requisites = True 'блок с реквизитами
                            ElseIf NameBlock Like "date_formation" Then
                                booldate_formation = True
                            ElseIf NameBlock Like "registration_number" Then
                                boolregistration_number = True
                                '=============================================================
                                'координаты
                            ElseIf NameBlock Like "ordinate" Then
                                countcoord += 1
                                ReDim Preserve NXYZ(8, countcoord) 'тип,номер, x,y,delta,desk,r
                                NXYZ(0, countcoord) = "Точка"
                            ElseIf NameBlock Like "x" Then
                                boolX = True
                            ElseIf NameBlock Like "y" Then
                                boolY = True
                            ElseIf NameBlock Like "geopoint_opred" Then
                                boolGeopoint_opred = True
                            ElseIf NameBlock Like "ord_nmb" Then
                                boolOrgNumb = True
                            ElseIf NameBlock Like "num_geopoint" Then
                                boolnum_geopoint = True
                            ElseIf NameBlock Like "delta_geopoint" Then
                                boolDelta = True
                            ElseIf NameBlock Like "r" Then
                                boolR = True
                                '==========================================================================
                                'блоки с объектами
                            ElseIf NameBlock Like "land_record" Then 'земельный участок
                                AddresParcel = ""
                                Cad_Number = ""
                                StateParcel = ""
                                Category = ""
                                CodCategory = ""
                                RazrIsp = ""
                                RazrIspByDoc = ""
                                CostParcel = ""
                                GlobalArea = ""
                                Cad_Number_Note = ""
                                deltaArea = ""
                                RightType = ""
                                encumbranceName = ""
                                ownerName = ""
                                okato = ""
                                Cad_Number_Short = ""
                                boolParcel = True
                                LocalID = rng.Next
                                GlobalID = 0
                                Erase ObjectCollections
                                CountObjectCollection = 0
                            ElseIf NameBlock Like "build_record" Then 'здание
                                Cad_Number = ""
                                AddresParcel = ""
                                okato = ""
                                RightType = ""
                                encumbranceName = ""
                                ownerName = ""
                                CostParcel = ""
                                GlobalArea = ""
                                Cad_Number_Note = ""
                                purpose_build = ""
                                pos_description_build = ""
                                Assignatuin = ""
                                MaterialWall = ""
                                CountFloat = ""
                                StateParcel = ""
                                boolBuild = True
                                purpose_build = ""
                                Cad_Number_Short = ""
                                LocalID = rng.Next
                                GlobalID = 0
                                Erase ObjectCollections
                                CountObjectCollection = 0
                            ElseIf NameBlock Like "construction_record" Then 'конструкция
                                Cad_Number = ""
                                AddresParcel = ""
                                CostParcel = ""
                                GlobalArea = ""
                                Cad_Number_Note = ""
                                baseParameters = ""
                                purpose_construction = ""
                                Cad_Number_Short = ""
                                boolconstruction = True
                                LocalID = rng.Next
                                GlobalID = 0
                                Erase ObjectCollections
                                CountObjectCollection = 0
                            ElseIf NameBlock Like "object_under_construction_record" Then 'объект незавершенного строительсва
                                Cad_Number = ""
                                AddresParcel = ""
                                okato = ""
                                RightType = ""
                                encumbranceName = ""
                                ownerName = ""
                                CostParcel = ""
                                GlobalArea = ""
                                Cad_Number_Note = ""
                                purpose_build = ""
                                pos_description_build = ""
                                Assignatuin = ""
                                MaterialWall = ""
                                CountFloat = ""
                                StateParcel = ""
                                Cad_Number_Short = ""
                                boolunderconstruction = True
                                LocalID = rng.Next
                                GlobalID = 0
                                Erase ObjectCollections
                                CountObjectCollection = 0
                            ElseIf NameBlock Like "inhabited_locality_boundary_record" Then 'граница населенного пункта
                                RegNumbZone = ""
                                type_boundary = ""
                                cod_boundary = ""
                                registr_data = ""
                                boolcityboundary = True
                                LocalID = rng.Next
                                GlobalID = 0
                                Erase ObjectCollections
                                CountObjectCollection = 0
                            ElseIf NameBlock Like "municipal_boundary_record" Then 'граница административная
                                RegNumbZone = ""
                                type_boundary = ""
                                cod_boundary = ""
                                registr_data = ""
                                boolboundary = True
                                LocalID = rng.Next
                                GlobalID = 0
                                Erase ObjectCollections
                                CountObjectCollection = 0
                            ElseIf NameBlock Like "spatial_data" Then 'квартал
                                boolspatial_data = True
                                boolKvartal = True
                                LocalID = rng.Next
                                GlobalID = 0
                                Erase ObjectCollections
                                CountObjectCollection = 0
                            ElseIf NameBlock Like "zones_and_territories_record" Then 'охранная зона
                                type_zona = ""
                                RegNumbZone = ""
                                code_zona = ""
                                registr_data = ""
                                boolzone = True
                                LocalID = rng.Next
                                GlobalID = 0
                                Erase ObjectCollections
                                CountObjectCollection = 0
                            ElseIf NameBlock Like "coastline_record" Then 'береговая линия
                                type_zona = ""
                                RegNumbZone = ""
                                code_zona = ""
                                registr_data = ""
                                ValTypeObject = ""
                                water_object_name = ""
                                codwater_object_type = ""
                                typewater_object_type = ""
                                type_boundary = ""
                                boolcoastline = True
                                LocalID = rng.Next
                                GlobalID = 0
                                Erase ObjectCollections
                                CountObjectCollection = 0
                            ElseIf NameBlock Like "land_records" Then
                                LayerObject = "0"
                                LineTypeObj = "Continuous"
                                IndexColorObj = 256
                                WidthrObj = 20
                                ScLine = 1
                                GlobalContLine = 0
                                NamePSSetTable = "OBJECT"
                                ValTypeObject = ""
                                boolNastr = False
                            ElseIf NameBlock Like "build_records" Then
                                LayerObject = "0"
                                LineTypeObj = "Continuous"
                                IndexColorObj = 256
                                WidthrObj = 20
                                ScLine = 1
                                GlobalContLine = 0
                                NamePSSetTable = "OBJECT"
                                ValTypeObject = ""
                                boolNastr = False
                            ElseIf NameBlock Like "construction_records" Then
                                LayerObject = "0"
                                LineTypeObj = "Continuous"
                                IndexColorObj = 256
                                WidthrObj = 20
                                ScLine = 1
                                GlobalContLine = 0
                                NamePSSetTable = "OBJECT"
                                ValTypeObject = ""
                                boolNastr = False
                            ElseIf NameBlock Like "object_under_construction_records" Then
                                LayerObject = "0"
                                LineTypeObj = "Continuous"
                                IndexColorObj = 256
                                WidthrObj = 20
                                ScLine = 1
                                GlobalContLine = 0
                                NamePSSetTable = "OBJECT"
                                ValTypeObject = ""
                                boolNastr = False
                            ElseIf NameBlock Like "municipal_boundaries" Then
                                LayerObject = "0"
                                LineTypeObj = "Continuous"
                                IndexColorObj = 256
                                WidthrObj = 20
                                ScLine = 1
                                GlobalContLine = 0
                                NamePSSetTable = "OBJECT"
                                ValTypeObject = ""
                                boolNastr = False
                            ElseIf NameBlock Like "inhabited_locality_boundaries" Then
                                LayerObject = "0"
                                LineTypeObj = "Continuous"
                                IndexColorObj = 256
                                WidthrObj = 20
                                ScLine = 1
                                GlobalContLine = 0
                                NamePSSetTable = "OBJECT"
                                ValTypeObject = ""
                                cod_boundary = ""
                                boolNastr = False
                            ElseIf NameBlock Like "zones_and_territories_boundaries" Then
                                LayerObject = "0"
                                LineTypeObj = "Continuous"
                                IndexColorObj = 256
                                WidthrObj = 20
                                ScLine = 1
                                GlobalContLine = 0
                                NamePSSetTable = "OBJECT"
                                ValTypeObject = ""
                                code_zona = ""
                                cod_boundary = ""
                                IndexZona = ""
                                NumberZona = ""
                                boolNastr = False
                            ElseIf NameBlock Like "coastline_boundaries" Then
                                LayerObject = "0"
                                LineTypeObj = "Continuous"
                                IndexColorObj = 256
                                WidthrObj = 20
                                ScLine = 1
                                GlobalContLine = 0
                                NamePSSetTable = "OBJECT"
                                boolNastr = False
                            End If
                           '==========================================================================================
                           'значения
                           '==========================================================================================
                        Case XmlNodeType.Text
                            'кадастровый квартал
                            If boolKvartalNumber = True Then
                                CadastralBlock = reader.Value 'номер кадастрового квартала
                                boolKvartalNumber = False
                            ElseIf boolarea_quarter = True And boolGlobalArea = True Then
                                'площадь квартала
                                Area_quarter = reader.Value
                            ElseIf boolarea_quarter = True And boolunit = True Then
                                'еденицы идмерения площади квартала
                                Dim unit As Integer = CInt(reader.Value)
                                Area_quarter = Area_quarter & " " & CadastralFunction.FuncFindCadastrUnitByCode(unit)
                            ElseIf boolcomdata = True And boolType = True Then
                                CodTypeObject = reader.Value
                                boolType = False
                            ElseIf boolcomdata = True And boolval = True Then
                                ValTypeObject = reader.Value
                                boolval = False
                            ElseIf boolCadNumber = True And boolcomdata = True Then
                                'кадастровый номер участка, сооружения, здания
                                Cad_Number = reader.Value
                                boolCadNumber = False
                                'отделяем краткий номер
                                Dim posN As Integer = Cad_Number.LastIndexOf(":")
                                If posN > 6 Then
                                    Cad_Number_Short = Mid(Cad_Number, posN + 1)
                                End If
                            ElseIf boolcommon_land_cad_number = True And boolCadNumber = True Then
                                'кадастровый номер участка, сооружения, здания
                                Cad_Number_Note = reader.Value
                                boolCadNumber = False
                            ElseIf boolNumberSubParcel = True Then
                                'номер части учаcтка
                                NumberSubParcel = reader.Value
                                boolNumberSubParcel = False
                            ElseIf boolNumberContour = True Then
                                'номер контура
                                NumberContour = reader.Value
                                boolNumberContour = False
                            ElseIf boolRegNumbZone = True Then
                                'номер территориальной зоны
                                RegNumbZone = reader.Value
                                boolRegNumbZone = False
                            ElseIf boolSubTypeParcel = True And boolval = True Then
                                'тип участка 01-землепользование, многоконтурный и т.п.
                                SubTypeParcel = reader.Value
                                If SubTypeParcel Like "Многоконтурный участок" Or SubTypeParcel Like "Единое землепользование" Then
                                    GlobalID = rng.Next
                                End If
                            ElseIf boolStateParcel = True And boolCode = True Then
                                StateParcel = reader.Value
                            ElseIf boolCategory = True And boolCode = True Then
                                'категория земель
                                CodCategory = reader.Value
                            ElseIf boolCategory = True And boolval = True Then
                                Category = reader.Value
                            ElseIf boolpermitted_use = True And byRazrIspByDoc = True Then
                                RazrIspByDoc = reader.Value
                                byRazrIspByDoc = False
                            ElseIf boolCostParcel = True And boolval = True Then
                                'кадастровая стоимость
                                CostParcel = reader.Value
                                boolCostParcel = False
                            ElseIf boolGlobalArea = True And boolval = True Then
                                'площадь участка
                                GlobalArea = reader.Value
                                boolval = False
                            ElseIf boolparams = True And boolGlobalArea = True And boolParcel = False Then
                                'площадь здания
                                GlobalArea = reader.Value
                                boolGlobalArea = False
                            ElseIf boolGlobalArea = True And booldeltaArea = True Then
                                deltaArea = reader.Value
                                booldeltaArea = False
                                '========================================================================
                                'адрес
                            ElseIf boolAdresParcel = True And boolokato = True Then
                                okato = reader.Value
                                boolokato = False
                            ElseIf boolAdresParcel = True And boolkladr = True Then
                                kladr = reader.Value
                                boolkladr = False
                            ElseIf boolAdresParcel = True And boolpostcode = True Then
                                postcode = reader.Value
                                boolpostcode = False
                            ElseIf boolAdresParcel = True And boolregion = True And boolval = True Then
                                region = reader.Value
                                boolregion = False
                                boolval = False
                            ElseIf boolAdresParcel = True And booltypeDistrict = True Then
                                typeDistrict = reader.Value
                                booltypeDistrict = False
                            ElseIf boolAdresParcel = True And boolnameDistrict = True Then
                                nameDistrict = reader.Value
                                boolnameDistrict = False
                            ElseIf boolAdresParcel = True And booltypeLocality = True Then
                                typeLocality = reader.Value
                                booltypeLocality = False
                            ElseIf boolAdresParcel = True And boolnameLocality = True Then
                                nameLocalyty = reader.Value
                                boolnameLocality = False
                            ElseIf boolAdresParcel = True And booltypeStreet = True Then
                                type_street = reader.Value
                                booltypeStreet = False
                            ElseIf boolAdresParcel = True And boolnameStreet = True Then
                                name_street = reader.Value
                                boolnameStreet = False
                            ElseIf boolAdresParcel = True And booltypeLevel1 = True Then
                                typeLevel1 = reader.Value
                                booltypeLevel1 = False
                            ElseIf boolAdresParcel = True And boolnameLevel1 = True Then
                                nameLevel1 = reader.Value
                                boolnameLevel1 = False
                            ElseIf boolAdresParcel = True And booltypeLevel2 = True Then
                                typeLevel2 = reader.Value
                                booltypeLevel2 = False
                            ElseIf boolAdresParcel = True And boolnameLevel2 = True Then
                                nameLevel2 = reader.Value
                                boolnameLevel2 = False
                            ElseIf boolAdresParcel = True And booltypeLevel3 = True Then
                                typeLevel3 = reader.Value
                                booltypeLevel3 = False
                            ElseIf boolAdresParcel = True And boolnameLevel3 = True Then
                                nameLevel3 = reader.Value
                                boolnameLevel3 = False
                            ElseIf boolreadable_address = True Then
                                readable_address = reader.Value
                                boolreadable_address = False
                                '============================================================================
                            ElseIf boolRightType = True And boolval = True Then
                                RightType = reader.Value
                                boolRightType = False
                                boolval = False
                            ElseIf boolOwner = True And boolOwnerName = True Then
                                ownerName = reader.Value
                                boolOwner = False
                                boolOwnerName = False
                            ElseIf boolresident = True And boolname = True Then
                                encumbranceName = reader.Value
                                boolresident = False
                                boolname = False
                                '============================================================
                                'здания
                            ElseIf boolpurpose = True And boolval = True And boolparams = True Then
                                purpose_build = reader.Value 'назначение
                                boolval = False
                            ElseIf boolpurpose = True And boolCode = True And boolparams = True Then
                                codeBuild = reader.Value 'код здания
                                boolCode = False
                            ElseIf boolposition_description = True Then
                                pos_description_build = reader.Value 'описание местоположения
                                boolposition_description = False
                            ElseIf boolpermitted_use = True And boolname = True Then
                                permitted_use = reader.Value
                                '===================================================================
                                'сооружения
                            ElseIf boolparams = True And boolpurpose = True Then
                                purpose_construction = reader.Value
                                boolpurpose = False
                            ElseIf boolbase_parameter = True And boolextension = True Then
                                baseParameters = baseParameters & "Длина " & reader.Value & " "
                                boolextension = False
                            ElseIf boolbase_parameter = True And booldepth = True Then
                                baseParameters = baseParameters & "Глубина " & reader.Value & " "
                                booldepth = False
                            ElseIf boolbase_parameter = True And booldepth = True Then
                                baseParameters = baseParameters & "Глубина " & reader.Value & " "
                                booldepth = False
                            ElseIf boolbase_parameter = True And boolvolume = True Then
                                baseParameters = baseParameters & "Объем " & reader.Value & " "
                                boolvolume = False
                            ElseIf boolbase_parameter = True And boolareaConstr = True Then
                                baseParameters = baseParameters & "Общая площадь " & reader.Value & " "
                                boolareaConstr = False
                            ElseIf boolbase_parameter = True And boolbuilt_up_area = True Then
                                baseParameters = baseParameters & "Общая площадь застройки " & reader.Value & " "
                                boolbuilt_up_area = False
                                '=======================================================================
                                'границы
                            ElseIf booltype_boundary = True And boolval = True Then
                                type_boundary = reader.Value
                                boolval = False
                            ElseIf booltype_boundary = True And boolCode = True Then
                                cod_boundary = reader.Value
                                boolCode = False
                                'береговые линии
                            ElseIf boolwater_object_type = True And boolCode = True Then
                                codwater_object_type = reader.Value
                                boolCode = False
                            ElseIf boolwater_object_type = True And boolType = True Then
                                typewater_object_type = reader.Value
                                boolType = False
                            ElseIf boolwater_object_name = True Then
                                water_object_name = reader.Value
                                boolwater_object_name = False
                                '==============================================================================
                                'территориальные зоны
                            ElseIf boolregistration_date = True And boolrecord_info = True Then
                                registr_data = reader.Value
                                boolregistration_date = False
                            ElseIf boolType_zone = True And boolval = True Then 'тип ЗОУИТ
                                type_zona = reader.Value
                                boolval = False
                            ElseIf boolType_zone = True And boolCode = True Then 'код ЗОУИТ
                                code_zona = reader.Value
                                boolCode = False
                            ElseIf boolRegNumbZone = True Then
                                RegNumbZone = reader.Value
                                boolRegNumbZone = False
                            ElseIf boolzone = True And boolIndexZona = True Then 'индекс зоны
                                IndexZona = reader.Value
                                boolIndexZona = False
                            ElseIf boolzone = True And boolNumberZona = True Then 'номер зоны
                                NumberZona = reader.Value
                                boolNumberZona = False
                                'координаты точек
                                '======================================================================================
                            ElseIf boolOrgNumb = True Then
                                If IsArray(NXYZ) = True Then
                                    NXYZ(1, countcoord) = reader.Value 'номер точки
                                    OrgNumb = reader.Value
                                End If
                                boolOrgNumb = False
                            ElseIf boolX = True Then
                                If IsArray(NXYZ) = True Then
                                    NXYZ(2, countcoord) = reader.Value 'X
                                    X = reader.Value
                                End If
                                boolX = False
                            ElseIf boolY = True Then
                                If IsArray(NXYZ) = True Then
                                    NXYZ(3, countcoord) = reader.Value 'Y
                                    Y = reader.Value
                                End If
                                boolY = False
                            ElseIf boolDelta = True Then
                                If IsArray(NXYZ) = True Then
                                    NXYZ(4, countcoord) = reader.Value 'Delta
                                    Delta = reader.Value
                                End If
                                boolDelta = False
                            ElseIf boolnum_geopoint = True Then
                                If IsArray(NXYZ) = True Then
                                    NXYZ(5, countcoord) = reader.Value 'старый номер
                                End If
                                boolnum_geopoint = False
                            ElseIf boolgeopoint_zacrep = True Then
                                If IsArray(NXYZ) = True Then
                                    NXYZ(6, countcoord) = reader.Value 'тип закрепления
                                End If
                                boolnum_geopoint = False
                            ElseIf boolGeopoint_opred = True And boolval = True Then
                                If IsArray(NXYZ) = True Then
                                    NXYZ(7, countcoord) = reader.Value
                                End If
                                boolGeopointValue = False
                            ElseIf boolR = True Then
                                If IsArray(NXYZ) = True Then
                                    NXYZ(8, countcoord) = reader.Value
                                    NXYZ(0, countcoord) = "Окружность"
                                End If
                                boolR = False
                                '===================================================================================
                                'реквизиты выписки
                            ElseIf booldate_formation = True Then
                                dataVP = reader.Value
                                booldate_formation = False
                            ElseIf boolregistration_number = True Then
                                NumberVP = reader.Value
                                boolregistration_number = False
                            End If
                            '====================================================================================================
                            'блоки закрытия
                        Case XmlNodeType.EndElement
                            '=======================================================================
                            'блок общих данных
                            If NameBlock Like "common_data" Then
                                boolcomdata = False
                            ElseIf NameBlock Like "code" Then
                                boolCode = False
                            ElseIf NameBlock Like "value" Then
                                boolval = False
                            ElseIf NameBlock Like "type" Then
                                boolType = False
                            ElseIf NameBlock Like "unit" Then
                                boolunit = False
                            ElseIf NameBlock Like "name" Then
                                boolname = False
                                '=========================================================
                            ElseIf NameBlock Like "area_quarter" Then
                                boolarea_quarter = False
                            ElseIf NameBlock Like "group_top_requisites" Then
                                boolgroup_top_requisites = False
                            ElseIf NameBlock Like "cad_number" Then
                                boolCadNumber = False
                            ElseIf NameBlock Like "common_land_cad_number" Then
                                boolcommon_land_cad_number = False
                            ElseIf NameBlock Like "subtype" Then
                                boolSubTypeParcel = False
                            ElseIf NameBlock Like "State" Then
                                boolStateParcel = False
                            ElseIf NameBlock Like "category" Then
                                boolCategory = False
                            ElseIf NameBlock Like "permitted_use" Then
                                boolRazrIsp = False
                            ElseIf NameBlock Like "cost" Then
                                boolCostParcel = False
                            ElseIf NameBlock Like "area" Then
                                boolGlobalArea = False
                                '=========================================================
                                'адрес
                            ElseIf NameBlock Like "address" Then
                                boolAdresParcel = False
                                If okato.Trim.Length > 0 Then
                                    AddresParcel = "ОКАТО " & okato & " "
                                    okato = ""
                                End If
                                If kladr.Trim.Length > 0 Then
                                    AddresParcel = AddresParcel & "КЛАДР " & kladr & " "
                                    kladr = ""
                                End If
                                If postcode.Trim.Length > 0 Then
                                    AddresParcel = AddresParcel & "Индекс " & postcode & " "
                                    postcode = ""
                                End If
                                If region.Trim.Length > 0 Then
                                    AddresParcel = AddresParcel & region & " "
                                    region = ""
                                End If
                                If typeDistrict.Trim.Length > 0 Then
                                    AddresParcel = AddresParcel & typeDistrict & " "
                                    typeDistrict = ""
                                End If
                                If nameDistrict.Trim.Length > 0 Then
                                    AddresParcel = AddresParcel & nameDistrict & " "
                                    nameDistrict = ""
                                End If
                                If typeLocality.Trim.Length > 0 Then
                                    AddresParcel = AddresParcel & typeLocality & " "
                                    typeLocality = ""
                                End If
                                If nameLocalyty.Trim.Length > 0 Then
                                    AddresParcel = AddresParcel & nameLocalyty & " "
                                    nameLocalyty = ""
                                End If
                                If type_street.Trim.Length > 0 Then
                                    AddresParcel = AddresParcel & type_street & " "
                                    type_street = ""
                                End If
                                If name_street.Trim.Length > 0 Then
                                    AddresParcel = AddresParcel & name_street & " "
                                    name_street = ""
                                End If
                                If typeLevel1.Trim.Length > 0 Then
                                    AddresParcel = AddresParcel & typeLevel1 & " "
                                    typeLevel1 = ""
                                End If
                                If nameLevel1.Trim.Length > 0 Then
                                    AddresParcel = AddresParcel & nameLevel1 & " "
                                    nameLevel1 = ""
                                End If
                                If typeLevel2.Trim.Length > 0 Then
                                    AddresParcel = AddresParcel & typeLevel2 & " "
                                    typeLevel2 = ""
                                End If
                                If nameLevel2.Trim.Length > 0 Then
                                    AddresParcel = AddresParcel & nameLevel2 & " "
                                    nameLevel2 = ""
                                End If
                                If typeLevel3.Trim.Length > 0 Then
                                    AddresParcel = AddresParcel & typeLevel3 & " "
                                    typeLevel3 = ""
                                End If
                                If nameLevel3.Trim.Length > 0 Then
                                    AddresParcel = AddresParcel & nameLevel3 & " "
                                    nameLevel3 = ""
                                End If
                            ElseIf NameBlock Like "okato" Then
                                boolokato = False
                            ElseIf NameBlock Like "kladr" Then
                                boolkladr = False
                            ElseIf NameBlock Like "region" Then
                                boolregion = False
                            ElseIf NameBlock Like "type_district" Then
                                booltypeDistrict = False
                            ElseIf NameBlock Like "name_district" Then
                                boolnameDistrict = False
                            ElseIf NameBlock Like "type_locality" Then
                                booltypeLocality = False
                            ElseIf NameBlock Like "name_locality" Then
                                boolnameLocality = False
                            ElseIf NameBlock Like "type_street" Then
                                booltypeStreet = False
                            ElseIf NameBlock Like "name_street" Then
                                boolnameStreet = False
                            ElseIf NameBlock Like "type_level1" Then
                                booltypeLevel1 = False
                            ElseIf NameBlock Like "name_level1" Then
                                boolnameLevel1 = False
                            ElseIf NameBlock Like "type_level2" Then
                                booltypeLevel2 = False
                            ElseIf NameBlock Like "name_level2" Then
                                boolnameLevel2 = False
                            ElseIf NameBlock Like "type_level3" Then
                                booltypeLevel3 = False
                            ElseIf NameBlock Like "name_level3" Then
                                boolnameLevel3 = False
                                '============================================================================
                            ElseIf NameBlock Like "purpose" Then
                                boolpurpose = False
                            ElseIf NameBlock Like "params" Then
                                boolparams = False
                            ElseIf NameBlock Like "permitted_use" Then
                                boolpermitted_use = False
                            ElseIf NameBlock Like "base_parameter" Then
                                boolbase_parameter = False
                            ElseIf NameBlock Like "type_boundary" Then
                                booltype_boundary = False
                            ElseIf NameBlock Like "type_zone" Then
                                boolType_zone = False
                            ElseIf NameBlock Like "record_info" Then
                                boolrecord_info = False
                            ElseIf NameBlock Like "registration_date" Then
                                boolregistration_date = False 'дата регистрации зоны
                            ElseIf NameBlock Like "object_part" Then
                                boolPart = False
                            ElseIf NameBlock Like "number_pp" Then 'номер контура
                                boolNumberContour = False
                            ElseIf NameBlock Like "area" Then
                                boolGlobalArea = False
                            ElseIf NameBlock Like "spatials_elements" Then
                                boolSpatialElements = False
                            ElseIf NameBlock Like "geopoint_opred" Then
                                boolGeopoint_opred = False
                            ElseIf NameBlock Like "x" Then
                                boolX = False
                            ElseIf NameBlock Like "y" Then
                                boolY = False
                            ElseIf NameBlock Like "geopoint_opred" Then
                                boolGeopoint_opred = False
                            ElseIf NameBlock Like "ord_nmb" Then
                                boolOrgNumb = False
                            ElseIf NameBlock Like "num_geopoint" Then
                                boolnum_geopoint = False
                            ElseIf NameBlock Like "delta_geopoint" Then
                                boolDelta = False
                            ElseIf NameBlock Like "r" Then
                                boolR = False
                                '===================================================================================================
                                'конец блока с координатами контура
                            ElseIf NameBlock Like "spatial_element" Then
                                '=====================================================================================
                                'делаем оформление
                                If boolParcel = True Then
                                    'не требуется выводить участки
                                    If FormKPTImport.CheckBox1.Checked = False Then
                                        Erase NXYZ
                                        countcoord = -1
                                        Continue While
                                    End If
                                ElseIf boolBuild = True Or boolunderconstruction = True Then
                                    If FormKPTImport.CheckBox2.Checked = False Then
                                        Erase NXYZ
                                        countcoord = -1
                                        Continue While
                                    End If
                                ElseIf boolconstruction = True Then
                                    If FormKPTImport.CheckBox3.Checked = False Then
                                        Erase NXYZ
                                        countcoord = -1
                                        Continue While
                                    End If
                                ElseIf boolboundary = True Or boolcityboundary Or boolcoastline Then
                                    If FormKPTImport.CheckBox4.Checked = False Then
                                        Erase NXYZ
                                        countcoord = -1
                                        Continue While
                                    End If
                                ElseIf boolzone = True Then
                                    If FormKPTImport.CheckBox5.Checked = False Then
                                        Erase NXYZ
                                        countcoord = -1
                                        Continue While
                                    End If
                                ElseIf boolKvartal = True Then
                                    If FormKPTImport.CheckBox6.Checked = False Then
                                        Erase NXYZ
                                        countcoord = -1
                                        Continue While
                                    End If
                                End If
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                'блок лля проверки дубликатов объектов
                                If boolParcel = True Or boolBuild = True Or boolconstruction = True Or boolunderconstruction = True Then
                                    If IsArray(ArrayOldCadastrNumbers) = True Then
                                        For k As Integer = 0 To ArrayOldCadastrNumbers.Length - 1
                                            Dim OldNumber As String = ArrayOldCadastrNumbers(k)
                                            If Cad_Number Like OldNumber Then
                                                Erase NXYZ
                                                countcoord = -1
                                                Continue While
                                            End If
                                        Next
                                    End If
                                End If
                                If boolcityboundary = True Or boolboundary = True Or boolzone = True Or boolcoastline = True Then
                                    If IsArray(ArrayOldCadastrRegNumbers) = True Then
                                        For k As Integer = 0 To ArrayOldCadastrRegNumbers.Length - 1
                                            Dim OldNumber As String = ArrayOldCadastrRegNumbers(k)
                                            If RegNumbZone Like OldNumber Then
                                                Erase NXYZ
                                                countcoord = -1
                                                Continue While
                                            End If
                                        Next
                                    End If
                                End If
                                If boolKvartal = True Then
                                    If IsArray(ArrayOldCadastrKVNumbers) = True Then
                                        For k As Integer = 0 To ArrayOldCadastrKVNumbers.Length - 1
                                            Dim OldNumber As String = ArrayOldCadastrKVNumbers(k)
                                            If CadastralBlock Like OldNumber Then
                                                Erase NXYZ
                                                countcoord = -1
                                                Continue While
                                            End If
                                        Next
                                    End If
                                End If

                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                'вставляем полилинию или окружность в качестве объекта
                                If IsArray(NXYZ) = True Then
                                    Dim areaObj As Double = 0
                                    Dim idObjectUser As UInteger = 0
                                    Dim centerPoint As Point3D = Nothing
                                    Dim boolOstrov As Boolean = False
                                    If NXYZ.GetUpperBound(1) > 0 Then
                                        'если число вершин больше 1 тогда рисуем полилинию
                                        Dim acPLine As DwgPolyline = CadastralFunction.FuncDrawPolylineToArrayCoord(ActivDocument, NXYZ)
                                        If IsNothing(acPLine) = False Then
                                            idObjectUser = acPLine.ObjectID
                                            'если объект имеет площадь тогда проеверяем контур на принадлежность к островку и ищем внутренню точку
                                            If Val(GlobalArea) > 0 And NXYZ.GetUpperBound(1) > 1 Then
                                                acPLine.Closed = True
                                                areaObj = acPLine.Area
                                                'проверяем вдруг это островок
                                                If IsArray(ObjectCollections) = True Then
                                                    For j1 As Integer = ObjectCollections.GetUpperBound(1) To 0 Step -1
                                                        Dim tempNumb As String = ObjectCollections(0, j1)
                                                        If tempNumb Like "Внутренний островок" Then
                                                            Continue For
                                                        End If
                                                        Dim tempArea As Double = GlobalArea
                                                        If tempArea = 0 Then Continue For
                                                        If tempArea >= acPLine.Area Then Continue For
                                                        Dim tempHandle As UInteger = Val(ObjectCollections(1, j1))
                                                        Dim tempEnt As DwgObject = Nothing
                                                        ActivDocument.ActiveSpace.Entities.TryGetObject(tempHandle, tempEnt) 'находим его id
                                                        If IsNothing(tempEnt) = False Then
                                                            If TypeOf tempEnt Is DwgPolyline Then
                                                                Dim tempPoly As DwgPolyline = tempEnt
                                                                If tempPoly.Bounds.Min.X > acPLine.Bounds.Min.X And tempPoly.Bounds.Min.Y > acPLine.Bounds.Min.Y Then
                                                                    If tempPoly.Bounds.Max.X < acPLine.Bounds.Max.X And tempPoly.Bounds.Max.Y < acPLine.Bounds.Max.Y Then
                                                                        Dim arrayVertexDbl As Double(,) = Nothing
                                                                        Dim boolVertDbl As Boolean = FuncConvertData.FuncConvertVertexPolylineToArrayDbl(tempPoly, arrayVertexDbl)
                                                                        If IsArray(arrayVertexDbl) = True Then
                                                                            Dim booIns As Boolean = MathFunction.FuncCrossPointPolyline(arrayVertexDbl, acPLine.Item(0).Vertex)
                                                                            If IsError(booIns) = True Then
                                                                                booIns = False
                                                                            End If
                                                                            If booIns = True Then
                                                                                boolOstrov = True
                                                                                Exit For
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    Next j1
                                                End If
                                            End If
                                        End If
                                    Else
                                        '==================================================================================
                                        'вставляем окружность нужного радиуса
                                        Dim acCirc As DwgCircle = New DwgCircle()
                                        acCirc.Center = New Topomatic.Cad.Foundation.Vector3D(Val(NXYZ(3, 0)), Val(NXYZ(2, 0)), 0)
                                        r = Val(NXYZ(8, 0))
                                        If r <= 0 Then
                                            r = 1
                                        End If
                                        acCirc.Radius = r
                                        ActivDocument.ActiveSpace.Add(acCirc)
                                        idObjectUser = acCirc.ObjectID
                                        areaObj = 0
                                    End If
                                    Erase NXYZ
                                    countcoord = -1
                                    '0-обозначение, 1 -ID , 2-
                                    ReDim Preserve ObjectCollections(2, CountObjectCollection)
                                    If boolOstrov = True Then
                                        ObjectCollections(2, CountObjectCollection) = "Внутренний островок"
                                        ObjectCollections(0, CountObjectCollection) = ""
                                    ElseIf NumberSubParcel.Trim.Length > 0 Then
                                        ObjectCollections(0, CountObjectCollection) = "\" & NumberSubParcel
                                        ObjectCollections(2, CountObjectCollection) = "Часть"
                                    ElseIf NumberContour.Trim.Length > 0 Then
                                        ObjectCollections(0, CountObjectCollection) = "(" & NumberContour & ")"
                                        ObjectCollections(2, CountObjectCollection) = "Контур"
                                    Else
                                        ObjectCollections(0, CountObjectCollection) = ""
                                        ObjectCollections(2, CountObjectCollection) = "Контур"
                                    End If
                                    ObjectCollections(1, CountObjectCollection) = idObjectUser 'ID объекта
                                    CountObjectCollection += 1
                                End If
                            ElseIf NameBlock Like "land_record" Then
                                boolParcel = False
                                OldCodObject = ""
                                If IsArray(ObjectCollections) = True Then
                                    'сортируем соллекцию в случае если там более 1 элемента
                                    For k As Integer = 0 To ObjectCollections.GetUpperBound(1)
                                        Dim numKounter As String = ObjectCollections(0, k) 'обозначение контура
                                        Dim HgStr As UInteger = ObjectCollections(1, k) 'handle контура
                                        Dim typeCounter As String = ObjectCollections(2, k) 'контур или часть
                                        Dim acEnt As DwgObject = Nothing
                                        ActivDocument.ActiveSpace.Entities.TryGetObject(HgStr, acEnt) 'находим его id
                                        If IsNothing(acEnt) = False Then
                                            If TypeOf acEnt Is DwgPolyline Then
                                                'делаем оформление
                                                Dim acPline As DwgPolyline = acEnt
                                            ElseIf TypeOf acEnt Is DwgCircle Then
                                                Dim acCircle As DwgCircle = acEnt
                                            End If
                                            'присваиваем расширенные данные
                                            If typeCounter Like "Часть" Or typeCounter Like "Контур" Then

                                                Dim ArrayData As String(,) = Nothing
                                                ReDim Preserve ArrayData(1, 15)
                                                ArrayData(0, 0) = "Cad_number"
                                                If Cad_Number_Note.Trim.Length > 0 Then
                                                    ArrayData(1, 0) = Cad_Number_Note & " вх.н. " & Cad_Number & numKounter
                                                Else
                                                    ArrayData(1, 0) = Cad_Number & numKounter
                                                End If

                                                ArrayData(0, 1) = "Short_cad_number"
                                                ArrayData(1, 1) = Cad_Number_Short & numKounter

                                                ArrayData(0, 2) = "Type"
                                                ArrayData(1, 2) = SubTypeParcel

                                                ArrayData(0, 3) = "Area"
                                                ArrayData(1, 3) = GlobalArea

                                                ArrayData(0, 4) = "Inaccuracy"
                                                ArrayData(1, 4) = deltaArea

                                                ArrayData(0, 5) = "OKATO"
                                                ArrayData(1, 5) = AddresParcel

                                                ArrayData(0, 6) = "Adress"
                                                ArrayData(1, 6) = readable_address

                                                ArrayData(0, 7) = "Category"
                                                ArrayData(1, 7) = Category

                                                ArrayData(0, 8) = "UtilizationByDoc"
                                                ArrayData(1, 8) = RazrIspByDoc

                                                ArrayData(0, 9) = "Utilization"
                                                ArrayData(1, 9) = RazrIsp

                                                ArrayData(0, 10) = "CadastralCost"
                                                ArrayData(1, 10) = CostParcel

                                                ArrayData(0, 11) = "Right_type"
                                                ArrayData(1, 11) = RightType

                                                ArrayData(0, 12) = "Patronymic"
                                                ArrayData(1, 12) = ownerName

                                                ArrayData(0, 13) = "Encumbrance"
                                                ArrayData(1, 13) = encumbranceName

                                                ArrayData(0, 14) = "LOCALID"
                                                ArrayData(1, 14) = rng.Next

                                                ArrayData(0, 15) = "GlobalID"
                                                ArrayData(1, 15) = GlobalID
                                                Dim boolWriteData As Boolean = FuncXRecords.FuncCreateXData(acEnt, "LAND_CADASTRAL_PARCEL", ArrayData, templateXML)

                                            Else
                                                Dim ArrayData As String(,) = Nothing
                                                ReDim Preserve ArrayData(1, 3)
                                                ArrayData(0, 0) = "IDKOUNTER"
                                                ArrayData(1, 0) = LocalID

                                                ArrayData(0, 1) = "LOCALID"
                                                ArrayData(1, 1) = rng.Next

                                                ArrayData(0, 2) = "ConditionalNumber"
                                                ArrayData(1, 2) = k + 1

                                                ArrayData(0, 3) = "NOTE"
                                                ArrayData(1, 3) = Cad_Number & numKounter
                                                Dim boolWriteData As Boolean = FuncXRecords.FuncCreateXData(acEnt, "CADASTRAL_OSTROV", ArrayData, templateXML)
                                            End If
                                        End If
                                    Next k
                                End If

                                Erase ObjectCollections
                                Cad_Number_Note = ""
                                GlobalArea = 0
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                'plfybz b cnhjtybz
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            ElseIf NameBlock Like "build_record" Then
                                boolBuild = False
                                OldCodObject = ""
                                'блок кода для определения внутренних островков
                                If IsArray(ObjectCollections) = True Then
                                    'сортируем соллекцию в случае если там более 1 элемента
                                    For k As Integer = 0 To ObjectCollections.GetUpperBound(1)
                                        Dim numKounter As String = ObjectCollections(0, k) 'обозначение контура
                                        Dim HgStr As UInteger = ObjectCollections(1, k) 'handle контура
                                        Dim typeCounter As String = ObjectCollections(2, k) 'контур или часть
                                        Dim acEnt As DwgObject = Nothing
                                        ActivDocument.ActiveSpace.Entities.TryGetObject(HgStr, acEnt) 'находим его id
                                        If IsNothing(acEnt) = False Then
                                            If TypeOf acEnt Is DwgPolyline Then
                                                'делаем оформление
                                                Dim acPline As DwgPolyline = acEnt
                                            ElseIf TypeOf acEnt Is DwgCircle Then
                                                Dim acCircle As DwgCircle = acEnt
                                            End If
                                            'присваиваем расширенные данные
                                            If typeCounter Like "Часть" Or typeCounter Like "Контур" Then
                                                Dim ArrayData As String(,) = Nothing
                                                ReDim Preserve ArrayData(1, 15)
                                                ArrayData(0, 0) = "Cad_number"
                                                ArrayData(1, 0) = Cad_Number & numKounter

                                                ArrayData(0, 1) = "Short_cad_number"
                                                ArrayData(1, 1) = Cad_Number_Short & numKounter

                                                ArrayData(0, 2) = "Assignation"
                                                ArrayData(1, 2) = permitted_use

                                                ArrayData(0, 3) = "Area"
                                                ArrayData(1, 3) = GlobalArea

                                                ArrayData(0, 4) = "OKATO"
                                                ArrayData(1, 4) = AddresParcel

                                                ArrayData(0, 5) = "Adress"
                                                ArrayData(1, 5) = readable_address

                                                ArrayData(0, 6) = "CadastralCost"
                                                ArrayData(1, 6) = CostParcel

                                                ArrayData(0, 7) = "KeyParameter"
                                                ArrayData(1, 7) = baseParameters

                                                ArrayData(0, 8) = "State"
                                                ArrayData(1, 8) = purpose_build

                                                ArrayData(0, 9) = "Right_type"
                                                ArrayData(1, 9) = RightType

                                                ArrayData(0, 10) = "Patronymic"
                                                ArrayData(1, 10) = ownerName

                                                ArrayData(0, 11) = "Encumbrance"
                                                ArrayData(1, 11) = encumbranceName

                                                ArrayData(0, 12) = "LOCALID"
                                                ArrayData(1, 12) = LocalID

                                                ArrayData(0, 13) = "GlobalID"
                                                ArrayData(1, 13) = GlobalID

                                                Dim boolWriteData As Boolean = FuncXRecords.FuncCreateXData(acEnt, "LAND_CADASTRAL_BUILDING", ArrayData, templateXML)
                                            Else
                                                Dim ArrayData As String(,) = Nothing
                                                ReDim Preserve ArrayData(1, 3)
                                                ArrayData(0, 0) = "IDKOUNTER"
                                                ArrayData(1, 0) = LocalID

                                                ArrayData(0, 1) = "LOCALID"
                                                ArrayData(1, 1) = rng.Next

                                                ArrayData(0, 2) = "ConditionalNumber"
                                                ArrayData(1, 2) = k + 1

                                                ArrayData(0, 3) = "NOTE"
                                                ArrayData(1, 3) = Cad_Number & numKounter
                                                Dim boolWriteData As Boolean = FuncXRecords.FuncCreateXData(acEnt, "CADASTRAL_OSTROV", ArrayData, templateXML)
                                            End If
                                        End If
                                    Next k
                                End If
                                Erase ObjectCollections
                                Cad_Number_Short = ""
                                purpose_build = ""
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            ElseIf NameBlock Like "municipal_boundary_record" Then
                                boolboundary = False
                                OldCodObject = ""
                                If IsArray(ObjectCollections) = True Then
                                    'сортируем соллекцию в случае если там более 1 элемента
                                    For k As Integer = 0 To ObjectCollections.GetUpperBound(1)
                                        Dim HgStr As UInteger = ObjectCollections(1, k) 'handle контура
                                        Dim typeCounter As String = ObjectCollections(2, k) 'контур или часть
                                        Dim acEnt As DwgObject = Nothing
                                        ActivDocument.ActiveSpace.Entities.TryGetObject(HgStr, acEnt) 'находим его id
                                        If IsNothing(acEnt) = False Then
                                            If TypeOf acEnt Is DwgPolyline Then
                                                'делаем оформление
                                                Dim acPline As DwgPolyline = acEnt
                                            End If
                                            'присваиваем расширенные данные
                                            If typeCounter Like "Часть" Or typeCounter Like "Контур" Then
                                                Dim ArrayData As String(,) = Nothing
                                                ReDim Preserve ArrayData(1, 4)
                                                ArrayData(0, 0) = "Reg_numb_border"
                                                ArrayData(1, 0) = RegNumbZone

                                                ArrayData(0, 1) = "Type_boundary"
                                                ArrayData(1, 1) = type_boundary

                                                ArrayData(0, 2) = "DateRegistration"
                                                ArrayData(1, 2) = registr_data

                                                ArrayData(0, 3) = "LOCALID"
                                                ArrayData(1, 3) = LocalID

                                                ArrayData(0, 4) = "GlobalID"
                                                ArrayData(1, 4) = GlobalID

                                                Dim boolWriteData As Boolean = FuncXRecords.FuncCreateXData(acEnt, "LAND_CADASTRAL_BOUNDARY", ArrayData, templateXML)

                                            Else
                                                Dim ArrayData As String(,) = Nothing
                                                ReDim Preserve ArrayData(1, 3)
                                                ArrayData(0, 0) = "IDKOUNTER"
                                                ArrayData(1, 0) = LocalID

                                                ArrayData(0, 1) = "LOCALID"
                                                ArrayData(1, 1) = rng.Next

                                                ArrayData(0, 2) = "ConditionalNumber"
                                                ArrayData(1, 2) = k + 1

                                                ArrayData(0, 3) = "NOTE"
                                                ArrayData(1, 3) = RegNumbZone
                                                Dim boolWriteData As Boolean = FuncXRecords.FuncCreateXData(acEnt, "CADASTRAL_OSTROV", ArrayData, templateXML)
                                            End If
                                        End If
                                    Next k
                                End If
                                Erase ObjectCollections
                                RegNumbZone = ""
                            ElseIf NameBlock Like "inhabited_locality_boundary_record" Then
                                boolcityboundary = False
                                OldCodObject = ""
                                If IsArray(ObjectCollections) = True Then
                                    'сортируем соллекцию в случае если там более 1 элемента
                                    For k As Integer = 0 To ObjectCollections.GetUpperBound(1)
                                        Dim numKounter As String = ObjectCollections(0, k) 'обозначение контура
                                        Dim HgStr As UInteger = ObjectCollections(1, k) 'handle контура
                                        Dim typeCounter As String = ObjectCollections(2, k) 'контур или часть
                                        Dim acEnt As DwgObject = Nothing
                                        ActivDocument.ActiveSpace.Entities.TryGetObject(HgStr, acEnt) 'находим его id
                                        If IsNothing(acEnt) = False Then
                                            If TypeOf acEnt Is DwgPolyline Then
                                                'делаем оформление
                                                Dim acPline As DwgPolyline = acEnt
                                            End If
                                            'присваиваем расширенные данные
                                            If typeCounter Like "Часть" Or typeCounter Like "Контур" Then
                                                Dim ArrayData As String(,) = Nothing
                                                ReDim Preserve ArrayData(1, 4)
                                                ArrayData(0, 0) = "Reg_numb_border"
                                                ArrayData(1, 0) = RegNumbZone

                                                ArrayData(0, 1) = "Type_boundary"
                                                ArrayData(1, 1) = type_boundary

                                                ArrayData(0, 2) = "DateRegistration"
                                                ArrayData(1, 2) = registr_data

                                                ArrayData(0, 3) = "LOCALID"
                                                ArrayData(1, 3) = LocalID

                                                ArrayData(0, 4) = "GlobalID"
                                                ArrayData(1, 4) = GlobalID

                                                Dim boolWriteData As Boolean = FuncXRecords.FuncCreateXData(acEnt, "LAND_CADASTRAL_BOUNDARY", ArrayData, templateXML)
                                            Else
                                                Dim ArrayData As String(,) = Nothing
                                                ReDim Preserve ArrayData(1, 3)
                                                ArrayData(0, 0) = "IDKOUNTER"
                                                ArrayData(1, 0) = LocalID

                                                ArrayData(0, 1) = "LOCALID"
                                                ArrayData(1, 1) = rng.Next

                                                ArrayData(0, 2) = "ConditionalNumber"
                                                ArrayData(1, 2) = k + 1

                                                ArrayData(0, 3) = "NOTE"
                                                ArrayData(1, 3) = RegNumbZone
                                                Dim boolWriteData As Boolean = FuncXRecords.FuncCreateXData(acEnt, "CADASTRAL_OSTROV", ArrayData, templateXML)
                                            End If
                                        End If
                                    Next k
                                End If
                                Erase ObjectCollections
                                RegNumbZone = ""
                            ElseIf NameBlock Like "construction_record" Then
                                boolconstruction = False
                                OldCodObject = ""
                                If IsArray(ObjectCollections) = True Then
                                    'сортируем соллекцию в случае если там более 1 элемента
                                    For k As Integer = 0 To ObjectCollections.GetUpperBound(1)
                                        Dim numKounter As String = ObjectCollections(0, k) 'обозначение контура
                                        Dim HgStr As UInteger = ObjectCollections(1, k) 'handle контура
                                        Dim typeCounter As String = ObjectCollections(2, k) 'контур или часть
                                        Dim acEnt As DwgObject = Nothing
                                        ActivDocument.ActiveSpace.Entities.TryGetObject(HgStr, acEnt) 'находим его id
                                        If IsNothing(acEnt) = False Then
                                            If TypeOf acEnt Is DwgPolyline Then
                                                'делаем оформление
                                                Dim acPline As DwgPolyline = acEnt
                                            ElseIf TypeOf acEnt Is DwgCircle Then
                                                Dim acCircle As DwgCircle = acEnt
                                            End If
                                            'присваиваем расширенные данные
                                            If typeCounter Like "Часть" Or typeCounter Like "Контур" Then

                                                Dim ArrayData As String(,) = Nothing
                                                ReDim Preserve ArrayData(1, 11)
                                                ArrayData(0, 0) = "Cad_number"
                                                ArrayData(1, 0) = Cad_Number & numKounter

                                                ArrayData(0, 1) = "Short_cad_number"
                                                ArrayData(1, 1) = Cad_Number_Short & numKounter

                                                ArrayData(0, 2) = "Assignation"
                                                ArrayData(1, 2) = purpose_construction

                                                ArrayData(0, 3) = "OKATO"
                                                ArrayData(1, 3) = AddresParcel

                                                ArrayData(0, 4) = "Adress"
                                                ArrayData(1, 4) = readable_address

                                                ArrayData(0, 5) = "CadastralCost"
                                                ArrayData(1, 5) = CostParcel

                                                ArrayData(0, 6) = "KeyParameter"
                                                ArrayData(1, 6) = baseParameters

                                                ArrayData(0, 7) = "Right_type"
                                                ArrayData(1, 7) = RightType

                                                ArrayData(0, 8) = "Patronymic"
                                                ArrayData(1, 8) = ownerName

                                                ArrayData(0, 9) = "Encumbrance"
                                                ArrayData(1, 9) = encumbranceName

                                                ArrayData(0, 10) = "LOCALID"
                                                ArrayData(1, 10) = LocalID

                                                ArrayData(0, 11) = "GlobalID"
                                                ArrayData(1, 11) = GlobalID

                                                Dim boolWriteData As Boolean = FuncXRecords.FuncCreateXData(acEnt, "LAND_CADASTRAL_CONSTRUCTION", ArrayData, templateXML)

                                            Else
                                                Dim ArrayData As String(,) = Nothing
                                                ReDim Preserve ArrayData(1, 3)
                                                ArrayData(0, 0) = "IDKOUNTER"
                                                ArrayData(1, 0) = LocalID

                                                ArrayData(0, 1) = "LOCALID"
                                                ArrayData(1, 1) = rng.Next

                                                ArrayData(0, 2) = "ConditionalNumber"
                                                ArrayData(1, 2) = k + 1

                                                ArrayData(0, 3) = "NOTE"
                                                ArrayData(1, 3) = Cad_Number & numKounter
                                                Dim boolWriteData As Boolean = FuncXRecords.FuncCreateXData(acEnt, "CADASTRAL_OSTROV", ArrayData, templateXML)
                                            End If
                                        End If
                                    Next k
                                End If
                                Erase ObjectCollections
                                Cad_Number_Short = ""
                            ElseIf NameBlock Like "object_under_construction_record" Then
                                'блок кода для определения внутренних островков
                                OldCodObject = ""
                                If IsArray(ObjectCollections) = True Then
                                    'сортируем соллекцию в случае если там более 1 элемента
                                    For k As Integer = 0 To ObjectCollections.GetUpperBound(1)
                                        Dim numKounter As String = ObjectCollections(0, k) 'обозначение контура
                                        Dim HgStr As UInteger = ObjectCollections(1, k) 'handle контура
                                        Dim typeCounter As String = ObjectCollections(2, k) 'контур или часть
                                        Dim acEnt As DwgObject = Nothing
                                        ActivDocument.ActiveSpace.Entities.TryGetObject(HgStr, acEnt) 'находим его id
                                        If IsNothing(acEnt) = False Then
                                            If TypeOf acEnt Is DwgPolyline Then
                                                'делаем оформление
                                                Dim acPline As DwgPolyline = acEnt
                                            ElseIf TypeOf acEnt Is DwgCircle Then
                                                Dim acCircle As DwgCircle = acEnt
                                            End If
                                            'присваиваем расширенные данные
                                            If typeCounter Like "Часть" Or typeCounter Like "Контур" Then

                                                Dim ArrayData As String(,) = Nothing
                                                ReDim Preserve ArrayData(1, 11)
                                                ArrayData(0, 0) = "Cad_number"
                                                ArrayData(1, 0) = Cad_Number & numKounter

                                                ArrayData(0, 1) = "Short_cad_number"
                                                ArrayData(1, 1) = Cad_Number_Short & numKounter

                                                ArrayData(0, 2) = "Assignation"
                                                ArrayData(1, 2) = permitted_use

                                                ArrayData(0, 3) = "OKATO"
                                                ArrayData(1, 3) = AddresParcel

                                                ArrayData(0, 4) = "Adress"
                                                ArrayData(1, 4) = readable_address

                                                ArrayData(0, 5) = "CadastralCost"
                                                ArrayData(1, 5) = CostParcel

                                                ArrayData(0, 6) = "KeyParameter"
                                                ArrayData(1, 6) = baseParameters

                                                ArrayData(0, 7) = "Right_type"
                                                ArrayData(1, 7) = RightType

                                                ArrayData(0, 8) = "Patronymic"
                                                ArrayData(1, 8) = ownerName

                                                ArrayData(0, 9) = "Encumbrance"
                                                ArrayData(1, 9) = encumbranceName

                                                ArrayData(0, 10) = "LOCALID"
                                                ArrayData(1, 10) = LocalID

                                                ArrayData(0, 11) = "GlobalID"
                                                ArrayData(1, 11) = GlobalID

                                                Dim boolWriteData As Boolean = FuncXRecords.FuncCreateXData(acEnt, "LAND_CADASTRAL_BUILDING", ArrayData, templateXML)

                                            Else
                                                Dim ArrayData As String(,) = Nothing
                                                ReDim Preserve ArrayData(1, 3)
                                                ArrayData(0, 0) = "IDKOUNTER"
                                                ArrayData(1, 0) = LocalID

                                                ArrayData(0, 1) = "LOCALID"
                                                ArrayData(1, 1) = rng.Next

                                                ArrayData(0, 2) = "ConditionalNumber"
                                                ArrayData(1, 2) = k + 1

                                                ArrayData(0, 3) = "NOTE"
                                                ArrayData(1, 3) = Cad_Number & numKounter
                                                Dim boolWriteData As Boolean = FuncXRecords.FuncCreateXData(acEnt, "CADASTRAL_OSTROV", ArrayData, templateXML)
                                            End If
                                        End If
                                    Next k
                                End If
                                Erase ObjectCollections
                                boolunderconstruction = False
                                Cad_Number_Short = ""
                            ElseIf NameBlock Like "zones_and_territories_record" Then
                                boolzone = False
                                OldCodObject = ""
                                If IsArray(ObjectCollections) = True Then
                                    'сортируем соллекцию в случае если там более 1 элемента
                                    For k As Integer = 0 To ObjectCollections.GetUpperBound(1)
                                        Dim numKounter As String = ObjectCollections(0, k) 'обозначение контура
                                        Dim HgStr As UInteger = ObjectCollections(1, k) 'handle контура
                                        Dim typeCounter As String = ObjectCollections(2, k) 'контур или часть
                                        Dim acEnt As DwgObject = Nothing
                                        ActivDocument.ActiveSpace.Entities.TryGetObject(HgStr, acEnt) 'находим его id
                                        If IsNothing(acEnt) = False Then
                                            If TypeOf acEnt Is DwgPolyline Then
                                                'делаем оформление
                                                Dim acPline As DwgPolyline = acEnt
                                            End If
                                            'присваиваем расширенные данные
                                            If typeCounter Like "Часть" Or typeCounter Like "Контур" Then

                                                Dim ArrayData As String(,) = Nothing
                                                ReDim Preserve ArrayData(1, 7)
                                                ArrayData(0, 0) = "Reg_numb_border"
                                                ArrayData(1, 0) = RegNumbZone

                                                ArrayData(0, 1) = "Type_boundary"
                                                ArrayData(1, 1) = type_boundary

                                                ArrayData(0, 2) = "Type_zone"
                                                ArrayData(1, 2) = type_zona

                                                ArrayData(0, 3) = "Index"
                                                ArrayData(1, 3) = IndexZona

                                                ArrayData(0, 4) = "Number"
                                                ArrayData(1, 4) = NumberZona

                                                ArrayData(0, 5) = "DateRegistration"
                                                ArrayData(1, 5) = registr_data

                                                ArrayData(0, 6) = "LOCALID"
                                                ArrayData(1, 6) = LocalID

                                                ArrayData(0, 7) = "GlobalID"
                                                ArrayData(1, 7) = GlobalID

                                                Dim boolWriteData As Boolean = FuncXRecords.FuncCreateXData(acEnt, "LAND_CADASTRAL_ZONE", ArrayData, templateXML)

                                            Else
                                                Dim ArrayData As String(,) = Nothing
                                                ReDim Preserve ArrayData(2, 3)
                                                ArrayData(0, 0) = "IDKOUNTER"
                                                ArrayData(1, 0) = LocalID

                                                ArrayData(0, 1) = "LOCALID"
                                                ArrayData(1, 1) = rng.Next

                                                ArrayData(0, 2) = "ConditionalNumber"
                                                ArrayData(1, 2) = k + 1

                                                ArrayData(0, 3) = "NOTE"
                                                ArrayData(1, 3) = RegNumbZone
                                                Dim boolWriteData As Boolean = FuncXRecords.FuncCreateXData(acEnt, "CADASTRAL_OSTROV", ArrayData, templateXML)
                                            End If
                                        End If
                                    Next k
                                End If
                                Erase ObjectCollections
                                RegNumbZone = ""
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                'кадастровый квартал
                            ElseIf NameBlock Like "spatial_data" Then
                                boolspatial_data = False
                                OldCodObject = ""
                                If IsArray(ObjectCollections) = True Then
                                    'сортируем соллекцию в случае если там более 1 элемента
                                    For k As Integer = 0 To ObjectCollections.GetUpperBound(1)
                                        Dim numKounter As String = ObjectCollections(0, k) 'обозначение контура
                                        Dim HgStr As UInteger = ObjectCollections(1, k) 'handle контура
                                        Dim typeCounter As String = ObjectCollections(2, k) 'контур или часть
                                        Dim acEnt As DwgObject = Nothing
                                        ActivDocument.ActiveSpace.Entities.TryGetObject(HgStr, acEnt) 'находим его id
                                        If IsNothing(acEnt) = False Then
                                            If TypeOf acEnt Is DwgPolyline Then
                                                'делаем оформление
                                                Dim acPline As DwgPolyline = acEnt

                                            End If
                                            'присваиваем расширенные данные
                                            If typeCounter Like "Часть" Or typeCounter Like "Контур" Then

                                                Dim ArrayData As String(,) = Nothing
                                                ReDim Preserve ArrayData(1, 3)
                                                ArrayData(0, 0) = "Cadastral_number"
                                                ArrayData(1, 0) = CadastralBlock

                                                ArrayData(0, 1) = "Area"
                                                ArrayData(1, 1) = Area_quarter

                                                ArrayData(0, 2) = "LOCALID"
                                                ArrayData(1, 2) = LocalID

                                                ArrayData(0, 3) = "GlobalID"
                                                ArrayData(1, 3) = GlobalID

                                                Dim boolWriteData As Boolean = FuncXRecords.FuncCreateXData(acEnt, "LAND_CADASTRAL_DIVISION", ArrayData, templateXML)

                                            Else
                                                Dim ArrayData As String(,) = Nothing
                                                ReDim Preserve ArrayData(1, 3)
                                                ArrayData(0, 0) = "IDKOUNTER"
                                                ArrayData(1, 0) = LocalID

                                                ArrayData(0, 1) = "LOCALID"
                                                ArrayData(1, 1) = rng.Next

                                                ArrayData(0, 2) = "ConditionalNumber"
                                                ArrayData(1, 2) = k + 1

                                                ArrayData(0, 3) = "NOTE"
                                                ArrayData(1, 3) = CadastralBlock
                                                Dim boolWriteData As Boolean = FuncXRecords.FuncCreateXData(acEnt, "CADASTRAL_OSTROV", ArrayData, templateXML)
                                            End If
                                        End If
                                    Next k
                                End If
                                Erase ObjectCollections
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                'береговые линии
                            ElseIf NameBlock Like "coastline_record" Then
                                boolcoastline = False
                                OldCodObject = ""
                                If IsArray(ObjectCollections) = True Then
                                    'сортируем соллекцию в случае если там более 1 элемента
                                    For k As Integer = 0 To ObjectCollections.GetUpperBound(1)
                                        Dim numKounter As String = ObjectCollections(0, k) 'обозначение контура
                                        Dim HgStr As UInteger = ObjectCollections(1, k) 'handle контура
                                        Dim typeCounter As String = ObjectCollections(2, k) 'контур или часть
                                        Dim acEnt As DwgObject = Nothing
                                        ActivDocument.ActiveSpace.Entities.TryGetObject(HgStr, acEnt) 'находим его id
                                        If IsNothing(acEnt) = False Then
                                            If TypeOf acEnt Is DwgPolyline Then
                                                'делаем оформление
                                                Dim acPline As DwgPolyline = acEnt
                                            End If
                                            'присваиваем расширенные данные
                                            If typeCounter Like "Часть" Or typeCounter Like "Контур" Then

                                                Dim ArrayData As String(,) = Nothing
                                                ReDim Preserve ArrayData(1, 6)
                                                ArrayData(0, 0) = "Reg_numb_border"
                                                ArrayData(1, 0) = RegNumbZone & numKounter

                                                ArrayData(0, 1) = "Type_boundary"
                                                ArrayData(1, 1) = type_boundary

                                                ArrayData(0, 2) = "Name"
                                                ArrayData(1, 2) = water_object_name

                                                ArrayData(0, 3) = "Water_object_type"
                                                ArrayData(1, 3) = typewater_object_type

                                                ArrayData(0, 4) = "DateRegistration"
                                                ArrayData(1, 4) = registr_data

                                                ArrayData(0, 5) = "LOCALID"
                                                ArrayData(1, 5) = LocalID

                                                ArrayData(0, 6) = "GlobalID"
                                                ArrayData(1, 6) = GlobalID

                                                Dim boolWriteData As Boolean = FuncXRecords.FuncCreateXData(acEnt, "LAND_CADASTRAL_ZONE", ArrayData, templateXML)

                                            Else
                                                Dim ArrayData As String(,) = Nothing
                                                ReDim Preserve ArrayData(1, 3)
                                                ArrayData(0, 0) = "IDKOUNTER"
                                                ArrayData(1, 0) = LocalID

                                                ArrayData(0, 1) = "LOCALID"
                                                ArrayData(1, 1) = rng.Next

                                                ArrayData(0, 2) = "ConditionalNumber"
                                                ArrayData(1, 2) = k + 1

                                                ArrayData(0, 3) = "NOTE"
                                                ArrayData(1, 3) = RegNumbZone & numKounter
                                                Dim boolWriteData As Boolean = FuncXRecords.FuncCreateXData(acEnt, "CADASTRAL_OSTROV", ArrayData, templateXML)
                                            End If
                                        End If
                                    Next k
                                End If
                                Erase ObjectCollections
                                RegNumbZone = ""
                                'Dim PosText As Boolean = AcadFunction.FuncDrawMText(InsertPointText, kanNumStr, HtextZone, LayerObject, idTTextStyle)
                                water_object_name = ""
                                typewater_object_type = ""
                            ElseIf NameBlock Like "water_object_type" Then
                                boolwater_object_type = False
                            ElseIf NameBlock Like "water_object_name" Then
                                boolwater_object_name = False
                            ElseIf NameBlock Like "extract_cadastral_plan_territory" Then
                            ElseIf NameBlock Like "land_records" Then
                                CodCategory = ""
                            End If
                    End Select
                End While
            End If
        End Sub
        '===================================================================================================================================
        'Сведения о наличии зон с особыми условиями использования территории
        <cmd("CreateSVCadastralPlan")>
        Public Sub CreateSVCadastralPlan()
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
            'получаем границу
            userCadView.SelectionSet.Clear()
            userCadView.SelectionSet.SelectOneObjectAtScreen(Function(obj) TypeOf obj Is DwgPolyline, "Выберите границу полосы отвода: ")
            Dim plineZone As DwgPolyline = Nothing
            Dim boundPolyline3d As Polyline3D = New Polyline3D()
            Dim boundBoxPolyline As BoundingBox2D = Nothing
            Dim l1stBoundVertex As List(Of Vector2D) = New List(Of Vector2D)
            If userCadView.SelectionSet.Count > 0 Then
                For Each acEnt As Object In userCadView.SelectionSet
                    If TypeOf acEnt Is DwgPolyline Then
                        Dim acPoly As DwgPolyline = acEnt
                        boundBoxPolyline = acPoly.Bounds
                        acPoly.GetPolyline(boundPolyline3d)
                        boundPolyline3d.ConvertToPosArray(l1stBoundVertex, ValueConverter.Eps)
                        plineZone = acEnt
                        Exit For
                    End If
                Next
            End If
            Dim arrayPositionsBound As Double() = {}
            Dim countArrayPositionsBound As Integer = 0
            If l1stBoundVertex.Count > 2 Then
                For i As Integer = 0 To l1stBoundVertex.Count - 1
                    ReDim Preserve arrayPositionsBound(countArrayPositionsBound + 1)
                    arrayPositionsBound(countArrayPositionsBound) = l1stBoundVertex.Item(i).X
                    arrayPositionsBound(countArrayPositionsBound + 1) = l1stBoundVertex.Item(i).Y
                    countArrayPositionsBound += 2
                Next i
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'создаем новый сводный план
            Dim userProject As ModelProject = ApplicationHost.Current.ActiveProject
            Dim folder As String = PluginCoreOps.FindModelPathId(PluginCoreOps.CreateFolder(New String() {"Модели", "LandCadastral", "CadastralPlan"}))
            Try
                Dim iUserProject As IProjectModel = ApplicationHost.Current.Plugins.Execute("mkitem", New Object() {folder, "arr"})
                Dim boolCreateModel As Boolean = ApplicationHost.Current.Plugins.Execute("activate", New Object() {iUserProject})
                If boolCreateModel = True Then
                    DrawLayer = DrawingLayer.GetDrawingLayer(userCadView)
                    ActivDocument = DrawLayer.Drawing
                    userProject = iUserProject.Project
                    Dim attrs As Attributes = New Attributes()
                    userProject.SetCommonModelAttributes(iUserProject.Uri, "SVPlan", attrs)
                End If
            Catch ex As System.Exception
            End Try
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Dim arrayObjectLocalID As String() = {}
            Dim countArrayObjectLocalID As Integer = 0
            userCadView.SelectionSet.SelectAll()
            For Each acCadastralObject As Object In userCadView.SelectionSet
                If TypeOf acCadastralObject Is DwgEntity Then
                    Dim drawigObject As Dwg.Drawing = acCadastralObject.Drawing
                    Dim fullName As String = drawigObject.Filename
                    Dim boolFindFolder As Boolean = FuncFiles.FuncIsFolderInPathSafe(fullName, "CadastralPlan")
                    If boolFindFolder = True Then
                        'ищем пересечения
                        If TypeOf acCadastralObject Is DwgPolyline Then
                            Dim plineEnt As DwgPolyline = acCadastralObject
                            Dim boolIntersect As Boolean = plineZone.IntersectWith(plineEnt.Bounds, 1)
                            If boolIntersect = True Then
                                Dim localID As String = FuncXRecords.FuncReadValueXData(plineEnt, "LOCALID")
                                If localID.Trim.Length > 1 Then
                                    Dim boolFindObject As Boolean = False
                                    For i As Integer = 0 To arrayObjectLocalID.Length - 1
                                        If localID Like arrayObjectLocalID(i) Then
                                            boolFindObject = True
                                            Exit For
                                        End If
                                    Next i
                                    If boolFindObject = False Then
                                        ReDim Preserve arrayObjectLocalID(countArrayObjectLocalID)
                                        arrayObjectLocalID(countArrayObjectLocalID) = localID
                                        countArrayObjectLocalID += 1

                                        Dim numberParcel As String = FuncXRecords.FuncReadValueXData(plineEnt, "Short_cad_number", "PROJECT_BRIDGE")
                                        If numberParcel.Trim.Length = 0 Then
                                            numberParcel = FuncXRecords.FuncReadValueXData(plineEnt, "Reg_numb_border", "PROJECT_BRIDGE")
                                        End If
                                        If numberParcel.Trim.Length > 0 Then
                                            Dim arrayVertexDbl As Double(,) = Nothing
                                            Dim boolVertDbl As Boolean = FuncConvertData.FuncConvertVertexPolylineToArrayDbl(plineEnt, arrayVertexDbl)
                                            If IsArray(arrayVertexDbl) = True Then
                                                Dim CrossPointFirstPline As Topomatic.Cad.Foundation.Vector3D = MathFunction.FuncFindCrossCenterPointPolyline(arrayVertexDbl)
                                                Dim acMText As DwgMText = RoburFunc.FuncDrawMText(ActivDocument, numberParcel, CrossPointFirstPline)
                                                If IsNothing(acMText) = False Then
                                                    'acMText.Height = HtextZone
                                                    'acMText.Layer = zonesLayer
                                                    acMText.Annotative = True
                                                End If
                                            End If



                                            'boundPolyline3d = New Polyline3D()
                                            'l1stBoundVertex = New List(Of Vector2D)
                                            'plineEnt.GetPolyline(boundPolyline3d)
                                            'boundPolyline3d.ConvertToPosArray(l1stBoundVertex, ValueConverter.Eps)

                                            'Dim arrayPositionsPline As Double() = {}
                                            'Dim countArrayPositionsPLine As Integer = 0
                                            'If l1stBoundVertex.Count > 2 Then
                                            '    For i As Integer = 0 To l1stBoundVertex.Count - 1
                                            '        ReDim Preserve arrayPositionsPline(countArrayPositionsPLine + 1)
                                            '        arrayPositionsPline(countArrayPositionsPLine) = l1stBoundVertex.Item(i).X
                                            '        arrayPositionsPline(countArrayPositionsPLine + 1) = l1stBoundVertex.Item(i).Y
                                            '        countArrayPositionsPLine += 2
                                            '    Next i
                                            'End If
                                            'If arrayPositionsPline.Length > 6 Then
                                            '    Dim intersections = MathFunction.FindPolygonIntersections(arrayPositionsBound, arrayPositionsPline)

                                            'If intersections IsNot Nothing Then
                                            '    MsgBox("Найдено " & intersections.Length & " областей пересечения")
                                            '    For Each polygon In intersections

                                            '        Console.WriteLine($"Площадь области пересечения: {polygon.Area}")
                                            '    Next
                                            'Else
                                            '    Console.WriteLine("Пересечений не найдено")
                                            'End If
                                        End If
                                    End If
                                End If
                                Dim plineObject = plineEnt.Clone()
                                ActivDocument.ActiveSpace.Add(plineObject)
                            End If
                        End If
                    End If
                End If
            Next


            'userProject = ApplicationHost.Current.ActiveProject
            ''// Получение всех моделей проекта
            'Dim childs As IProjectModel() = userProject.Model.GetChilds()
            'For Each child As IProjectModel In childs
            '    Dim modelUri As URI = child.Uri
            '    If modelUri.Extension Like ".arrx" Then
            '        Dim boolFindFolder As Boolean = FuncFiles.FuncIsFolderInPathSafe(fullName, "CadastralPlan")
            '        If boolFindFolder = True Then
            '            DrawLayer = DrawingLayer.GetDrawingLayer(userCadView)
            '            ActivDocument = DrawLayer.Drawing
            '            For Each acCadastralObject As DwgEntity In ActivDocument.ActiveSpace.Entities
            '                Dim drawigObject As Dwg.Drawing = acCadastralObject.Drawing
            '                Dim fullName As String = drawigObject.Filename
            '                Dim boolFindFolder As Boolean = FuncFiles.FuncIsFolderInPathSafe(fullName, "CadastralPlan")
            '                If boolFindFolder = True Then
            '                    'ищем пересечения
            '                    If TypeOf acCadastralObject Is DwgPolyline Then
            '                        Dim plineEnt As DwgPolyline = acCadastralObject
            '                        Dim boolIntersect As Boolean = plineZone.IntersectWith(plineEnt.Bounds, 1)
            '                        Dim plineObject = plineEnt.Clone()
            '                        ActivDocument.ActiveSpace.Add(plineObject)
            '                    End If
            '                End If
            '            Next

            '        End If
            '    End If
            'Next


            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'перебираем элементы




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
                Dim boolCreateStyle As Boolean = FuncDecorPlane.FuncCreateStyleLinearObject(ActivDocument, templateXML, arrayPropertiesParcels, parcelsLayer, parcelsTypeLine, parcelsColor, parcelsScaleTypeLine, parcelsWightLine, parcelsPSTable, parcelsGWidth)
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
                FuncDecorPlane.FuncCreateStylePointObject(ActivDocument, templateXML, arrayPropertiesPoints, pointName, pointLayer, pointTypeLine, pointColor, pointScaleLineType, pointWightLine, pointScale, pointAngle, pointAttributeHeight, pointPSTable, pointNote)
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
            Dim globID As String = ""
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'получаем все видимые модели проекта

            '1.\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'сброс выделения
            userCadView.SelectionSet.Clear()
            Dim formCreateParcels As FormCreateParcels = New FormCreateParcels
            formCreateParcels.userCadView = userCadView
            formCreateParcels.ShowDialog()
            If formCreateParcels.boolShow = False Then
                Exit Sub
            End If
            '=========================================================================================================================
            Dim activUserModel As String = formCreateParcels.ComboBox4.Text
            Dim Project As ModelProject = ApplicationHost.Current.ActiveProject
            '// Получение всех моделей проекта
            Dim childs As IProjectModel() = Project.Model.GetChilds()
            For Each child As IProjectModel In childs
                Dim modelUri As URI = child.Uri
                If modelUri.Extension Like ".arrx" Then
                    Dim fileNameSfcx As String = Path.GetFileNameWithoutExtension(modelUri.LastPathComponent)
                    If fileNameSfcx.Trim Like activUserModel.Trim Then
                        ApplicationHost.Current.Plugins.Execute("activate", New Object() {child})
                        Exit For
                    End If
                End If
            Next
            Dim DrawShowLayers As IEnumerable(Of DrawingLayer) = DrawingLayer.GetDrawingLayers(userCadView)
            If IsNothing(activUserModel) = False Then
                If activUserModel.Trim.Length > 0 Then
                    For k1 As Integer = 0 To DrawShowLayers.Count - 1
                        Dim TempDrawLayer As DrawingLayer = DrawShowLayers.ElementAt(k1)
                        Dim nameUserFile As String = TempDrawLayer.Drawing.Filename
                        nameUserFile = IO.Path.GetFileNameWithoutExtension(nameUserFile)
                        If nameUserFile Like activUserModel Then
                            DrawLayer = TempDrawLayer
                            ActivDocument = DrawLayer.Drawing
                            Exit For
                        End If
                    Next k1
                End If
            End If
            If formCreateParcels.CheckBox1.Checked = True Then
                globID = rng.Next
            End If
            If IsArray(formCreateParcels.arrayObjectProperties) = True Then
                'участок
                Dim parcelsLayer As DwgLayer = ActivDocument.ActiveLayer
                Dim parcelsLineType As DwgLinetype = ActivDocument.ActiveLinetype
                Dim parcelsColor As CadColor = New CadColor(7)
                Dim parcelsScaleLineType As Double = 1
                Dim parcelsWidthLine As Double = 20
                Dim parcelsGWidthLine As Double = 0
                Dim parcelsPSData As String = formCreateParcels.DataGridView1.Tag
                Dim deskParcel As String = formCreateParcels.ComboBox2.Text
                Dim boolCreateStyle As Boolean = FuncDecorPlane.FuncCreateStyleLinearObject(ActivDocument, formCreateParcels.templateXML, formCreateParcels.arrayObjectProperties, parcelsLayer, parcelsLineType, parcelsColor, parcelsScaleLineType, parcelsWidthLine, parcelsPSData, parcelsGWidthLine)
                Dim arrayData As String(,) = Nothing
                Dim countArrayData As Integer = 0
                If formCreateParcels.DataGridView1.RowCount > 0 Then
                    For i As Integer = 0 To formCreateParcels.DataGridView1.RowCount - 1
                        Dim nameFields As String = formCreateParcels.DataGridView1.Rows(i).Cells(0).Value
                        Dim strVal As String = formCreateParcels.DataGridView1.Rows(i).Cells(1).Value
                        ReDim Preserve arrayData(1, countArrayData)
                        arrayData(0, countArrayData) = nameFields
                        arrayData(1, countArrayData) = strVal
                        countArrayData += 1
                    Next i
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'выбираем объекты
                userCadView.SelectionSet.SelectObjectsAtScreen(Function(obj) TypeOf obj Is DwgPolyline, "Выборите полилинии для преобразования их в островки: ")
                For Each acEnt As Object In userCadView.SelectionSet
                    If TypeOf (acEnt) Is DwgPolyline Then
                        Dim plineParcels As DwgPolyline = acEnt
                        Dim fileObjectName As String = plineParcels.Drawing.Layers.Drawing.Filename
                        Dim nameUserFile As String = IO.Path.GetFileNameWithoutExtension(fileObjectName)
                        If Not (nameUserFile Like activUserModel) Then
                            plineParcels = plineParcels.Clone()
                            If IsNothing(ActivDocument) = False Then
                                ActivDocument.ActiveSpace.Add(plineParcels)
                            End If
                        End If
                        If plineParcels.Count > 2 Then
                            Dim boolDblVert As Boolean = MathFunction.FuncDeleteDubleVertexPline(plineParcels, 0.01)
                            If plineParcels.Closed = False Then
                                If parcelsPSData Like "PROJECT_PARCEL" Or parcelsPSData Like "TEMPORARY_PARCEL" Or parcelsPSData Like "PROJECT_ZONA" Or parcelsPSData Like "PROJECT_BUILDING" Then
                                    plineParcels.Closed = True
                                End If
                            End If
                        End If
                        plineParcels.Color = parcelsColor
                        plineParcels.Layer = parcelsLayer
                        plineParcels.Linetype = parcelsLineType
                        plineParcels.Lineweight = parcelsWidthLine
                        plineParcels.LinetypeScale = parcelsScaleLineType
                        plineParcels.Width = parcelsGWidthLine
                        Dim boolAreaLenght As Boolean = False
                        If IsArray(arrayData) = True Then
                            For i As Integer = 0 To arrayData.GetUpperBound(1)
                                If arrayData(0, i) Like "Area" Or arrayData(0, i) Like "AREA" Or arrayData(0, i) Like "area" Then
                                    Dim areaParcel As Double = Math.Abs(MathFunction.GetArea(plineParcels, 2))
                                    arrayData(1, i) = Math.Round(areaParcel, 0)
                                    boolAreaLenght = True
                                ElseIf arrayData(0, i) Like "Lenght" Or arrayData(0, i) Like "LENGHT" Or arrayData(0, i) Like "lenght" Then
                                    arrayData(1, i) = Math.Round(plineParcels.Length, 2)
                                    boolAreaLenght = True
                                ElseIf arrayData(0, i) Like "GLOBALID" Then
                                    arrayData(1, i) = globID
                                ElseIf arrayData(0, i) Like "LOCALID" Then
                                    arrayData(1, i) = rng.Next
                                End If
                            Next i
                            If boolAreaLenght = False Then
                                For i As Integer = 0 To arrayData.GetUpperBound(1)
                                    If arrayData(0, i) Like "KeyParameters" Then
                                        If plineParcels.Closed = True Then
                                            Dim areaParcel As Double = Math.Abs(MathFunction.GetArea(plineParcels, 2))
                                            Dim lenghParcel As Double = Math.Round(plineParcels.Length, 2)
                                            arrayData(1, i) = "Площадь: " & Math.Round(areaParcel, 0) & " кв.м. Периметр: " & lenghParcel & " м."
                                        Else
                                            Dim lenghParcel As Double = Math.Round(plineParcels.Length, 2)
                                            arrayData(1, i) = "Периметр: " & lenghParcel & " м."
                                        End If
                                    End If
                                Next i
                            End If
                            Dim boolWriteData As Boolean = FuncXRecords.FuncCreateXData(plineParcels, parcelsPSData, arrayData, formCreateParcels.templateXML)
                        End If
                    End If
                Next
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
            'сброс выделения
            userCadView.SelectionSet.Clear()
            userCadView.SelectionSet.SelectOneObjectAtScreen(Function(obj) TypeOf obj Is DwgPolyline, "Выборите объект для которого требуется создать островок: ")
            Dim localID As String = ""
            Dim nameTablePS As String = ""
            Dim parentObject As DwgPolyline = Nothing
            Dim parentSQRT As Double = 0
            Dim newParentSQRT As Double = 0
            Dim parentVertex As List(Of Vector2D) = New List(Of Vector2D)
            Dim countOstrov As Integer = 1
            Dim nameTablePSOstrov As String = "CADASTRAL_OSTROV"
            Dim arrayData As String(,) = Nothing
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

                    Dim arrayTablePs As String() = Nothing
                    Dim boolTablePs As Boolean = FuncXRecords.FuncReadTablesPS(parentObject, arrayTablePs)
                    If IsArray(arrayTablePs) = True Then
                        For i As Integer = 0 To arrayTablePs.Length - 1
                            Dim tempTable = arrayTablePs(i)
                            If Not (tempTable Like "CADASTRAL_OSTROV" Or tempTable Like "IS_SYSTEM") Then
                                Dim boolData As Boolean = FuncXRecords.FuncReadXData(parentObject, tempTable, arrayData)
                                localID = FuncXML.FuncReturnValueByArrayProperties("LOCALID", arrayData)
                                Exit For
                            End If
                        Next i
                    End If
                    'продолжаем если найден localID
                    If localID.Trim.Length > 2 Then
                        Dim arrayOstrov As String(,) = Nothing
                        If localID.Trim.Length > 1 Then
                            'вычисляем площадь заново
                            parentSQRT = Math.Abs(MathFunction.GetArea(parentObject, 2))
                            parentSQRT = Math.Abs(parentSQRT)
                            newParentSQRT = parentSQRT
                            'ищем уже имеющиеся островки у участка
                            Dim boolFindOstrov As Boolean = CadastralFunction.FuncFindCadastralOstrov(ActivDocument, parentObject, localID, arrayOstrov)
                            If IsArray(arrayOstrov) = True Then
                                For i As Integer = 0 To arrayOstrov.GetUpperBound(1)
                                    Dim sqrtOstrov As Double = Val(arrayOstrov(2, i))
                                    newParentSQRT = newParentSQRT - Math.Abs(sqrtOstrov)
                                Next i
                                'следующий номер островка
                                countOstrov = arrayOstrov(0, arrayOstrov.GetUpperBound(1)) + 1
                            End If
                        End If
                        Exit For
                    End If
                End If
            Next

            If localID.Trim.Length > 2 And parentObject.Count > 2 Then
                userCadView.SelectionSet.Clear()
                userCadView.SelectionSet.SelectObjectsAtScreen(Function(obj) TypeOf obj Is DwgPolyline, "Выборите полилинии для преобразования их в островки: ")
                For Each acEnt As Object In userCadView.SelectionSet
                    If TypeOf (acEnt) Is DwgPolyline Then
                        Dim plineOstrov As DwgPolyline = acEnt
                        If Not (plineOstrov.Drawing.Layers.Drawing.Filename = parentObject.Drawing.Layers.Drawing.Filename) Then
                            plineOstrov = plineOstrov.Clone
                            parentObject.Drawing.ActiveSpace.Add(plineOstrov)
                        End If
                        Dim boolFindTablePS As Boolean = FuncXRecords.FuncFindTablePS(acEnt, nameTablePSOstrov)
                        If boolFindTablePS = False Then
                            Dim boolDelVertex As Boolean = MathFunction.FuncDeleteDubleVertexPline(plineOstrov, 0.01)
                            If plineOstrov.Count > 2 Then
                                If plineOstrov.Closed = False Then plineOstrov.Closed = True
                                Dim boolCrossVertex As Boolean = False
                                For i As Integer = 0 To plineOstrov.Count - 1
                                    Dim buildVertexOstrov As BugleVector2D = plineOstrov.Item(i)
                                    Dim ostrovVert As Vector2D = New Vector2D(buildVertexOstrov.Vertex.X, buildVertexOstrov.Vertex.Y)
                                    'boolCrossVertex = Topomatic.Cad.Foundation.CadLibrary.PosInPolygon(ostrovVert, parentVertex, False)
                                Next i
                                If boolCrossVertex = False Then
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'записываем расширенные данные
                                    Dim arrayRec As String(,) = Nothing
                                    ReDim Preserve arrayRec(1, 3)
                                    arrayRec(0, 0) = "LABEL"
                                    arrayRec(1, 0) = "Внутренний островок"
                                    arrayRec(0, 1) = "ConditionalNumber"
                                    arrayRec(1, 1) = countOstrov
                                    arrayRec(0, 2) = "IDKOUNTER"
                                    arrayRec(1, 2) = localID
                                    arrayRec(0, 3) = "LOCALID"
                                    arrayRec(1, 3) = rng.Next
                                    'записываем массив в островок
                                    If IsArray(arrayRec) = True Then
                                        Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXData(plineOstrov, nameTablePSOstrov, arrayRec, templateXML)
                                        If boolInsDataPS = True Then
                                            plineOstrov.Layer = parentObject.Layer
                                            plineOstrov.Color = parentObject.Color
                                            plineOstrov.Linetype = parentObject.Linetype
                                            plineOstrov.LinetypeScale = parentObject.LinetypeScale
                                            plineOstrov.Lineweight = parentObject.Lineweight
                                        End If
                                    End If
                                    Dim sqrtOstrov As Double = MathFunction.GetArea(plineOstrov, 2)
                                    sqrtOstrov = Math.Abs(sqrtOstrov)
                                    newParentSQRT = newParentSQRT - sqrtOstrov
                                    countOstrov += 1
                                End If
                            End If
                        End If
                    End If
                Next
                'меняем исходный участок
                If newParentSQRT < parentSQRT Then
                    'записываем расширенные данные
                    Dim arrayRec As String(,) = Nothing
                    ReDim Preserve arrayRec(1, 0)
                    arrayRec(0, 0) = "Area"
                    arrayRec(1, 0) = Math.Round(newParentSQRT, 0)
                    If IsArray(arrayRec) = True Then
                        Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXData(parentObject, nameTablePS, arrayRec, templateXML)
                    End If
                    '============================================================================================================
                    'обновляем метку
                    'ищем метку
                    Dim acLeader As DwgLeader = CadastralFunction.FuncFindCadastralLabel(ActivDocument, localID)
                    If IsNothing(acLeader) = False Then
                        Dim boolUpdateLabel As Boolean = CadastralFunction.FuncUpdateCadastralLabel(ActivDocument, acLeader, arrayData)
                    End If
                End If
            End If
        End Sub
        '================================================================================================================================
        'добавить контур
        <cmd("CreateCounterObject")>
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
            'находим каталог с файлами (c:\Users\mzaharov\AppData\Roaming\Civil3DToolsUtility\TOPOTemplates\)
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
            'сброс выделения
            userCadView.SelectionSet.Clear()
            userCadView.SelectionSet.SelectOneObjectAtScreen(Function(obj) TypeOf obj Is DwgPolyline, "Выборите исходный полигон к котому следует добавить контур: ")
            Dim localID As String = ""
            Dim globalID As String = ""
            Dim fullNumberParcel As String = "1"
            Dim numberParcel As String = "1"
            Dim numberCounter As Integer = 0
            Dim parentObject As DwgEntity = Nothing
            Dim arrayData As String(,) = Nothing
            Dim nameTablesPS As String = ""
            Dim boolFirstNumberCounter As Boolean = False
            Dim acParentLeader As DwgLeader = Nothing
            Dim tableLeaderPs As String = "CADASTRAL_OBJECT"
            For Each acEnt As Object In userCadView.SelectionSet
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                parentObject = acEnt
                Dim arrayTables As String() = Nothing
                Dim boolFindTables As Boolean = FuncXRecords.FuncReadTablesPS(acEnt, arrayTables)
                If IsArray(arrayTables) = True Then
                    For i As Integer = 0 To arrayTables.Length - 1
                        nameTablesPS = arrayTables(i)
                        If Not (nameTablesPS Like "CADASTRAL_OSTROV" Or nameTablesPS Like "IS_SYSTEM") Then
                            Dim boolData As Boolean = FuncXRecords.FuncReadXData(parentObject, nameTablesPS, arrayData)
                            If IsArray(arrayData) = True Then
                                localID = FuncXML.FuncReturnValueByArrayProperties("LOCALID", arrayData)
                                If localID.Trim.Length > 1 Then
                                    Exit For
                                End If
                            End If
                            Exit For
                        End If
                    Next i
                End If
            Next
            '============================================================================================
            If localID.Trim.Length > 1 Then
                globalID = FuncXML.FuncReturnValueByArrayProperties("GLOBALID", arrayData)
                Dim arrayCounterParcels As String(,) = Nothing
                If globalID.Trim.Length < 3 Then
                    globalID = rng.Next
                    numberCounter = 1
                Else
                    Dim boolFindCounters As Boolean = CadastralFunction.FuncFindCadastralCounter(ActivDocument, globalID, arrayCounterParcels)
                End If
                'номер контура
                If IsArray(arrayCounterParcels) = True Then
                    numberCounter = Val(arrayCounterParcels(0, arrayCounterParcels.GetUpperBound(1))) + 1
                    If numberCounter > 0 Then
                        boolFirstNumberCounter = True
                    End If
                End If
                'полный номер участка
                fullNumberParcel = FuncXML.FuncReturnValueByArrayProperties("ConditionalNumber", arrayData)
                If fullNumberParcel.Trim.Length > 0 Then
                    numberParcel = FuncFormatZn.FuncFindFullConditionalNumberCadastralObject(fullNumberParcel)
                    If numberParcel.Trim.Length = 0 Then
                        numberParcel = 1
                    End If
                End If
                If numberCounter = 0 Then
                    numberCounter = FuncFormatZn.FuncFindNumberCounterCadastralObject(fullNumberParcel)
                    If numberCounter = 0 Then
                        numberCounter = 1
                    Else
                        boolFirstNumberCounter = True
                    End If
                End If
                'запрашиваем номер контура
                Dim rez As GetPointResult = CadCursors.GetInteger(userCadView, numberCounter, "Следующий номер контура (" & numberCounter & "):")

                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'записываем расширенные данные
                If boolFirstNumberCounter = False Then
                    Dim newArrayData As String(,) = Nothing
                    Dim countNewArrayData As Integer = 0
                    For i As Integer = 0 To arrayData.GetUpperBound(1)
                        Dim nameField As String = arrayData(0, i)
                        If nameField Like "GLOBALID" Then
                            If arrayData(1, i).Trim.Length < 3 Then
                                ReDim Preserve newArrayData(1, countNewArrayData)
                                newArrayData(0, countNewArrayData) = "GLOBALID"
                                newArrayData(1, countNewArrayData) = globalID
                                countNewArrayData += 1
                            End If
                        ElseIf nameField Like "ConditionalNumber" Then
                            ReDim Preserve newArrayData(1, countNewArrayData)
                            newArrayData(0, countNewArrayData) = "ConditionalNumber"
                            newArrayData(1, countNewArrayData) = numberParcel & "(" & numberCounter & ")"
                            countNewArrayData += 1
                        End If
                    Next i
                    If IsArray(newArrayData) = True Then
                        Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXData(parentObject, nameTablesPS, newArrayData, templateXML)
                        'обновляем метку
                        acParentLeader = CadastralFunction.FuncFindCadastralLabel(ActivDocument, localID)
                        If IsNothing(acParentLeader) = False Then
                            Dim boolUpdateLeader As Boolean = CadastralFunction.FuncUpdateCadastralLabel(ActivDocument, acParentLeader, newArrayData)
                        End If
                        numberCounter += 1
                    End If
                Else
                    acParentLeader = CadastralFunction.FuncFindCadastralLabel(ActivDocument, localID)
                End If
            Else
                MsgBox("Выбранный объект не преобразован в Участок. Макро прервано!!!!")
                Exit Sub
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'выбираем новые объекты
            userCadView.SelectionSet.Clear()
            userCadView.SelectionSet.SelectObjectsAtScreen(Function(obj) TypeOf obj Is DwgPolyline, "Выборите контура для присоединения: ")
            For Each acEnt As Object In userCadView.SelectionSet
                If TypeOf (acEnt) Is DwgPolyline Then
                    Dim plineCounter As DwgPolyline = acEnt
                    If Not (plineCounter.Drawing.Layers.Drawing.Filename Like parentObject.Drawing.Layers.Drawing.Filename) Then
                        plineCounter = plineCounter.Clone
                        parentObject.Drawing.ActiveSpace.Add(plineCounter)
                    End If
                    Dim arrayDataCounter As String(,) = Nothing
                    Dim boolData As Boolean = FuncXRecords.FuncReadXData(plineCounter, nameTablesPS, arrayDataCounter)
                    Dim localIDCounter As String = ""
                    If IsArray(arrayDataCounter) = True Then
                        For i As Integer = 0 To arrayDataCounter.GetUpperBound(1)
                            Dim nameField As String = arrayDataCounter(0, i)
                            If nameField Like "GLOBALID" Then
                                arrayDataCounter(1, i) = globalID
                            ElseIf nameField Like "ConditionalNumber" Then
                                arrayDataCounter(1, i) = numberParcel & "(" & numberCounter & ")"
                            ElseIf nameField Like "LOCALID" Then
                                localIDCounter = arrayDataCounter(1, i)
                            End If
                        Next i
                        'записываем данные
                        Dim boolInsDataPSCounter As Boolean = FuncXRecords.FuncCreateXData(plineCounter, nameTablesPS, arrayDataCounter, templateXML)
                        If boolInsDataPSCounter = True Then
                            'обновляем метку
                            Dim acLeaderCounter As DwgLeader = CadastralFunction.FuncFindCadastralLabel(ActivDocument, localIDCounter)
                            If IsNothing(acLeaderCounter) = False Then
                                Dim boolUpdateLeader As Boolean = CadastralFunction.FuncUpdateCadastralLabel(ActivDocument, acLeaderCounter, arrayDataCounter)
                            End If
                            plineCounter.Layer = parentObject.Layer
                            plineCounter.Color = parentObject.Color
                            plineCounter.Linetype = parentObject.Linetype
                            plineCounter.LinetypeScale = parentObject.LinetypeScale
                            plineCounter.Lineweight = parentObject.Lineweight
                            numberCounter += 1
                        End If
                    Else
                        'записываем расширенные данные
                        Dim boolDelVertex As Boolean = MathFunction.FuncDeleteDubleVertexPline(plineCounter, 0.01)
                        localIDCounter = rng.Next
                        For i As Integer = 0 To arrayData.GetUpperBound(1)
                            Dim nameField As String = arrayData(0, i)
                            If nameField Like "LOCALID" Then
                                arrayData(1, i) = localIDCounter
                            ElseIf nameField Like "ConditionalNumber" Then
                                arrayData(1, i) = numberParcel & "(" & numberCounter & ")"
                            ElseIf nameField Like "Area" Then
                                Dim sqrtCounters As Double = MathFunction.GetArea(plineCounter, 2)
                                sqrtCounters = Math.Abs(sqrtCounters)
                                arrayData(1, i) = Math.Round(sqrtCounters, 0)
                            ElseIf nameField Like "GLOBALID" Then
                                If arrayData(1, i).Trim.Length < 3 Then
                                    arrayData(1, i) = globalID
                                End If
                            End If
                        Next i
                        Dim boolInsDataPSCounter As Boolean = FuncXRecords.FuncCreateXData(plineCounter, nameTablesPS, arrayData, templateXML)
                        If boolInsDataPSCounter = True Then
                            plineCounter.Layer = parentObject.Layer
                            plineCounter.Color = parentObject.Color
                            plineCounter.Linetype = parentObject.Linetype
                            plineCounter.LinetypeScale = parentObject.LinetypeScale
                            plineCounter.Lineweight = parentObject.Lineweight
                            numberCounter += 1
                            If IsNothing(acParentLeader) = False Then
                                Dim centerPolygon As Vector2D = New Vector2D(-1, -1)
                                If plineCounter.Closed = True Then
                                    Dim listVertex As List(Of BugleVector2D) = plineCounter.ToList
                                    Dim arrayVertex As Double(,) = Nothing
                                    For k As Integer = 0 To listVertex.Count - 1
                                        ReDim Preserve arrayVertex(1, k)
                                        arrayVertex(0, k) = listVertex.Item(k).Vertex.X
                                        arrayVertex(1, k) = listVertex.Item(k).Vertex.Y
                                    Next k
                                    centerPolygon = MathFunction.FuncFindCrossCenterPointPolyline(arrayVertex)
                                End If
                                Dim pointInsText As Vector2D = plineCounter.Bounds.Max
                                Dim acLeaderCounter As DwgLeader = acParentLeader.Clone
                                ActivDocument.ActiveSpace.Add(acLeaderCounter)
                                acLeaderCounter.Item(0) = centerPolygon
                                acLeaderCounter.Position = pointInsText
                                Dim noteLabel As String = FuncXRecords.FuncReadValueXData(acParentLeader, "NOTE")
                                Dim boolUpdateLeader As Boolean = CadastralFunction.FuncUpdateCadastralLabel(ActivDocument, acLeaderCounter, arrayData)
                                Dim arrayRec As String(,) = Nothing
                                ReDim Preserve arrayRec(1, 0)
                                arrayRec(0, 0) = "IDRNG"
                                arrayRec(1, 0) = localIDCounter
                                If IsArray(arrayRec) = True Then
                                    Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXData(acLeaderCounter, tableLeaderPs, arrayRec, templateXML)
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        End Sub
        '================================================================================================================================
        'заменить контур
        <cmd("ReplaseParcels")>
        Public Sub ReplaseParcels()
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
            'сброс выделения
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            userCadView.SelectionSet.Clear()
            userCadView.SelectionSet.SelectOneObjectAtScreen(Function(obj) TypeOf obj Is Object, "Выборите исходный полигон который требуется заменить: ")
            Dim parentParcel As DwgPolyline = Nothing
            Dim parentDict As DwgDictionary = Nothing
            Dim arrayTables As String() = Nothing
            For Each acEnt As Object In userCadView.SelectionSet
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                If TypeOf acEnt Is DwgPolyline Then
                    parentParcel = acEnt
                    parentDict = parentParcel.GetExtensionDictionary()
                    If IsNothing(parentDict) = False Then
                        Dim boolFindTables As Boolean = FuncXRecords.FuncReadTablesPS(parentParcel, arrayTables)
                    End If
                    Exit For
                End If
            Next

            userCadView.SelectionSet.Clear()
            userCadView.SelectionSet.SelectObjectsAtScreen(Function(obj) TypeOf obj Is Object, "Выборите новый полигон: ")
            For Each acEnt As Object In userCadView.SelectionSet
                If TypeOf (acEnt) Is DwgPolyline Then
                    Dim plineCounter As DwgPolyline = acEnt
                    If IsNothing(parentDict) = False Then
                        If IsNothing(plineCounter.GetExtensionDictionary) = True Then
                            plineCounter.CreateExtensionDictionary()
                        End If
                        Dim userDict As DwgDictionary = plineCounter.GetExtensionDictionary()
                        If IsArray(arrayTables) = True Then
                            For i As Integer = 0 To arrayTables.Count - 1
                                Dim arrayData As String(,) = Nothing
                                Dim nameTablePS As String = arrayTables(i)
                                Dim boolReadXRec As Boolean = FuncXRecords.FuncReadXData(parentParcel, nameTablePS, arrayData)
                                If IsArray(arrayData) = True Then
                                    Dim boolWriteXData As Boolean = FuncXRecords.FuncCreateXData(plineCounter, nameTablePS, arrayData, templateXML)
                                End If
                            Next i
                        End If
                    End If
                    plineCounter.Layer = parentParcel.Layer
                    plineCounter.Color = parentParcel.Color
                    plineCounter.Linetype = parentParcel.Linetype
                    plineCounter.LinetypeScale = parentParcel.LinetypeScale
                    plineCounter.Lineweight = parentParcel.Lineweight
                End If
            Next
        End Sub
        '===============================================================================================================================
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '2 Нумеровать полигон
        <cmd("NumberPolygon")>
        Public Sub NumberPolygon()
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
                layerPoint = FuncDecorPlane.FuncCreateLayerByXML(ActivDocument, templateXML, nameLayerPoint)
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
            If formInsertPoint.RadioButton1.Checked = True Then
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'вручную
                Do

                    CadView.SelectionSet.Clear()
                    CadView.SelectionSet.SelectOneObjectAtScreen(Function(obj) TypeOf obj Is DwgPolyline, "Выборите вершину полигона: ")
                    If CadView.SelectionSet.Count = 0 Then Exit Do
                    For Each acEnt As Object In cadWiever.SelectionSet
                        If TypeOf acEnt Is DwgPolyline Then
                            Dim acPoly As DwgPolyline = acEnt
                            If acPoly.HasExtensionDictionary = True Then
                                Dim localIDCounter As String = FuncXRecords.FuncReadValueXData(acPoly, "LOCALID")
                                If localIDCounter.Trim.Length > 2 Then
                                    Dim category As String = FuncXRecords.FuncReadValueXData(acPoly, "Category")
                                    Dim delta_geopoint As Double = CadastralFunction.FuncReturnDeltaGeopointByCategory(category)
                                    Dim v3d As Vector3D = cadWiever.LastPoint
                                    Dim v2d As Vector2D = New Vector2D(v3d.X, v3d.Y)
                                    Dim l As Double = 9999999
                                    Dim numVertex As Integer = 0
                                    For i As Integer = 0 To acPoly.Count - 1
                                        Dim buildVector As Vector2D = acPoly.Item(i).Vertex
                                        If (buildVector - v2d).Length < l Then
                                            numVertex = i
                                            l = (buildVector - v2d).Length
                                        End If
                                    Next
                                    'вставляем блок
                                    Dim posInsertBlk As Vector2D = acPoly.Item(numVertex).Vertex
                                    If IsArray(arrayPointCollections) = True Then
                                        Dim idFindPoint As String = CadastralFunction.FuncFindCadastralPointToVertex(ActivDocument, posInsertBlk, arrayPointCollections, 0.01)
                                        If idFindPoint.Trim.Length > 0 Then
                                            Continue For
                                        End If
                                    End If
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
                                        Dim countColl As Integer = 0
                                        If IsArray(arrayPointCollections) = True Then
                                            countColl = arrayPointCollections.GetUpperBound(1) + 1
                                        End If

                                        ReDim Preserve arrayPointCollections(3, countColl)
                                        arrayPointCollections(0, countColl) = posInsertBlk.X
                                        arrayPointCollections(1, countColl) = posInsertBlk.Y
                                        arrayPointCollections(2, countColl) = insertBlk.ObjectID
                                        arrayPointCollections(3, countColl) = localIDCounter
                                        startNumber += 1
                                    End If
                                End If
                            End If
                        End If
                    Next
                Loop
            Else
                '2.\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'выбираем объекты
                CadView.SelectionSet.SelectObjectsAtScreen(Function(obj) TypeOf obj Is DwgPolyline, "Выборите полилинии: ")
                '=======================================================================================================
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                Dim arrayObject As String(,) = Nothing '0-Номер участка,1-handle,2-GLobalID,2-LocalID,3-ha
                Dim countArrayObject As Integer = 0
                Dim labelParcel As String = "Земельный участок"
                For Each acEnt As Object In cadWiever.SelectionSet
                    If TypeOf acEnt Is DwgPolyline Then
                        Dim acPoly As DwgPolyline = acEnt
                        If acPoly.HasExtensionDictionary = True Then
                            Dim LABEL As String = FuncXRecords.FuncReadValueXData(acPoly, "LABEL")
                            If LABEL Like "Внутренний островок" Then
                                Continue For
                            Else
                                Dim conditionalParcel As String = FuncXRecords.FuncReadValueXData(acPoly, "ConditionalNumber")
                                Dim pos1 As Integer = conditionalParcel.IndexOf("(")
                                Dim numberParcel As String = FuncFormatZn.FuncFindFullConditionalNumberCadastralObject(conditionalParcel)
                                Dim globalID As String = FuncXRecords.FuncReadValueXData(acPoly, "GLOBALID")
                                Dim localID As String = FuncXRecords.FuncReadValueXData(acPoly, "LOCALID")
                                Dim boolFindGLOBALID As Boolean = False
                                If globalID.Trim.Length > 2 Then
                                    boolFindGLOBALID = MathFunction.FuncFindValueToArray(globalID, 2, arrayObject)
                                End If
                                If boolFindGLOBALID = False Then
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
                    End If
                Next
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
                        Dim numberParcel As Integer = Val(arrayObject(0, i))
                        If globalID.Trim.Length > 2 Then
                            'ищем входящие контура
                            Dim arrayDataCounters As String(,) = Nothing
                            Dim countCounters As Integer = CadastralFunction.FuncFindCadastralCounter(ActivDocument, numberParcel, arrayDataCounters)
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
                                                        'проверяем на наличие блока в вершине
                                                        If formInsertPoint.CheckBox2.Checked = True Then
                                                            If IsArray(arrayPointCollections) = True Then
                                                                Dim idFindPoint As String = CadastralFunction.FuncFindCadastralPointToVertex(ActivDocument, posInsertBlk, arrayPointCollections, 0.01)
                                                                If idFindPoint.Trim.Length > 0 Then
                                                                    Continue For
                                                                End If
                                                            End If
                                                        End If
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
                                                            Dim countColl As Integer = 0
                                                            If IsArray(arrayPointCollections) = True Then
                                                                countColl = arrayPointCollections.GetUpperBound(1) + 1
                                                            End If
                                                            ReDim Preserve arrayPointCollections(3, countColl)
                                                            arrayPointCollections(0, countColl) = x
                                                            arrayPointCollections(1, countColl) = y
                                                            arrayPointCollections(2, countColl) = insertBlk.ObjectID
                                                            arrayPointCollections(3, countColl) = localIDCounter
                                                            startNumber += 1
                                                        End If
                                                    Next k
                                                    'ищем внутренние островки
                                                    Dim arrayDataOstrov As String(,) = Nothing
                                                    Dim countOstrov As Integer = CadastralFunction.FuncFindCadastralOstrov(ActivDocument, acPoly, localIDCounter, arrayDataOstrov)
                                                    If IsArray(arrayDataOstrov) = True Then
                                                        For k As Integer = 0 To arrayDataOstrov.GetUpperBound(1)
                                                            Dim hgOstrov As String = arrayDataOstrov(1, k)
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
                                                                            If formInsertPoint.CheckBox2.Checked = True Then
                                                                                If IsArray(arrayPointCollections) = True Then
                                                                                    Dim idFindPoint As String = CadastralFunction.FuncFindCadastralPointToVertex(ActivDocument, posInsertBlk, arrayPointCollections, 0.01)
                                                                                    If idFindPoint.Trim.Length > 0 Then
                                                                                        Continue For
                                                                                    End If
                                                                                End If
                                                                            End If
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
                                                                                arrayRec(1, 5) = arrayDataOstrov(3, k)
                                                                                If IsArray(arrayRec) = True Then
                                                                                    Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXData(insertBlk, PSTable, arrayRec, templateXML)
                                                                                End If
                                                                                Dim countColl As Integer = 0
                                                                                If IsArray(arrayPointCollections) = True Then
                                                                                    countColl = arrayPointCollections.GetUpperBound(1) + 1
                                                                                End If
                                                                                ReDim Preserve arrayPointCollections(3, countColl)
                                                                                arrayPointCollections(0, countColl) = x
                                                                                arrayPointCollections(1, countColl) = y
                                                                                arrayPointCollections(2, countColl) = insertBlk.ObjectID
                                                                                arrayPointCollections(3, countColl) = arrayDataOstrov(3, k)
                                                                                startNumber += 1
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
                                        'проверяем на наличие блока
                                        If formInsertPoint.CheckBox2.Checked = True Then
                                            If IsArray(arrayPointCollections) = True Then
                                                Dim idFindPoint As String = CadastralFunction.FuncFindCadastralPointToVertex(ActivDocument, posInsertBlk, arrayPointCollections, 0.01)
                                                If idFindPoint.Trim.Length > 0 Then
                                                    Continue For
                                                End If
                                            End If
                                        End If
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
                                            arrayRec(1, 5) = arrayObject(3, i)
                                            If IsArray(arrayRec) = True Then
                                                Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXData(insertBlk, PSTable, arrayRec, templateXML)
                                            End If
                                            Dim countColl As Integer = 0
                                            If IsArray(arrayPointCollections) = True Then
                                                countColl = arrayPointCollections.GetUpperBound(1) + 1
                                            End If

                                            ReDim Preserve arrayPointCollections(3, countColl)
                                            arrayPointCollections(0, countColl) = x
                                            arrayPointCollections(1, countColl) = y
                                            arrayPointCollections(2, countColl) = insertBlk.ObjectID
                                            arrayPointCollections(3, countColl) = arrayObject(3, i)
                                            startNumber += 1
                                        End If
                                    Next k
                                    'ищем внутренние островки
                                    Dim arrayDataOstrov As String(,) = Nothing
                                    Dim localID As String = FuncXRecords.FuncReadValueXData(acPoly, "LOCALID")
                                    Dim countOstrov As Integer = CadastralFunction.FuncFindCadastralOstrov(ActivDocument, acPoly, localID, arrayDataOstrov)
                                    If IsArray(arrayDataOstrov) = True Then
                                        For k As Integer = 0 To arrayDataOstrov.GetUpperBound(1)
                                            Dim hgOstrov As String = arrayDataOstrov(1, k)
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
                                                            'проверяем на наличие блока
                                                            If formInsertPoint.CheckBox2.Checked = True Then
                                                                If IsArray(arrayPointCollections) = True Then
                                                                    Dim idFindPoint As String = CadastralFunction.FuncFindCadastralPointToVertex(ActivDocument, posInsertBlk, arrayPointCollections, 0.01)
                                                                    If idFindPoint.Trim.Length > 0 Then
                                                                        Continue For
                                                                    End If
                                                                End If
                                                            End If
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
                                                                arrayRec(1, 5) = arrayDataOstrov(3, k)
                                                                If IsArray(arrayRec) = True Then
                                                                    Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXData(insertBlk, PSTable, arrayRec, templateXML)
                                                                End If
                                                                Dim countColl As Integer = arrayPointCollections.GetUpperBound(1) + 1
                                                                ReDim Preserve arrayPointCollections(3, countColl)
                                                                arrayPointCollections(0, countColl) = x
                                                                arrayPointCollections(1, countColl) = y
                                                                arrayPointCollections(2, countColl) = insertBlk.ObjectID
                                                                arrayPointCollections(3, countColl) = arrayDataOstrov(3, k)
                                                                startNumber += 1
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
                layerPoint = FuncDecorPlane.FuncCreateLayerByXML(ActivDocument, templateXML, nameLayerPoint)
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
                FuncDecorPlane.FuncCreateStyleLeaderObject(ActivDocument, templateXML, arrayPropertyMLeader, nameCreateObject, objectLayer, objectUserTypeLine, objectColor, scaleUserTypeLine, widthUserTypeLine, TextStyleUser, hText, nameArrow, arrowSize, tablePs, note)
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
                                            Dim styleTextUser As DwgStyle = FuncStyles.FuncReturnTextStyle(ActivDocument, TextStyleUser)
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
    End Class
End Namespace
