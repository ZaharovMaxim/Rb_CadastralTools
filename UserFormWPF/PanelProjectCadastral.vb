Imports System.Windows.Forms
Imports System.Drawing
Imports System.Xml
Imports Topomatic.Dwg.Entities
Imports Topomatic.ApplicationPlatform.Plugins
Imports Topomatic.Cad.View
Imports Topomatic.Dwg.Layer
Imports Topomatic.Dwg
Imports Topomatic.Cad.Foundation
Imports Topomatic.Sfc
Imports Topomatic.FoundationClasses
Imports Topomatic.Visualization
Imports Topomatic.Visualization.Runtime
Imports Topomatic.Alg
Imports Topomatic.Cad.View.Hints
Imports System.Windows.Controls
Imports System.IO
Imports Topomatic.ApplicationPlatform.Core
Imports Topomatic.ApplicationPlatform
Imports Topomatic.Arrangements
Imports Topomatic.Alg.Road.Core
Imports Topomatic.Visualization.Constructions
Imports System.Windows.Media.Animation
Imports Topomatic.FoundationClasses.Undo
Imports Topomatic.Alg.Survey.Core
Imports Topomatic.Pipes.Layers.Caches.Plan.Segments.SegmentPlanCache.DrawingData
Imports Topomatic.Alg.Prf
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Net
Imports System.Runtime.Remoting.Metadata.W3cXsd2001
Imports Microsoft.Office.Interop
Imports System.Windows
Imports Topomatic.Acax.Export
Imports Acax.Import
Imports Topomatic.Alg.Bridges
Imports System.Windows.Markup
Imports System.Reflection.Emit
Imports Microsoft.Office.Interop.Word
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.ListView
Imports System.Windows.Documents
Imports Topomatic.Cad.View.Design


Public Class PanelProjectCadastral
    Public cadViewPanel As CadView = Nothing
    Public activDocumentPanel As Topomatic.Dwg.Drawing = Nothing
    Public generalDirectoryPrograms As String = ""
    Public putchTemplatesXML As String = ""
    Public putchTemplateDwg As String = ""
    Public putchProjectDirectories As String = ""
    Public putchProjectFileData As String = ""

    Public Sub New()

        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()



        ' Добавить код инициализации после вызова InitializeComponent().

    End Sub

    Private Sub PanelProjectCadastral_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If IsNothing(activDocumentPanel) = False Then
            Dim fileNameProject As String = activDocumentPanel.Filename
            Dim directoryNameProject As String = IO.Path.GetDirectoryName(fileNameProject)
            Dim parentDirectoryProject As DirectoryInfo = IO.Directory.GetParent(directoryNameProject)
            If Directory.Exists(parentDirectoryProject.FullName & "\LandCadastral") = True Then
                putchProjectDirectories = parentDirectoryProject.FullName & "\LandCadastral"
                If File.Exists(putchProjectDirectories & "\LandProjectData.xml") = True Then
                    putchProjectFileData = putchProjectDirectories & "\LandProjectData.xml"
                    Me.Button1.Enabled = False
                End If
            End If
        End If
    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub


    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'создать новый проект
    Public Sub CreateCadastralProject()
        If IsNothing(activDocumentPanel) = False Then
            Dim fileNameProject As String = activDocumentPanel.Filename
            Dim directoryNameProject As String = IO.Path.GetDirectoryName(fileNameProject)
            Dim nameProject As String = IO.Path.GetFileNameWithoutExtension(fileNameProject)
            Dim userProject As ModelProject = ApplicationHost.Current.ActiveProject
            If IsNothing(userProject) = False Then
                Dim folderProject As String = PluginCoreOps.FindModelPathId(PluginCoreOps.CreateFolder(New String() {"Модели", "LandCadastral"}))
                directoryNameProject = directoryNameProject & "\LandCadastral"
                Using userXMLFile As XmlWriter = XmlWriter.Create(directoryNameProject & "\LandProjectData.xml")
                    userXMLFile.WriteStartDocument()
                    userXMLFile.WriteStartElement("LandPlanningDocumentation")
                    userXMLFile.WriteAttributeString("guid", Guid.NewGuid.ToString)
                    userXMLFile.WriteAttributeString("version", "01")
                    '1 уровень. Настройки проекта

                    userXMLFile.WriteStartElement("LandProjectInfo")
                    userXMLFile.WriteAttributeString("KeyParameters", "")
                    userXMLFile.WriteAttributeString("Note", "")
                    userXMLFile.WriteEndElement()
                    '2 уровень. Данные о зонах
                    userXMLFile.WriteStartElement("Object_placement_zones")
                    userXMLFile.WriteAttributeString("Description", "Зоны размещения линейного объекта")
                    userXMLFile.WriteAttributeString("Note", "")
                    userXMLFile.WriteEndElement()
                    '3 уровень. красные линии
                    userXMLFile.WriteStartElement("Red_line")
                    userXMLFile.WriteAttributeString("Description", "Красные линии")
                    userXMLFile.WriteAttributeString("Note", "")
                    userXMLFile.WriteEndElement()
                    '4 уровень. Образуемые земельные участки
                    userXMLFile.WriteStartElement("Project_parcel")
                    userXMLFile.WriteAttributeString("Description", "Образуемые земельные участки")
                    userXMLFile.WriteAttributeString("Note", "")
                    '4.1 Изъятие
                    userXMLFile.WriteStartElement("Confiscation_parcel")
                    userXMLFile.WriteAttributeString("Description", "Земельные участки изымаемые")
                    userXMLFile.WriteAttributeString("Note", "")
                    '4.1.1 Раздел с исходным ЗУ
                    userXMLFile.WriteStartElement("Division_with_original_parcel")
                    userXMLFile.WriteAttributeString("Description", "Раздел с исходным земельным участком")
                    userXMLFile.WriteAttributeString("Note", "")
                    userXMLFile.WriteEndElement()
                    '4.1.2 Раздел ЗУ
                    userXMLFile.WriteStartElement("Division_parcel")
                    userXMLFile.WriteAttributeString("Description", "Раздел участка")
                    userXMLFile.WriteAttributeString("Note", "")
                    userXMLFile.WriteEndElement()
                    '4.1.3 Без проведения казастровых работ
                    userXMLFile.WriteStartElement("Without_cadastral_works")
                    userXMLFile.WriteAttributeString("Description", "Без проведения кадастровых работ")
                    userXMLFile.WriteAttributeString("Note", "")
                    userXMLFile.WriteEndElement()
                    userXMLFile.WriteEndElement()
                    '4.2 Резервирование
                    userXMLFile.WriteStartElement("Reservation_parcel")
                    userXMLFile.WriteAttributeString("Description", "Земельные участки резервируемые")
                    userXMLFile.WriteAttributeString("Note", "")
                    '4.2.1 Раздел с исходным ЗУ
                    userXMLFile.WriteStartElement("Division_with_original_parcel")
                    userXMLFile.WriteAttributeString("Description", "Раздел с исходным земельным участком")
                    userXMLFile.WriteAttributeString("Note", "")
                    userXMLFile.WriteEndElement()
                    '4.2.2 Раздел ЗУ
                    userXMLFile.WriteStartElement("Division_parcel")
                    userXMLFile.WriteAttributeString("Description", "Раздел участка")
                    userXMLFile.WriteAttributeString("Note", "")
                    userXMLFile.WriteEndElement()
                    '4.2.3 Без проведения казастровых работ
                    userXMLFile.WriteStartElement("Without_cadastral_works")
                    userXMLFile.WriteAttributeString("Description", "Без проведения кадастровых работ")
                    userXMLFile.WriteAttributeString("Note", "")
                    userXMLFile.WriteEndElement()

                    userXMLFile.WriteEndElement()
                    '5 Сервитуты/Публичные сервитуты
                    userXMLFile.WriteStartElement("Public_servitude")
                    userXMLFile.WriteAttributeString("Description", "Сервитуты/Публичные сервитуты")
                    userXMLFile.WriteAttributeString("Note", "")
                    '5.1 Сервитут публичный
                    userXMLFile.WriteStartElement("Public_servitude")
                    userXMLFile.WriteAttributeString("Description", "Публичный сервитут")
                    userXMLFile.WriteAttributeString("Note", "")
                    userXMLFile.WriteEndElement()
                    '5.2 Сервитут 
                    userXMLFile.WriteStartElement("Servitude")
                    userXMLFile.WriteAttributeString("Description", "Сервитут")
                    userXMLFile.WriteAttributeString("Note", "")
                    userXMLFile.WriteEndElement()
                    '5.3 Сервитут 
                    userXMLFile.WriteStartElement("Land_part")
                    userXMLFile.WriteAttributeString("Description", "Часть участка")
                    userXMLFile.WriteAttributeString("Note", "")
                    userXMLFile.WriteEndElement()
                    userXMLFile.WriteEndElement()







                    userXMLFile.WriteEndElement()
                    userXMLFile.WriteEndDocument()
                End Using

            End If

        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'создать новый проект
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        CreateCadastralProject()
    End Sub
End Class
