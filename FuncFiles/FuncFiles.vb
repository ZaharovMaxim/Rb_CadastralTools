Imports System.IO
Imports Topomatic.Alg
Imports Topomatic.Alg.Offsets
Imports Topomatic.Cad.Foundation
Imports Topomatic.Cad.View
Imports Topomatic.Cad.View.Hints
Imports Topomatic.Dwg
Imports System.Text.RegularExpressions
Imports Microsoft.Win32
Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Interop.Excel
Imports System.Windows.Forms

Public Class FuncFiles
    Public Shared namePrograms As String = "Civil3DToolsUtility"

    'поиск файла во вспомогательных папках поддержки
    Public Shared Function funcFindFileToDirectorySupport(ByVal shortNameFile As String) As String
        funcFindFileToDirectorySupport = ""
        Dim pr As Topomatic.ApplicationEnvironment.ProcessEnvironment = Topomatic.ApplicationEnvironment.ProcessEnvironment.Current
        Dim fileShpUser As String() = pr.SearchFiles("LookInFolders", shortNameFile)
        If IsNothing(fileShpUser) = False Then
            If fileShpUser.Length > 0 Then
                For i As Integer = 0 To fileShpUser.Length - 1
                    Dim tempFile As String = fileShpUser(i)
                    If File.Exists(tempFile) = True Then
                        Return tempFile
                    End If
                Next i
            End If
        End If
    End Function
    'поиск каталога во вспомогательных папках поддержки
    Public Shared Function funcFindDirectorySupport(ByVal directory As String) As String
        funcFindDirectorySupport = ""
        Dim pr As Topomatic.ApplicationEnvironment.ProcessEnvironment = Topomatic.ApplicationEnvironment.ProcessEnvironment.Current
        Dim fileShpUser As String() = pr.GetDirectories("LookInFolders")
        If IsNothing(fileShpUser) = False Then
            If fileShpUser.Length > 0 Then
                For i As Integer = 0 To fileShpUser.Length - 1
                    Dim tempFile As String = fileShpUser(i)
                    Dim pos As Integer = tempFile.LastIndexOf(directory)
                    If pos > -1 Then
                        Return tempFile
                    End If
                Next i
            End If
        End If
    End Function
    'записать каталог в файлы поддержки
    Public Shared Function funcWriteDirectoryToDirectorySupport(ByVal putchLookDirectory As String) As Boolean
        funcWriteDirectoryToDirectorySupport = False
        Try
            If Directory.Exists(putchLookDirectory) = True Then
                Dim pr As Topomatic.ApplicationEnvironment.ProcessEnvironment = Topomatic.ApplicationEnvironment.ProcessEnvironment.Current
                Dim nameDirectories As String() = pr.GetDirectories("LookInFolders")
                Dim boolWrite As Boolean = False
                If IsArray(nameDirectories) = True Then
                    For i As Integer = 0 To nameDirectories.Length - 1
                        Dim tempDir As String = nameDirectories(i)
                        If tempDir Like putchLookDirectory Then
                            boolWrite = True
                            Exit For
                        End If
                    Next
                End If
                If boolWrite = False Then
                    Dim userModif As Topomatic.ApplicationEnvironment.EnvironmentModifier = Topomatic.ApplicationEnvironment.EnvironmentModifier.Process
                    pr.Insert(0, userModif, "LookInFolders", putchLookDirectory)
                End If
                Return True
            End If
        Catch ex As Exception
        End Try
    End Function
    '
    Public Shared Function funcReadDirectoriesSupport(ByRef nameDirectories As String()) As Boolean
        funcReadDirectoriesSupport = False
        Try
            Dim pr As Topomatic.ApplicationEnvironment.ProcessEnvironment = Topomatic.ApplicationEnvironment.ProcessEnvironment.Current
            nameDirectories = pr.GetDirectories("LookInFolders")
            funcReadDirectoriesSupport = True
        Catch ex As Exception
        End Try
    End Function

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'вывод диалогового окна выбор каталога
    Public Shared Function FuncGetAbsoluteFolderPathName(ByVal strTitle As String, Optional lngRegim As Long = 1) As String
        ' Определяем переменные
        Dim objShell As Object
        Dim objFolder As Object
        Dim objFolderItem As Object
        ' Создаём переменную Shell
        objShell = CreateObject("Shell.Application")
        ' Выводим диалоговое окно выбора папки с нужными параметрами
        objFolder = objShell.BrowseForFolder(0, strTitle, lngRegim)
        ' Если объект создан не удачно (НЕ выбрали какую-то папку), то
        ' возвращаем пустую строку в качестве результата работы функции...
        If objFolder Is Nothing Then
            FuncGetAbsoluteFolderPathName = ""
            Exit Function
        End If
        ' Получаем объект, у которого "можно спросить" его path
        objFolderItem = objFolder.Self
        ' Получаем значение Path
        FuncGetAbsoluteFolderPathName = objFolderItem.Path
        ' Удаляем все использованные объекты
        objFolderItem = Nothing
        objFolder = Nothing
        objShell = Nothing
    End Function

    Public Shared Function FuncGetAbsoluteFilePathName(ByVal strTitle As String, Optional ByVal strFilter As String = "All files (*.*)|*.*") As String
        '"txt files (*.txt)|*.txt|All files (*.*)|*.*";
        FuncGetAbsoluteFilePathName = ""
        Dim openFileDlg As System.Windows.Forms.OpenFileDialog = New System.Windows.Forms.OpenFileDialog()
        openFileDlg.InitialDirectory = "c:\\"
        openFileDlg.Filter = strFilter
        openFileDlg.RestoreDirectory = True
        If openFileDlg.ShowDialog() = DialogResult.OK Then
            Return openFileDlg.FileName
        End If
    End Function

    Public Shared Function FuncSaveFilePathName(ByVal strTitle As String, Optional ByVal strFilter As String = "All files (*.*)|*.*") As String
        '"txt files (*.txt)|*.txt|All files (*.*)|*.*";
        FuncSaveFilePathName = ""
        Dim saveFileDlg As System.Windows.Forms.SaveFileDialog = New System.Windows.Forms.SaveFileDialog()
        saveFileDlg.InitialDirectory = "c:\\"
        saveFileDlg.Filter = strFilter
        saveFileDlg.RestoreDirectory = True
        If saveFileDlg.ShowDialog() = DialogResult.OK Then
            Return saveFileDlg.FileName
        End If
    End Function

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Public Shared Function FuncReadProperties(ByVal nameDirectoriesProgram As String, ByRef generalDirectoryPrograms As String, ByRef putchTemplameXML As String, ByRef putchTemplateDwg As String) As Boolean
        FuncReadProperties = False
        Dim arrayDirSupport As String() = Nothing
        Dim boolFindDirSupport As Boolean = FuncFiles.funcReadDirectoriesSupport(arrayDirSupport)
        If IsArray(arrayDirSupport) = True Then
            For i As Integer = 0 To arrayDirSupport.Length - 1
                Dim nameFolder As String = arrayDirSupport(i)
                Dim pos As Integer = nameFolder.IndexOf(nameDirectoriesProgram)
                If pos > -1 Then
                    generalDirectoryPrograms = Mid(nameFolder, 1, pos)
                    generalDirectoryPrograms = generalDirectoryPrograms & nameDirectoriesProgram
                    Exit For
                End If
            Next
        End If
        'шаблон оформления
        Dim directorySupport As String = generalDirectoryPrograms & "\FileResources\Sample\"
        If Directory.Exists(directorySupport) = True Then
            Dim xmlFiles As String() = Directory.GetFiles(directorySupport, "*.xml")
            If IsArray(xmlFiles) = True Then
                For i As Integer = 0 To xmlFiles.Length - 1
                    putchTemplameXML = xmlFiles(i)
                    Dim shortNameFile As String = IO.Path.GetFileNameWithoutExtension(putchTemplameXML)
                    Dim tempNameFileDwg As String = directorySupport & shortNameFile & ".dwg"
                    If IO.File.Exists(tempNameFileDwg) = True Then
                        putchTemplateDwg = tempNameFileDwg
                        Exit For
                    End If
                Next
            End If
        End If
        Return True
    End Function

    Public Shared Function FuncIsFolderInPathSafe(filePath As String, folderName As String) As Boolean
        Dim dir As String = Path.GetDirectoryName(filePath) & Path.DirectorySeparatorChar
        Dim searchFor As String = Path.DirectorySeparatorChar & folderName & Path.DirectorySeparatorChar
        Return dir.IndexOf(searchFor, StringComparison.OrdinalIgnoreCase) >= 0
    End Function
End Class
