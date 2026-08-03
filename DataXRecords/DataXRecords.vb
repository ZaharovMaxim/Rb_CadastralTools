Imports CadastralTools
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities

Public Class DataXRecords
    'чтение значений xData с примитива
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Public Shared Function getXRecords(ByVal entity As DwgEntity, ByRef xData As StructureElement, Optional nameTableXRecords As StructureElement.tableXRecords = StructureElement.tableXRecords.PROJECT_STRUCTURES) As Boolean
        getXRecords = False
        xData = New StructureElement()
        Try
            Dim userDictionary As DwgDictionary = entity.GetExtensionDictionary()
            If IsNothing(userDictionary) = True Then
                Return False
            Else
                Dim countRecDictionary As Integer = userDictionary.Count
                If countRecDictionary > 0 Then
                    For i As Integer = 0 To countRecDictionary - 1
                        'ищем словари
                        Dim tempUserDictionary As KeyValuePair(Of String, Object) = userDictionary.ElementAt(i)
                        Dim dictionaryName As String = tempUserDictionary.Key
                        If dictionaryName Like nameTableXRecords.ToString Then
                            'словарь найден
                            Dim tempStructureDictionary As DwgDictionary = tempUserDictionary.Value
                            Dim tempFieldDictionary As DwgDictionary = tempStructureDictionary.Item("Field")
                            'читаем имена полей
                            If tempFieldDictionary.Count > 1 Then
                                For k As Integer = 0 To tempFieldDictionary.Count - 1
                                    Dim rez As KeyValuePair(Of String, Object) = tempFieldDictionary.ElementAt(k)
                                    Dim nameF As String = rez.Key
                                    Dim valF As String = rez.Value
                                    If nameF Like "Label" Or nameF Like "LABEL" Then
                                        xData.Label = valF
                                    ElseIf nameF Like "ClassObject" Then
                                        xData.NameClassObject = CType([Enum].Parse(GetType(StructureElement.classObject), valF), StructureElement.classObject)
                                    ElseIf nameF Like "Name" Or nameF Like "NAME" Then
                                        xData.Name = valF
                                    ElseIf nameF Like "Description" Then
                                        xData.Description = valF
                                    ElseIf nameF Like "KeyParameters" Then
                                        xData.KeyParameter = valF
                                    ElseIf nameF Like "IdStructure" Then
                                        xData.IdStructure = valF
                                    ElseIf nameF Like "IdElement" Then
                                        xData.IdElement = valF
                                    ElseIf nameF Like "Note" Or nameF Like "NOTE" Then
                                        xData.Note = valF
                                    End If
                                Next k
                                xData.IdObject = entity.ObjectID
                                Return True
                            End If
                        End If
                    Next i
                End If
            End If
        Catch ex As System.Exception
        End Try
    End Function

    'запись значений xData в примитив
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Public Shared Function setXRecords(ByVal entity As DwgEntity, ByVal nameTable As StructureElement.tableXRecords, ByVal xData As StructureElement, Optional deskTableXRecords As String = "") As Boolean
        If entity Is Nothing OrElse xData Is Nothing Then
            Return False
        End If
        Try
            ' Получаем или создаем словарь расширений
            Dim userDictionary As DwgDictionary = entity.GetExtensionDictionary()
            If userDictionary Is Nothing Then
                userDictionary = CreateExtensionDictionary(entity)
            End If
            ' Ищем существующую таблицу
            Dim targetDictionary As DwgDictionary = FindTableDictionary(userDictionary, nameTable.ToString())
            If targetDictionary IsNot Nothing Then
                ' Обновляем существующую таблицу
                UpdateDictionaryData(targetDictionary, entity, xData)
            Else
                ' Создаем новую таблицу
                CreateNewTableDictionary(userDictionary, nameTable, deskTableXRecords, entity, xData)
            End If
            Return True
        Catch ex As System.Exception
            ' Логирование ошибки (рекомендуется добавить систему логирования)
            'Debug.WriteLine($"Ошибка при установке XRecords: {ex.Message}")
            Return False
        End Try
    End Function
    ' Вспомогательные приватные методы для разделения логики
    Private Shared Function CreateExtensionDictionary(ByVal entity As DwgEntity) As DwgDictionary
        entity.CreateExtensionDictionary()
        Return entity.GetExtensionDictionary()
    End Function

    Private Shared Function FindTableDictionary(ByVal userDictionary As DwgDictionary, ByVal tableName As String) As DwgDictionary
        If userDictionary Is Nothing OrElse userDictionary.Count = 0 Then
            Return Nothing
        End If
        ' Используем LINQ для поиска
        For Each kvp As KeyValuePair(Of String, Object) In userDictionary
            If String.Equals(kvp.Key, tableName, StringComparison.OrdinalIgnoreCase) Then
                Try
                    Return DirectCast(kvp.Value, DwgDictionary)
                Catch
                    Continue For
                End Try
            End If
        Next
        Return Nothing
    End Function

    Private Shared Sub UpdateDictionaryData(ByVal targetDictionary As DwgDictionary, ByVal entity As DwgEntity, ByVal xData As StructureElement)
        ' Ищем словарь с данными полей
        For Each kvp As KeyValuePair(Of String, Object) In targetDictionary
            If String.Equals(kvp.Key, "Field", StringComparison.OrdinalIgnoreCase) Then
                Dim fieldDict As DwgDictionary = DirectCast(kvp.Value, DwgDictionary)
                SetFieldData(fieldDict, entity, xData)
                Exit For
            End If
        Next
    End Sub

    Private Shared Sub CreateNewTableDictionary(ByVal userDictionary As DwgDictionary, ByVal nameTable As StructureElement.tableXRecords, ByVal deskTableXRecords As String, ByVal entity As DwgEntity, ByVal xData As StructureElement)
        Dim tableDictionary As DwgDictionary = userDictionary.AddDictionary(nameTable.ToString())
        ' Устанавливаем метаданные таблицы
        tableDictionary.SetString("TableName", nameTable.ToString())
        If Not String.IsNullOrEmpty(deskTableXRecords) Then
            tableDictionary.SetString("TableDescription", deskTableXRecords)
        End If
        ' Создаем словарь для полей данных
        Dim fieldDictionary As DwgDictionary = tableDictionary.AddDictionary("Field")
        SetFieldData(fieldDictionary, entity, xData)
    End Sub

    Private Shared Sub SetFieldData(ByVal fieldDictionary As DwgDictionary, ByVal entity As DwgEntity, ByVal xData As StructureElement)
        ' Устанавливаем данные полей с проверкой на пустые значения
        SetStringIfNotEmpty(fieldDictionary, "Label", xData.Label)
        SetStringIfNotEmpty(fieldDictionary, "ClassObject", xData.NameClassObject)
        SetStringIfNotEmpty(fieldDictionary, "Name", xData.Name)
        SetStringIfNotEmpty(fieldDictionary, "Description", xData.Description)
        SetStringIfNotEmpty(fieldDictionary, "KeyParameters", xData.KeyParameter)
        SetStringIfNotEmpty(fieldDictionary, "IdStructure", xData.IdStructure)
        SetStringIfNotEmpty(fieldDictionary, "IdElement", xData.IdElement)
        SetStringIfNotEmpty(fieldDictionary, "Note", xData.Note)
        ' ObjectID всегда должен быть установлен
        fieldDictionary.SetString("IdObject", entity.ObjectID)
    End Sub
    Private Shared Sub SetStringIfNotEmpty(ByVal dictionary As DwgDictionary, ByVal key As String, ByVal value As String)
        If Not String.IsNullOrEmpty(value) Then
            dictionary.SetString(key, value)
        Else
            dictionary.SetString(key, "")
        End If
    End Sub

End Class
