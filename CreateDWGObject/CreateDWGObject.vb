Imports Microsoft.Office.Interop.Word
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities

Public Class CreateDWGObject
    Private _activeDocument As Drawing
    Public Sub New()
        _activeDocument = Nothing
    End Sub
    Public Property ActiveDrawing As Drawing
        Get
            Return _activeDocument
        End Get
        Set(value As Drawing)
            _activeDocument = value
        End Set
    End Property

    '================================================================================================
    'создание выноски с текстом
    Public Function createLeader(ByVal value As String, ByVal InsertFirstVertex As Vector2D, ByVal InserSecondVertex As Vector2D, ByRef arrayObjectProperties As String(,), ByRef arrayMleaderProperties As String(,), ByVal templateXML As String, ByVal templateDWG As String, Optional annoScale As Double = 1) As DwgLeader
        Dim result As DwgLeader = Nothing
        If IsNothing(ActiveDrawing) = True Then
            Return result
        End If

        Try
            result = New Topomatic.Dwg.Entities.DwgLeader
            result.Add(InsertFirstVertex)
            result.Position = InserSecondVertex
            result.Content = value
            'назначаем свойства
            If IsArray(arrayObjectProperties) = True Then
                If arrayObjectProperties.Length > 0 Then
                    For i As Integer = 0 To arrayObjectProperties.GetUpperBound(1)
                        Dim nameField As String = arrayObjectProperties(0, i)
                        Dim valueField As String = arrayObjectProperties(1, i)
                        If nameField Like "Layer" Then
                            Dim mleaderLayer As DwgLayer = RoburFunc.FuncAddLayer(ActiveDrawing, valueField)
                            If IsNothing(mleaderLayer) = True Then
                                mleaderLayer = ActiveDrawing.ActiveLayer
                            End If
                            result.Layer = mleaderLayer
                        ElseIf nameField Like "Color" Then
                            Dim color As CadColor = New CadColor(256)
                            If valueField Like "ПоБлоку" Then
                                color = New CadColor(0)
                            Else
                                If IsNumeric(valueField) = True Then
                                    Dim indColor As Integer = CInt(valueField)
                                    If indColor < 0 Then indColor = 0
                                    If indColor > 256 Then indColor = 256
                                    color = New CadColor(CInt(indColor))
                                End If
                            End If
                            result.Color = color
                        ElseIf nameField Like "ArrowSize" Then
                            If IsNumeric(valueField) = True Then
                                result.ArrowheadSize = CInt(valueField) * annoScale
                            End If
                        ElseIf nameField Like "TextStyle" Then
                            Dim textStyle As DwgStyle = ActiveDrawing.ActiveStyle
                            If ActiveDrawing.Styles.Names.ContainsKey(valueField) = True Then
                                textStyle = ActiveDrawing.Styles.Item(valueField)
                            Else
                                Dim arrayPropertiesTextStyle As String(,) = {}
                                Dim boolReadTextStyleProperties As Boolean = CadastralTools.TemplateXML.FuncReadTextStylePropertiesInfoToXML(templateXML, valueField, arrayPropertiesTextStyle)
                                If arrayPropertiesTextStyle.Length > 0 Then
                                    textStyle = FuncStyles.FuncCreateTextStyleByArray(ActiveDrawing, valueField, arrayPropertiesTextStyle)
                                End If
                            End If
                            result.Style = textStyle
                        ElseIf nameField Like "TextHeight" Then
                            result.Height = Val(valueField)
                        End If
                    Next i
                End If
            End If
            result.Annotative = True
            result.ArrowheadType = AcDimArrowheadType.ClosedBlank
            If IsNothing(arrayMleaderProperties) = False Then
                If IsArray(arrayMleaderProperties) = True Then
                    If arrayMleaderProperties.Length > 0 Then
                        For i As Integer = 0 To arrayMleaderProperties.GetUpperBound(1)
                            Dim nameField As String = arrayMleaderProperties(0, i)
                            Dim nameArrow As String = arrayMleaderProperties(1, i)
                            If nameField Like "NameArrow" Then
                                If nameArrow Like "" Then
                                    result.ArrowheadType = AcDimArrowheadType.None
                                ElseIf nameArrow Like "_ArchTick" Then
                                    result.ArrowheadType = AcDimArrowheadType.ArchTick
                                ElseIf nameArrow Like "_BoxBlank" Then
                                    result.ArrowheadType = AcDimArrowheadType.BoxBlank
                                ElseIf nameArrow Like "_BoxFilled" Then
                                    result.ArrowheadType = AcDimArrowheadType.BoxFilled
                                ElseIf nameArrow Like "_Closed" Then
                                    result.ArrowheadType = AcDimArrowheadType.Closed
                                ElseIf nameArrow Like "_ClosedBlank" Then
                                    result.ArrowheadType = AcDimArrowheadType.ClosedBlank
                                ElseIf nameArrow Like "_DatumBlank" Then
                                    result.ArrowheadType = AcDimArrowheadType.DatumBlank
                                ElseIf nameArrow Like "_DatumFilled" Then
                                    result.ArrowheadType = AcDimArrowheadType.DatumFilled
                                ElseIf nameArrow Like "_Default" Then
                                    result.ArrowheadType = AcDimArrowheadType.Default
                                ElseIf nameArrow Like "_Dot" Then
                                    result.ArrowheadType = AcDimArrowheadType.Dot
                                ElseIf nameArrow Like "_DotBlank" Then
                                    result.ArrowheadType = AcDimArrowheadType.DotBlank
                                ElseIf nameArrow Like "_DotSmall" Then
                                    result.ArrowheadType = AcDimArrowheadType.DotSmall
                                ElseIf nameArrow Like "_Integral" Then
                                    result.ArrowheadType = AcDimArrowheadType.Integral
                                ElseIf nameArrow Like "_None" Then
                                    result.ArrowheadType = AcDimArrowheadType.None
                                ElseIf nameArrow Like "_Oblique" Then
                                    result.ArrowheadType = AcDimArrowheadType.Oblique
                                ElseIf nameArrow Like "_Open" Then
                                    result.ArrowheadType = AcDimArrowheadType.Open
                                ElseIf nameArrow Like "_Open30" Then
                                    result.ArrowheadType = AcDimArrowheadType.Open30
                                ElseIf nameArrow Like "_Open90" Then
                                    result.ArrowheadType = AcDimArrowheadType.Open90
                                ElseIf nameArrow Like "_Origin" Then
                                    result.ArrowheadType = AcDimArrowheadType.Origin
                                ElseIf nameArrow Like "_Origin2" Then
                                    result.ArrowheadType = AcDimArrowheadType.Origin2
                                ElseIf nameArrow Like "_Small" Then
                                    result.ArrowheadType = AcDimArrowheadType.Small
                                Else
                                    Dim userBlock As DwgBlock = FuncStyles.containsBlock(ActiveDrawing, nameArrow)
                                    If IsNothing(userBlock) = True Then
                                        userBlock = RoburFunc.FuncInsertBlockByDWGFile(ActiveDrawing, templateDWG, nameArrow)
                                    End If
                                    If IsNothing(userBlock) = False Then
                                        result.ArrowheadType = AcDimArrowheadType.UserDefined
                                        result.ArrowheadName = nameArrow
                                    Else
                                        result.ArrowheadType = AcDimArrowheadType.None
                                    End If
                                End If
                            End If
                        Next i
                    End If
                End If
            End If
            ActiveDrawing.ActiveSpace.Add(result)
        Catch ex As Exception
        End Try
        Return result
    End Function
End Class
