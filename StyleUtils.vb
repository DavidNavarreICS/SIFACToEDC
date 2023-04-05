Imports System.Diagnostics
Imports System.Drawing
Imports Microsoft.Office.Interop.Excel

Public Class StyleUtils
    Private Shared ReadOnly BASE_COLOR As Object = RGB(102, 102, 153)
    Private Shared ReadOnly BASE_HIGHLIGHT_COLOR As Object = RGB(255, 177, 63)
    Private Shared ReadOnly RECAP_COLOR As Object = RGB(166, 166, 166)
    Private Shared ReadOnly FROM_COLOR As Object = RGB(209, 54, 33)
    Private Shared ReadOnly COMMENTS_COLOR As Object = RGB(255, 232, 197)
    Private Shared ReadOnly PATTERN_MONNAIE As String = "# ##0,00"
    Private Shared ReadOnly PATTERN_FULL_MONNAIE As String = "# ##0,00 €"
    Private Shared ReadOnly FONT_SIZE_DEFAULT = 11
    Private Shared ReadOnly FONT_SIZE_SUMMARY = 9
    ''' <summary>
    ''' Statically prepares a set of styles used by all worsheets.
    ''' </summary>
    ''' <param name="workbook"></param>
    Public Shared Sub PrepareStyles(workbook As Excel._Workbook)
        If workbook Is Nothing Then
            Throw New System.ArgumentNullException(NameOf(workbook))
        End If

        CreateStyle("WarningDetailStyle", workbook, BASE_COLOR, Color.White, FONT_SIZE_DEFAULT, False, XlHAlign.xlHAlignLeft, XlVAlign.xlVAlignBottom, False)

        CreateStyle("WarningHeaderStyle", workbook, BASE_HIGHLIGHT_COLOR, Color.White, FONT_SIZE_DEFAULT, True, XlHAlign.xlHAlignLeft, XlVAlign.xlVAlignBottom, False)

        CreateStyle("HeaderStyle", workbook, BASE_COLOR, Color.White, FONT_SIZE_DEFAULT, False, XlHAlign.xlHAlignLeft, XlVAlign.xlVAlignBottom, False)

        CreateStyle("HeaderStyleComment", workbook, BASE_HIGHLIGHT_COLOR, Color.White, FONT_SIZE_DEFAULT, False, XlHAlign.xlHAlignLeft, XlVAlign.xlVAlignBottom, False)

        CreateStyle("HeaderStyleFrom", workbook, FROM_COLOR, Color.White, FONT_SIZE_DEFAULT, False, XlHAlign.xlHAlignLeft, XlVAlign.xlVAlignBottom, False)

        CreateStyle("MtEngStyle", workbook, BASE_COLOR, Color.White, FONT_SIZE_DEFAULT, False, XlHAlign.xlHAlignRight, XlVAlign.xlVAlignBottom, False, PATTERN_MONNAIE)

        CreateStyle("SIFACCommentaires", workbook, COMMENTS_COLOR, Color.Black, FONT_SIZE_DEFAULT, False, XlHAlign.xlHAlignLeft, XlVAlign.xlVAlignBottom, False)

        CreateStyle("MtPAStyle", workbook, 0, Color.Black, FONT_SIZE_DEFAULT, False, XlHAlign.xlHAlignRight, XlVAlign.xlVAlignBottom, False, PATTERN_MONNAIE)

        CreateStyle("SumStyle", workbook, BASE_COLOR, Color.White, FONT_SIZE_DEFAULT, True, XlHAlign.xlHAlignRight, XlVAlign.xlVAlignBottom, False, PATTERN_MONNAIE)

        CreateStyle("RecapHeaderStyle", workbook, BASE_COLOR, Color.White, FONT_SIZE_SUMMARY, True, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignBottom, True)

        CreateStyle("RecapHeaderStyle2", workbook, RECAP_COLOR, Color.White, FONT_SIZE_SUMMARY, True, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignBottom, True)

        CreateStyle("RecapHeaderStyle3", workbook, BASE_COLOR, Color.White, FONT_SIZE_SUMMARY, True, XlHAlign.xlHAlignLeft, XlVAlign.xlVAlignBottom, True)

        CreateStyle("RecapHeaderStyle4", workbook, BASE_COLOR, Color.White, FONT_SIZE_SUMMARY, False, XlHAlign.xlHAlignLeft, XlVAlign.xlVAlignBottom, True)

        CreateStyle("RecapHeaderStyle5", workbook, BASE_COLOR, Color.White, FONT_SIZE_SUMMARY, False, XlHAlign.xlHAlignRight, XlVAlign.xlVAlignBottom, True)

        CreateStyle("RecapHeaderStyle6", workbook, BASE_COLOR, Color.Black, FONT_SIZE_SUMMARY, True, XlHAlign.xlHAlignCenter, XlVAlign.xlVAlignBottom, True)

        CreateStyle("RecapNumberStyle", workbook, BASE_COLOR, Color.White, FONT_SIZE_SUMMARY, False, XlHAlign.xlHAlignRight, XlVAlign.xlVAlignBottom, False, PATTERN_FULL_MONNAIE)

        CreateStyle("RecapNumberStyle2", workbook, BASE_COLOR, Color.White, FONT_SIZE_SUMMARY, True, XlHAlign.xlHAlignRight, XlVAlign.xlVAlignBottom, False, PATTERN_FULL_MONNAIE)

        CreateStyle("RecapNumberStyle3", workbook, RECAP_COLOR, Color.White, FONT_SIZE_SUMMARY, False, XlHAlign.xlHAlignRight, XlVAlign.xlVAlignBottom, False, PATTERN_FULL_MONNAIE)

        CreateStyle("RecapNumberStyle4", workbook, RECAP_COLOR, Color.White, FONT_SIZE_SUMMARY, True, XlHAlign.xlHAlignRight, XlVAlign.xlVAlignBottom, False, PATTERN_FULL_MONNAIE)

        CreateStyle("RecapNumberStyle5", workbook, BASE_COLOR, Color.Black, FONT_SIZE_SUMMARY, True, XlHAlign.xlHAlignRight, XlVAlign.xlVAlignBottom, False, PATTERN_FULL_MONNAIE)
    End Sub

    Private Shared Sub CreateStyle(aStyle As String, workbook As Excel._Workbook, interiorColor As Object, fontColor As Color, fontSize As Single, fontBold As Boolean, horizontalAlign As XlHAlign, verticalAlign As XlHAlign, wrap As Boolean, Optional numFormat As String = Nothing)
        Dim newStyle As Excel.Style
        For Each style As Style In workbook.Styles
            If style.Name = aStyle Then
                style.Delete()
            End If
        Next
        newStyle = workbook.Styles.Add(aStyle)
        ConfigureStyle(newStyle, interiorColor, fontColor, fontSize, fontBold, horizontalAlign, verticalAlign, wrap, numFormat)
    End Sub
    Private Shared Sub ConfigureStyle(style As Excel.Style, interiorColor As Object, fontColor As Color, fontSize As Single, fontBold As Boolean, horizontalAlign As XlHAlign, verticalAlign As XlHAlign, wrap As Boolean, Optional numFormat As String = Nothing)
        If interiorColor = 0 Then
            style.Interior.ColorIndex = 0
        Else
            style.Interior.Color = interiorColor
        End If
        style.Font.Color = System.Drawing.ColorTranslator.ToOle(fontColor)
        style.Font.Size = fontSize
        style.Font.Bold = fontBold
        style.HorizontalAlignment = horizontalAlign
        style.VerticalAlignment = verticalAlign
        style.WrapText = wrap
        If numFormat IsNot Nothing Then
            style.NumberFormatLocal = numFormat
        End If
    End Sub
End Class
