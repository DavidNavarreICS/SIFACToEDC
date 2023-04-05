Imports System.Collections.ObjectModel
Imports System.Globalization
Imports Microsoft.Office.Interop.Excel

Public NotInheritable Class Utils
    Private Sub New()
    End Sub
    Public Shared Function GetMessage(resourceName As String) As String
        Return My.Resources.ResourceManager.GetString(resourceName, CultureInfo.CurrentCulture)
    End Function
    Public Shared Sub CreateWarning(message As String, baseRange As Range, currentLine As Integer, style As String)
        If baseRange Is Nothing Then
            Throw New System.ArgumentNullException(NameOf(baseRange))
        End If
        Dim startRange As Excel.Range = baseRange.Cells(currentLine, 1)
        Dim startAddress As String = startRange.Address
        Dim endAddress As String = startRange.Offset(0, 13).Address
        Dim mergedRange As Range = baseRange.Range(startAddress, endAddress)
        mergedRange.Merge()
        SetCellValue2(mergedRange, message, style)
    End Sub
    Public Shared Function IsNewHeaderVersion(cell As Range) As Boolean
        If cell Is Nothing Then
            Throw New System.ArgumentNullException(NameOf(cell))
        End If
        Return IsHeader(cell.Cells(1, 1)) AndAlso IsComment(cell.Cells(1, 13))
    End Function

    Public Shared Function IsHeader(cell As Range) As Boolean
        If cell Is Nothing Then
            Throw New System.ArgumentNullException(NameOf(cell))
        End If
        Return CStr(cell.Value2) <> "" AndAlso Not IsNumeric(cell.Value2)
    End Function

    Public Shared Function IsComment(cell As Range) As Boolean
        If cell Is Nothing Then
            Throw New System.ArgumentNullException(NameOf(cell))
        End If
        Return CStr(cell.Value2) = "Commentaires"
    End Function

    Public Shared Function IsSum(cell As Range) As Boolean
        If cell Is Nothing Then
            Throw New System.ArgumentNullException(NameOf(cell))
        End If
        Return CStr(cell.Cells(1, 5).Value2) = "Somme :"
    End Function

    Public Shared Function IsLineWithComment(cell As Range) As Boolean
        If cell Is Nothing Then
            Throw New System.ArgumentNullException(NameOf(cell))
        End If
        Return CStr(cell.Cells(1, 13).Value2) <> ""
    End Function

    Public Shared Function IsSumWithComment(cell As Range) As Boolean
        If cell Is Nothing Then
            Throw New System.ArgumentNullException(NameOf(cell))
        End If
        For I As Integer = 1 To 4
            If CStr(cell.Cells(1, I).Value2) <> "" Then
                Return True
            End If
        Next
        For I As Integer = 7 To 14
            If CStr(cell.Cells(1, I).Value2) <> "" Then
                Return True
            End If
        Next
        Return False
    End Function
    Private Shared Sub SetCellContentAndStyle(cell As Excel.Range, content As String, isArea As Boolean, isStyled As Boolean, isValue2 As Boolean, isFormula As Boolean, Optional aStyle As String = "")
        If cell Is Nothing Then
            Throw New System.ArgumentNullException(NameOf(cell))
        End If
        If isStyled Then
            If isArea Then
                cell.MergeArea.Style = aStyle
            Else
                cell.Style = aStyle
            End If
        End If
        If isFormula Then
            cell.Formula = content
        ElseIf isValue2 Then
            cell.Value2 = content
        Else
            cell.Value = content
        End If
    End Sub
    Public Shared Sub SetCellValue(cell As Excel.Range, aValue As String, aStyle As String)
        SetCellContentAndStyle(cell, aValue, False, True, False, False, aStyle)
    End Sub
    Public Shared Sub SetCellValue2(cell As Excel.Range, aValue As String, aStyle As String)
        SetCellContentAndStyle(cell, aValue, False, True, True, False, aStyle)
    End Sub
    Public Shared Sub SetCellFormula(cell As Excel.Range, aValue As String, aStyle As String)
        SetCellContentAndStyle(cell, aValue, False, True, False, True, aStyle)
    End Sub
    Public Shared Sub SetCellRawFormula(cell As Excel.Range, aValue As String)
        SetCellContentAndStyle(cell, aValue, False, False, False, True)
    End Sub
    Public Shared Sub SetAreaValue2(cell As Excel.Range, aValue As String, aStyle As String)
        SetCellContentAndStyle(cell, aValue, True, True, True, False, aStyle)
    End Sub
    Public Shared Function GetFormattedString(format As String, ParamArray value() As Object) As String
        Return String.Format(CultureInfo.CurrentCulture, Utils.GetMessage(format), value)
    End Function
    Public Shared Sub AutoFit(newWorsheet As Excel._Worksheet)
        If newWorsheet Is Nothing Then
            Throw New System.ArgumentNullException(NameOf(newWorsheet))
        End If
        newWorsheet.Range("A:N").EntireColumn.AutoFit()
    End Sub
    Public Shared Function IsInvest(line As BookLine) As Boolean
        If line Is Nothing Then
            Throw New System.ArgumentNullException(NameOf(line))
        End If
        Return line.ACptegen.Trim.StartsWith("2", StringComparison.CurrentCulture)
    End Function
    Public Shared Sub AddLineFromTableToTable(destination As Collection(Of BookLine), source As Collection(Of BookLine))
        If source Is Nothing Then
            Throw New System.ArgumentNullException(NameOf(source))
        End If
        If destination Is Nothing Then
            Throw New System.ArgumentNullException(NameOf(destination))
        End If
        For Each line As BookLine In source
            destination.Add(line)
        Next
    End Sub
    Public Shared Sub AddLineToTable(key As String, preparedLines As Dictionary(Of String, Collection(Of BookLine)), line As BookLine)
        If preparedLines Is Nothing Then
            Throw New System.ArgumentNullException(NameOf(preparedLines))
        End If
        If Not preparedLines.ContainsKey(key) Then
            Dim NewLines As New Collection(Of BookLine) From {
                line
            }
            preparedLines.Add(key, NewLines)
        Else
            preparedLines.Item(key).Add(line)
        End If
    End Sub
    Public Shared Sub CopyHeaders(baseWorksheet As Excel._Worksheet, newWorksheet As Excel._Worksheet)
        If baseWorksheet Is Nothing Then
            Throw New System.ArgumentNullException(NameOf(baseWorksheet))
        End If
        If newWorksheet Is Nothing Then
            Throw New System.ArgumentNullException(NameOf(newWorksheet))
        End If
        Dim SourceRange As Excel.Range = baseWorksheet.UsedRange.Rows(1)
        Dim DestRange As Excel.Range = newWorksheet.Range("A1")
        SourceRange.Copy(DestRange)
        For I As Integer = 1 To 12
            Dim RDest As Excel.Range = DestRange.Cells(1, I)
            Dim RSource As Excel.Range = SourceRange.Cells(1, I)
            RDest.ColumnWidth = RSource.ColumnWidth
        Next
    End Sub
    Public Shared Sub DumpData(worksheet As Excel._Worksheet, dataTable As Dictionary(Of String, Collection(Of BookLine)))
        If worksheet Is Nothing Then
            Throw New System.ArgumentNullException(NameOf(worksheet))
        End If
        If dataTable Is Nothing Then
            Throw New System.ArgumentNullException(NameOf(dataTable))
        End If
        Dim CurrentLine As Integer = 1
        Dim StartRange As Excel.Range = worksheet.Range("A3")
        For Each LineCollection As Collection(Of BookLine) In dataTable.Values
            For Each Line As BookLine In LineCollection
                StartRange.Cells(CurrentLine, 1).Value2 = Line.ACptegen
                StartRange.Cells(CurrentLine, 2).Value2 = Line.BRubrique
                StartRange.Cells(CurrentLine, 3).Value2 = Line.CNumeroFlux
                StartRange.Cells(CurrentLine, 4).Value2 = Line.DNom
                StartRange.Cells(CurrentLine, 5).Value2 = Line.ELibelle
                StartRange.Cells(CurrentLine, 6).Value2 = Line.FMntEngHtr
                StartRange.Cells(CurrentLine, 7).Value2 = Line.GMontantPA
                StartRange.Cells(CurrentLine, 8).Value2 = Line.HRapprochmt
                StartRange.Cells(CurrentLine, 9).Value2 = Line.IRefFactF
                StartRange.Cells(CurrentLine, 10).Value2 = Line.JDatePce
                StartRange.Cells(CurrentLine, 11).Value2 = GetDateCompteAsText(Line)
                StartRange.Cells(CurrentLine, 12).Value2 = Line.LNumPiece
                CurrentLine += 1
                Globals.ThisAddIn.NextStep()
                Globals.ThisAddIn.NextStep()
                Globals.ThisAddIn.NextStep()
                Globals.ThisAddIn.NextStep()
                Globals.ThisAddIn.NextStep()
            Next
        Next
    End Sub
    Public Shared Function GetDateCompteAsText(line As BookLine) As String
        If line IsNot Nothing AndAlso line.KDCompt IsNot Nothing Then
            Return line.KDCompt.AsText
        Else
            Return ""
        End If
    End Function
    Public Shared Function ReadLine(fullRange As Range, rowNum As Integer) As BookLine
        If fullRange IsNot Nothing Then
            Return New BookLine With {
            .ACptegen = fullRange.Cells(rowNum, 1).Value2,
            .BRubrique = fullRange.Cells(rowNum, 2).Value2,
            .CNumeroFlux = fullRange.Cells(rowNum, 3).Value2,
            .DNom = fullRange.Cells(rowNum, 4).Value2,
            .ELibelle = fullRange.Cells(rowNum, 5).Value2,
            .FMntEngHtr = GetNumber(fullRange, rowNum, 6),
            .GMontantPA = GetNumber(fullRange, rowNum, 7),
            .HRapprochmt = fullRange.Cells(rowNum, 8).Value2,
            .IRefFactF = fullRange.Cells(rowNum, 9).Value2,
            .JDatePce = fullRange.Cells(rowNum, 10).Value2,
            .KDCompt = GetDateCompte(fullRange.Cells(rowNum, 11).Value2),
            .LNumPiece = fullRange.Cells(rowNum, 12).Value2
        }
        Else
            Return Nothing
        End If
    End Function
    Public Shared Function GetDateCompte(textValue As String) As DateCompte
        If textValue <> "" Then
            Return New DateCompte(textValue)
        Else
            Return Nothing
        End If
    End Function
    Public Shared Function GetNumber(fullRange As Range, rowNum As Integer, colNum As Integer) As Double
        If fullRange Is Nothing Then
            Throw New System.ArgumentNullException(NameOf(fullRange))
        End If
        Dim TextToConvert As String = fullRange.Cells(rowNum, colNum).Value2
        Dim IndexVirgule As Integer = TextToConvert.IndexOf(",", StringComparison.Ordinal)
        Dim IndexPoint As Integer = TextToConvert.IndexOf(".", StringComparison.Ordinal)
        If IndexPoint = -1 Then
            'French format
            Return Double.Parse(TextToConvert, CultureInfo.CurrentCulture)
        ElseIf IndexVirgule = -1 Then
            'Invariant culture format
            Return Double.Parse(TextToConvert, CultureInfo.InvariantCulture)
        ElseIf IndexVirgule < IndexPoint Then
            'Invariant culture format
            Return Double.Parse(TextToConvert, CultureInfo.InvariantCulture)
        Else
            'Space format
            Dim FirstSplit As String() = TextToConvert.Split(".")
            Dim SecondSplit As String() = FirstSplit(1).Split(",")
            Dim NewTextToConvert As String = FirstSplit(0) & "," & SecondSplit(0) & "." & SecondSplit(1)
            Return Double.Parse(NewTextToConvert, CultureInfo.InvariantCulture)
        End If
    End Function
End Class
