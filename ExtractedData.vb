Imports System.Collections.ObjectModel
Imports System.Diagnostics
Imports System.Globalization
Imports Microsoft.Office.Interop.Excel
Public Class ExtractedData
    Private ReadOnly FullTable As New Dictionary(Of String, Collection(Of BookLine))
    Private ReadOnly ReducedTable As New Dictionary(Of String, Collection(Of BookLine))
    Private ReadOnly PreReducedTable As New Dictionary(Of String, Dictionary(Of String, Collection(Of BookLine)))
    Private ReadOnly TrashTable As New Dictionary(Of String, Collection(Of BookLine))
    Private FirstDataRow As Integer
    Private LastDataRow As Integer
    Private aNumberOfLines As Integer
    Private Const CASE_COMMANDE_PAIEMENT As String = "COMMANDE/FACTURE/PAIEMENT"
    Private Const CASE_COMMANDE_AJUSTEMENT As String = "COMMANDES/FACT-AJUSTEMENT/PAIEMENT"
    Private Const CASE_COMMANDE As String = "COMMANDE"
    Private Const CASE_COMMANDE_PAIEMENT_INVEST As String = "COMMANDE_INVEST/FACTURE/PAIEMENT"
    Private Const CASE_COMMANDE_AJUSTEMENT_INVEST As String = "COMMANDES_INVEST/FACT-AJUSTEMENT/PAIEMENT"
    Private Const CASE_COMMANDE_INVEST As String = "COMMANDE_INVEST"
    Private Const CASE_MISSION As String = "MISSION"
    Private Const CASE_MISSION_PAIEMENT As String = "MISSION/COMPTABILISATION/PAIEMENT"
    Private Const CASE_REIMPUTATION As String = "REIMPUTATION"
    Private Const CASE_AVOIR_PAIEMENT As String = "COMMANDE/AVOIR/PAIEMENT"
    Private Const CASE_REIMPUTATION_INVEST As String = "REIMPUTATION_INVEST"
    Private Const CASE_AVOIR_PAIEMENT_INVEST As String = "COMMANDE_INVEST/AVOIR/PAIEMENT"
    Private Const CASE_UNUSED As String = "UNUSED"
    Private Const KEY_ORDER As String = "COMMANDE"
    Private Const KEY_ORDER_PENDING As String = "COMMANDE/PENDING"
    Private Const KEY_INVEST As String = "INVESTISSEMENT"
    Private Const KEY_INVEST_PENDING As String = "INVESTISSEMENT/PENDING"
    Private Const KEY_MISSION As String = "MISSION"
    Private Const KEY_MISSION_PENDING As String = "MISSION/PENDING"
    Private Const KEY_TRASH As String = "TRASH"
    Private ReadOnly aSourceName As String
    Private ReadOnly DestinationName As String
    Private ReadOnly AddinApplication As Application
    Private NewWorkbook As Excel.Workbook
    Private BaseWorksheet As Excel.Worksheet
    Private NewWorksheet As Excel.Worksheet
    Private TrashWorksheet As Excel.Worksheet
    Private aSheetYear As Integer
    Public ReadOnly Property SheetYear As Integer
        Get
            Return aSheetYear
        End Get
    End Property
    Public ReadOnly Property NumberOfLines As Integer
        Get
            Return aNumberOfLines
        End Get
    End Property
    Public ReadOnly Property SourceName As String
        Get
            Return aSourceName
        End Get
    End Property
    Public ReadOnly Property Orders As Collection(Of BookLine)
        Get
            Return ReducedTable.Item(KEY_ORDER)
        End Get
    End Property
    Public ReadOnly Property Invests As Collection(Of BookLine)
        Get
            Return ReducedTable.Item(KEY_INVEST)
        End Get
    End Property
    Public ReadOnly Property Missions As Collection(Of BookLine)
        Get
            Return ReducedTable.Item(KEY_MISSION)
        End Get
    End Property
    Public ReadOnly Property PendingOrders As Collection(Of BookLine)
        Get
            Return ReducedTable.Item(KEY_ORDER_PENDING)
        End Get
    End Property
    Public ReadOnly Property PendingInvests As Collection(Of BookLine)
        Get
            Return ReducedTable.Item(KEY_INVEST_PENDING)
        End Get
    End Property
    Public ReadOnly Property PendingMissions As Collection(Of BookLine)
        Get
            Return ReducedTable.Item(KEY_MISSION_PENDING)
        End Get
    End Property
    Public Sub New(aSourceName As String, aDestinationName As String, aAddinApplication As Application)
        If aAddinApplication Is Nothing Then
            Throw New ArgumentNullException(NameOf(aAddinApplication))
        End If

        Me.aSourceName = aSourceName
        Me.DestinationName = aDestinationName
        Me.AddinApplication = aAddinApplication
        PrepareWorkbook()
    End Sub
    Public Sub PrepareWorkbook()
        aSheetYear = 0
        NewWorkbook = CopyWorkbook()
        BaseWorksheet = CType(AddinApplication.ActiveSheet, Excel.Worksheet)
        Dim MaxNameLength = Math.Min(20, BaseWorksheet.Name.Length)
        NewWorksheet = CreateWorksheet(BaseWorksheet, BaseWorksheet.Name.Substring(0, MaxNameLength) & " Extracted")
        TrashWorksheet = CreateWorksheet(NewWorksheet, BaseWorksheet.Name.Substring(0, MaxNameLength) & " Trash")

        Dim R1 As Excel.Range = BaseWorksheet.UsedRange
        FirstDataRow = R1.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).Row
        LastDataRow = R1.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).Row
        aNumberOfLines = LastDataRow - FirstDataRow + 1
    End Sub
    Public Sub DoPrepareExtract()
        TrashTable.Clear()
        FullTable.Clear()
        ReducedTable.Clear()
        ReducedTable.Add(KEY_ORDER_PENDING, New Collection(Of BookLine))
        ReducedTable.Add(KEY_ORDER, New Collection(Of BookLine))
        ReducedTable.Add(KEY_INVEST_PENDING, New Collection(Of BookLine))
        ReducedTable.Add(KEY_INVEST, New Collection(Of BookLine))
        ReducedTable.Add(KEY_MISSION_PENDING, New Collection(Of BookLine))
        ReducedTable.Add(KEY_MISSION, New Collection(Of BookLine))
        TrashTable.Add(KEY_TRASH, New Collection(Of BookLine))
        PrepareExtractFromTo(BaseWorksheet)
    End Sub
    Public Sub DoExtract(previousExtraction As ExtractedData)
        ExtractFromTo(BaseWorksheet, NewWorksheet, TrashWorksheet, previousExtraction)
        NewWorkbook.Save()
        NewWorkbook.Close()
    End Sub
    Private Function CopyWorkbook() As Excel.Workbook
        Dim oldWorkbook As Excel.Workbook = AddinApplication.Workbooks.Open(aSourceName)
        'oldWorkbook.IsAddin = True
        Dim OriginalWorksheet As Excel.Worksheet = oldWorkbook.ActiveSheet
        Dim newWorkbook As Excel.Workbook = AddinApplication.Workbooks.Add()
        'newWorkbook.IsAddin = True
        Dim WorksheetToRemove As Excel.Worksheet = newWorkbook.ActiveSheet
        OriginalWorksheet.Copy(WorksheetToRemove)
        WorksheetToRemove.Delete()
        Dim CopiedWorksheet As Excel.Worksheet = newWorkbook.ActiveSheet
        CopiedWorksheet.Range("1:1").Delete(XlDeleteShiftDirection.xlShiftUp)
        oldWorkbook.Close()
        newWorkbook.SaveAs(DestinationName)

        Return newWorkbook
    End Function
    Private Function CreateWorksheet(baseWorksheet As Excel.Worksheet, sheetName As String) As Excel.Worksheet
        Dim worksheet = CType(AddinApplication.Worksheets.Add(After:=baseWorksheet), Excel.Worksheet)
        worksheet.Name = sheetName
        Return worksheet
    End Function
    Private Sub PrepareExtractFromTo(baseWorksheet As Excel.Worksheet)
        Globals.ThisAddIn.NameStep("Lecture des lignes")
        FeedTable(baseWorksheet)
        Globals.ThisAddIn.NameStep("Traitement des lignes")
        ReduceTable()
    End Sub
    Private Sub ExtractFromTo(baseWorksheet As Excel.Worksheet, newWorksheet As Excel.Worksheet, trashWorksheet As Excel.Worksheet, previousExtraction As ExtractedData)
        ReduceLines(previousExtraction)
        Globals.ThisAddIn.NameStep("Générations des feuilles")
        CopyHeaders(baseWorksheet, trashWorksheet)
        CopyHeaders(baseWorksheet, newWorksheet)
        DumpData(newWorksheet, ReducedTable)
        DumpData(trashWorksheet, TrashTable)
    End Sub
    Private Sub ReduceTable()
        For Each Key As String In FullTable.Keys
            Dim PreparedLines As New Dictionary(Of String, Collection(Of BookLine)) From {
            {CASE_AVOIR_PAIEMENT, New Collection(Of BookLine)},
            {CASE_COMMANDE, New Collection(Of BookLine)},
            {CASE_COMMANDE_AJUSTEMENT, New Collection(Of BookLine)},
            {CASE_COMMANDE_PAIEMENT, New Collection(Of BookLine)},
            {CASE_AVOIR_PAIEMENT_INVEST, New Collection(Of BookLine)},
            {CASE_COMMANDE_INVEST, New Collection(Of BookLine)},
            {CASE_COMMANDE_AJUSTEMENT_INVEST, New Collection(Of BookLine)},
            {CASE_COMMANDE_PAIEMENT_INVEST, New Collection(Of BookLine)},
            {CASE_MISSION, New Collection(Of BookLine)},
            {CASE_MISSION_PAIEMENT, New Collection(Of BookLine)},
            {CASE_REIMPUTATION, New Collection(Of BookLine)},
            {CASE_REIMPUTATION_INVEST, New Collection(Of BookLine)},
            {CASE_UNUSED, New Collection(Of BookLine)}
            }
            Dim MAX_K_DCompt As DateCompte = FullTable.Item(Key).Max().KDCompt
            For Each Line As BookLine In FullTable.Item(Key)
                Dim Kind As String = Line.BRubrique
                Select Case Kind
                    Case CASE_COMMANDE_PAIEMENT
                        Line.KDCompt = MAX_K_DCompt
                        If IsInvest(Line) Then
                            AddLineToTable(CASE_COMMANDE_PAIEMENT_INVEST, PreparedLines, Line)
                        Else
                            AddLineToTable(CASE_COMMANDE_PAIEMENT, PreparedLines, Line)
                        End If
                    Case CASE_COMMANDE_AJUSTEMENT
                        Line.KDCompt = MAX_K_DCompt
                        If IsInvest(Line) Then
                            AddLineToTable(CASE_COMMANDE_AJUSTEMENT_INVEST, PreparedLines, Line)
                        Else
                            AddLineToTable(CASE_COMMANDE_AJUSTEMENT, PreparedLines, Line)
                        End If
                    Case CASE_COMMANDE
                        Line.KDCompt = MAX_K_DCompt
                        If IsInvest(Line) Then
                            AddLineToTable(CASE_COMMANDE_INVEST, PreparedLines, Line)
                        Else
                            AddLineToTable(CASE_COMMANDE, PreparedLines, Line)
                        End If
                    Case CASE_MISSION_PAIEMENT
                        Line.KDCompt = MAX_K_DCompt
                        AddLineToTable(CASE_MISSION_PAIEMENT, PreparedLines, Line)
                    Case CASE_MISSION
                        Line.KDCompt = MAX_K_DCompt
                        AddLineToTable(CASE_MISSION, PreparedLines, Line)
                    Case CASE_REIMPUTATION
                        If IsInvest(Line) Then
                            AddLineToTable(CASE_REIMPUTATION_INVEST, PreparedLines, Line)
                        Else
                            AddLineToTable(CASE_REIMPUTATION, PreparedLines, Line)
                        End If
                    Case CASE_AVOIR_PAIEMENT
                        If IsInvest(Line) Then
                            AddLineToTable(CASE_AVOIR_PAIEMENT_INVEST, PreparedLines, Line)
                        Else
                            AddLineToTable(CASE_AVOIR_PAIEMENT, PreparedLines, Line)
                        End If
                    Case Else
                        AddLineToTable(CASE_UNUSED, PreparedLines, Line)
                End Select
                Globals.ThisAddIn.NextStep()
            Next
            PreReducedTable.Add(Key, PreparedLines)
        Next
    End Sub

    Private Shared Function IsInvest(line As BookLine) As Boolean
        Return line.ACptegen.Trim.StartsWith("2", StringComparison.CurrentCulture)
    End Function

    Private Sub ReduceLines(previousExtraction As ExtractedData)
        PreparePreviousData(previousExtraction)

        For Each Key As String In PreReducedTable.Keys
            Dim CommandePaiementFound As Boolean = PreReducedTable.Item(Key).ContainsKey(CASE_COMMANDE_PAIEMENT)
            If CommandePaiementFound Then
                CommandePaiementFound = PreReducedTable.Item(Key).Item(CASE_COMMANDE_PAIEMENT).Count > 0
            End If
            Dim MissionPaiementFound As Boolean = PreReducedTable.Item(Key).ContainsKey(CASE_MISSION_PAIEMENT)
            If MissionPaiementFound Then
                MissionPaiementFound = PreReducedTable.Item(Key).Item(CASE_MISSION_PAIEMENT).Count > 0
            End If
            Dim InvestissementPaiementFound As Boolean = PreReducedTable.Item(Key).ContainsKey(CASE_COMMANDE_PAIEMENT_INVEST)
            If InvestissementPaiementFound Then
                InvestissementPaiementFound = PreReducedTable.Item(Key).Item(CASE_COMMANDE_PAIEMENT_INVEST).Count > 0
            End If
            Dim PreparedLines As Dictionary(Of String, Collection(Of BookLine)) = PreReducedTable.Item(Key)

            If CommandePaiementFound Then
                AddLineToReducedTable(KEY_ORDER, CASE_COMMANDE_PAIEMENT, PreparedLines)
                AddLineToTrashTable(KEY_TRASH, CASE_COMMANDE, PreparedLines)
            ElseIf PreparedLines.ContainsKey(CASE_COMMANDE) Then
                AddLineToReducedTable(KEY_ORDER_PENDING, CASE_COMMANDE, PreparedLines)
            End If
            If InvestissementPaiementFound Then
                AddLineToReducedTable(KEY_INVEST, CASE_COMMANDE_PAIEMENT_INVEST, PreparedLines)
                AddLineToTrashTable(KEY_TRASH, CASE_COMMANDE_INVEST, PreparedLines)
            ElseIf PreparedLines.ContainsKey(CASE_COMMANDE_INVEST) Then
                AddLineToReducedTable(KEY_INVEST_PENDING, CASE_COMMANDE_INVEST, PreparedLines)
            End If
            If PreparedLines.ContainsKey(CASE_COMMANDE_AJUSTEMENT) Then
                AddLineToReducedTable(KEY_ORDER, CASE_COMMANDE_AJUSTEMENT, PreparedLines)
            End If
            If PreparedLines.ContainsKey(CASE_COMMANDE_AJUSTEMENT_INVEST) Then
                AddLineToReducedTable(KEY_INVEST, CASE_COMMANDE_AJUSTEMENT_INVEST, PreparedLines)
            End If
            If MissionPaiementFound Then
                AddLineToReducedTable(KEY_MISSION, CASE_MISSION_PAIEMENT, PreparedLines)
                AddLineToTrashTable(KEY_TRASH, CASE_MISSION, PreparedLines)
            ElseIf PreparedLines.ContainsKey(CASE_MISSION) Then
                AddLineToReducedTable(KEY_MISSION_PENDING, CASE_MISSION, PreparedLines)
            End If
            If PreparedLines.ContainsKey(CASE_REIMPUTATION) Then
                AddLineToReducedTable(KEY_ORDER, CASE_REIMPUTATION, PreparedLines)
            End If
            If PreparedLines.ContainsKey(CASE_REIMPUTATION_INVEST) Then
                AddLineToReducedTable(KEY_INVEST, CASE_REIMPUTATION_INVEST, PreparedLines)
            End If
            If PreparedLines.ContainsKey(CASE_AVOIR_PAIEMENT) Then
                AddLineToReducedTable(KEY_ORDER, CASE_AVOIR_PAIEMENT, PreparedLines)
            End If
            If PreparedLines.ContainsKey(CASE_AVOIR_PAIEMENT_INVEST) Then
                AddLineToReducedTable(KEY_INVEST, CASE_AVOIR_PAIEMENT_INVEST, PreparedLines)
            End If
            If PreparedLines.ContainsKey(CASE_UNUSED) Then
                AddLineToTrashTable(KEY_TRASH, CASE_UNUSED, PreparedLines)
            End If
        Next
    End Sub

    Private Sub PreparePreviousData(previousExtraction As ExtractedData)
        If previousExtraction IsNot Nothing Then
            For Each Line As BookLine In previousExtraction.PendingOrders
                If Line.NFrom = "" Then
                    Line.NFrom = String.Format(CultureInfo.CurrentCulture, My.Resources.ResourceManager.GetString("LineFrom"), previousExtraction.SheetYear)
                End If
                If PreReducedTable.ContainsKey(Line.CNumeroFlux) Then
                    PreReducedTable.Item(Line.CNumeroFlux).Item(CASE_COMMANDE).Add(Line)
                Else
                    PreReducedTable.Add(Line.CNumeroFlux, New Dictionary(Of String, Collection(Of BookLine)) From {
                    {CASE_COMMANDE, New Collection(Of BookLine) From {Line}}})
                End If
            Next
            previousExtraction.PendingOrders.Clear()
            For Each Line As BookLine In previousExtraction.PendingMissions
                If Line.NFrom = "" Then
                    Line.NFrom = String.Format(CultureInfo.CurrentCulture, My.Resources.ResourceManager.GetString("LineFrom"), previousExtraction.SheetYear)
                End If
                If PreReducedTable.ContainsKey(Line.CNumeroFlux) Then
                    PreReducedTable.Item(Line.CNumeroFlux).Item(CASE_MISSION).Add(Line)
                Else
                    PreReducedTable.Add(Line.CNumeroFlux, New Dictionary(Of String, Collection(Of BookLine)) From {
                    {CASE_MISSION, New Collection(Of BookLine) From {Line}}})
                End If
            Next
            previousExtraction.PendingMissions.Clear()
            For Each Line As BookLine In previousExtraction.PendingInvests
                If Line.NFrom = "" Then
                    Line.NFrom = String.Format(CultureInfo.CurrentCulture, My.Resources.ResourceManager.GetString("LineFrom"), previousExtraction.SheetYear)
                End If
                If PreReducedTable.ContainsKey(Line.CNumeroFlux) Then
                    PreReducedTable.Item(Line.CNumeroFlux).Item(CASE_AVOIR_PAIEMENT_INVEST).Add(Line)
                Else
                    PreReducedTable.Add(Line.CNumeroFlux, New Dictionary(Of String, Collection(Of BookLine)) From {
                    {CASE_COMMANDE_INVEST, New Collection(Of BookLine) From {Line}}})
                End If
            Next
            previousExtraction.PendingInvests.Clear()
        End If
    End Sub

    Private Sub AddLineToReducedTable(keyReduced As String, keyPrepared As String, preparedLines As Dictionary(Of String, Collection(Of BookLine)))
        AddLineFromTableToTable(keyReduced, ReducedTable, keyPrepared, preparedLines)
    End Sub
    Private Sub AddLineToTrashTable(keyTrashed As String, keyPrepared As String, preparedLines As Dictionary(Of String, Collection(Of BookLine)))
        AddLineFromTableToTable(keyTrashed, TrashTable, keyPrepared, preparedLines)
    End Sub
    Private Shared Sub AddLineFromTableToTable(keyReduced As String, destTable As Dictionary(Of String, Collection(Of BookLine)), keyPrepared As String, sourceTable As Dictionary(Of String, Collection(Of BookLine)))
        For Each line As BookLine In sourceTable.Item(keyPrepared)
            destTable.Item(keyReduced).Add(line)
        Next
    End Sub
    Private Shared Sub AddLineToTable(key As String, preparedLines As Dictionary(Of String, Collection(Of BookLine)), line As BookLine)
        If Not preparedLines.ContainsKey(key) Then
            Dim NewLines As New Collection(Of BookLine) From {
                line
            }
            preparedLines.Add(key, NewLines)
        Else
            preparedLines.Item(key).Add(line)
        End If
    End Sub
    Private Shared Sub CopyHeaders(baseWorksheet As Excel.Worksheet, newWorksheet As Excel.Worksheet)
        Dim SourceRange As Excel.Range = baseWorksheet.UsedRange.Rows(1)
        Dim DestRange As Excel.Range = newWorksheet.Range("A1")
        SourceRange.Copy(DestRange)
        For I As Integer = 1 To 12
            Dim RDest As Excel.Range = DestRange.Cells(1, I)
            Dim RSource As Excel.Range = SourceRange.Cells(1, I)
            RDest.ColumnWidth = RSource.ColumnWidth
        Next
    End Sub
    Private Shared Sub DumpData(worksheet As Excel.Worksheet, dataTable As Dictionary(Of String, Collection(Of BookLine)))
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
    Private Sub FeedTable(baseWorksheet As Excel.Worksheet)
        Dim FullRange As Excel.Range = baseWorksheet.UsedRange.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown)
        For RowNum As Integer = 1 To NumberOfLines
            Dim Line As Excel.Range = FullRange.Item(RowNum, 3)
            Dim Key As String = Line.Value2
            Dim Data = New Collection(Of BookLine)
            If FullTable.ContainsKey(Key) Then
                FullTable.TryGetValue(Key, Data)
            Else
                FullTable.Add(Key, Data)
            End If

            Dim NewLine As BookLine = ReadLine(FullRange, RowNum)
            NewLine.MComment = ""
            NewLine.NFrom = ""
            Data.Add(NewLine)
            If NewLine.KDCompt IsNot Nothing Then
                aSheetYear = Math.Max(aSheetYear, NewLine.KDCompt.Year)
            End If
            Globals.ThisAddIn.NextStep()
        Next
    End Sub
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
    Private Shared Function GetDateCompte(textValue As String) As DateCompte
        If textValue <> "" Then
            Return New DateCompte(textValue)
        Else
            Return Nothing
        End If
    End Function
    Private Shared Function GetNumber(fullRange As Range, rowNum As Integer, colNum As Integer) As Double
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
