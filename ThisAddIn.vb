Imports System.Diagnostics
Imports System.Globalization
Imports System.IO
Imports Microsoft.Office.Interop.Excel
<Assembly: CLSCompliant(False)>
Public Class ThisAddIn
    Private Const FILE_NAME_PATTERN As String = "*.xls*"
    Private Const DEST_NAME_EXTENSION As String = ".xlsx"
    Private Const SOURCE_DIRECTORY_MODIFIER As String = "\sources\"
    Private Const EXTRACT_DIRECTORY_MODIFIER As String = "\extractions\"
    Private Const PROGRESS_STEP_CDD As Double = 0.05
    Private Const PROGRESS_STEP_CREATE_FILES As Double = 0.1
    Private Const PROGRESS_STEP_READ_FILES As Double = 0.65
    Private Const PROGRESS_STEP_ASSEMBLE As Double = 0.995
    Private Const PROGRESS_REFERENCE As Double = 1000
    Private Const KEY_SHEET_MODE_EMPLOI As String = "Mode d'emploi"
    Private Const KEY_SHEET_RECAPITULATIF As String = "RECAPITULATIF"
    Private Const KEY_SHEET_YEARS As String = "Years"
    Private Const KEY_SHEET_WARNINGS As String = "Warnings"
    Private Const KEY_SHEET_ELSE As String = "Divers"
    Private BaseDirectory As String
    Private SourcesDirectory As String
    Private ExtractionDirectory As String
    Private ProgressIncrement As Double
    Private CurrentProgrees As Double
    Private CurrentWorkbook As Excel.Workbook
    Private SummaryWorkSheet As Excel.Worksheet
    Private ReadOnly ProgressDialog As New ExecutionStatus
    Private ReadOnly Data As New Dictionary(Of Integer, ExtractedData)
    Private ReadOnly AllWorksheets As New Dictionary(Of Integer, Excel.Worksheet)
    Private ReadOnly LinesWithComment As New Dictionary(Of Integer, List(Of BookLine))
    Private ReadOnly OutOfRangeComments As New Dictionary(Of Integer, Boolean)
    Private ReadOnly OutOfTableComments As New Dictionary(Of Integer, Boolean)
    Private ReadOnly OutOfSumComments As New Dictionary(Of Integer, Boolean)
    Private Enum SummaryCellKind
        TOTAL
        TOTAL_NET
        CUMUL
        BUDGET
        ENGAGED
    End Enum
    Private ReadOnly SummaryCellsNotFound As New Dictionary(Of SummaryCellKind, Boolean)
    Private Const KEY_FONCT As String = "Fonctionnement"
    Private Const KEY_INVEST As String = "Investissement"
    Private Const KEY_MISSION As String = "Missions"
    Private Const KEY_SALARY As String = "Salaires"
    Private Const SUM_COL As Integer = 6
    Private ReadOnly SUM_COL_LETTER As String = Encoding.ASCII.GetString(New Byte() {64 + SUM_COL})
    Private Const LABEL_SUM As String = "Somme :"
    Private RecapNumber As String
    Private ReadOnly CDDMap As New Dictionary(Of Integer, List(Of BookLine))
    Private ReadOnly ImportantCells As New Dictionary(Of Integer, Dictionary(Of String, String))
    Private ReadOnly ExistingSheets As New Dictionary(Of String, Dictionary(Of String, Excel.Worksheet))
    Private ReadOnly HEADERS As New List(Of String) From {
    "Cpte gén.",
    "Rubrique de la pièce",
    "Numéro de flux",
    "Nom du tiers",
    "Libellé du flux",
    "MntEng.HTR",
    "Montant pa",
    "D.paiement",
    "Réf. FactF",
    "D. Piéce F",
    "D. compt.",
    "Nº pièce",
    "Commentaires",
    "Provenance de la ligne"
    }
    Private ReadOnly Property YearList As List(Of Integer)
        Get
            Dim yList As New List(Of Integer)
            yList.AddRange(Data.Keys)
            yList.Sort()
            Return yList
        End Get
    End Property
    Public Sub ExtractAndMerge()
        Data.Clear()
        CDDMap.Clear()
        ImportantCells.Clear()
        AllWorksheets.Clear()
        ExistingSheets.Clear()
        LinesWithComment.Clear()
        OutOfRangeComments.Clear()
        OutOfSumComments.Clear()
        OutOfTableComments.Clear()
        SummaryCellsNotFound.Clear()
        CurrentWorkbook = CreateCopy(CType(Me.Application.ActiveWorkbook, Excel.Workbook))
        SummaryWorkSheet = CType(CurrentWorkbook.Sheets(2), Excel.Worksheet)
        RecapNumber = SummaryWorkSheet.Range("A2").Value2
        GetExistingData()
        DoIntegration()
        ProduceWarnings()
        CurrentWorkbook.Save()
    End Sub
    Private Sub ProduceWarnings()
        If IsWorkbookWithWarnings() Then
            NameStep("Traitement des commentaires hors format")
            Dim warningSheet As Excel.Worksheet
            If ExistingSheets.ContainsKey(KEY_SHEET_WARNINGS) Then
                warningSheet = ExistingSheets.Item(KEY_SHEET_WARNINGS).Item(KEY_SHEET_WARNINGS)
                warningSheet.UsedRange.Rows.Delete(XlDeleteShiftDirection.xlShiftToLeft)
            Else
                warningSheet = CType(CurrentWorkbook.Worksheets.Add(After:=SummaryWorkSheet), Excel.Worksheet)
                warningSheet.Name = KEY_SHEET_WARNINGS
            End If
            Dim baseRange As Excel.Range = warningSheet.Range("A1")
            Dim currentLine As Integer = 1
            For Each year As Integer In Data.Keys
                If IsSheetWithWarnings(year) Then
                    currentLine = DumpWarnings(year, baseRange, currentLine) + 4
                End If
            Next
            If IsSummaryWithWarnings() Then
                DumpSummaryWarnings(baseRange, currentLine)
            End If
        ElseIf ExistingSheets.ContainsKey(KEY_SHEET_WARNINGS) Then
            Application.DisplayAlerts = False
            ExistingSheets.Item(KEY_SHEET_WARNINGS).Item(KEY_SHEET_WARNINGS).Delete()
            Application.DisplayAlerts = True
        End If
    End Sub
    Private Function DumpSummaryWarnings(baseRange As Range, currentLine As Integer) As Integer
        Utils.CreateWarning(Utils.GetMessage("RecapProblem"), baseRange, currentLine, "WarningHeaderStyle")
        currentLine += 1
        If SummaryCellsNotFound.Item(SummaryCellKind.TOTAL) Then
            currentLine += 1
            Utils.CreateWarning(Utils.GetMessage("TotalAmountProblem"), baseRange, currentLine, "WarningDetailStyle")
            currentLine += 1
        End If
        If SummaryCellsNotFound.Item(SummaryCellKind.TOTAL_NET) Then
            currentLine += 1
            Utils.CreateWarning(Utils.GetMessage("TotalNetAmountProblem"), baseRange, currentLine, "WarningDetailStyle")
            currentLine += 1
        End If
        If SummaryCellsNotFound.Item(SummaryCellKind.CUMUL) Then
            currentLine += 1
            Utils.CreateWarning(Utils.GetMessage("TotalCumulProblem"), baseRange, currentLine, "WarningDetailStyle")
            currentLine += 1
        End If
        If SummaryCellsNotFound.Item(SummaryCellKind.BUDGET) Then
            currentLine += 1
            Utils.CreateWarning(Utils.GetMessage("TotalBudgetProblem"), baseRange, currentLine, "WarningDetailStyle")
            currentLine += 1
        End If
        If SummaryCellsNotFound.Item(SummaryCellKind.ENGAGED) Then
            currentLine += 1
            Utils.CreateWarning(Utils.GetMessage("TotalEngagedProblem"), baseRange, currentLine, "WarningDetailStyle")
            currentLine += 1
        End If
        Return currentLine
    End Function



    Private Function DumpWarnings(year As Integer, baseRange As Range, currentLine As Integer) As Integer
        Utils.CreateWarning(String.Format(CultureInfo.CurrentCulture, Utils.GetMessage("YearProblem"), year), baseRange, currentLine, "WarningHeaderStyle")
        currentLine += 1
        Dim currentPbNum As Integer = 1
        If LinesWithComment.ContainsKey(year) AndAlso LinesWithComment.Item(year).Count > 0 Then
            currentLine += 1
            Utils.CreateWarning(String.Format(CultureInfo.CurrentCulture, Utils.GetMessage("Problem1"), currentPbNum), baseRange, currentLine, "WarningDetailStyle")
            currentLine += 1
            currentPbNum += 1
            currentLine = DumpLineThatDesapeared(year, baseRange, currentLine)
            currentLine += 1
        End If
        If OutOfRangeComments.Item(year) Then
            Utils.CreateWarning(String.Format(CultureInfo.CurrentCulture, Utils.GetMessage("Problem2"), currentPbNum), baseRange, currentLine, "WarningDetailStyle")
            currentLine += 2
            currentPbNum += 1
        End If
        If OutOfSumComments.Item(year) Then
            Utils.CreateWarning(String.Format(CultureInfo.CurrentCulture, Utils.GetMessage("Problem3"), currentPbNum), baseRange, currentLine, "WarningDetailStyle")
            currentLine += 2
            currentPbNum += 1
        End If
        If OutOfTableComments.Item(year) Then
            Utils.CreateWarning(String.Format(CultureInfo.CurrentCulture, Utils.GetMessage("Problem4"), currentPbNum), baseRange, currentLine, "WarningDetailStyle")
            currentLine += 2
        End If
        Return currentLine
    End Function

    Private Function DumpLineThatDesapeared(year As Integer, baseRange As Range, currentLine As Integer) As Integer
        For Each line As BookLine In LinesWithComment.Item(year)
            baseRange.Cells(currentLine, 1).Value2 = line.ACptegen
            baseRange.Cells(currentLine, 2).Value2 = line.BRubrique
            baseRange.Cells(currentLine, 3).Value2 = line.CNumeroFlux
            baseRange.Cells(currentLine, 4).Value2 = line.DNom
            baseRange.Cells(currentLine, 5).Value2 = line.ELibelle
            baseRange.Cells(currentLine, 6).Value2 = line.FMntEngHtr
            baseRange.Cells(currentLine, 7).Value2 = line.GMontantPA
            baseRange.Cells(currentLine, 8).Value2 = line.HRapprochmt
            baseRange.Cells(currentLine, 9).Value2 = line.IRefFactF
            baseRange.Cells(currentLine, 10).Value2 = line.JDatePce
            baseRange.Cells(currentLine, 11).Value2 = Utils.GetDateCompteAsText(line)
            baseRange.Cells(currentLine, 12).Value2 = line.LNumPiece
            baseRange.Cells(currentLine, 13).Value2 = line.MComment
            currentLine += 1
        Next

        Return currentLine
    End Function

    Private Function IsSheetWithWarnings(year As Integer) As Boolean
        If LinesWithComment.ContainsKey(year) AndAlso LinesWithComment.Item(year).Count > 0 Then
            Return True
        End If
        If OutOfRangeComments.ContainsKey(year) AndAlso OutOfRangeComments.Item(year) Then
            Return True
        End If
        If OutOfSumComments.ContainsKey(year) AndAlso OutOfSumComments.Item(year) Then
            Return True
        End If
        If OutOfTableComments.ContainsKey(year) AndAlso OutOfTableComments.Item(year) Then
            Return True
        End If
        Return False
    End Function

    Private Function IsWorkbookWithWarnings() As Boolean
        If IsSummaryWithWarnings() Then
            Return True
        End If
        For Each Year As Integer In Data.Keys
            If IsSheetWithWarnings(Year) Then
                Return True
            End If
        Next
        Return False
    End Function

    Private Function IsSummaryWithWarnings() As Boolean
        For Each missingTotal As Boolean In SummaryCellsNotFound.Values
            If missingTotal Then
                Return True
            End If
        Next
        Return False
    End Function
    Private Sub GetExistingData()
        ExistingSheets.Add(KEY_SHEET_MODE_EMPLOI, New Dictionary(Of String, Excel.Worksheet) From {
                {CType(CurrentWorkbook.Sheets(1), Excel.Worksheet).Name, CType(CurrentWorkbook.Sheets(1), Excel.Worksheet)}})
        ExistingSheets.Add(KEY_SHEET_RECAPITULATIF, New Dictionary(Of String, Excel.Worksheet) From {
                {CType(CurrentWorkbook.Sheets(2), Excel.Worksheet).Name, CType(CurrentWorkbook.Sheets(2), Excel.Worksheet)}})
        For I As Integer = 3 To CurrentWorkbook.Sheets.Count
            Dim worksheet As Worksheet = CType(CurrentWorkbook.Sheets(I), Excel.Worksheet)
            If worksheet.Name Like "####" Then
                If ExistingSheets.ContainsKey(KEY_SHEET_YEARS) Then
                    ExistingSheets.Item(KEY_SHEET_YEARS).Add(worksheet.Name, worksheet)
                Else
                    ExistingSheets.Add(KEY_SHEET_YEARS, New Dictionary(Of String, Excel.Worksheet) From {
                {worksheet.Name, worksheet}})
                End If
            ElseIf worksheet.Name = KEY_SHEET_WARNINGS Then
                If ExistingSheets.ContainsKey(KEY_SHEET_WARNINGS) Then
                    ExistingSheets.Item(KEY_SHEET_WARNINGS).Add(worksheet.Name, worksheet)
                Else
                    ExistingSheets.Add(KEY_SHEET_WARNINGS, New Dictionary(Of String, Excel.Worksheet) From {
                {KEY_SHEET_WARNINGS, worksheet}})
                End If
            Else
                If ExistingSheets.ContainsKey(KEY_SHEET_ELSE) Then
                    ExistingSheets.Item(KEY_SHEET_ELSE).Add(worksheet.Name, worksheet)
                Else
                    ExistingSheets.Add(KEY_SHEET_ELSE, New Dictionary(Of String, Excel.Worksheet) From {
                {worksheet.Name, worksheet}})
                End If
            End If
        Next
        If ExistingSheets.ContainsKey(KEY_SHEET_YEARS) Then
            Application.DisplayAlerts = False
            For Each yearSheet As Excel.Worksheet In ExistingSheets.Item(KEY_SHEET_YEARS).Values
                LinesWithComment.Add(CInt(yearSheet.Name), New List(Of BookLine))
                ExtractOldDataFromExistingSheet(yearSheet)
                yearSheet.Delete()
            Next
            Application.DisplayAlerts = True
        End If
    End Sub

    Private Sub ExtractOldDataFromExistingSheet(yearSheet As Worksheet)
        Dim FullRange As Excel.Range = yearSheet.UsedRange
        If Utils.IsNewHeaderVersion(FullRange) Then
            OutOfRangeComments.Add(CInt(yearSheet.Name), IsCommentOutOfRange(FullRange))
            OutOfTableComments.Add(CInt(yearSheet.Name), False)
            OutOfSumComments.Add(CInt(yearSheet.Name), False)
            Dim inTable As Boolean = False
            For Each cell As Excel.Range In FullRange.Rows
                If Not inTable AndAlso Utils.IsNewHeaderVersion(cell) Then
                    inTable = True
                ElseIf inTable AndAlso Utils.IsSum(cell) Then
                    inTable = False
                    If Utils.IsSumWithComment(cell) Then
                        OutOfSumComments.Item(CInt(yearSheet.Name)) = True
                    End If
                ElseIf inTable AndAlso Not (Utils.IsNewHeaderVersion(cell) OrElse IsEmptyLine(cell) OrElse Utils.IsSum(cell)) AndAlso Utils.IsLineWithComment(cell) Then
                    Dim newLine As BookLine = Utils.ReadLine(cell, 1)
                    LinesWithComment.Item(CInt(yearSheet.Name)).Add(newLine)
                    newLine.MComment = CStr(CType(cell.Cells(1, 13), Range).Value2)
                ElseIf Not inTable AndAlso Not IsEmptyLine(cell) Then
                    OutOfTableComments.Item(CInt(yearSheet.Name)) = True
                End If
            Next
        End If
    End Sub

    Private Function IsCommentOutOfRange(fullRange As Range) As Boolean
        Return fullRange.Columns.Count > HEADERS.Count
    End Function

    Private Function IsEmptyLine(cell As Range) As Boolean
        For I As Integer = 1 To HEADERS.Count
            If CStr(cell.Cells(1, I).value2) <> "" Then
                Return False
            End If
        Next
        Return True
    End Function
    Public Sub DoIntegration()
        Dim CDDWorkbook As Workbook = Nothing
        Dim dialog As Microsoft.Office.Core.FileDialog
        dialog = Application.FileDialog(Microsoft.Office.Core.MsoFileDialogType.msoFileDialogFilePicker)
            dialog.Filters.Add("fichier salaires CDD *.xlsx", "*.xlsx", 1)
            dialog.AllowMultiSelect = False
            If dialog.Show = -1 Then
                CDDWorkbook = GetCDDWorkbook(dialog.SelectedItems.Item(1))
            End If
        If CDDWorkbook IsNot Nothing Then
            ProgressDialog.Show()
            SetProgress(0)
            GetCDD(CDDWorkbook)
            CDDWorkbook.Close()
            SetProgress(PROGRESS_STEP_CDD * PROGRESS_REFERENCE)
            DoExtractAndMerge()
            SetProgress(PROGRESS_STEP_ASSEMBLE * PROGRESS_REFERENCE)
            PrepareRecap()
            SetProgress(PROGRESS_REFERENCE)
            CurrentWorkbook.Save()
            ProgressDialog.Hide()
        End If
    End Sub

    Private Sub PrepareRecap()
        NameStep("Préparation du récapitulatif")
        SummaryWorkSheet.Activate()
        SummaryWorkSheet.Range("B18:BB28").Rows.Delete(XlDeleteShiftDirection.xlShiftToLeft)

        Dim BaseRange As Excel.Range = SummaryWorkSheet.Range("B18")
        Dim TotalNetAddress As String = SearchAddressWithPattern("montant total net", SummaryCellKind.TOTAL_NET, True)
        Dim TotalAddress As String = SearchAddressWithPattern("montant total", SummaryCellKind.TOTAL, True)
        Dim CumulAddress As String = SearchAddressWithPattern("cumul", SummaryCellKind.CUMUL, False)
        Dim BudgetAddress As String = SearchAddressWithPattern("budget", SummaryCellKind.BUDGET, False)
        Dim EngagedAddress As String = SearchAddressWithPattern("engagé", SummaryCellKind.ENGAGED, False)

        DumpSummaryHeaders(BaseRange)
        Dim yList As List(Of Integer) = YearList
        Dim ClosedYearsStart As Range = SummaryWorkSheet.Range("$C$18")
        If yList.Count > 2 Then
            SummaryWorkSheet.Range(ClosedYearsStart.Address, ClosedYearsStart.Offset(0, yList.Count - 2).Address).Merge()
            Utils.SetAreaValue2(ClosedYearsStart, Utils.GetFormattedString("DepenseUntil", yList.Max - 1), "RecapHeaderStyle")
        ElseIf yList.Count = 2 Then
            Utils.SetAreaValue2(ClosedYearsStart, "", "RecapHeaderStyle")
        End If

        For Each Year As Integer In yList
            If Year = yList.Max Then
                DumpSummaryForLastYear(BaseRange, yList, Year)
            Else
                DumpSummaryForPreviousYear(BaseRange, yList, Year)
            End If
        Next
        Dim FirstCol As Integer = 1
        Dim LastClosedCol As Integer = yList.Count - 1
        Dim CurrentYearAmountCol As Integer = LastClosedCol + 2
        Dim CurrentCol As Integer = CurrentYearAmountCol + 3
        If FirstCol > LastClosedCol Then
            DumpSummaryFinalRecapWithNoPreviousYear(BaseRange, TotalNetAddress, TotalAddress, CumulAddress, BudgetAddress, EngagedAddress, yList, CurrentYearAmountCol, CurrentCol)
        Else
            DumpSummaryFinalRecapWithPreviousYear(BaseRange, TotalNetAddress, TotalAddress, CumulAddress, BudgetAddress, EngagedAddress, yList, FirstCol, LastClosedCol, CurrentYearAmountCol, CurrentCol)
        End If
    End Sub

    Private Sub DumpSummaryFinalRecapWithPreviousYear(BaseRange As Range, TotalNetAddress As String, TotalAddress As String, CumulAddress As String, BudgetAddress As String, EngagedAddress As String, yList As List(Of Integer), FirstCol As Integer, LastClosedCol As Integer, CurrentYearAmountCol As Integer, CurrentCol As Integer)
        DumpSummaryFinalRecapCommonParts(BaseRange, BudgetAddress, EngagedAddress, yList, CurrentYearAmountCol, CurrentCol)
        DumpSummaryFinalRecapSpecificParts(True, BaseRange, TotalNetAddress, TotalAddress, CumulAddress, FirstCol, LastClosedCol, CurrentYearAmountCol, CurrentCol)
    End Sub
    Private Sub DumpSummaryFinalRecapWithNoPreviousYear(BaseRange As Range, TotalNetAddress As String, TotalAddress As String, CumulAddress As String, BudgetAddress As String, EngagedAddress As String, yList As List(Of Integer), CurrentYearAmountCol As Integer, CurrentCol As Integer)
        DumpSummaryFinalRecapCommonParts(BaseRange, BudgetAddress, EngagedAddress, yList, CurrentYearAmountCol, CurrentCol)
        DumpSummaryFinalRecapSpecificParts(False, BaseRange, TotalNetAddress, TotalAddress, CumulAddress, 0, 0, CurrentYearAmountCol, CurrentCol)
    End Sub
    Private Sub DumpSummaryFinalRecapSpecificParts(fullRecap As Boolean, BaseRange As Range, TotalNetAddress As String, TotalAddress As String, CumulAddress As String, FirstCol As Integer, LastClosedCol As Integer, CurrentYearAmountCol As Integer, CurrentCol As Integer)
        If Not SummaryCellsNotFound.Item(SummaryCellKind.TOTAL) Then
            Dim specificFormula As String
            If fullRecap Then
                specificFormula = Utils.GetFormattedString("DirectDiff2", TotalAddress, BaseRange.Offset(7, CurrentYearAmountCol).Address(False, False), BaseRange.Offset(7, FirstCol).Address(False, False), BaseRange.Offset(7, LastClosedCol).Address(False, False))
            Else
                specificFormula = Utils.GetFormattedString("DirectDiff", TotalAddress, BaseRange.Offset(7, CurrentYearAmountCol).Address(False, False))
            End If
            Utils.SetCellFormula(BaseRange.Offset(7, CurrentCol), specificFormula, "RecapNumberStyle5")
        Else
            Utils.SetCellValue2(BaseRange.Offset(7, CurrentCol), "", "RecapNumberStyle5")
        End If
        If Not SummaryCellsNotFound.Item(SummaryCellKind.TOTAL_NET) Then
            Dim specificFormula As String
            If fullRecap Then
                specificFormula = Utils.GetFormattedString("DirectDiff2", TotalNetAddress, BaseRange.Offset(7, CurrentYearAmountCol).Address(False, False), BaseRange.Offset(7, FirstCol).Address(False, False), BaseRange.Offset(7, LastClosedCol).Address(False, False))
            Else
                specificFormula = Utils.GetFormattedString("DirectDiff", TotalNetAddress, BaseRange.Offset(7, CurrentYearAmountCol).Address(False, False))
            End If
            Utils.SetCellFormula(BaseRange.Offset(7, CurrentCol + 1), specificFormula, "RecapNumberStyle2")
        Else
            Utils.SetCellValue2(BaseRange.Offset(7, CurrentCol + 1), "", "RecapNumberStyle2")
        End If
        If Not SummaryCellsNotFound.Item(SummaryCellKind.CUMUL) Then
            Dim specificFormula As String
            If fullRecap Then
                specificFormula = Utils.GetFormattedString("DirectCellSum2", BaseRange.Offset(7, CurrentYearAmountCol + 1).Address(False, False), BaseRange.Offset(7, FirstCol).Address(False, False), BaseRange.Offset(7, LastClosedCol).Address(False, False))
            Else
                specificFormula = Utils.GetFormattedString("CellRef2", BaseRange.Offset(7, CurrentYearAmountCol + 1).Address(False, False))
            End If
            Utils.SetCellRawFormula(SummaryWorkSheet.Range(CumulAddress), specificFormula)
        End If
    End Sub

    Private Sub DumpSummaryFinalRecapCommonParts(BaseRange As Range, BudgetAddress As String, EngagedAddress As String, yList As List(Of Integer), CurrentYearAmountCol As Integer, CurrentCol As Integer)
        SummaryWorkSheet.Range(BaseRange.Offset(0, CurrentCol).Address, BaseRange.Offset(0, CurrentCol + 1).Address).Merge()
        Utils.SetAreaValue2(BaseRange.Offset(0, CurrentCol), Utils.GetFormattedString("HeaderOverYear", yList.Max), "RecapHeaderStyle")
        SummaryWorkSheet.Range(BaseRange.Offset(1, CurrentCol).Address, BaseRange.Offset(3, CurrentCol).Address).Merge()
        Utils.SetAreaValue2(BaseRange.Offset(1, CurrentCol), "Reste Total à dépenser (y compris frais de Gestion UPS)", "RecapHeaderStyle6")
        Utils.SetAreaValue2(BaseRange.Offset(4, CurrentCol), "", "RecapHeaderStyle6")
        Utils.SetAreaValue2(BaseRange.Offset(5, CurrentCol), "", "RecapHeaderStyle6")
        Utils.SetAreaValue2(BaseRange.Offset(6, CurrentCol), "", "RecapHeaderStyle6")
        BaseRange.Offset(6, CurrentCol).ColumnWidth = 16

        Utils.SetAreaValue2(BaseRange.Offset(0, CurrentCol + 1), "", "RecapHeaderStyle")
        SummaryWorkSheet.Range(BaseRange.Offset(1, CurrentCol + 1).Address, BaseRange.Offset(3, CurrentCol + 1).Address).Merge()
        Utils.SetAreaValue2(BaseRange.Offset(1, CurrentCol + 1), "Reste Net à dépenser (net des Frais de Gestion UPS)", "RecapHeaderStyle")
        BaseRange.Offset(1, CurrentCol + 1).MergeArea.Rows.RowHeight = BaseRange.Offset(0, CurrentCol + 1).RowHeight * 1.1
        Utils.SetAreaValue2(BaseRange.Offset(4, CurrentCol + 1), "", "RecapNumberStyle2")
        Utils.SetAreaValue2(BaseRange.Offset(5, CurrentCol + 1), "", "RecapNumberStyle2")
        Utils.SetAreaValue2(BaseRange.Offset(6, CurrentCol + 1), "", "RecapNumberStyle2")
        BaseRange.Offset(6, CurrentCol + 1).ColumnWidth = 16

        SummaryWorkSheet.Range(BaseRange.Offset(9, CurrentCol + 1).Address, BaseRange.Offset(9, CurrentCol + 2).Address).Merge()
        Utils.SetAreaValue2(BaseRange.Offset(9, CurrentCol + 1), "Montant total disponible :", "RecapHeaderStyle5")
        Utils.SetCellFormula(BaseRange.Offset(9, CurrentCol + 3), Utils.GetFormattedString("DirectCellSum3", BaseRange.Offset(7, CurrentYearAmountCol + 2).Address(False, False), BaseRange.Offset(7, CurrentCol + 1).Address(False, False)), "RecapNumberStyle2")
        If Not SummaryCellsNotFound.Item(SummaryCellKind.BUDGET) Then
            Utils.SetCellRawFormula(SummaryWorkSheet.Range(BudgetAddress), Utils.GetFormattedString("CellRef2", BaseRange.Offset(7, CurrentYearAmountCol).Address(False, False)))
        End If
        If Not SummaryCellsNotFound.Item(SummaryCellKind.ENGAGED) Then
            Utils.SetCellRawFormula(SummaryWorkSheet.Range(EngagedAddress), Utils.GetFormattedString("CellRef2", BaseRange.Offset(7, CurrentYearAmountCol + 1).Address(False, False)))
        End If
    End Sub

    Private Sub DumpSummaryHeaders(BaseRange As Range)
        Utils.SetCellValue2(BaseRange.Offset(0, 0), "", "RecapHeaderStyle3")
        SummaryWorkSheet.Range(BaseRange.Offset(1, 0).Address, BaseRange.Offset(3, 0).Address).Merge()
        Utils.SetAreaValue2(BaseRange.Offset(1, 0), "Masse", "RecapHeaderStyle3")
        Utils.SetCellValue2(BaseRange.Offset(4, 0), "1 Personnel", "RecapHeaderStyle4")
        Utils.SetCellValue2(BaseRange.Offset(5, 0), "2 Fonctionnement hors amort.", "RecapHeaderStyle4")
        Utils.SetCellValue2(BaseRange.Offset(6, 0), "3 Investissement", "RecapHeaderStyle4")
        Utils.SetCellValue2(BaseRange.Offset(7, 0), "Total", "RecapHeaderStyle5")
        BaseRange.Offset(7, 0).ColumnWidth = 26.5
    End Sub

    Private Sub DumpSummaryForPreviousYear(ByRef BaseRange As Range, yList As List(Of Integer), Year As Integer)
        Dim CurrentCol As Integer = yList.IndexOf(Year) + 1
        SummaryWorkSheet.Range(BaseRange.Offset(1, CurrentCol).Address, BaseRange.Offset(3, CurrentCol).Address).Merge()
        Utils.SetAreaValue2(BaseRange.Offset(1, CurrentCol), CStr(Year), "RecapHeaderStyle")
        If ImportantCells.Item(Year).ContainsKey(KEY_SALARY) Then
            Utils.SetCellFormula(BaseRange.Offset(4, CurrentCol), Utils.GetFormattedString("CellRef", Year, ImportantCells.Item(Year).Item(KEY_SALARY)), "RecapNumberStyle")
        Else
            Utils.SetCellValue(BaseRange.Offset(4, CurrentCol), 0, "RecapNumberStyle")
        End If
        If ImportantCells.Item(Year).ContainsKey(KEY_FONCT) AndAlso ImportantCells.Item(Year).ContainsKey(KEY_MISSION) Then
            Utils.SetCellFormula(BaseRange.Offset(5, CurrentCol), Utils.GetFormattedString("DirectCellSum", Year, ImportantCells.Item(Year).Item(KEY_FONCT), ImportantCells.Item(Year).Item(KEY_MISSION)), "RecapNumberStyle")
        ElseIf ImportantCells.Item(Year).ContainsKey(KEY_FONCT) Then
            Utils.SetCellFormula(BaseRange.Offset(5, CurrentCol), Utils.GetFormattedString("CellRef", Year, ImportantCells.Item(Year).Item(KEY_FONCT)), "RecapNumberStyle")
        ElseIf ImportantCells.Item(Year).ContainsKey(KEY_MISSION) Then
            Utils.SetCellFormula(BaseRange.Offset(5, CurrentCol), Utils.GetFormattedString("CellRef", Year, ImportantCells.Item(Year).Item(KEY_MISSION)), "RecapNumberStyle")
        Else
            Utils.SetCellValue(BaseRange.Offset(5, CurrentCol), 0, "RecapNumberStyle")
        End If
        If ImportantCells.Item(Year).ContainsKey(KEY_INVEST) Then
            Utils.SetCellFormula(BaseRange.Offset(6, CurrentCol), Utils.GetFormattedString("CellRef", Year, ImportantCells.Item(Year).Item(KEY_INVEST)), "RecapNumberStyle")
        Else
            Utils.SetCellValue(BaseRange.Offset(6, CurrentCol), 0, "RecapNumberStyle")
        End If
        Utils.SetCellFormula(BaseRange.Offset(7, CurrentCol), Utils.GetFormattedString("SumRange", BaseRange.Offset(4, CurrentCol).Address(False, False), BaseRange.Offset(6, CurrentCol).Address(False, False)), "RecapNumberStyle2")
        If yList.Count > 2 Then
            BaseRange.Offset(7, CurrentCol).ColumnWidth = 11.5
        Else
            BaseRange.Offset(7, CurrentCol).ColumnWidth = 26.5
        End If
    End Sub

    Private Sub DumpSummaryForLastYear(ByRef baseRange As Range, yList As List(Of Integer), year As Integer)
        Dim currentCol As Integer = yList.Count
        SummaryWorkSheet.Range(baseRange.Offset(0, currentCol).Address, baseRange.Offset(0, currentCol + 3).Address).Merge()
        Utils.SetAreaValue2(baseRange.Offset(0, currentCol), year, "RecapHeaderStyle2")
        SummaryWorkSheet.Range(baseRange.Offset(1, currentCol).Address, baseRange.Offset(3, currentCol).Address).Merge()
        Utils.SetAreaValue2(baseRange.Offset(1, currentCol), Utils.GetFormattedString("InitialBudget", year), "RecapHeaderStyle2")
        Utils.SetCellValue(baseRange.Offset(4, currentCol), 0, "RecapNumberStyle3")
        Utils.SetCellValue(baseRange.Offset(5, currentCol), 0, "RecapNumberStyle3")
        Utils.SetCellValue(baseRange.Offset(6, currentCol), 0, "RecapNumberStyle3")
        Utils.SetCellFormula(baseRange.Offset(7, currentCol), Utils.GetFormattedString("SumRange", baseRange.Offset(4, currentCol).Address(False, False), baseRange.Offset(6, currentCol).Address(False, False)), "RecapNumberStyle3")
        baseRange.Offset(7, currentCol).ColumnWidth = 11.5

        SummaryWorkSheet.Range(baseRange.Offset(1, currentCol + 1).Address, baseRange.Offset(3, currentCol + 1).Address).Merge()
        Utils.SetAreaValue2(baseRange.Offset(1, currentCol + 1), Utils.GetFormattedString("ModifiedBudget", year), "RecapHeaderStyle2")
        Utils.SetCellValue(baseRange.Offset(4, currentCol + 1), 0, "RecapNumberStyle4")
        Utils.SetCellValue(baseRange.Offset(5, currentCol + 1), 0, "RecapNumberStyle4")
        Utils.SetCellValue(baseRange.Offset(6, currentCol + 1), 0, "RecapNumberStyle4")
        Utils.SetCellFormula(baseRange.Offset(7, currentCol + 1), Utils.GetFormattedString("SumRange", baseRange.Offset(4, currentCol + 1).Address(False, False), baseRange.Offset(6, currentCol + 1).Address(False, False)), "RecapNumberStyle4")
        baseRange.Offset(7, currentCol + 1).ColumnWidth = 11.5

        SummaryWorkSheet.Range(baseRange.Offset(1, currentCol + 2).Address, baseRange.Offset(3, currentCol + 2).Address).Merge()
        Utils.SetAreaValue2(baseRange.Offset(1, currentCol + 2), Utils.GetFormattedString("EngagedBudget", year), "RecapHeaderStyle2")
        WriteSalaryForLastYear(baseRange, year, currentCol)
        WriteFonctForLastYear(baseRange, year, currentCol)
        WriteInvestForLastYear(baseRange, year, currentCol)
        Utils.SetCellFormula(baseRange.Offset(7, currentCol + 2), Utils.GetFormattedString("SumRange", baseRange.Offset(4, currentCol + 2).Address(False, False), baseRange.Offset(6, currentCol + 2).Address(False, False)), "RecapNumberStyle3")
        baseRange.Offset(7, currentCol + 2).ColumnWidth = 11.5

        SummaryWorkSheet.Range(baseRange.Offset(1, currentCol + 3).Address, baseRange.Offset(3, currentCol + 3).Address).Merge()
        Utils.SetAreaValue2(baseRange.Offset(1, currentCol + 3), Utils.GetFormattedString("AvailableBudget", year), "RecapHeaderStyle2")
        Utils.SetCellFormula(baseRange.Offset(4, currentCol + 3), Utils.GetFormattedString("DirectDiff", baseRange.Offset(4, currentCol + 1).Address(False, False), baseRange.Offset(4, currentCol + 2).Address(False, False)), "RecapNumberStyle4")
        Utils.SetCellFormula(baseRange.Offset(5, currentCol + 3), Utils.GetFormattedString("DirectDiff", baseRange.Offset(5, currentCol + 1).Address(False, False), baseRange.Offset(5, currentCol + 2).Address(False, False)), "RecapNumberStyle4")
        Utils.SetCellFormula(baseRange.Offset(6, currentCol + 3), Utils.GetFormattedString("DirectDiff", baseRange.Offset(6, currentCol + 1).Address(False, False), baseRange.Offset(6, currentCol + 2).Address(False, False)), "RecapNumberStyle4")
        Utils.SetCellFormula(baseRange.Offset(7, currentCol + 3), Utils.GetFormattedString("DirectDiff", baseRange.Offset(7, currentCol + 1).Address(False, False), baseRange.Offset(7, currentCol + 2).Address(False, False)), "RecapNumberStyle4")
        baseRange.Offset(7, currentCol + 3).ColumnWidth = 11.5
    End Sub

    Private Sub WriteInvestForLastYear(baseRange As Range, year As Integer, currentCol As Integer)
        Dim isInvest As Boolean = ImportantCells.Item(year).ContainsKey(KEY_INVEST)
        If isInvest Then
            Utils.SetCellFormula(baseRange.Offset(6, currentCol + 2), Utils.GetFormattedString("CellRef", year, ImportantCells.Item(year).Item(KEY_INVEST)), "RecapNumberStyle3")
        Else
            Utils.SetCellValue(baseRange.Offset(6, currentCol + 2), 0, "RecapNumberStyle3")
        End If
    End Sub

    Private Sub WriteFonctForLastYear(baseRange As Range, year As Integer, currentCol As Integer)
        Dim isFonct As Boolean = ImportantCells.Item(year).ContainsKey(KEY_FONCT)
        Dim isMission As Boolean = ImportantCells.Item(year).ContainsKey(KEY_MISSION)
        If isFonct AndAlso isMission Then
            Utils.SetCellFormula(baseRange.Offset(5, currentCol + 2), Utils.GetFormattedString("DirectCellSum", year, ImportantCells.Item(year).Item(KEY_FONCT), ImportantCells.Item(year).Item(KEY_MISSION)), "RecapNumberStyle3")
        ElseIf isFonct Then
            Utils.SetCellFormula(baseRange.Offset(5, currentCol + 2), Utils.GetFormattedString("CellRef", year, ImportantCells.Item(year).Item(KEY_FONCT)), "RecapNumberStyle3")
        ElseIf isMission Then
            Utils.SetCellFormula(baseRange.Offset(5, currentCol + 2), Utils.GetFormattedString("CellRef", year, ImportantCells.Item(year).Item(KEY_MISSION)), "RecapNumberStyle3")
        Else
            Utils.SetCellValue(baseRange.Offset(5, currentCol + 2), 0, "RecapNumberStyle3")
        End If
    End Sub

    Private Sub WriteSalaryForLastYear(baseRange As Range, year As Integer, currentCol As Integer)
        If ImportantCells.Item(year).ContainsKey(KEY_SALARY) Then
            Utils.SetCellFormula(baseRange.Offset(4, currentCol + 2), Utils.GetFormattedString("CellRef", year, ImportantCells.Item(year).Item(KEY_SALARY)), "RecapNumberStyle3")
        Else
            Utils.SetCellValue(baseRange.Offset(4, currentCol + 2), 0, "RecapNumberStyle3")
        End If
    End Sub
    Private Function SearchAddressWithPattern(v As String, kind As SummaryCellKind, exact As Boolean) As String
        For Each cell As Excel.Range In SummaryWorkSheet.Range("A1:A18").Cells
            If exact AndAlso String.Equals(v, cell.Value2, StringComparison.OrdinalIgnoreCase) OrElse Not exact AndAlso cell.Value2 IsNot Nothing AndAlso CStr(cell.Value2).StartsWith(v, StringComparison.OrdinalIgnoreCase) Then
                SummaryCellsNotFound.Add(kind, False)
                Return cell.Offset(0, 1).Address
            End If
        Next
        SummaryCellsNotFound.Add(kind, True)
        Return Nothing
    End Function
    Private Sub GetCDD(cddWorkbook As Workbook)
        NameStep("Récupération des CDD")
        ProgressIncrement = PROGRESS_STEP_CDD / 2
        Dim CDDWorksheet As Worksheet = CType(cddWorkbook.Sheets.Item(2), Excel.Worksheet)
        Dim DataRange As Excel.Range = CDDWorksheet.UsedRange
        Dim FirstRow As Integer = 7
        Dim LastRow As Integer = DataRange.Rows.Count
        NextStep()
        ProgressIncrement = PROGRESS_STEP_CDD / 2 / LastRow

        For NumRow As Integer = FirstRow To LastRow
            Dim TempRecapNum As String = CDDWorksheet.Range(String.Format(CultureInfo.CurrentCulture, "Q{0}", NumRow)).Value2
            If String.Equals(RecapNumber, TempRecapNum, StringComparison.OrdinalIgnoreCase) Then
                Dim NewLine As New BookLine With {
                .ACptegen = "",
                .BRubrique = "SALAIRE",
                .CNumeroFlux = CDDWorksheet.Range(String.Format(CultureInfo.CurrentCulture, "B{0}", NumRow)).Value2,
                .DNom = CDDWorksheet.Range(String.Format(CultureInfo.CurrentCulture, "I{0}", NumRow)).Value2,
                .ELibelle = Utils.GetFormattedString("LiasseDate", Format(CDDWorksheet.Range(String.Format(CultureInfo.CurrentCulture, "D{0}", NumRow)).Value, "dd/MM/yyyy"), Format(CDDWorksheet.Range(String.Format(CultureInfo.CurrentCulture, "E{0}", NumRow)).Value, "dd/MM/yyyy")),
                .FMntEngHtr = CDbl(CDDWorksheet.Range(String.Format(CultureInfo.CurrentCulture, "N{0}", NumRow)).Value2),
                .GMontantPA = 0,
                .HRapprochmt = "",
                .IRefFactF = "",
                .JDatePce = "",
                .KDCompt = Nothing,
                .LNumPiece = "",
                .MComment = "",
                .NFrom = ""
            }
                Dim Year As Integer = CInt(CDDWorksheet.Range(String.Format(CultureInfo.CurrentCulture, "C{0}", NumRow)).Value2)
                If CDDMap.ContainsKey(Year) Then
                    CDDMap.Item(Year).Add(NewLine)
                Else
                    Dim NewList As New List(Of BookLine) From {
                    NewLine
                }
                    CDDMap.Add(Year, NewList)
                End If
            End If
            NextStep()
        Next
    End Sub

    Private Function GetCDDWorkbook(fileName As String) As Workbook
        Dim CDDWorkbook As Excel.Workbook = Nothing
        For Each wb As Excel.Workbook In Me.Application.Workbooks
            If String.Equals(fileName, wb.FullName) Then
                CDDWorkbook = wb
            End If
        Next
        If CDDWorkbook Is Nothing Then
            CDDWorkbook = Me.Application.Workbooks.Open(fileName, [ReadOnly]:=True)
        End If
        Return CDDWorkbook
    End Function

    ''' <summary>
    ''' Extracts Data from Excel files and merge them into an EDC.
    ''' </summary>
    Private Sub DoExtractAndMerge()
        ProgressDialog.ProgressTraitement.Maximum = CInt(PROGRESS_REFERENCE)
        ProgressDialog.ProgressTraitement.Minimum = 0
        ProgressDialog.ProgressTraitement.Value = 0
        CurrentProgrees = 0
        BaseDirectory = Path.GetDirectoryName(CurrentWorkbook.FullName)
        SourcesDirectory = String.Format(CultureInfo.InvariantCulture, "{0}{1}", BaseDirectory, SOURCE_DIRECTORY_MODIFIER)
        ExtractionDirectory = String.Format(CultureInfo.InvariantCulture, "{0}{1}", BaseDirectory, EXTRACT_DIRECTORY_MODIFIER)
        StyleUtils.PrepareStyles(CurrentWorkbook)
        Dim Extractions As New List(Of ExtractedData)
        If Not Directory.Exists(ExtractionDirectory) Then
            Directory.CreateDirectory(ExtractionDirectory)
        Else
            For Each FileName In Directory.EnumerateFiles(ExtractionDirectory)
                File.Delete(FileName)
            Next
        End If
        Dim TotalNumberOfLines = 0
        NameStep("Préparation des fichiers")
        Try
            Dim ExcelFiles As String() = Directory.GetFiles(SourcesDirectory, FILE_NAME_PATTERN)
            ProgressIncrement = PROGRESS_STEP_CREATE_FILES * PROGRESS_REFERENCE / ExcelFiles.Length
            For Each Name As String In ExcelFiles
                Dim extracted As ExtractedData = ExtractDataFromFile(Name)
                Extractions.Add(extracted)
                TotalNumberOfLines += extracted.NumberOfLines
                NextStep()
            Next
        Catch FileException As System.IO.DirectoryNotFoundException
            Debug.WriteLine("No directory found: " & FileException.Message)
        End Try
        NameStep("Extraction des données")
        ProgressIncrement = (PROGRESS_STEP_READ_FILES - PROGRESS_STEP_CREATE_FILES) * PROGRESS_REFERENCE / (TotalNumberOfLines * 9)
        Dim TotalNumberOfLinesToRecap As Integer = 0
        For Each Extraction As ExtractedData In Extractions
            Extraction.DoPrepareExtract()
            Data.Add(Extraction.SheetYear, Extraction)
        Next
        Dim YearList As New List(Of Integer)
        YearList.AddRange(Data.Keys)
        YearList.Sort()
        For Each Year As Integer In YearList
            Dim Extraction As ExtractedData = Data.Item(Year)
            If Year = YearList.Min Then
                Extraction.DoExtract(Nothing)
            Else
                Dim PreviousExtraction As ExtractedData = Data.Item(Year - 1)
                Extraction.DoExtract(PreviousExtraction)
            End If
            TotalNumberOfLinesToRecap += Extraction.Orders.Count
            TotalNumberOfLinesToRecap += Extraction.Missions.Count
            TotalNumberOfLinesToRecap += Extraction.PendingOrders.Count
            TotalNumberOfLinesToRecap += Extraction.PendingMissions.Count
        Next
        NameStep("Assemblage des données")
        ProgressIncrement = (PROGRESS_STEP_ASSEMBLE - PROGRESS_STEP_READ_FILES) * PROGRESS_REFERENCE / TotalNumberOfLinesToRecap
        For Each Year As Integer In YearList
            CreateSheetForYear(Year, Year = YearList.Min, Year = YearList.Max)
        Next
    End Sub
    Private Function CreateCopy(aWorkbook As Excel.Workbook) As Excel.Workbook
        Dim NewFileName As String = String.Format(CultureInfo.InvariantCulture, "{0}{1}{2} - New version {3}", Path.GetDirectoryName(aWorkbook.FullName), Path.DirectorySeparatorChar, Path.GetFileNameWithoutExtension(aWorkbook.FullName), DEST_NAME_EXTENSION)
        aWorkbook.SaveCopyAs(NewFileName)
        Dim NewWorkbook As Excel.Workbook = Me.Application.Workbooks.Open(NewFileName)
        Application.DisplayAlerts = False
        aWorkbook.Close()
        Application.DisplayAlerts = True
        Return NewWorkbook
    End Function

    Private Sub CreateSheetForYear(year As Integer, firstYear As Boolean, lastYear As Boolean)
        Dim NewWorsheet As Excel.Worksheet
        If firstYear Then
            'No pending order
            NewWorsheet = CType(CurrentWorkbook.Worksheets.Add(After:=SummaryWorkSheet), Excel.Worksheet)
            NewWorsheet.Name = year
            AllWorksheets.Add(year, NewWorsheet)
            FeedWorkSheet(NewWorsheet, year)
        ElseIf Not lastYear Then
            'Potential pending orders
            NewWorsheet = CType(CurrentWorkbook.Worksheets.Add(Before:=AllWorksheets.Item(year - 1)), Excel.Worksheet)
            NewWorsheet.Name = year
            AllWorksheets.Add(year, NewWorsheet)
            FeedWorkSheetWithPendings(NewWorsheet, year)
        Else
            'Potential pending orders past and present years
            NewWorsheet = CType(CurrentWorkbook.Worksheets.Add(Before:=AllWorksheets.Item(year - 1)), Excel.Worksheet)
            NewWorsheet.Name = year
            AllWorksheets.Add(year, NewWorsheet)
            FeedWorkSheetWithAllPendings(NewWorsheet, year)
        End If
        Utils.AutoFit(NewWorsheet)
    End Sub
    Private Sub FeedWorkSheet(newWorsheet As Excel.Worksheet, year As Integer)
        Dim MergedData As New Dictionary(Of String, List(Of BookLine)) From {
        {KEY_FONCT, New List(Of BookLine)},
        {KEY_INVEST, New List(Of BookLine)},
        {KEY_MISSION, New List(Of BookLine)},
        {KEY_SALARY, New List(Of BookLine)}
    }
        MergedData.Item(KEY_FONCT).AddRange(Data.Item(year).Orders)
        MergedData.Item(KEY_INVEST).AddRange(Data.Item(year).Invests)
        MergedData.Item(KEY_MISSION).AddRange(Data.Item(year).Missions)
        If CDDMap.ContainsKey(year) Then
            MergedData.Item(KEY_SALARY).AddRange(CDDMap.Item(year))
            MergedData.Item(KEY_SALARY).Sort(New SalaryLineComparison)
        End If
        Dump(MergedData, newWorsheet, year, False)
    End Sub

    Private Sub FeedWorkSheetWithPendings(newWorsheet As Excel.Worksheet, year As Integer)
        Dim MergedData As New Dictionary(Of String, List(Of BookLine)) From {
        {KEY_FONCT, New List(Of BookLine)},
        {KEY_INVEST, New List(Of BookLine)},
        {KEY_MISSION, New List(Of BookLine)},
        {KEY_SALARY, New List(Of BookLine)}
    }
        MergedData.Item(KEY_FONCT).AddRange(Data.Item(year - 1).PendingOrders)
        MergedData.Item(KEY_FONCT).AddRange(Data.Item(year).Orders)
        MergedData.Item(KEY_INVEST).AddRange(Data.Item(year - 1).PendingInvests)
        MergedData.Item(KEY_INVEST).AddRange(Data.Item(year).Invests)
        MergedData.Item(KEY_MISSION).AddRange(Data.Item(year - 1).PendingMissions)
        MergedData.Item(KEY_MISSION).AddRange(Data.Item(year).Missions)
        If CDDMap.ContainsKey(year) Then
            MergedData.Item(KEY_SALARY).AddRange(CDDMap.Item(year))
            MergedData.Item(KEY_SALARY).Sort(New SalaryLineComparison)
        End If
        Dump(MergedData, newWorsheet, year, False)
    End Sub
    Private Sub FeedWorkSheetWithAllPendings(newWorsheet As Excel.Worksheet, year As Integer)
        Dim MergedData As New Dictionary(Of String, List(Of BookLine)) From {
        {KEY_FONCT, New List(Of BookLine)},
        {KEY_INVEST, New List(Of BookLine)},
        {KEY_MISSION, New List(Of BookLine)},
        {KEY_SALARY, New List(Of BookLine)}}

        MergedData.Item(KEY_FONCT).AddRange(Data.Item(year - 1).PendingOrders)
        MergedData.Item(KEY_FONCT).AddRange(Data.Item(year).PendingOrders)
        MergedData.Item(KEY_FONCT).AddRange(Data.Item(year).Orders)
        MergedData.Item(KEY_INVEST).AddRange(Data.Item(year - 1).PendingInvests)
        MergedData.Item(KEY_INVEST).AddRange(Data.Item(year).PendingInvests)
        MergedData.Item(KEY_INVEST).AddRange(Data.Item(year).Invests)
        MergedData.Item(KEY_MISSION).AddRange(Data.Item(year - 1).PendingMissions)
        MergedData.Item(KEY_MISSION).AddRange(Data.Item(year).PendingMissions)
        MergedData.Item(KEY_MISSION).AddRange(Data.Item(year).Missions)
        If CDDMap.ContainsKey(year) Then
            MergedData.Item(KEY_SALARY).AddRange(CDDMap.Item(year))
            MergedData.Item(KEY_SALARY).Sort(New SalaryLineComparison)
        End If
        Dump(MergedData, newWorsheet, year, True)
    End Sub
    Private Sub Dump(mergedData As Dictionary(Of String, List(Of BookLine)), newWorsheet As Excel.Worksheet, year As Integer, fullHeader As Boolean)
        ImportantCells.Add(year, New Dictionary(Of String, String))
        Dim CurrentLine As Integer = 1
        Dim StartRange As Excel.Range = newWorsheet.Range("A1")
        For Each Key In mergedData.Keys
            Dim LineList As List(Of BookLine) = mergedData.Item(Key)
            If LineList.Count > 0 Then
                CurrentLine += 1
                DumpHeaders(StartRange, CurrentLine, fullHeader)
                CurrentLine += 1
                Dim FirstLine As Integer = CurrentLine
                For Each Line As BookLine In LineList
                    StartRange.Cells(CurrentLine, 1).Value2 = Line.ACptegen
                    StartRange.Cells(CurrentLine, 2).Value2 = Line.BRubrique
                    StartRange.Cells(CurrentLine, 3).Value2 = Line.CNumeroFlux
                    StartRange.Cells(CurrentLine, 4).Value2 = Line.DNom
                    StartRange.Cells(CurrentLine, 5).Value2 = Line.ELibelle
                    If Line.FMntEngHtr = 0 Then
                        Line.FMntEngHtr = Line.GMontantPA
                    End If
                    StartRange.Cells(CurrentLine, 6).Value2 = Line.FMntEngHtr
                    CType(StartRange.Cells(CurrentLine, 6), Excel.Range).Style = "MtEngStyle"
                    StartRange.Cells(CurrentLine, 7).Value2 = Line.GMontantPA
                    CType(StartRange.Cells(CurrentLine, 7), Excel.Range).Style = "MtPAStyle"
                    StartRange.Cells(CurrentLine, 8).Value2 = Line.HRapprochmt
                    StartRange.Cells(CurrentLine, 9).Value2 = Line.IRefFactF
                    StartRange.Cells(CurrentLine, 10).Value2 = Line.JDatePce
                    StartRange.Cells(CurrentLine, 11).Value2 = Utils.GetDateCompteAsText(Line)
                    StartRange.Cells(CurrentLine, 12).Value2 = Line.LNumPiece
                    AddPossibleCommentToLine(Line)
                    StartRange.Cells(CurrentLine, 13).Value2 = Line.MComment
                    CType(StartRange.Cells(CurrentLine, 13), Excel.Range).Style = "SIFACCommentaires"
                    StartRange.Cells(CurrentLine, 14).Value2 = Line.NFrom
                    CurrentLine += 1
                    Globals.ThisAddIn.NextStep()
                Next
                Dim LastLine As Integer = CurrentLine - 1
                StartRange.Cells(CurrentLine, SUM_COL - 1).Value2 = LABEL_SUM
                CType(StartRange.Cells(CurrentLine, SUM_COL), Excel.Range).Formula = Utils.GetFormattedString("SumRange", Utils.GetFormattedString("CellAddress", SUM_COL_LETTER, FirstLine), Utils.GetFormattedString("CellAddress", SUM_COL_LETTER, LastLine))
                CType(StartRange.Cells(CurrentLine, SUM_COL), Excel.Range).Style = "SumStyle"
                ImportantCells.Item(year).Add(Key, CType(StartRange.Cells(CurrentLine, SUM_COL), Excel.Range).Address(False, False))
                CurrentLine += 1
            End If
        Next
    End Sub

    Private Sub AddPossibleCommentToLine(line As BookLine)
        If line.BRubrique <> "SALAIRE" Then
            If line.MComment <> "" Then
                Return
            End If
            For Each year As Integer In LinesWithComment.Keys
                For Each commentedLine As BookLine In LinesWithComment.Item(year)
                    If BookLine.Equals(line, commentedLine) Then
                        line.MComment = commentedLine.MComment
                        LinesWithComment.Item(year).Remove(commentedLine)
                        Return
                    End If
                Next
            Next
        Else
            For Each year As Integer In LinesWithComment.Keys
                For Each commentedLine As BookLine In LinesWithComment.Item(year)
                    If commentedLine.BRubrique = "SALAIRE" AndAlso commentedLine.CNumeroFlux = line.CNumeroFlux Then
                        line.MComment = commentedLine.MComment
                        LinesWithComment.Item(year).Remove(commentedLine)
                        Return
                    End If
                Next
            Next
        End If
    End Sub

    Private Sub DumpHeaders(startRange As Excel.Range, currentLine As Integer, fullHeader As Boolean)
        For NumCol As Integer = 1 To HEADERS.Count - 2
            Dim Cell As Excel.Range = CType(startRange.Cells(currentLine, NumCol), Excel.Range)
            Cell.Value2 = HEADERS.Item(NumCol - 1)
            Cell.Style = "HeaderStyle"
        Next
        Dim CellComment As Excel.Range = CType(startRange.Cells(currentLine, HEADERS.Count - 1), Excel.Range)
        CellComment.Value2 = HEADERS.Item(HEADERS.Count - 2)
        CellComment.Style = "HeaderStyleComment"
        If fullHeader Then
            Dim CellFrom As Excel.Range = CType(startRange.Cells(currentLine, HEADERS.Count), Excel.Range)
            CellFrom.Value2 = HEADERS.Item(HEADERS.Count - 1)
            CellFrom.Style = "HeaderStyleFrom"
        End If
    End Sub

    Public Sub NextStep()
        CurrentProgrees += ProgressIncrement
        ProgressDialog.ProgressTraitement.Value = Math.Min(CInt(CurrentProgrees), PROGRESS_REFERENCE)
    End Sub
    Public Sub SetProgress(progressValue As Double)
        CurrentProgrees = progressValue
        ProgressDialog.ProgressTraitement.Value = CInt(CurrentProgrees)
    End Sub
    Public Sub NameStep(stepName As String)
        ProgressDialog.LblPhase.Text = stepName
        ProgressDialog.LblPhase.Refresh()
    End Sub
    Private Function ExtractDataFromFile(name As String) As ExtractedData
        Dim NewWorkbookPath As String = String.Format(CultureInfo.InvariantCulture, "{0}{1}{2}", ExtractionDirectory, Path.GetFileNameWithoutExtension(name), DEST_NAME_EXTENSION)
        Return New ExtractedData(name, NewWorkbookPath, Me.Application)
    End Function

    Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
        Return Globals.Factory.GetRibbonFactory().CreateRibbonManager(New Ribbon.IRibbonExtension() {New RibbonEdc()})
    End Function
End Class
