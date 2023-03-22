Imports System.Diagnostics
Imports System.Drawing
Imports System.Globalization
Imports System.IO
Imports Microsoft.Office.Interop.Excel
Imports SIFACToEDC.ExtractedData
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
    Private Class SalaryLineComparison : Implements IComparer(Of BookLine)
        Public Function Compare(x As BookLine, y As BookLine) As Integer Implements IComparer(Of BookLine).Compare
            If x IsNot Nothing Then
                If y IsNot Nothing Then
                    Return String.CompareOrdinal(x.DNom, y.DNom)
                Else
                    Return String.CompareOrdinal(x.DNom, Nothing)
                End If
            Else
                If y IsNot Nothing Then
                    Return String.CompareOrdinal(Nothing, y.DNom)
                Else
                    Return 0
                End If
            End If
        End Function
    End Class
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
        CreateWarning(My.Resources.ResourceManager.GetString("RecapProblem", CultureInfo.CurrentCulture), baseRange, currentLine, "WarningHeaderStyle")
        currentLine += 1
        If SummaryCellsNotFound.Item(SummaryCellKind.TOTAL) Then
            currentLine += 1
            CreateWarning(My.Resources.ResourceManager.GetString("TotalAmountProblem", CultureInfo.CurrentCulture), baseRange, currentLine, "WarningDetailStyle")
            currentLine += 1
        End If
        If SummaryCellsNotFound.Item(SummaryCellKind.TOTAL_NET) Then
            currentLine += 1
            CreateWarning(My.Resources.ResourceManager.GetString("TotalNetAmountProblem", CultureInfo.CurrentCulture), baseRange, currentLine, "WarningDetailStyle")
            currentLine += 1
        End If
        If SummaryCellsNotFound.Item(SummaryCellKind.CUMUL) Then
            currentLine += 1
            CreateWarning(My.Resources.ResourceManager.GetString("TotalCumulProblem", CultureInfo.CurrentCulture), baseRange, currentLine, "WarningDetailStyle")
            currentLine += 1
        End If
        If SummaryCellsNotFound.Item(SummaryCellKind.BUDGET) Then
            currentLine += 1
            CreateWarning(My.Resources.ResourceManager.GetString("TotalBudgetProblem", CultureInfo.CurrentCulture), baseRange, currentLine, "WarningDetailStyle")
            currentLine += 1
        End If
        If SummaryCellsNotFound.Item(SummaryCellKind.ENGAGED) Then
            currentLine += 1
            CreateWarning(My.Resources.ResourceManager.GetString("TotalEngagedProblem", CultureInfo.CurrentCulture), baseRange, currentLine, "WarningDetailStyle")
            currentLine += 1
        End If
        Return currentLine
    End Function

    Private Function DumpWarnings(year As Integer, baseRange As Range, currentLine As Integer) As Integer
        CreateWarning(String.Format(CultureInfo.CurrentCulture, My.Resources.ResourceManager.GetString("YearProblem"), year), baseRange, currentLine, "WarningHeaderStyle")
        currentLine += 1
        Dim currentPbNum As Integer = 1
        If LinesWithComment.ContainsKey(year) AndAlso LinesWithComment.Item(year).Count > 0 Then
            currentLine += 1
            CreateWarning(String.Format(CultureInfo.CurrentCulture, My.Resources.ResourceManager.GetString("Problem1"), currentPbNum), baseRange, currentLine, "WarningDetailStyle")
            currentLine += 1
            currentPbNum += 1
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
                baseRange.Cells(currentLine, 11).Value2 = ExtractedData.GetDateCompteAsText(line)
                baseRange.Cells(currentLine, 12).Value2 = line.LNumPiece
                baseRange.Cells(currentLine, 13).Value2 = line.MComment
                currentLine += 1
            Next
            currentLine += 1
        End If
        If OutOfRangeComments.Item(year) Then
            CreateWarning(String.Format(CultureInfo.CurrentCulture, My.Resources.ResourceManager.GetString("Problem2"), currentPbNum), baseRange, currentLine, "WarningDetailStyle")
            currentLine += 2
            currentPbNum += 1
        End If
        If OutOfSumComments.Item(year) Then
            CreateWarning(String.Format(CultureInfo.CurrentCulture, My.Resources.ResourceManager.GetString("Problem3"), currentPbNum), baseRange, currentLine, "WarningDetailStyle")
            currentLine += 2
            currentPbNum += 1
        End If
        If OutOfTableComments.Item(year) Then
            CreateWarning(String.Format(CultureInfo.CurrentCulture, My.Resources.ResourceManager.GetString("Problem4"), currentPbNum), baseRange, currentLine, "WarningDetailStyle")
            currentLine += 2
        End If
        Return currentLine
    End Function

    Private Shared Sub CreateWarning(message As String, baseRange As Range, currentLine As Integer, style As String)
        Dim startRange As Excel.Range = baseRange.Cells(currentLine, 1)
        Dim startAddress As String = startRange.Address
        Dim endAddress As String = startRange.Offset(0, 13).Address
        Dim mergedRange As Range = baseRange.Range(startAddress, endAddress)
        'Debug.WriteLine($"{baseRange.Address} + {currentLine} = {startRange.Address} and {mergedRange.Address}")
        mergedRange.Merge()
        mergedRange.Value2 = message
        mergedRange.Style = style
    End Sub

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
        If IsNewHeaderVersion(FullRange) Then
            OutOfRangeComments.Add(CInt(yearSheet.Name), IsCommentOutOfRange(FullRange))
            OutOfTableComments.Add(CInt(yearSheet.Name), False)
            OutOfSumComments.Add(CInt(yearSheet.Name), False)
            Dim inTable As Boolean = False
            For Each cell As Excel.Range In FullRange.Rows
                If Not inTable AndAlso IsNewHeaderVersion(cell) Then
                    inTable = True
                ElseIf inTable AndAlso IsSum(cell) Then
                    inTable = False
                    If IsSumWithComment(cell) Then
                        OutOfSumComments.Item(CInt(yearSheet.Name)) = True
                    End If
                ElseIf inTable AndAlso Not (IsNewHeaderVersion(cell) OrElse IsEmptyLine(cell) OrElse IsSum(cell)) AndAlso IsLineWithComment(cell) Then
                    Dim newLine As BookLine = ExtractedData.ReadLine(cell, 1)
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
    Private Shared Function IsNewHeaderVersion(cell As Range) As Boolean
        Return IsHeader(cell.Cells(1, 1)) AndAlso IsComment(cell.Cells(1, 13))
    End Function

    Private Shared Function IsHeader(firstCell As Range) As Boolean
        Return CStr(firstCell.Value2) <> "" AndAlso Not IsNumeric(firstCell.Value2)
    End Function

    Private Shared Function IsComment(cell As Range) As Boolean
        Return CStr(cell.Value2) = "Commentaires"
    End Function

    Private Shared Function IsSum(cell As Range) As Boolean
        Return CStr(cell.Cells(1, 5).Value2) = "Somme :"
    End Function

    Private Shared Function IsLineWithComment(cell As Range) As Boolean
        Return CStr(cell.Cells(1, 13).Value2) <> ""
    End Function

    Private Shared Function IsSumWithComment(cell As Range) As Boolean
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

    Public Sub DoIntegration()
        Dim dialog As Microsoft.Office.Core.FileDialog
        dialog = Application.FileDialog(Microsoft.Office.Core.MsoFileDialogType.msoFileDialogFilePicker)
        dialog.Filters.Add("fichier salaires CDD *.xlsx", "*.xlsx", 1)
        dialog.AllowMultiSelect = False
        If dialog.Show = -1 Then
            ProgressDialog.Show()
            SetProgress(0)
            GetCDD(dialog.SelectedItems.Item(1))
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
        Dim TotalNetAddress As String = SearchTotalNetAddress()
        SummaryCellsNotFound.Add(SummaryCellKind.TOTAL_NET, String.IsNullOrEmpty(TotalNetAddress))
        Dim TotalAddress As String = SearchTotalAddress()
        SummaryCellsNotFound.Add(SummaryCellKind.TOTAL, String.IsNullOrEmpty(TotalAddress))
        Dim CumulAddress As String = SearchCumulAddress()
        SummaryCellsNotFound.Add(SummaryCellKind.CUMUL, String.IsNullOrEmpty(CumulAddress))
        Dim BudgetAddress As String = SearchBudgetAddress()
        SummaryCellsNotFound.Add(SummaryCellKind.BUDGET, String.IsNullOrEmpty(BudgetAddress))
        Dim EngagedAddress As String = SearchEngagedAddress()
        SummaryCellsNotFound.Add(SummaryCellKind.ENGAGED, String.IsNullOrEmpty(EngagedAddress))
        Dim CurrentCol As Integer = 0

        BaseRange.Offset(0, CurrentCol).Value2 = ""
        BaseRange.Offset(0, CurrentCol).Style = "RecapHeaderStyle3"
        SummaryWorkSheet.Range(BaseRange.Offset(1, CurrentCol).Address, BaseRange.Offset(3, CurrentCol).Address).Merge()
        BaseRange.Offset(1, CurrentCol).Value2 = "Masse"
        BaseRange.Offset(1, CurrentCol).MergeArea.Style = "RecapHeaderStyle3"
        BaseRange.Offset(4, CurrentCol).Value2 = "1 Personnel"
        BaseRange.Offset(4, CurrentCol).Style = "RecapHeaderStyle4"
        BaseRange.Offset(5, CurrentCol).Value2 = "2 Fonctionnement hors amort."
        BaseRange.Offset(5, CurrentCol).Style = "RecapHeaderStyle4"
        BaseRange.Offset(6, CurrentCol).Value2 = "3 Investissement"
        BaseRange.Offset(6, CurrentCol).Style = "RecapHeaderStyle4"
        BaseRange.Offset(7, CurrentCol).Value2 = "Total"
        BaseRange.Offset(7, CurrentCol).Style = "RecapHeaderStyle5"


        Dim YearList As New List(Of Integer)
        YearList.AddRange(Data.Keys)
        YearList.Sort()
        If YearList.Count > 2 Then
            Dim ClosedYearsStart As Range = SummaryWorkSheet.Range("$C$18")
            ClosedYearsStart.Value2 = GetFormattedString("DepenseUntil", YearList.Max - 1)
            SummaryWorkSheet.Range(ClosedYearsStart.Address, ClosedYearsStart.Offset(0, YearList.Count - 2).Address).Merge()
            ClosedYearsStart.MergeArea.Style = "RecapHeaderStyle"
        Else
            Dim ClosedYearsStart As Range = SummaryWorkSheet.Range("$C$18")
            ClosedYearsStart.Value2 = ""
            ClosedYearsStart.MergeArea.Style = "RecapHeaderStyle"
        End If

        For Each Year As Integer In YearList
            If Year = YearList.Max Then
                'Last year
                CurrentCol = YearList.Count
                SummaryWorkSheet.Range(BaseRange.Offset(0, CurrentCol).Address, BaseRange.Offset(0, CurrentCol + 3).Address).Merge()
                BaseRange.Offset(0, CurrentCol).Value2 = Year
                BaseRange.Offset(0, CurrentCol).MergeArea.Style = "RecapHeaderStyle2"
                SummaryWorkSheet.Range(BaseRange.Offset(1, CurrentCol).Address, BaseRange.Offset(3, CurrentCol).Address).Merge()
                BaseRange.Offset(1, CurrentCol).Value2 = GetFormattedString("InitialBudget", Year)
                BaseRange.Offset(1, CurrentCol).MergeArea.Style = "RecapHeaderStyle2"
                BaseRange.Offset(4, CurrentCol).Value = 0
                BaseRange.Offset(4, CurrentCol).Style = "RecapNumberStyle3"
                BaseRange.Offset(5, CurrentCol).Value = 0
                BaseRange.Offset(5, CurrentCol).Style = "RecapNumberStyle3"
                BaseRange.Offset(6, CurrentCol).Value = 0
                BaseRange.Offset(6, CurrentCol).Style = "RecapNumberStyle3"
                BaseRange.Offset(7, CurrentCol).Formula = GetFormattedString("SumRange", BaseRange.Offset(4, CurrentCol).Address(False, False), BaseRange.Offset(6, CurrentCol).Address(False, False))
                BaseRange.Offset(7, CurrentCol).Style = "RecapNumberStyle3"

                SummaryWorkSheet.Range(BaseRange.Offset(1, CurrentCol + 1).Address, BaseRange.Offset(3, CurrentCol + 1).Address).Merge()
                BaseRange.Offset(1, CurrentCol + 1).Value2 = GetFormattedString("ModifiedBudget", Year)
                BaseRange.Offset(1, CurrentCol + 1).MergeArea.Style = "RecapHeaderStyle2"
                BaseRange.Offset(4, CurrentCol + 1).Value = 0
                BaseRange.Offset(4, CurrentCol + 1).Style = "RecapNumberStyle4"
                BaseRange.Offset(5, CurrentCol + 1).Value = 0
                BaseRange.Offset(5, CurrentCol + 1).Style = "RecapNumberStyle4"
                BaseRange.Offset(6, CurrentCol + 1).Value = 0
                BaseRange.Offset(6, CurrentCol + 1).Style = "RecapNumberStyle4"
                BaseRange.Offset(7, CurrentCol + 1).Formula = GetFormattedString("SumRange", BaseRange.Offset(4, CurrentCol + 1).Address(False, False), BaseRange.Offset(6, CurrentCol + 1).Address(False, False))
                BaseRange.Offset(7, CurrentCol + 1).Style = "RecapNumberStyle4"

                SummaryWorkSheet.Range(BaseRange.Offset(1, CurrentCol + 2).Address, BaseRange.Offset(3, CurrentCol + 2).Address).Merge()
                BaseRange.Offset(1, CurrentCol + 2).Value2 = GetFormattedString("EngagedBudget", Year)
                BaseRange.Offset(1, CurrentCol + 2).MergeArea.Style = "RecapHeaderStyle2"
                If ImportantCells.Item(Year).ContainsKey(KEY_SALARY) Then
                    BaseRange.Offset(4, CurrentCol + 2).Formula = GetFormattedString("CellRef", Year, ImportantCells.Item(Year).Item(KEY_SALARY))
                Else
                    BaseRange.Offset(4, CurrentCol + 2).Value = 0
                End If
                BaseRange.Offset(4, CurrentCol + 2).Style = "RecapNumberStyle3"
                If ImportantCells.Item(Year).ContainsKey(KEY_FONCT) AndAlso ImportantCells.Item(Year).ContainsKey(KEY_MISSION) Then
                    BaseRange.Offset(5, CurrentCol + 2).Formula = GetFormattedString("DirectCellSum", Year, ImportantCells.Item(Year).Item(KEY_FONCT), ImportantCells.Item(Year).Item(KEY_MISSION))
                ElseIf ImportantCells.Item(Year).ContainsKey(KEY_FONCT) Then
                    BaseRange.Offset(5, CurrentCol + 2).Formula = GetFormattedString("CellRef", Year, ImportantCells.Item(Year).Item(KEY_FONCT))
                ElseIf ImportantCells.Item(Year).ContainsKey(KEY_MISSION) Then
                    BaseRange.Offset(5, CurrentCol + 2).Formula = GetFormattedString("CellRef", Year, ImportantCells.Item(Year).Item(KEY_MISSION))
                Else
                    BaseRange.Offset(5, CurrentCol + 2).Value = 0
                End If
                BaseRange.Offset(5, CurrentCol + 2).Style = "RecapNumberStyle3"
                If ImportantCells.Item(Year).ContainsKey(KEY_INVEST) Then
                    BaseRange.Offset(6, CurrentCol + 2).Formula = GetFormattedString("CellRef", Year, ImportantCells.Item(Year).Item(KEY_INVEST))
                Else
                    BaseRange.Offset(6, CurrentCol + 2).Value = 0
                End If
                BaseRange.Offset(6, CurrentCol + 2).Style = "RecapNumberStyle3"
                BaseRange.Offset(7, CurrentCol + 2).Formula = GetFormattedString("SumRange", BaseRange.Offset(4, CurrentCol + 2).Address(False, False), BaseRange.Offset(6, CurrentCol + 2).Address(False, False))
                BaseRange.Offset(7, CurrentCol + 2).Style = "RecapNumberStyle3"

                SummaryWorkSheet.Range(BaseRange.Offset(1, CurrentCol + 3).Address, BaseRange.Offset(3, CurrentCol + 3).Address).Merge()
                BaseRange.Offset(1, CurrentCol + 3).Value2 = GetFormattedString("AvailableBudget", Year)
                BaseRange.Offset(1, CurrentCol + 3).MergeArea.Style = "RecapHeaderStyle2"
                BaseRange.Offset(4, CurrentCol + 3).Formula = GetFormattedString("DirectDiff", BaseRange.Offset(4, CurrentCol + 1).Address(False, False), BaseRange.Offset(4, CurrentCol + 2).Address(False, False))
                BaseRange.Offset(4, CurrentCol + 3).Style = "RecapNumberStyle4"
                BaseRange.Offset(5, CurrentCol + 3).Formula = GetFormattedString("DirectDiff", BaseRange.Offset(5, CurrentCol + 1).Address(False, False), BaseRange.Offset(5, CurrentCol + 2).Address(False, False))
                BaseRange.Offset(5, CurrentCol + 3).Style = "RecapNumberStyle4"
                BaseRange.Offset(6, CurrentCol + 3).Formula = GetFormattedString("DirectDiff", BaseRange.Offset(6, CurrentCol + 1).Address(False, False), BaseRange.Offset(6, CurrentCol + 2).Address(False, False))
                BaseRange.Offset(6, CurrentCol + 3).Style = "RecapNumberStyle4"
                BaseRange.Offset(7, CurrentCol + 3).Formula = GetFormattedString("DirectDiff", BaseRange.Offset(7, CurrentCol + 1).Address(False, False), BaseRange.Offset(7, CurrentCol + 2).Address(False, False))
                BaseRange.Offset(7, CurrentCol + 3).Style = "RecapNumberStyle4"

                CurrentCol += 4
            Else
                CurrentCol = YearList.IndexOf(Year) + 1
                SummaryWorkSheet.Range(BaseRange.Offset(1, CurrentCol).Address, BaseRange.Offset(3, CurrentCol).Address).Merge()
                BaseRange.Offset(1, CurrentCol).Value2 = CStr(Year)
                BaseRange.Offset(1, CurrentCol).MergeArea.Style = "RecapHeaderStyle"
                If ImportantCells.Item(Year).ContainsKey(KEY_SALARY) Then
                    BaseRange.Offset(4, CurrentCol).Formula = GetFormattedString("CellRef", Year, ImportantCells.Item(Year).Item(KEY_SALARY))
                Else
                    BaseRange.Offset(4, CurrentCol).Value = 0
                End If
                BaseRange.Offset(4, CurrentCol).Style = "RecapNumberStyle"
                If ImportantCells.Item(Year).ContainsKey(KEY_FONCT) AndAlso ImportantCells.Item(Year).ContainsKey(KEY_MISSION) Then
                    BaseRange.Offset(5, CurrentCol).Formula = GetFormattedString("DirectCellSum", Year, ImportantCells.Item(Year).Item(KEY_FONCT), ImportantCells.Item(Year).Item(KEY_MISSION))
                ElseIf ImportantCells.Item(Year).ContainsKey(KEY_FONCT) Then
                    BaseRange.Offset(5, CurrentCol).Formula = GetFormattedString("CellRef", Year, ImportantCells.Item(Year).Item(KEY_FONCT))
                ElseIf ImportantCells.Item(Year).ContainsKey(KEY_MISSION) Then
                    BaseRange.Offset(5, CurrentCol).Formula = GetFormattedString("CellRef", Year, ImportantCells.Item(Year).Item(KEY_MISSION))
                Else
                    BaseRange.Offset(5, CurrentCol).Value = 0
                End If
                BaseRange.Offset(5, CurrentCol).Style = "RecapNumberStyle"
                If ImportantCells.Item(Year).ContainsKey(KEY_INVEST) Then
                    BaseRange.Offset(6, CurrentCol).Formula = GetFormattedString("CellRef", Year, ImportantCells.Item(Year).Item(KEY_INVEST))
                Else
                    BaseRange.Offset(6, CurrentCol).Value = 0
                End If
                BaseRange.Offset(6, CurrentCol).Style = "RecapNumberStyle"
                BaseRange.Offset(7, CurrentCol).Formula = GetFormattedString("SumRange", BaseRange.Offset(4, CurrentCol).Address(False, False), BaseRange.Offset(6, CurrentCol).Address(False, False))
                BaseRange.Offset(7, CurrentCol).Style = "RecapNumberStyle2"
            End If
        Next
        Dim FirstCol As Integer = 1
        Dim LastClosedCol As Integer = CurrentCol - 5
        Dim CurrentYearAmountCol As Integer = CurrentCol - 3
        If FirstCol > LastClosedCol Then
            SummaryWorkSheet.Range(BaseRange.Offset(0, CurrentCol).Address, BaseRange.Offset(0, CurrentCol + 1).Address).Merge()
            BaseRange.Offset(0, CurrentCol).Value2 = GetFormattedString("HeaderOverYear", YearList.Max)
            BaseRange.Offset(0, CurrentCol).MergeArea.Style = "RecapHeaderStyle"
            SummaryWorkSheet.Range(BaseRange.Offset(1, CurrentCol).Address, BaseRange.Offset(3, CurrentCol).Address).Merge()
            BaseRange.Offset(1, CurrentCol).Value2 = "Reste Total à dépenser (y compris frais de Gestion UPS)"
            BaseRange.Offset(1, CurrentCol).MergeArea.Style = "RecapHeaderStyle6"
            BaseRange.Offset(4, CurrentCol).Value2 = ""
            BaseRange.Offset(4, CurrentCol).MergeArea.Style = "RecapHeaderStyle6"
            BaseRange.Offset(5, CurrentCol).Value2 = ""
            BaseRange.Offset(5, CurrentCol).MergeArea.Style = "RecapHeaderStyle6"
            BaseRange.Offset(6, CurrentCol).Value2 = ""
            BaseRange.Offset(6, CurrentCol).MergeArea.Style = "RecapHeaderStyle6"
            If Not SummaryCellsNotFound.Item(SummaryCellKind.TOTAL) Then
                BaseRange.Offset(7, CurrentCol).Formula = GetFormattedString("DirectDiff", TotalAddress, BaseRange.Offset(7, CurrentYearAmountCol).Address(False, False))
            Else
                BaseRange.Offset(7, CurrentCol).Value2 = ""
            End If
            BaseRange.Offset(7, CurrentCol).Style = "RecapNumberStyle5"

            BaseRange.Offset(0, CurrentCol + 1).Value2 = ""
            SummaryWorkSheet.Range(BaseRange.Offset(1, CurrentCol + 1).Address, BaseRange.Offset(3, CurrentCol + 1).Address).Merge()
            BaseRange.Offset(1, CurrentCol + 1).Value2 = "Reste Net à dépenser (net des Frais de Gestion UPS)"
            BaseRange.Offset(1, CurrentCol + 1).MergeArea.Style = "RecapHeaderStyle"
            BaseRange.Offset(1, CurrentCol + 1).MergeArea.Rows.RowHeight = BaseRange.Offset(0, CurrentCol + 1).RowHeight * 1.1
            BaseRange.Offset(4, CurrentCol + 1).Value2 = ""
            BaseRange.Offset(4, CurrentCol + 1).Style = "RecapNumberStyle2"
            BaseRange.Offset(5, CurrentCol + 1).Value2 = ""
            BaseRange.Offset(5, CurrentCol + 1).Style = "RecapNumberStyle2"
            BaseRange.Offset(6, CurrentCol + 1).Value2 = ""
            BaseRange.Offset(6, CurrentCol + 1).Style = "RecapNumberStyle2"
            If Not SummaryCellsNotFound.Item(SummaryCellKind.TOTAL_NET) Then
                BaseRange.Offset(7, CurrentCol + 1).Formula = GetFormattedString("DirectDiff", TotalNetAddress, BaseRange.Offset(7, CurrentYearAmountCol).Address(False, False))
            Else
                BaseRange.Offset(7, CurrentCol + 1).Value2 = ""
            End If
            BaseRange.Offset(7, CurrentCol + 1).Style = "RecapNumberStyle2"

            SummaryWorkSheet.Range(BaseRange.Offset(9, CurrentCol + 1).Address, BaseRange.Offset(9, CurrentCol + 2).Address).Merge()
            BaseRange.Offset(9, CurrentCol + 1).Value2 = "Montant total disponible :"
            BaseRange.Offset(9, CurrentCol + 1).MergeArea.Style = "RecapHeaderStyle5"

            BaseRange.Offset(9, CurrentCol + 3).Formula = GetFormattedString("DirectCellSum3", BaseRange.Offset(7, CurrentYearAmountCol + 2).Address(False, False), BaseRange.Offset(7, CurrentCol + 1).Address(False, False))
            BaseRange.Offset(9, CurrentCol + 3).Style = "RecapNumberStyle2"
            If Not SummaryCellsNotFound.Item(SummaryCellKind.CUMUL) Then
                SummaryWorkSheet.Range(CumulAddress).Formula = GetFormattedString("CellRef2", BaseRange.Offset(7, CurrentYearAmountCol + 1).Address(False, False))
            End If
            If Not SummaryCellsNotFound.Item(SummaryCellKind.BUDGET) Then
                SummaryWorkSheet.Range(BudgetAddress).Formula = GetFormattedString("CellRef2", BaseRange.Offset(7, CurrentYearAmountCol).Address(False, False))
            End If
            If Not SummaryCellsNotFound.Item(SummaryCellKind.ENGAGED) Then
                SummaryWorkSheet.Range(EngagedAddress).Formula = GetFormattedString("CellRef2", BaseRange.Offset(7, CurrentYearAmountCol + 1).Address(False, False))
            End If
        Else
            SummaryWorkSheet.Range(BaseRange.Offset(0, CurrentCol).Address, BaseRange.Offset(0, CurrentCol + 1).Address).Merge()
            BaseRange.Offset(0, CurrentCol).Value2 = GetFormattedString("HeaderOverYear", YearList.Max)
            BaseRange.Offset(0, CurrentCol).MergeArea.Style = "RecapHeaderStyle"
            SummaryWorkSheet.Range(BaseRange.Offset(1, CurrentCol).Address, BaseRange.Offset(3, CurrentCol).Address).Merge()
            BaseRange.Offset(1, CurrentCol).Value2 = "Reste Total à dépenser (y compris frais de Gestion UPS)"
            BaseRange.Offset(1, CurrentCol).MergeArea.Style = "RecapHeaderStyle6"
            BaseRange.Offset(4, CurrentCol).Value2 = ""
            BaseRange.Offset(4, CurrentCol).MergeArea.Style = "RecapHeaderStyle6"
            BaseRange.Offset(5, CurrentCol).Value2 = ""
            BaseRange.Offset(5, CurrentCol).MergeArea.Style = "RecapHeaderStyle6"
            BaseRange.Offset(6, CurrentCol).Value2 = ""
            BaseRange.Offset(6, CurrentCol).MergeArea.Style = "RecapHeaderStyle6"
            If Not SummaryCellsNotFound.Item(SummaryCellKind.TOTAL) Then
                BaseRange.Offset(7, CurrentCol).Formula = GetFormattedString("DirectDiff2", TotalAddress, BaseRange.Offset(7, CurrentYearAmountCol).Address(False, False), BaseRange.Offset(7, FirstCol).Address(False, False), BaseRange.Offset(7, LastClosedCol).Address(False, False))
            Else
                BaseRange.Offset(7, CurrentCol).Value2 = ""
            End If
            BaseRange.Offset(7, CurrentCol).Style = "RecapNumberStyle5"

            BaseRange.Offset(0, CurrentCol + 1).Value2 = ""
            SummaryWorkSheet.Range(BaseRange.Offset(1, CurrentCol + 1).Address, BaseRange.Offset(3, CurrentCol + 1).Address).Merge()
            BaseRange.Offset(1, CurrentCol + 1).Value2 = "Reste Net à dépenser (net des Frais de Gestion UPS)"
            BaseRange.Offset(1, CurrentCol + 1).MergeArea.Style = "RecapHeaderStyle"
            BaseRange.Offset(1, CurrentCol + 1).MergeArea.Rows.RowHeight = BaseRange.Offset(0, CurrentCol + 1).RowHeight * 1.1
            BaseRange.Offset(4, CurrentCol + 1).Value2 = ""
            BaseRange.Offset(4, CurrentCol + 1).Style = "RecapNumberStyle2"
            BaseRange.Offset(5, CurrentCol + 1).Value2 = ""
            BaseRange.Offset(5, CurrentCol + 1).Style = "RecapNumberStyle2"
            BaseRange.Offset(6, CurrentCol + 1).Value2 = ""
            BaseRange.Offset(6, CurrentCol + 1).Style = "RecapNumberStyle2"
            If Not SummaryCellsNotFound.Item(SummaryCellKind.TOTAL_NET) Then
                BaseRange.Offset(7, CurrentCol + 1).Formula = GetFormattedString("DirectDiff2", TotalNetAddress, BaseRange.Offset(7, CurrentYearAmountCol).Address(False, False), BaseRange.Offset(7, FirstCol).Address(False, False), BaseRange.Offset(7, LastClosedCol).Address(False, False))
            Else
                BaseRange.Offset(7, CurrentCol + 1).Value2 = ""
            End If
            BaseRange.Offset(7, CurrentCol + 1).Style = "RecapNumberStyle2"

            SummaryWorkSheet.Range(BaseRange.Offset(9, CurrentCol + 1).Address, BaseRange.Offset(9, CurrentCol + 2).Address).Merge()
            BaseRange.Offset(9, CurrentCol + 1).Value2 = "Montant total disponible :"
            BaseRange.Offset(9, CurrentCol + 1).MergeArea.Style = "RecapHeaderStyle5"

            BaseRange.Offset(9, CurrentCol + 3).Formula = GetFormattedString("DirectCellSum3", BaseRange.Offset(7, CurrentYearAmountCol + 2).Address(False, False), BaseRange.Offset(7, CurrentCol + 1).Address(False, False))
            BaseRange.Offset(9, CurrentCol + 3).Style = "RecapNumberStyle2"
            If Not SummaryCellsNotFound.Item(SummaryCellKind.CUMUL) Then
                SummaryWorkSheet.Range(CumulAddress).Formula = GetFormattedString("DirectCellSum2", BaseRange.Offset(7, CurrentYearAmountCol + 1).Address(False, False), BaseRange.Offset(7, FirstCol).Address(False, False), BaseRange.Offset(7, LastClosedCol).Address(False, False))
            End If
            If Not SummaryCellsNotFound.Item(SummaryCellKind.BUDGET) Then
                SummaryWorkSheet.Range(BudgetAddress).Formula = GetFormattedString("CellRef2", BaseRange.Offset(7, CurrentYearAmountCol).Address(False, False))
            End If
            If Not SummaryCellsNotFound.Item(SummaryCellKind.ENGAGED) Then
                SummaryWorkSheet.Range(EngagedAddress).Formula = GetFormattedString("CellRef2", BaseRange.Offset(7, CurrentYearAmountCol + 1).Address(False, False))
            End If
        End If
    End Sub

    Private Function SearchTotalNetAddress() As String
        Return SearchAddressWithPattern("montant total net", True)
    End Function

    Private Function SearchTotalAddress() As String
        Return SearchAddressWithPattern("montant total", True)
    End Function

    Private Function SearchCumulAddress() As String
        Return SearchAddressWithPattern("cumul", False)
    End Function
    Private Function SearchBudgetAddress() As String
        Return SearchAddressWithPattern("budget", False)
    End Function
    Private Function SearchEngagedAddress() As String
        Return SearchAddressWithPattern("engagé", False)
    End Function

    Private Function SearchAddressWithPattern(v As String, exact As Boolean) As String
        For Each cell As Excel.Range In SummaryWorkSheet.Range("A1:A18").Cells
            If exact AndAlso String.Equals(v, cell.Value2, StringComparison.OrdinalIgnoreCase) OrElse Not exact AndAlso cell.Value2 IsNot Nothing AndAlso CStr(cell.Value2).StartsWith(v, StringComparison.OrdinalIgnoreCase) Then
                Return cell.Offset(0, 1).Address
            End If
        Next
        Return Nothing
    End Function

    Private Shared Function GetFormattedString(Format As String, ParamArray Value() As Object) As String
        Return String.Format(CultureInfo.CurrentCulture, My.Resources.ResourceManager.GetString(Format), Value)
    End Function
    Private Sub GetCDD(fileName As String)
        NameStep("Récupération des CDD")
        ProgressIncrement = PROGRESS_STEP_CDD / 2
        Dim CDDWorkbook As Excel.Workbook = Me.Application.Workbooks.Open(fileName)
        'CDDWorkbook.IsAddin = True
        Dim CDDWorksheet As Worksheet = CType(CDDWorkbook.Sheets.Item(2), Excel.Worksheet)
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
                .ELibelle = GetFormattedString("LiasseDate", Format(CDDWorksheet.Range(String.Format(CultureInfo.CurrentCulture, "D{0}", NumRow)).Value, "dd/MM/yyyy"), Format(CDDWorksheet.Range(String.Format(CultureInfo.CurrentCulture, "E{0}", NumRow)).Value, "dd/MM/yyyy")),
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
        CDDWorkbook.Close()
    End Sub
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
        PrepareStyles()
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

    Private Sub PrepareStyles()
        If Not ContainsStyle("WarningDetailStyle") Then
            Dim NewStyle As Excel.Style = CurrentWorkbook.Styles.Add("WarningDetailStyle")
            NewStyle.Interior.Color = RGB(102, 102, 153)
            NewStyle.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)
            NewStyle.Font.Size = 11
            NewStyle.Font.Bold = True
        End If

        If Not ContainsStyle("WarningHeaderStyle") Then
            Dim NewStyle As Excel.Style = CurrentWorkbook.Styles.Add("WarningHeaderStyle")
            NewStyle.Interior.Color = RGB(255, 177, 63)
            NewStyle.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)
        End If

        If Not ContainsStyle("HeaderStyle") Then
            Dim NewStyle As Excel.Style = CurrentWorkbook.Styles.Add("HeaderStyle")
            NewStyle.Interior.Color = RGB(102, 102, 153)
            NewStyle.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)
        End If

        If Not ContainsStyle("HeaderStyleComment") Then
            Dim NewStyle As Excel.Style = CurrentWorkbook.Styles.Add("HeaderStyleComment")
            NewStyle.Interior.Color = RGB(255, 177, 63)
            NewStyle.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)
        End If

        If Not ContainsStyle("HeaderStyleFrom") Then
            Dim NewStyle As Excel.Style = CurrentWorkbook.Styles.Add("HeaderStyleFrom")
            NewStyle.Interior.Color = RGB(209, 54, 33)
            NewStyle.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)
        End If

        If Not ContainsStyle("MtEngStyle") Then
            Dim NewStyle As Excel.Style = CurrentWorkbook.Styles.Add("MtEngStyle")
            NewStyle.Interior.Color = RGB(102, 102, 153)
            NewStyle.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)
            NewStyle.NumberFormatLocal = "# ##0,00"
        End If

        If Not ContainsStyle("SIFACCommentaires") Then
            Dim NewStyle As Excel.Style = CurrentWorkbook.Styles.Add("SIFACCommentaires")
            NewStyle.Interior.Color = RGB(255, 232, 197)
            NewStyle.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black)
        End If

        If Not ContainsStyle("MtPAStyle") Then
            Dim NewStyle As Excel.Style = CurrentWorkbook.Styles.Add("MtPAStyle")
            NewStyle.NumberFormatLocal = "# ##0,00"
        End If

        If Not ContainsStyle("SumStyle") Then
            Dim NewStyle As Excel.Style = CurrentWorkbook.Styles.Add("SumStyle")
            NewStyle.Interior.Color = RGB(102, 102, 153)
            NewStyle.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)
            NewStyle.Font.Bold = True
            NewStyle.NumberFormatLocal = "# ##0,00"
        End If

        If Not ContainsStyle("RecapHeaderStyle") Then
            Dim NewStyle As Excel.Style = CurrentWorkbook.Styles.Add("RecapHeaderStyle")
            NewStyle.Interior.Color = RGB(102, 102, 153)
            NewStyle.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)
            NewStyle.Font.Size = 9
            NewStyle.Font.Bold = True
            NewStyle.HorizontalAlignment = XlHAlign.xlHAlignCenter
            NewStyle.VerticalAlignment = XlVAlign.xlVAlignBottom
            NewStyle.WrapText = True
        End If

        If Not ContainsStyle("RecapHeaderStyle2") Then
            Dim NewStyle As Excel.Style = CurrentWorkbook.Styles.Add("RecapHeaderStyle2")
            NewStyle.Interior.Color = RGB(166, 166, 166)
            NewStyle.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)
            NewStyle.Font.Size = 9
            NewStyle.Font.Bold = True
            NewStyle.HorizontalAlignment = XlHAlign.xlHAlignCenter
            NewStyle.VerticalAlignment = XlVAlign.xlVAlignBottom
            NewStyle.WrapText = True
        End If

        If Not ContainsStyle("RecapHeaderStyle3") Then
            Dim NewStyle As Excel.Style = CurrentWorkbook.Styles.Add("RecapHeaderStyle3")
            NewStyle.Interior.Color = RGB(102, 102, 153)
            NewStyle.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)
            NewStyle.Font.Size = 9
            NewStyle.Font.Bold = True
            NewStyle.HorizontalAlignment = XlHAlign.xlHAlignLeft
            NewStyle.VerticalAlignment = XlVAlign.xlVAlignBottom
            NewStyle.WrapText = True
        End If

        If Not ContainsStyle("RecapHeaderStyle4") Then
            Dim NewStyle As Excel.Style = CurrentWorkbook.Styles.Add("RecapHeaderStyle4")
            NewStyle.Interior.Color = RGB(102, 102, 153)
            NewStyle.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)
            NewStyle.Font.Size = 9
            NewStyle.Font.Bold = False
            NewStyle.HorizontalAlignment = XlHAlign.xlHAlignLeft
            NewStyle.VerticalAlignment = XlVAlign.xlVAlignBottom
            NewStyle.WrapText = True
        End If

        If Not ContainsStyle("RecapHeaderStyle5") Then
            Dim NewStyle As Excel.Style = CurrentWorkbook.Styles.Add("RecapHeaderStyle5")
            NewStyle.Interior.Color = RGB(102, 102, 153)
            NewStyle.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)
            NewStyle.Font.Size = 9
            NewStyle.Font.Bold = False
            NewStyle.HorizontalAlignment = XlHAlign.xlHAlignRight
            NewStyle.VerticalAlignment = XlVAlign.xlVAlignBottom
            NewStyle.WrapText = True
        End If

        If Not ContainsStyle("RecapHeaderStyle6") Then
            Dim NewStyle As Excel.Style = CurrentWorkbook.Styles.Add("RecapHeaderStyle6")
            NewStyle.Interior.Color = RGB(102, 102, 153)
            NewStyle.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black)
            NewStyle.Font.Size = 9
            NewStyle.Font.Bold = True
            NewStyle.HorizontalAlignment = XlHAlign.xlHAlignCenter
            NewStyle.VerticalAlignment = XlVAlign.xlVAlignBottom
            NewStyle.WrapText = True
        End If

        If Not ContainsStyle("RecapNumberStyle") Then
            Dim NewStyle As Excel.Style = CurrentWorkbook.Styles.Add("RecapNumberStyle")
            NewStyle.Interior.Color = RGB(102, 102, 153)
            NewStyle.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)
            NewStyle.Font.Size = 9
            NewStyle.Font.Bold = False
            NewStyle.HorizontalAlignment = XlHAlign.xlHAlignRight
            NewStyle.VerticalAlignment = XlVAlign.xlVAlignBottom
            NewStyle.NumberFormatLocal = "# ##0,00 €"
        End If

        If Not ContainsStyle("RecapNumberStyle2") Then
            Dim NewStyle As Excel.Style = CurrentWorkbook.Styles.Add("RecapNumberStyle2")
            NewStyle.Interior.Color = RGB(102, 102, 153)
            NewStyle.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)
            NewStyle.Font.Size = 9
            NewStyle.Font.Bold = True
            NewStyle.HorizontalAlignment = XlHAlign.xlHAlignRight
            NewStyle.VerticalAlignment = XlVAlign.xlVAlignBottom
            NewStyle.NumberFormatLocal = "# ##0,00 €"
        End If

        If Not ContainsStyle("RecapNumberStyle3") Then
            Dim NewStyle As Excel.Style = CurrentWorkbook.Styles.Add("RecapNumberStyle3")
            NewStyle.Interior.Color = RGB(166, 166, 166)
            NewStyle.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)
            NewStyle.Font.Size = 9
            NewStyle.Font.Bold = False
            NewStyle.HorizontalAlignment = XlHAlign.xlHAlignRight
            NewStyle.VerticalAlignment = XlVAlign.xlVAlignBottom
            NewStyle.NumberFormatLocal = "# ##0,00 €"
        End If

        If Not ContainsStyle("RecapNumberStyle4") Then
            Dim NewStyle As Excel.Style = CurrentWorkbook.Styles.Add("RecapNumberStyle4")
            NewStyle.Interior.Color = RGB(166, 166, 166)
            NewStyle.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White)
            NewStyle.Font.Size = 9
            NewStyle.Font.Bold = True
            NewStyle.HorizontalAlignment = XlHAlign.xlHAlignRight
            NewStyle.VerticalAlignment = XlVAlign.xlVAlignBottom
            NewStyle.NumberFormatLocal = "# ##0,00 €"
        End If

        If Not ContainsStyle("RecapNumberStyle5") Then
            Dim NewStyle As Excel.Style = CurrentWorkbook.Styles.Add("RecapNumberStyle5")
            NewStyle.Interior.Color = RGB(102, 102, 153)
            NewStyle.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black)
            NewStyle.Font.Size = 9
            NewStyle.Font.Bold = True
            NewStyle.HorizontalAlignment = XlHAlign.xlHAlignRight
            NewStyle.VerticalAlignment = XlVAlign.xlVAlignBottom
            NewStyle.NumberFormatLocal = "# ##0,00 €"
        End If
    End Sub

    Private Function ContainsStyle(aStyle As String) As Boolean
        For Each Style As Style In CurrentWorkbook.Styles
            If Style.Name = aStyle Then
                Return True
            End If
        Next
        Return False
    End Function

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
        If firstYear Then
            'No pending order
            Dim NewWorsheet As Excel.Worksheet = CType(CurrentWorkbook.Worksheets.Add(After:=SummaryWorkSheet), Excel.Worksheet)
            NewWorsheet.Name = year
            AllWorksheets.Add(year, NewWorsheet)
            FeedWorkSheet(NewWorsheet, year)
            AutoFit(NewWorsheet)
        ElseIf Not lastYear Then
            'Potential pending orders
            Dim NewWorsheet As Excel.Worksheet = CType(CurrentWorkbook.Worksheets.Add(Before:=AllWorksheets.Item(year - 1)), Excel.Worksheet)
            NewWorsheet.Name = year
            AllWorksheets.Add(year, NewWorsheet)
            FeedWorkSheetWithPendings(NewWorsheet, year)
            AutoFit(NewWorsheet)
        Else
            'Potential pending orders past and present years
            Dim NewWorsheet As Excel.Worksheet = CType(CurrentWorkbook.Worksheets.Add(Before:=AllWorksheets.Item(year - 1)), Excel.Worksheet)
            NewWorsheet.Name = year
            AllWorksheets.Add(year, NewWorsheet)
            FeedWorkSheetWithAllPendings(NewWorsheet, year)
            AutoFit(NewWorsheet)
        End If
    End Sub

    Private Shared Sub AutoFit(newWorsheet As Excel.Worksheet)
        newWorsheet.Range("A:N").EntireColumn.AutoFit()
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
        {KEY_SALARY, New List(Of BookLine)}
    }
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
                    StartRange.Cells(CurrentLine, 11).Value2 = ExtractedData.GetDateCompteAsText(Line)
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
                CType(StartRange.Cells(CurrentLine, SUM_COL), Excel.Range).Formula = GetFormattedString("SumRange", GetFormattedString("CellAddress", SUM_COL_LETTER, FirstLine), GetFormattedString("CellAddress", SUM_COL_LETTER, LastLine))
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
