Imports System.Diagnostics
Imports System.Globalization
Imports System.IO
Imports Microsoft.Office.Interop.Excel
Imports SIFACToEDC.ExtractedData

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
    Private Class FooLine
        Public ReadOnly data As New List(Of String)

        Public Shared Function ReadLine(FullRange As Range, RowNum As Integer) As FooLine
            Dim answer As New FooLine
            For I As Integer = 1 To FullRange.Columns.Count
                answer.data.Add(FullRange.Cells(RowNum, 1).Value2)
            Next
            Return answer
        End Function
        Private Shared Function GetDateCompte(TextValue As String) As DateCompte
            If Not TextValue = "" Then
                Return New DateCompte(TextValue)
            Else
                Return Nothing
            End If
        End Function

        Private Shared Function GetNumber(FullRange As Range, RowNum As Integer, ColNum As Integer) As Double
            Dim TextToConvert As String = FullRange.Cells(RowNum, ColNum).Value2
            Dim IndexVirgule As Integer = TextToConvert.IndexOf(",")
            Dim IndexPoint As Integer = TextToConvert.IndexOf(".")
            If IndexPoint = -1 Then
                'French format
                Return Double.Parse(TextToConvert)
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
    Private Sub ThisAddIn_Startup() Handles Me.Startup

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

    Public Sub ExtractAndMerge()
        Data.Clear()
        CDDMap.Clear()
        ImportantCells.Clear()
        AllWorksheets.Clear()
        ExistingSheets.Clear()
        CurrentWorkbook = CreateCopy(CType(Me.Application.ActiveWorkbook, Excel.Workbook))
        SummaryWorkSheet = CType(CurrentWorkbook.Sheets(2), Excel.Worksheet)
        RecapNumber = SummaryWorkSheet.Range("A2").Value2
        GetExistingData()
        'DoIntegration()
    End Sub

    Private Sub GetExistingData()
        ExistingSheets.Add(KEY_SHEET_MODE_EMPLOI, New Dictionary(Of String, Excel.Worksheet) From {
                    {CType(CurrentWorkbook.Sheets(1), Excel.Worksheet).Name, CType(CurrentWorkbook.Sheets(1), Excel.Worksheet)}})
        ExistingSheets.Add(KEY_SHEET_RECAPITULATIF, New Dictionary(Of String, Excel.Worksheet) From {
                    {CType(CurrentWorkbook.Sheets(2), Excel.Worksheet).Name, CType(CurrentWorkbook.Sheets(2), Excel.Worksheet)}})
        For I As Integer = 3 To CurrentWorkbook.Sheets.Count
            Dim worksheet As Worksheet = CType(CurrentWorkbook.Sheets(I), Excel.Worksheet)
            If IsNumeric(worksheet.Name) Then
                If ExistingSheets.ContainsKey(KEY_SHEET_YEARS) Then
                    ExistingSheets.Item(KEY_SHEET_YEARS).Add(worksheet.Name, worksheet)
                Else
                    ExistingSheets.Add(KEY_SHEET_YEARS, New Dictionary(Of String, Excel.Worksheet) From {
                    {worksheet.Name, worksheet}})
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

        Dim testWorksheet As Excel.Worksheet = ExistingSheets.Item(KEY_SHEET_YEARS).Values().ElementAt(1)
        For Each FooRange As Excel.Range In testWorksheet.UsedRange.Rows
            Debug.WriteLine($"--{FooRange.Cells(1, 1).value2}")
            If FooRange.Cells(1, 2).value2 = "COMMANDE" Then
                Dim Line As FooLine = FooLine.ReadLine(FooRange, 1)
            End If
        Next
    End Sub

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
        SummaryWorkSheet.Range("B18:BB27").Rows.Delete(XlDeleteShiftDirection.xlShiftToLeft)

        Dim BaseRange As Excel.Range = SummaryWorkSheet.Range("B18")
        Dim TableCol As Excel.Range = SummaryWorkSheet.Range("B18:B25")
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
            ClosedYearsStart.Value2 = $"Dépenses jusqu'en {YearList.Max - 1}"
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
                BaseRange.Offset(1, CurrentCol).Value2 = $"Budget initial {Year}"
                BaseRange.Offset(1, CurrentCol).MergeArea.Style = "RecapHeaderStyle2"
                BaseRange.Offset(4, CurrentCol).Value = 0
                BaseRange.Offset(4, CurrentCol).Style = "RecapNumberStyle3"
                BaseRange.Offset(5, CurrentCol).Value = 0
                BaseRange.Offset(5, CurrentCol).Style = "RecapNumberStyle3"
                BaseRange.Offset(6, CurrentCol).Value = 0
                BaseRange.Offset(6, CurrentCol).Style = "RecapNumberStyle3"
                BaseRange.Offset(7, CurrentCol).Formula = $"=SUM({BaseRange.Offset(4, CurrentCol).Address(False, False)}:{BaseRange.Offset(6, CurrentCol).Address(False, False)})"
                BaseRange.Offset(7, CurrentCol).Style = "RecapNumberStyle3"

                SummaryWorkSheet.Range(BaseRange.Offset(1, CurrentCol + 1).Address, BaseRange.Offset(3, CurrentCol + 1).Address).Merge()
                BaseRange.Offset(1, CurrentCol + 1).Value2 = $"Budget modifié {Year}"
                BaseRange.Offset(1, CurrentCol + 1).MergeArea.Style = "RecapHeaderStyle2"
                BaseRange.Offset(4, CurrentCol + 1).Value = 0
                BaseRange.Offset(4, CurrentCol + 1).Style = "RecapNumberStyle4"
                BaseRange.Offset(5, CurrentCol + 1).Value = 0
                BaseRange.Offset(5, CurrentCol + 1).Style = "RecapNumberStyle4"
                BaseRange.Offset(6, CurrentCol + 1).Value = 0
                BaseRange.Offset(6, CurrentCol + 1).Style = "RecapNumberStyle4"
                BaseRange.Offset(7, CurrentCol + 1).Formula = $"=SUM({BaseRange.Offset(4, CurrentCol + 1).Address(False, False)}:{BaseRange.Offset(6, CurrentCol + 1).Address(False, False)})"
                BaseRange.Offset(7, CurrentCol + 1).Style = "RecapNumberStyle4"

                SummaryWorkSheet.Range(BaseRange.Offset(1, CurrentCol + 2).Address, BaseRange.Offset(3, CurrentCol + 2).Address).Merge()
                BaseRange.Offset(1, CurrentCol + 2).Value2 = $"Montant engagé + liquidé {Year}"
                BaseRange.Offset(1, CurrentCol + 2).MergeArea.Style = "RecapHeaderStyle2"
                If ImportantCells.Item(Year).ContainsKey(KEY_SALARY) Then
                    BaseRange.Offset(4, CurrentCol + 2).Formula = $"='{Year}'!{ImportantCells.Item(Year).Item(KEY_SALARY)}"
                Else
                    BaseRange.Offset(4, CurrentCol + 2).Value = 0
                End If
                BaseRange.Offset(4, CurrentCol + 2).Style = "RecapNumberStyle3"
                If ImportantCells.Item(Year).ContainsKey(KEY_FONCT) And ImportantCells.Item(Year).ContainsKey(KEY_MISSION) Then
                    BaseRange.Offset(5, CurrentCol + 2).Formula = $"='{Year}'!{ImportantCells.Item(Year).Item(KEY_FONCT)}+'{Year}'!{ImportantCells.Item(Year).Item(KEY_MISSION)}"
                ElseIf ImportantCells.Item(Year).ContainsKey(KEY_FONCT) Then
                    BaseRange.Offset(5, CurrentCol + 2).Formula = $"='{Year}'!{ImportantCells.Item(Year).Item(KEY_FONCT)}"
                ElseIf ImportantCells.Item(Year).ContainsKey(KEY_MISSION) Then
                    BaseRange.Offset(5, CurrentCol + 2).Formula = $"='{Year}'!{ImportantCells.Item(Year).Item(KEY_MISSION)}"
                Else
                    BaseRange.Offset(5, CurrentCol + 2).Value = 0
                End If
                BaseRange.Offset(5, CurrentCol + 2).Style = "RecapNumberStyle3"
                If ImportantCells.Item(Year).ContainsKey(KEY_INVEST) Then
                    BaseRange.Offset(6, CurrentCol + 2).Formula = $"='{Year}'!{ImportantCells.Item(Year).Item(KEY_INVEST)}"
                Else
                    BaseRange.Offset(6, CurrentCol + 2).Value = 0
                End If
                BaseRange.Offset(6, CurrentCol + 2).Style = "RecapNumberStyle3"
                BaseRange.Offset(7, CurrentCol + 2).Formula = $"=SUM({BaseRange.Offset(4, CurrentCol + 2).Address(False, False)}:{BaseRange.Offset(6, CurrentCol + 2).Address(False, False)})"
                BaseRange.Offset(7, CurrentCol + 2).Style = "RecapNumberStyle3"

                SummaryWorkSheet.Range(BaseRange.Offset(1, CurrentCol + 3).Address, BaseRange.Offset(3, CurrentCol + 3).Address).Merge()
                BaseRange.Offset(1, CurrentCol + 3).Value2 = $"Montant disponible {Year}"
                BaseRange.Offset(1, CurrentCol + 3).MergeArea.Style = "RecapHeaderStyle2"
                BaseRange.Offset(4, CurrentCol + 3).Formula = $"={BaseRange.Offset(4, CurrentCol + 1).Address(False, False)} - {BaseRange.Offset(4, CurrentCol + 2).Address(False, False)}"
                BaseRange.Offset(4, CurrentCol + 3).Style = "RecapNumberStyle4"
                BaseRange.Offset(5, CurrentCol + 3).Formula = $"={BaseRange.Offset(5, CurrentCol + 1).Address(False, False)} - {BaseRange.Offset(5, CurrentCol + 2).Address(False, False)}"
                BaseRange.Offset(5, CurrentCol + 3).Style = "RecapNumberStyle4"
                BaseRange.Offset(6, CurrentCol + 3).Formula = $"={BaseRange.Offset(6, CurrentCol + 1).Address(False, False)} - {BaseRange.Offset(6, CurrentCol + 2).Address(False, False)}"
                BaseRange.Offset(6, CurrentCol + 3).Style = "RecapNumberStyle4"
                BaseRange.Offset(7, CurrentCol + 3).Formula = $"={BaseRange.Offset(7, CurrentCol + 1).Address(False, False)} - {BaseRange.Offset(7, CurrentCol + 2).Address(False, False)}"
                BaseRange.Offset(7, CurrentCol + 3).Style = "RecapNumberStyle4"

                CurrentCol += 4
            Else
                CurrentCol = YearList.IndexOf(Year) + 1
                SummaryWorkSheet.Range(BaseRange.Offset(1, CurrentCol).Address, BaseRange.Offset(3, CurrentCol).Address).Merge()
                BaseRange.Offset(1, CurrentCol).Value2 = CStr(Year)
                BaseRange.Offset(1, CurrentCol).MergeArea.Style = "RecapHeaderStyle"
                If ImportantCells.Item(Year).ContainsKey(KEY_SALARY) Then
                    BaseRange.Offset(4, CurrentCol).Formula = $"='{Year}'!{ImportantCells.Item(Year).Item(KEY_SALARY)}"
                Else
                    BaseRange.Offset(4, CurrentCol).Value = 0
                End If
                BaseRange.Offset(4, CurrentCol).Style = "RecapNumberStyle"
                If ImportantCells.Item(Year).ContainsKey(KEY_FONCT) And ImportantCells.Item(Year).ContainsKey(KEY_MISSION) Then
                    BaseRange.Offset(5, CurrentCol).Formula = $"='{Year}'!{ImportantCells.Item(Year).Item(KEY_FONCT)}+'{Year}'!{ImportantCells.Item(Year).Item(KEY_MISSION)}"
                ElseIf ImportantCells.Item(Year).ContainsKey(KEY_FONCT) Then
                    BaseRange.Offset(5, CurrentCol).Formula = $"='{Year}'!{ImportantCells.Item(Year).Item(KEY_FONCT)}"
                ElseIf ImportantCells.Item(Year).ContainsKey(KEY_MISSION) Then
                    BaseRange.Offset(5, CurrentCol).Formula = $"='{Year}'!{ImportantCells.Item(Year).Item(KEY_MISSION)}"
                Else
                    BaseRange.Offset(5, CurrentCol).Value = 0
                End If
                BaseRange.Offset(5, CurrentCol).Style = "RecapNumberStyle"
                If ImportantCells.Item(Year).ContainsKey(KEY_INVEST) Then
                    BaseRange.Offset(6, CurrentCol).Formula = $"='{Year}'!{ImportantCells.Item(Year).Item(KEY_INVEST)}"
                Else
                    BaseRange.Offset(6, CurrentCol).Value = 0
                End If
                BaseRange.Offset(6, CurrentCol).Style = "RecapNumberStyle"
                BaseRange.Offset(7, CurrentCol).Formula = $"=SUM({BaseRange.Offset(4, CurrentCol).Address(False, False)}:{BaseRange.Offset(6, CurrentCol).Address(False, False)})"
                BaseRange.Offset(7, CurrentCol).Style = "RecapNumberStyle2"
            End If
        Next
        Dim FirstCol As Integer = 1
        Dim LastClosedCol As Integer = CurrentCol - 5
        Dim CurrentYearAmountCol As Integer = CurrentCol - 3
        Dim TotalNetAddress As String = "$B$9"
        Dim TotalNetUPSAddress As String = "$B$7"

        SummaryWorkSheet.Range(BaseRange.Offset(0, CurrentCol).Address, BaseRange.Offset(0, CurrentCol + 1).Address).Merge()
        BaseRange.Offset(0, CurrentCol).Value2 = $"> {YearList.Max}"
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
        BaseRange.Offset(7, CurrentCol).Formula = $"={TotalNetUPSAddress}-{BaseRange.Offset(7, CurrentYearAmountCol).Address(False, False)}-SUM({BaseRange.Offset(7, FirstCol).Address(False, False)}:{BaseRange.Offset(7, LastClosedCol).Address(False, False)})"
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
        BaseRange.Offset(7, CurrentCol + 1).Formula = $"={TotalNetAddress}-{BaseRange.Offset(7, CurrentYearAmountCol).Address(False, False)}-SUM({BaseRange.Offset(7, FirstCol).Address(False, False)}:{BaseRange.Offset(7, LastClosedCol).Address(False, False)})"
        BaseRange.Offset(7, CurrentCol + 1).Style = "RecapNumberStyle2"


        SummaryWorkSheet.Range(BaseRange.Offset(9, CurrentCol + 1).Address, BaseRange.Offset(9, CurrentCol + 2).Address).Merge()
        BaseRange.Offset(9, CurrentCol + 1).Value2 = "Montant total disponible :"
        BaseRange.Offset(9, CurrentCol + 1).MergeArea.Style = "RecapHeaderStyle5"

        BaseRange.Offset(9, CurrentCol + 3).Formula = $"={BaseRange.Offset(7, CurrentYearAmountCol + 2).Address(False, False)}+{BaseRange.Offset(7, CurrentCol + 1).Address(False, False)}"
        BaseRange.Offset(9, CurrentCol + 3).Style = "RecapNumberStyle2"

        SummaryWorkSheet.Range("$B10").Formula = $"={BaseRange.Offset(7, CurrentYearAmountCol + 1).Address(False, False)}+SUM({BaseRange.Offset(7, FirstCol).Address(False, False)}:{BaseRange.Offset(7, LastClosedCol).Address(False, False)})"
        SummaryWorkSheet.Range("$B12").Formula = $"={BaseRange.Offset(7, CurrentYearAmountCol).Address(False, False)}"
        SummaryWorkSheet.Range("$B13").Formula = $"={BaseRange.Offset(7, CurrentYearAmountCol + 1).Address(False, False)}"
    End Sub

    Private Sub GetCDD(FileName As String)
        NameStep("Récupération des CDD")
        ProgressIncrement = PROGRESS_STEP_CDD / 2
        Dim CDDWorkbook As Excel.Workbook = Me.Application.Workbooks.Open(FileName)
        'CDDWorkbook.IsAddin = True
        Dim CDDWorksheet As Worksheet = CType(CDDWorkbook.Sheets.Item(2), Excel.Worksheet)
        Dim DataRange As Excel.Range = CDDWorksheet.UsedRange
        Dim FirstRow As Integer = 7
        Dim LastRow As Integer = DataRange.Rows.Count
        NextStep()
        ProgressIncrement = PROGRESS_STEP_CDD / 2 / LastRow

        For NumRow As Integer = FirstRow To LastRow
            Dim TempRecapNum As String = CDDWorksheet.Range($"Q{NumRow}").Value2
            If String.Equals(RecapNumber, TempRecapNum, StringComparison.OrdinalIgnoreCase) Then
                Dim NewLine As New BookLine With {
                    .A_Cptegen = "",
                    .B_Rubrique = "SALAIRE",
                    .C_NumeroFlux = CDDWorksheet.Range($"B{NumRow}").Value2,
                    .D_Nom = CDDWorksheet.Range($"I{NumRow}").Value2,
                    .E_Libelle = $"liasse du {Format(CDDWorksheet.Range($"D{NumRow}").Value, "dd/MM/yyyy")} au {Format(CDDWorksheet.Range($"E{NumRow}").Value, "dd/MM/yyyy")}",
                    .F_MntEngHTR = CDbl(CDDWorksheet.Range($"N{NumRow}").Value2),
                    .G_MontantPa = 0,
                    .H_Rapprochmt = "",
                    .I_RefFactF = "",
                    .J_DatePce = "",
                    .K_DCompt = Nothing,
                    .L_NumPiece = "",
                    .M_Comment = "",
                    .N_From = ""
                }
                Dim Year As Integer = CInt(CDDWorksheet.Range($"C{NumRow}").Value2)
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
        SourcesDirectory = $"{BaseDirectory}{SOURCE_DIRECTORY_MODIFIER}"
        ExtractionDirectory = $"{BaseDirectory}{EXTRACT_DIRECTORY_MODIFIER}"
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

    Private Function ContainsStyle(AStyle As String) As Boolean
        For Each Style As Style In CurrentWorkbook.Styles
            If Style.Name = AStyle Then
                Return True
            End If
        Next
        Return False
    End Function

    Private Function CreateCopy(AWorkbook As Excel.Workbook) As Excel.Workbook
        Dim NewFileName As String = $"{Path.GetDirectoryName(AWorkbook.FullName)}{Path.DirectorySeparatorChar}{Path.GetFileNameWithoutExtension(AWorkbook.FullName)} - New version {DEST_NAME_EXTENSION}"
        AWorkbook.SaveCopyAs(NewFileName)
        Dim NewWorkbook As Excel.Workbook = Me.Application.Workbooks.Open(NewFileName)
        AWorkbook.Close()
        Return NewWorkbook
    End Function

    Private Sub CreateSheetForYear(Year As Integer, FirstYear As Boolean, LastYear As Boolean)
        If FirstYear Then
            'No pending order
            Dim NewWorsheet As Excel.Worksheet = CType(CurrentWorkbook.Worksheets.Add(After:=SummaryWorkSheet), Excel.Worksheet)
            NewWorsheet.Name = Year
            AllWorksheets.Add(Year, NewWorsheet)
            FeedWorkSheet(NewWorsheet, Year)
            AutoFit(NewWorsheet)
        ElseIf Not LastYear Then
            'Potential pending orders
            Dim NewWorsheet As Excel.Worksheet = CType(CurrentWorkbook.Worksheets.Add(Before:=AllWorksheets.Item(Year - 1)), Excel.Worksheet)
            NewWorsheet.Name = Year
            AllWorksheets.Add(Year, NewWorsheet)
            FeedWorkSheetWithPendings(NewWorsheet, Year)
            AutoFit(NewWorsheet)
        Else
            'Potential pending orders past and present years
            Dim NewWorsheet As Excel.Worksheet = CType(CurrentWorkbook.Worksheets.Add(Before:=AllWorksheets.Item(Year - 1)), Excel.Worksheet)
            NewWorsheet.Name = Year
            AllWorksheets.Add(Year, NewWorsheet)
            FeedWorkSheetWithAllPendings(NewWorsheet, Year)
            AutoFit(NewWorsheet)
        End If
    End Sub

    Private Sub AutoFit(newWorsheet As Excel.Worksheet)
        newWorsheet.Range("A:N").EntireColumn.AutoFit()
    End Sub

    Private Sub FeedWorkSheet(NewWorsheet As Excel.Worksheet, Year As Integer)
        Dim MergedData As New Dictionary(Of String, List(Of ExtractedData.BookLine)) From {
            {KEY_FONCT, New List(Of ExtractedData.BookLine)},
            {KEY_INVEST, New List(Of ExtractedData.BookLine)},
            {KEY_MISSION, New List(Of ExtractedData.BookLine)},
            {KEY_SALARY, New List(Of ExtractedData.BookLine)}
        }
        MergedData.Item(KEY_FONCT).AddRange(Data.Item(Year).Orders)
        MergedData.Item(KEY_INVEST).AddRange(Data.Item(Year).Invests)
        MergedData.Item(KEY_MISSION).AddRange(Data.Item(Year).Missions)
        If CDDMap.ContainsKey(Year) Then
            MergedData.Item(KEY_SALARY).AddRange(CDDMap.Item(Year))
        End If
        Dump(MergedData, NewWorsheet, Year, False)
    End Sub

    Private Sub FeedWorkSheetWithPendings(NewWorsheet As Excel.Worksheet, Year As Integer)
        Dim MergedData As New Dictionary(Of String, List(Of ExtractedData.BookLine)) From {
            {KEY_FONCT, New List(Of ExtractedData.BookLine)},
            {KEY_INVEST, New List(Of ExtractedData.BookLine)},
            {KEY_MISSION, New List(Of ExtractedData.BookLine)},
            {KEY_SALARY, New List(Of ExtractedData.BookLine)}
        }
        MergedData.Item(KEY_FONCT).AddRange(Data.Item(Year - 1).PendingOrders)
        MergedData.Item(KEY_FONCT).AddRange(Data.Item(Year).Orders)
        MergedData.Item(KEY_INVEST).AddRange(Data.Item(Year - 1).PendingInvests)
        MergedData.Item(KEY_INVEST).AddRange(Data.Item(Year).Invests)
        MergedData.Item(KEY_MISSION).AddRange(Data.Item(Year - 1).PendingMissions)
        MergedData.Item(KEY_MISSION).AddRange(Data.Item(Year).Missions)
        If CDDMap.ContainsKey(Year) Then
            MergedData.Item(KEY_SALARY).AddRange(CDDMap.Item(Year))
        End If
        Dump(MergedData, NewWorsheet, Year, False)
    End Sub
    Private Sub FeedWorkSheetWithAllPendings(NewWorsheet As Excel.Worksheet, Year As Integer)
        Dim MergedData As New Dictionary(Of String, List(Of ExtractedData.BookLine)) From {
            {KEY_FONCT, New List(Of ExtractedData.BookLine)},
            {KEY_INVEST, New List(Of ExtractedData.BookLine)},
            {KEY_MISSION, New List(Of ExtractedData.BookLine)},
            {KEY_SALARY, New List(Of ExtractedData.BookLine)}
        }
        MergedData.Item(KEY_FONCT).AddRange(Data.Item(Year - 1).PendingOrders)
        MergedData.Item(KEY_FONCT).AddRange(Data.Item(Year).PendingOrders)
        MergedData.Item(KEY_FONCT).AddRange(Data.Item(Year).Orders)
        MergedData.Item(KEY_INVEST).AddRange(Data.Item(Year - 1).PendingInvests)
        MergedData.Item(KEY_INVEST).AddRange(Data.Item(Year).PendingInvests)
        MergedData.Item(KEY_INVEST).AddRange(Data.Item(Year).Invests)
        MergedData.Item(KEY_MISSION).AddRange(Data.Item(Year - 1).PendingMissions)
        MergedData.Item(KEY_MISSION).AddRange(Data.Item(Year).PendingMissions)
        MergedData.Item(KEY_MISSION).AddRange(Data.Item(Year).Missions)
        If CDDMap.ContainsKey(Year) Then
            MergedData.Item(KEY_SALARY).AddRange(CDDMap.Item(Year))
        End If
        Dump(MergedData, NewWorsheet, Year, True)
    End Sub
    Private Sub Dump(mergedData As Dictionary(Of String, List(Of ExtractedData.BookLine)), newWorsheet As Excel.Worksheet, Year As Integer, FullHeader As Boolean)
        ImportantCells.Add(Year, New Dictionary(Of String, String))
        Dim CurrentLine As Integer = 1
        Dim StartRange As Excel.Range = newWorsheet.Range("A1")
        For Each Key In mergedData.Keys
            Dim LineList As List(Of BookLine) = mergedData.Item(Key)
            If LineList.Count > 0 Then
                CurrentLine += 1
                DumpHeaders(StartRange, CurrentLine, FullHeader)
                CurrentLine += 1
                Dim FirstLine As Integer = CurrentLine
                For Each Line As BookLine In LineList
                    StartRange.Cells(CurrentLine, 1).Value2 = Line.A_Cptegen
                    StartRange.Cells(CurrentLine, 2).Value2 = Line.B_Rubrique
                    StartRange.Cells(CurrentLine, 3).Value2 = Line.C_NumeroFlux
                    StartRange.Cells(CurrentLine, 4).Value2 = Line.D_Nom
                    StartRange.Cells(CurrentLine, 5).Value2 = Line.E_Libelle
                    If Line.F_MntEngHTR = 0 Then
                        StartRange.Cells(CurrentLine, 6).Value2 = Line.G_MontantPa
                    Else
                        StartRange.Cells(CurrentLine, 6).Value2 = Line.F_MntEngHTR
                    End If
                    CType(StartRange.Cells(CurrentLine, 6), Excel.Range).Style = "MtEngStyle"
                    StartRange.Cells(CurrentLine, 7).Value2 = Line.G_MontantPa
                    CType(StartRange.Cells(CurrentLine, 7), Excel.Range).Style = "MtPAStyle"
                    StartRange.Cells(CurrentLine, 8).Value2 = Line.H_Rapprochmt
                    StartRange.Cells(CurrentLine, 9).Value2 = Line.I_RefFactF
                    StartRange.Cells(CurrentLine, 10).Value2 = Line.J_DatePce
                    StartRange.Cells(CurrentLine, 11).Value2 = ExtractedData.GetDateCompteAsText(Line)
                    StartRange.Cells(CurrentLine, 12).Value2 = Line.L_NumPiece
                    StartRange.Cells(CurrentLine, 13).Value2 = Line.M_Comment
                    StartRange.Cells(CurrentLine, 14).Value2 = Line.N_From
                    CurrentLine += 1
                    Globals.ThisAddIn.NextStep()
                Next
                Dim LastLine As Integer = CurrentLine - 1
                StartRange.Cells(CurrentLine, SUM_COL - 1).Value2 = LABEL_SUM
                CType(StartRange.Cells(CurrentLine, SUM_COL), Excel.Range).Formula = $"=SUM({SUM_COL_LETTER}{FirstLine}:{SUM_COL_LETTER}{LastLine})"
                CType(StartRange.Cells(CurrentLine, SUM_COL), Excel.Range).Style = "SumStyle"
                ImportantCells.Item(Year).Add(Key, CType(StartRange.Cells(CurrentLine, SUM_COL), Excel.Range).Address(False, False))
                CurrentLine += 1
            End If
        Next
    End Sub

    Private Sub DumpHeaders(startRange As Excel.Range, currentLine As Integer, FullHeader As Boolean)
        For NumCol As Integer = 1 To HEADERS.Count - 2
            Dim Cell As Excel.Range = CType(startRange.Cells(currentLine, NumCol), Excel.Range)
            Cell.Value2 = HEADERS.Item(NumCol - 1)
            Cell.Style = "HeaderStyle"
        Next
        Dim CellComment As Excel.Range = CType(startRange.Cells(currentLine, HEADERS.Count - 1), Excel.Range)
        CellComment.Value2 = HEADERS.Item(HEADERS.Count - 2)
        CellComment.Style = "HeaderStyleComment"
        If FullHeader Then
            Dim CellFrom As Excel.Range = CType(startRange.Cells(currentLine, HEADERS.Count), Excel.Range)
            CellFrom.Value2 = HEADERS.Item(HEADERS.Count - 1)
            CellFrom.Style = "HeaderStyleFrom"
        End If
    End Sub

    Public Sub NextStep()
        CurrentProgrees += ProgressIncrement
        ProgressDialog.ProgressTraitement.Value = Math.Min(CInt(CurrentProgrees), PROGRESS_REFERENCE)
    End Sub
    Public Sub SetProgress(ProgressValue As Double)
        CurrentProgrees = ProgressValue
        ProgressDialog.ProgressTraitement.Value = CInt(CurrentProgrees)
    End Sub
    Public Sub NameStep(StepName As String)
        ProgressDialog.LblPhase.Text = StepName
        ProgressDialog.LblPhase.Refresh()
    End Sub
    Private Function ExtractDataFromFile(Name As String) As ExtractedData
        Dim NewWorkbookPath As String = $"{ExtractionDirectory}{Path.GetFileNameWithoutExtension(Name)}{DEST_NAME_EXTENSION}"
        Return New ExtractedData(Name, NewWorkbookPath, Me.Application)
    End Function

    Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
        Return Globals.Factory.GetRibbonFactory().CreateRibbonManager(New Ribbon.IRibbonExtension() {New RibbonEDC()})
    End Function
End Class
