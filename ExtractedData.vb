Imports System.Collections.ObjectModel
Imports System.Diagnostics
Imports System.Globalization
Imports Microsoft.Office.Interop.Excel
Public Class ExtractedData
    Private Delegate Sub TableReducer(PreparedLines As Dictionary(Of String, Collection(Of BookLine)), MAX_K_DCompt As DateCompte, Line As BookLine)
    Private Delegate Sub LinePreparer(keyMain As String, preparedLines As Dictionary(Of String, Collection(Of BookLine)))
    Private ReadOnly FullTable As New Dictionary(Of String, Collection(Of BookLine))
    Private ReadOnly ReducedTable As New Dictionary(Of String, Collection(Of BookLine))
    Private ReadOnly PreReducedTable As New Dictionary(Of String, Dictionary(Of String, Collection(Of BookLine)))
    Private ReadOnly TrashTable As New Dictionary(Of String, Collection(Of BookLine))
    Private ReadOnly TableReducers As New Dictionary(Of String, TableReducer) From {
                        {CASE_COMMANDE_PAIEMENT, AddressOf HandleCommandePaiement},
                        {CASE_COMMANDE_AJUSTEMENT, AddressOf HandleCommandeAjustement},
                        {CASE_COMMANDE, AddressOf HandleCommande},
                        {CASE_MISSION_PAIEMENT, AddressOf HandleMissionPaiement},
                        {CASE_MISSION, AddressOf HandleMission},
                        {CASE_REIMPUTATION, AddressOf HandleReimputation},
                        {CASE_AVOIR_PAIEMENT, AddressOf HandleAvoirPaiement},
                        {CASE_DEFAULT, AddressOf ExtractedData.HandleOthers}}
    Private ReadOnly LinePreparers As New Dictionary(Of String, LinePreparer) From {
                        {CASE_COMMANDE_PAIEMENT, AddressOf PrepareLineTemplate1},
                        {CASE_COMMANDE_PAIEMENT_INVEST, AddressOf PrepareLineTemplate1},
                        {CASE_MISSION_PAIEMENT, AddressOf PrepareLineTemplate1},
                        {CASE_COMMANDE, AddressOf PrepareLineTemplate2},
                        {CASE_COMMANDE_INVEST, AddressOf PrepareLineTemplate2},
                        {CASE_MISSION, AddressOf PrepareLineTemplate2},
                        {CASE_COMMANDE_AJUSTEMENT, AddressOf PrepareLineTemplate3},
                        {CASE_REIMPUTATION, AddressOf PrepareLineTemplate3},
                        {CASE_AVOIR_PAIEMENT, AddressOf PrepareLineTemplate3},
                        {CASE_COMMANDE_AJUSTEMENT_INVEST, AddressOf PrepareLineTemplate3},
                        {CASE_REIMPUTATION_INVEST, AddressOf PrepareLineTemplate3},
                        {CASE_AVOIR_PAIEMENT_INVEST, AddressOf PrepareLineTemplate3},
                        {CASE_UNUSED, AddressOf PrepareLineTemplate4}}
    Private Shared ReadOnly PREPARED_LINES_MAPPING2 As New Dictionary(Of String, String) From {
        {CASE_COMMANDE_AJUSTEMENT, KEY_ORDER},
        {CASE_REIMPUTATION, KEY_ORDER},
        {CASE_AVOIR_PAIEMENT, KEY_ORDER},
        {CASE_COMMANDE_AJUSTEMENT_INVEST, KEY_INVEST},
        {CASE_REIMPUTATION_INVEST, KEY_INVEST},
        {CASE_AVOIR_PAIEMENT_INVEST, KEY_INVEST}}
    Private Shared ReadOnly PREPARED_LINES_MAPPING1 As New Dictionary(Of String, String()) From {
        {CASE_COMMANDE_PAIEMENT, {KEY_ORDER, CASE_COMMANDE, KEY_ORDER_PENDING}},
        {CASE_COMMANDE_PAIEMENT_INVEST, {KEY_INVEST, CASE_COMMANDE_INVEST, KEY_INVEST_PENDING}},
        {CASE_MISSION_PAIEMENT, {KEY_MISSION, CASE_MISSION, KEY_MISSION_PENDING}}}
    Private Shared ReadOnly PREPARED_LINES_MAPPING3 As New Dictionary(Of String, String) From {
        {CASE_COMMANDE, CASE_COMMANDE_PAIEMENT},
        {CASE_COMMANDE_INVEST, CASE_COMMANDE_PAIEMENT_INVEST},
        {CASE_MISSION, CASE_MISSION_PAIEMENT}}
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
    Private Const CASE_DEFAULT As String = "DEFAULT"
    Private Shared ReadOnly CASE_LIST As New Collection(Of String) From {
        CASE_AVOIR_PAIEMENT,
        CASE_COMMANDE,
        CASE_COMMANDE_AJUSTEMENT,
        CASE_COMMANDE_PAIEMENT,
        CASE_AVOIR_PAIEMENT_INVEST,
        CASE_COMMANDE_INVEST,
        CASE_COMMANDE_AJUSTEMENT_INVEST,
        CASE_COMMANDE_PAIEMENT_INVEST,
        CASE_MISSION,
        CASE_MISSION_PAIEMENT,
        CASE_REIMPUTATION,
        CASE_REIMPUTATION_INVEST,
        CASE_UNUSED
        }
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
    Private NewWorkbook As Workbook
    Private BaseWorksheet As Worksheet
    Private NewWorksheet As Worksheet
    Private TrashWorksheet As Worksheet
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
        BaseWorksheet = CType(AddinApplication.ActiveSheet, Worksheet)
        Dim MaxNameLength = Math.Min(20, BaseWorksheet.Name.Length)
        NewWorksheet = CreateWorksheet(BaseWorksheet, BaseWorksheet.Name.Substring(0, MaxNameLength) & " Extracted")
        TrashWorksheet = CreateWorksheet(NewWorksheet, BaseWorksheet.Name.Substring(0, MaxNameLength) & " Trash")

        Dim R1 As Range = BaseWorksheet.UsedRange
        FirstDataRow = R1.End(XlDirection.xlDown).Row
        LastDataRow = R1.End(XlDirection.xlDown).End(XlDirection.xlDown).Row
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
    Private Function CopyWorkbook() As Workbook
        Dim oldWorkbook As Workbook = AddinApplication.Workbooks.Open(aSourceName)
        Dim OriginalWorksheet As Worksheet = oldWorkbook.ActiveSheet
        Dim newWorkbook As Workbook = AddinApplication.Workbooks.Add()
        Dim WorksheetToRemove As Worksheet = newWorkbook.ActiveSheet
        OriginalWorksheet.Copy(WorksheetToRemove)
        WorksheetToRemove.Delete()
        Dim CopiedWorksheet As Worksheet = newWorkbook.ActiveSheet
        CopiedWorksheet.Range("1:1").Delete(XlDeleteShiftDirection.xlShiftUp)
        oldWorkbook.Close()
        newWorkbook.SaveAs(DestinationName)

        Return newWorkbook
    End Function
    Private Function CreateWorksheet(baseWorksheet As Worksheet, sheetName As String) As Worksheet
        Dim worksheet = CType(AddinApplication.Worksheets.Add(After:=baseWorksheet), Worksheet)
        worksheet.Name = sheetName
        Return worksheet
    End Function
    Private Sub PrepareExtractFromTo(baseWorksheet As Worksheet)
        Globals.ThisAddIn.NameStep("Lecture des lignes")
        FeedTable(baseWorksheet)
        Globals.ThisAddIn.NameStep("Traitement des lignes")
        ReduceTable()
    End Sub
    Private Sub ExtractFromTo(baseWorksheet As Worksheet, newWorksheet As Worksheet, trashWorksheet As Worksheet, previousExtraction As ExtractedData)
        ReduceLines(previousExtraction)
        Globals.ThisAddIn.NameStep("Générations des feuilles")
        Utils.CopyHeaders(baseWorksheet, trashWorksheet)
        Utils.CopyHeaders(baseWorksheet, newWorksheet)
        Utils.DumpData(newWorksheet, ReducedTable)
        Utils.DumpData(trashWorksheet, TrashTable)
    End Sub
    Private Sub ReduceTable()
        For Each Key As String In FullTable.Keys
            ReduceEntry(Key)
        Next
    End Sub

    Private Sub ReduceEntry(Key As String)
        Dim PreparedLines As Dictionary(Of String, Collection(Of BookLine)) = GetFreshEmptyLines()
        Dim MAX_K_DCompt As DateCompte = FullTable.Item(Key).Max().KDCompt
        For Each line As BookLine In FullTable.Item(Key)
            ReduceLine(PreparedLines, MAX_K_DCompt, line)
        Next
        PreReducedTable.Add(Key, PreparedLines)
    End Sub

    Private Sub ReduceLine(PreparedLines As Dictionary(Of String, Collection(Of BookLine)), MAX_K_DCompt As DateCompte, line As BookLine)
        Dim kind As String = line.BRubrique
        If TableReducers.ContainsKey(kind) Then
            TableReducers.Item(kind).Invoke(PreparedLines, MAX_K_DCompt, line)
        Else
            TableReducers.Item(CASE_DEFAULT).Invoke(PreparedLines, MAX_K_DCompt, line)
        End If
        Globals.ThisAddIn.NextStep()
    End Sub

    Private Shared Function GetFreshEmptyLines() As Dictionary(Of String, Collection(Of BookLine))
        Dim PreparedLines As New Dictionary(Of String, Collection(Of BookLine))
        For Each lineCase As String In CASE_LIST
            PreparedLines.Add(lineCase, New Collection(Of BookLine))
        Next
        Return PreparedLines
    End Function

    Private Shared Sub HandleOthers(PreparedLines As Dictionary(Of String, Collection(Of BookLine)), MAX_K_DCompt As DateCompte, Line As BookLine)
        If Line.KDCompt Is Nothing Then
            Line.KDCompt = MAX_K_DCompt
        End If
        Utils.AddLineToTable(CASE_UNUSED, PreparedLines, Line)
    End Sub

    Private Shared Sub HandleAvoirPaiement(PreparedLines As Dictionary(Of String, Collection(Of BookLine)), MAX_K_DCompt As DateCompte, Line As BookLine)
        If Line.KDCompt Is Nothing Then
            Line.KDCompt = MAX_K_DCompt
        End If
        If Utils.IsInvest(Line) Then
            Utils.AddLineToTable(CASE_AVOIR_PAIEMENT_INVEST, PreparedLines, Line)
        Else
            Utils.AddLineToTable(CASE_AVOIR_PAIEMENT, PreparedLines, Line)
        End If
    End Sub

    Private Shared Sub HandleReimputation(PreparedLines As Dictionary(Of String, Collection(Of BookLine)), MAX_K_DCompt As DateCompte, Line As BookLine)
        If Line.KDCompt Is Nothing Then
            Line.KDCompt = MAX_K_DCompt
        End If
        If Utils.IsInvest(Line) Then
            Utils.AddLineToTable(CASE_REIMPUTATION_INVEST, PreparedLines, Line)
        Else
            Utils.AddLineToTable(CASE_REIMPUTATION, PreparedLines, Line)
        End If
    End Sub

    Private Shared Sub HandleMission(PreparedLines As Dictionary(Of String, Collection(Of BookLine)), MAX_K_DCompt As DateCompte, Line As BookLine)
        Line.KDCompt = MAX_K_DCompt
        Utils.AddLineToTable(CASE_MISSION, PreparedLines, Line)
    End Sub

    Private Shared Sub HandleMissionPaiement(PreparedLines As Dictionary(Of String, Collection(Of BookLine)), MAX_K_DCompt As DateCompte, Line As BookLine)
        Line.KDCompt = MAX_K_DCompt
        Utils.AddLineToTable(CASE_MISSION_PAIEMENT, PreparedLines, Line)
    End Sub

    Private Shared Sub HandleCommande(PreparedLines As Dictionary(Of String, Collection(Of BookLine)), MAX_K_DCompt As DateCompte, Line As BookLine)
        Line.KDCompt = MAX_K_DCompt
        If Utils.IsInvest(Line) Then
            Utils.AddLineToTable(CASE_COMMANDE_INVEST, PreparedLines, Line)
        Else
            Utils.AddLineToTable(CASE_COMMANDE, PreparedLines, Line)
        End If
    End Sub

    Private Shared Sub HandleCommandeAjustement(PreparedLines As Dictionary(Of String, Collection(Of BookLine)), MAX_K_DCompt As DateCompte, Line As BookLine)
        Line.KDCompt = MAX_K_DCompt
        If Utils.IsInvest(Line) Then
            Utils.AddLineToTable(CASE_COMMANDE_AJUSTEMENT_INVEST, PreparedLines, Line)
        Else
            Utils.AddLineToTable(CASE_COMMANDE_AJUSTEMENT, PreparedLines, Line)
        End If
    End Sub

    Private Shared Sub HandleCommandePaiement(PreparedLines As Dictionary(Of String, Collection(Of BookLine)), MAX_K_DCompt As DateCompte, Line As BookLine)
        Line.KDCompt = MAX_K_DCompt
        If Utils.IsInvest(Line) Then
            Utils.AddLineToTable(CASE_COMMANDE_PAIEMENT_INVEST, PreparedLines, Line)
        Else
            Utils.AddLineToTable(CASE_COMMANDE_PAIEMENT, PreparedLines, Line)
        End If
    End Sub

    Private Sub ReduceLines(previousExtraction As ExtractedData)
        PreparePreviousData(previousExtraction)

        For Each preparedLines As Dictionary(Of String, Collection(Of BookLine)) In PreReducedTable.Values
            PrepareLines(preparedLines)
        Next
    End Sub

    Private Sub PrepareLines(preparedLines As Dictionary(Of String, Collection(Of BookLine)))
        For Each key As String In preparedLines.Keys
            If LinePreparers.ContainsKey(key) Then
                LinePreparers.Item(key).Invoke(key, preparedLines)
            Else
                Throw New ArgumentException($"{key} is not handled.")
            End If
        Next

    End Sub
    Private Sub PrepareLineTemplate1(keyMain As String, preparedLines As Dictionary(Of String, Collection(Of BookLine)))
        Dim paiementFound As Boolean = preparedLines.ContainsKey(keyMain) AndAlso preparedLines.Item(keyMain).Count > 0
        Dim mapping As String() = PREPARED_LINES_MAPPING1.Item(keyMain)
        If paiementFound Then
            AddLineToReducedTable(mapping(0), keyMain, preparedLines)
            AddLineToTrashTable(KEY_TRASH, mapping(1), preparedLines)
        Else
            AddLineToReducedTable(mapping(2), mapping(1), preparedLines)
        End If
    End Sub
    Private Sub PrepareLineTemplate2(keyMain As String, preparedLines As Dictionary(Of String, Collection(Of BookLine)))
        Dim keyAlt As String = PREPARED_LINES_MAPPING3.Item(keyMain)
        Dim paiementFound As Boolean = preparedLines.ContainsKey(keyAlt)
        Dim mapping As String() = PREPARED_LINES_MAPPING1.Item(keyAlt)
        If Not paiementFound Then
            AddLineToReducedTable(mapping(2), mapping(1), preparedLines)
        End If
    End Sub
    Private Sub PrepareLineTemplate3(keyMain As String, preparedLines As Dictionary(Of String, Collection(Of BookLine)))
        AddLineToReducedTable(PREPARED_LINES_MAPPING2.Item(keyMain), keyMain, preparedLines)
    End Sub
    Private Sub PrepareLineTemplate4(keyMain As String, preparedLines As Dictionary(Of String, Collection(Of BookLine)))
        AddLineToTrashTable(KEY_TRASH, keyMain, preparedLines)
    End Sub

    Private Sub PreparePreviousData(previousExtraction As ExtractedData)
        If previousExtraction IsNot Nothing Then
            PreparePendings(previousExtraction.PendingOrders, previousExtraction.SheetYear, CASE_COMMANDE)
            PreparePendings(previousExtraction.PendingMissions, previousExtraction.SheetYear, CASE_MISSION)
            PreparePendings(previousExtraction.PendingInvests, previousExtraction.SheetYear, CASE_COMMANDE_INVEST)
        End If
    End Sub

    Private Sub PreparePendings(pendings As Collection(Of BookLine), year As Integer, concernedCase As String)
        For Each line As BookLine In pendings
            If line.NFrom = "" Then
                line.NFrom = String.Format(CultureInfo.CurrentCulture, My.Resources.ResourceManager.GetString("LineFrom"), year)
            End If
            If PreReducedTable.ContainsKey(line.CNumeroFlux) Then
                PreReducedTable.Item(line.CNumeroFlux).Item(concernedCase).Add(line)
            Else
                PreReducedTable.Add(line.CNumeroFlux, New Dictionary(Of String, Collection(Of BookLine)) From {
                {concernedCase, New Collection(Of BookLine) From {line}}})
            End If
        Next
        pendings.Clear()
    End Sub

    Private Sub AddLineToReducedTable(keyReduced As String, keyPrepared As String, preparedLines As Dictionary(Of String, Collection(Of BookLine)))
        Utils.AddLineFromTableToTable(ReducedTable.Item(keyReduced), preparedLines.Item(keyPrepared))
    End Sub
    Private Sub AddLineToTrashTable(keyTrashed As String, keyPrepared As String, preparedLines As Dictionary(Of String, Collection(Of BookLine)))
        Utils.AddLineFromTableToTable(TrashTable.Item(keyTrashed), preparedLines.Item(keyPrepared))
    End Sub
    Private Sub FeedTable(baseWorksheet As Worksheet)
        Dim FullRange As Range = baseWorksheet.UsedRange.End(XlDirection.xlDown)
        For RowNum As Integer = 1 To NumberOfLines
            Dim Line As Range = FullRange.Item(RowNum, 3)
            Dim Key As String = Line.Value2
            Dim Data = New Collection(Of BookLine)
            If FullTable.ContainsKey(Key) Then
                FullTable.TryGetValue(Key, Data)
            Else
                FullTable.Add(Key, Data)
            End If

            Dim NewLine As BookLine = Utils.ReadLine(FullRange, RowNum)
            NewLine.MComment = ""
            NewLine.NFrom = ""
            Data.Add(NewLine)
            If NewLine.KDCompt IsNot Nothing Then
                aSheetYear = Math.Max(aSheetYear, NewLine.KDCompt.Year)
            End If
            Globals.ThisAddIn.NextStep()
        Next
    End Sub
End Class
