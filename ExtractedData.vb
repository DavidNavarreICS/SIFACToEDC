Imports System.Diagnostics
Imports System.Globalization
Imports Microsoft.Office.Interop.Excel

Public Class ExtractedData
    Public Class DateCompte : Implements IComparable(Of DateCompte)
        Public ReadOnly Day As Integer
        Public ReadOnly Month As Integer
        Public ReadOnly Year As Integer
        Public ReadOnly AsText As String
        Public Sub New(TextValue As String)
            Dim ThisDate As String() = TextValue.Split(".")
            Year = CInt(ThisDate(2))
            Month = CInt(ThisDate(1))
            Day = CInt(ThisDate(0))
            AsText = TextValue
        End Sub
        Public Function CompareTo(Other As DateCompte) As Integer Implements IComparable(Of DateCompte).CompareTo
            If Other Is Nothing Then
                Return String.Compare(AsText, Nothing)
            Else
                If Year <> Other.Year Then
                    Return Year - Other.Year
                ElseIf Month <> Other.Month Then
                    Return Month - Other.Month
                Else
                    Return Day - Other.Day
                End If
            End If
        End Function
    End Class
    Public Class BookLine : Implements IComparable(Of BookLine)
        Public A_Cptegen As String
        Public B_Rubrique As String
        Public C_NumeroFlux As String
        Public D_Nom As String
        Public E_Libelle As String
        Public F_MntEngHTR As Double
        Public G_MontantPa As Double
        Public H_Rapprochmt As String
        Public I_RefFactF As String
        Public J_DatePce As String
        Public K_DCompt As DateCompte
        Public L_NumPiece As String
        Public M_Comment As String
        Public N_From As String

        Public Function CompareTo(other As BookLine) As Integer Implements IComparable(Of BookLine).CompareTo
            If K_DCompt IsNot Nothing Then
                Return K_DCompt.CompareTo(other.K_DCompt)
            ElseIf other.K_DCompt IsNot Nothing Then
                Return other.K_DCompt.CompareTo(K_DCompt)
            Else
                Return 0
            End If
        End Function
    End Class
    Private ReadOnly FullTable As New Dictionary(Of String, List(Of BookLine))
    Private ReadOnly ReducedTable As New Dictionary(Of String, List(Of BookLine))
    Private ReadOnly PreReducedTable As New Dictionary(Of String, Dictionary(Of String, List(Of BookLine)))
    Private ReadOnly TrashTable As New Dictionary(Of String, List(Of BookLine))
    Private FirstDataRow As Integer
    Private LastDataRow As Integer
    Private LastDataColumn As Integer
    Private FirstDataColumn As Integer
    Public NumberOfLines As Integer
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
    Public ReadOnly SourceName As String
    Private ReadOnly DestinationName As String
    Private ReadOnly AddinApplication As Application
    Private NewWorkbook As Excel.Workbook
    Private BaseWorksheet As Excel.Worksheet
    Private NewWorksheet As Excel.Worksheet
    Private TrashWorksheet As Excel.Worksheet
    Public SheetYear As Integer
    Public ReadOnly Property Orders As List(Of BookLine)
        Get
            Return ReducedTable.Item(KEY_ORDER)
        End Get
    End Property
    Public ReadOnly Property Invests As List(Of BookLine)
        Get
            Return ReducedTable.Item(KEY_INVEST)
        End Get
    End Property
    Public ReadOnly Property Missions As List(Of BookLine)
        Get
            Return ReducedTable.Item(KEY_MISSION)
        End Get
    End Property
    Public ReadOnly Property PendingOrders As List(Of BookLine)
        Get
            Return ReducedTable.Item(KEY_ORDER_PENDING)
        End Get
    End Property
    Public ReadOnly Property PendingInvests As List(Of BookLine)
        Get
            Return ReducedTable.Item(KEY_INVEST_PENDING)
        End Get
    End Property
    Public ReadOnly Property PendingMissions As List(Of BookLine)
        Get
            Return ReducedTable.Item(KEY_MISSION_PENDING)
        End Get
    End Property
    Public Sub New(ASourceName As String, ADestinationName As String, AAddinApplication As Application)
        Me.SourceName = ASourceName
        Me.DestinationName = ADestinationName
        Me.AddinApplication = AAddinApplication
        PrepareWorkbook()
    End Sub
    Public Sub PrepareWorkbook()
        SheetYear = 0
        NewWorkbook = CopyWorkbook()
        BaseWorksheet = CType(AddinApplication.ActiveSheet, Excel.Worksheet)
        Dim MaxNameLength = Math.Min(20, BaseWorksheet.Name.Length)
        NewWorksheet = CreateWorksheet(BaseWorksheet, BaseWorksheet.Name.Substring(0, MaxNameLength) & " Extracted")
        TrashWorksheet = CreateWorksheet(NewWorksheet, BaseWorksheet.Name.Substring(0, MaxNameLength) & " Trash")

        Dim R1 As Excel.Range = BaseWorksheet.UsedRange
        FirstDataRow = R1.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).Row
        LastDataRow = R1.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).Row
        LastDataColumn = R1.End(Microsoft.Office.Interop.Excel.XlDirection.xlToRight).Column
        FirstDataColumn = R1.Column
        NumberOfLines = LastDataRow - FirstDataRow + 1
    End Sub
    Public Sub DoPrepareExtract()
        TrashTable.Clear()
        FullTable.Clear()
        ReducedTable.Clear()
        ReducedTable.Add(KEY_ORDER_PENDING, New List(Of BookLine))
        ReducedTable.Add(KEY_ORDER, New List(Of BookLine))
        ReducedTable.Add(KEY_INVEST_PENDING, New List(Of BookLine))
        ReducedTable.Add(KEY_INVEST, New List(Of BookLine))
        ReducedTable.Add(KEY_MISSION_PENDING, New List(Of BookLine))
        ReducedTable.Add(KEY_MISSION, New List(Of BookLine))
        TrashTable.Add(KEY_TRASH, New List(Of BookLine))
        PrepareExtractFromTo(BaseWorksheet)
    End Sub
    Public Sub DoExtract(PreviousExtraction As ExtractedData)
        ExtractFromTo(BaseWorksheet, NewWorksheet, TrashWorksheet, PreviousExtraction)
        NewWorkbook.Save()
        NewWorkbook.Close()
    End Sub
    Private Function CopyWorkbook() As Excel.Workbook
        Dim oldWorkbook As Excel.Workbook = AddinApplication.Workbooks.Open(SourceName)
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

    Private Function CreateWorksheet(baseWorksheet As Excel.Worksheet, SheetName As String) As Excel.Worksheet
        Dim NewWorksheet = CType(AddinApplication.Worksheets.Add(After:=baseWorksheet), Excel.Worksheet)
        NewWorksheet.Name = SheetName
        Return NewWorksheet
    End Function
    Private Sub PrepareExtractFromTo(BaseWorksheet As Excel.Worksheet)
        Globals.ThisAddIn.NameStep("Lecture des lignes")
        FeedTable(BaseWorksheet)
        Globals.ThisAddIn.NameStep("Traitement des lignes")
        ReduceTable()
    End Sub
    Private Sub ExtractFromTo(BaseWorksheet As Excel.Worksheet, NewWorksheet As Excel.Worksheet, TrashWorksheet As Excel.Worksheet, previousExtraction As ExtractedData)
        ReduceLines(previousExtraction)
        Globals.ThisAddIn.NameStep("Générations des feuilles")
        CopyHeaders(BaseWorksheet, TrashWorksheet)
        CopyHeaders(BaseWorksheet, NewWorksheet)
        DumpData(NewWorksheet, ReducedTable)
        DumpData(TrashWorksheet, TrashTable)
    End Sub

    Private Sub ReduceTable()
        For Each Key As String In FullTable.Keys
            Dim PreparedLines As New Dictionary(Of String, List(Of BookLine)) From {
            {CASE_AVOIR_PAIEMENT, New List(Of BookLine)},
            {CASE_COMMANDE, New List(Of BookLine)},
            {CASE_COMMANDE_AJUSTEMENT, New List(Of BookLine)},
            {CASE_COMMANDE_PAIEMENT, New List(Of BookLine)},
            {CASE_AVOIR_PAIEMENT_INVEST, New List(Of BookLine)},
            {CASE_COMMANDE_INVEST, New List(Of BookLine)},
            {CASE_COMMANDE_AJUSTEMENT_INVEST, New List(Of BookLine)},
            {CASE_COMMANDE_PAIEMENT_INVEST, New List(Of BookLine)},
            {CASE_MISSION, New List(Of BookLine)},
            {CASE_MISSION_PAIEMENT, New List(Of BookLine)},
            {CASE_REIMPUTATION, New List(Of BookLine)},
            {CASE_REIMPUTATION_INVEST, New List(Of BookLine)},
            {CASE_UNUSED, New List(Of BookLine)}
            }
            Dim MAX_K_DCompt As DateCompte = FullTable.Item(Key).Max().K_DCompt
            For Each Line As BookLine In FullTable.Item(Key)
                Dim Kind As String = Line.B_Rubrique
                Select Case Kind
                    Case CASE_COMMANDE_PAIEMENT
                        Line.K_DCompt = MAX_K_DCompt
                        If Line.A_Cptegen.Trim.StartsWith("2") Then
                            AddLineToTable(CASE_COMMANDE_PAIEMENT_INVEST, PreparedLines, Line)
                        Else
                            AddLineToTable(CASE_COMMANDE_PAIEMENT, PreparedLines, Line)
                        End If
                    Case CASE_COMMANDE_AJUSTEMENT
                        Line.K_DCompt = MAX_K_DCompt
                        If Line.A_Cptegen.Trim.StartsWith("2") Then
                            AddLineToTable(CASE_COMMANDE_AJUSTEMENT_INVEST, PreparedLines, Line)
                        Else
                            AddLineToTable(CASE_COMMANDE_AJUSTEMENT, PreparedLines, Line)
                        End If
                    Case CASE_COMMANDE
                        Line.K_DCompt = MAX_K_DCompt
                        If Line.A_Cptegen.Trim.StartsWith("2") Then
                            AddLineToTable(CASE_COMMANDE_INVEST, PreparedLines, Line)
                        Else
                            AddLineToTable(CASE_COMMANDE, PreparedLines, Line)
                        End If
                    Case CASE_MISSION_PAIEMENT
                        Line.K_DCompt = MAX_K_DCompt
                        AddLineToTable(CASE_MISSION_PAIEMENT, PreparedLines, Line)
                    Case CASE_MISSION
                        Line.K_DCompt = MAX_K_DCompt
                        AddLineToTable(CASE_MISSION, PreparedLines, Line)
                    Case CASE_REIMPUTATION
                        If Line.A_Cptegen.Trim.StartsWith("2") Then
                            AddLineToTable(CASE_REIMPUTATION_INVEST, PreparedLines, Line)
                        Else
                            AddLineToTable(CASE_REIMPUTATION, PreparedLines, Line)
                        End If
                    Case CASE_AVOIR_PAIEMENT
                        If Line.A_Cptegen.Trim.StartsWith("2") Then
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
    Private Sub ReduceLines(previousExtraction As ExtractedData)
        If previousExtraction IsNot Nothing Then
            For Each Line As BookLine In previousExtraction.PendingOrders
                If Line.N_From = "" Then
                    Line.N_From = $"Ligne venant de {previousExtraction.SheetYear}"
                End If
                If PreReducedTable.ContainsKey(Line.C_NumeroFlux) Then
                    PreReducedTable.Item(Line.C_NumeroFlux).Item(CASE_COMMANDE).Add(Line)
                Else
                    PreReducedTable.Add(Line.C_NumeroFlux, New Dictionary(Of String, List(Of BookLine)) From {
                    {CASE_COMMANDE, New List(Of BookLine) From {Line}}})
                End If
            Next
            previousExtraction.PendingOrders.Clear()
            For Each Line As BookLine In previousExtraction.PendingMissions
                If Line.N_From = "" Then
                    Line.N_From = $"Ligne venant de {previousExtraction.SheetYear}"
                End If
                If PreReducedTable.ContainsKey(Line.C_NumeroFlux) Then
                    PreReducedTable.Item(Line.C_NumeroFlux).Item(CASE_MISSION).Add(Line)
                Else
                    PreReducedTable.Add(Line.C_NumeroFlux, New Dictionary(Of String, List(Of BookLine)) From {
                    {CASE_MISSION, New List(Of BookLine) From {Line}}})
                End If
            Next
            previousExtraction.PendingMissions.Clear()
            For Each Line As BookLine In previousExtraction.PendingInvests
                If Line.N_From = "" Then
                    Line.N_From = $"Ligne venant de {previousExtraction.SheetYear}"
                End If
                If PreReducedTable.ContainsKey(Line.C_NumeroFlux) Then
                    PreReducedTable.Item(Line.C_NumeroFlux).Item(CASE_AVOIR_PAIEMENT_INVEST).Add(Line)
                Else
                    PreReducedTable.Add(Line.C_NumeroFlux, New Dictionary(Of String, List(Of BookLine)) From {
                    {CASE_COMMANDE_INVEST, New List(Of BookLine) From {Line}}})
                End If
            Next
            previousExtraction.PendingInvests.Clear()
        End If
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
            Dim PreparedLines As Dictionary(Of String, List(Of BookLine)) = PreReducedTable.Item(Key)

            If CommandePaiementFound Then
                ReducedTable.Item(KEY_ORDER).AddRange(PreparedLines.Item(CASE_COMMANDE_PAIEMENT))
                TrashTable.Item(KEY_TRASH).AddRange(PreparedLines.Item(CASE_COMMANDE))
            ElseIf PreparedLines.ContainsKey(CASE_COMMANDE) Then
                ReducedTable.Item(KEY_ORDER_PENDING).AddRange(PreparedLines.Item(CASE_COMMANDE))
            End If
            If InvestissementPaiementFound Then
                ReducedTable.Item(KEY_INVEST).AddRange(PreparedLines.Item(CASE_COMMANDE_PAIEMENT_INVEST))
                TrashTable.Item(KEY_TRASH).AddRange(PreparedLines.Item(CASE_COMMANDE_INVEST))
            ElseIf PreparedLines.ContainsKey(CASE_COMMANDE_INVEST) Then
                ReducedTable.Item(KEY_INVEST_PENDING).AddRange(PreparedLines.Item(CASE_COMMANDE_INVEST))
            End If
            If PreparedLines.ContainsKey(CASE_COMMANDE_AJUSTEMENT) Then
                ReducedTable.Item(KEY_ORDER).AddRange(PreparedLines.Item(CASE_COMMANDE_AJUSTEMENT))
            End If
            If PreparedLines.ContainsKey(CASE_COMMANDE_AJUSTEMENT_INVEST) Then
                ReducedTable.Item(KEY_INVEST).AddRange(PreparedLines.Item(CASE_COMMANDE_AJUSTEMENT_INVEST))
            End If
            If MissionPaiementFound Then
                ReducedTable.Item(KEY_MISSION).AddRange(PreparedLines.Item(CASE_MISSION_PAIEMENT))
                TrashTable.Item(KEY_TRASH).AddRange(PreparedLines.Item(CASE_MISSION))
            ElseIf PreparedLines.ContainsKey(CASE_MISSION) Then
                ReducedTable.Item(KEY_MISSION_PENDING).AddRange(PreparedLines.Item(CASE_MISSION))
            End If
            If PreparedLines.ContainsKey(CASE_REIMPUTATION) Then
                ReducedTable.Item(KEY_ORDER).AddRange(PreparedLines.Item(CASE_REIMPUTATION))
            End If
            If PreparedLines.ContainsKey(CASE_REIMPUTATION_INVEST) Then
                ReducedTable.Item(KEY_INVEST).AddRange(PreparedLines.Item(CASE_REIMPUTATION_INVEST))
            End If
            If PreparedLines.ContainsKey(CASE_AVOIR_PAIEMENT) Then
                ReducedTable.Item(KEY_ORDER).AddRange(PreparedLines.Item(CASE_AVOIR_PAIEMENT))
            End If
            If PreparedLines.ContainsKey(CASE_AVOIR_PAIEMENT_INVEST) Then
                ReducedTable.Item(KEY_INVEST).AddRange(PreparedLines.Item(CASE_AVOIR_PAIEMENT_INVEST))
            End If
            If PreparedLines.ContainsKey(CASE_UNUSED) Then
                TrashTable.Item(KEY_TRASH).AddRange(PreparedLines.Item(CASE_UNUSED))
            End If
        Next
    End Sub
    Private Sub AddLineToTable(Key As String, PreparedLines As Dictionary(Of String, List(Of BookLine)), Line As BookLine)
        If Not PreparedLines.ContainsKey(Key) Then
            Dim NewLines As New List(Of BookLine) From {
                Line
            }
            PreparedLines.Add(Key, NewLines)
        Else
            PreparedLines.Item(Key).Add(Line)
        End If
    End Sub

    Private Sub CopyHeaders(baseWorksheet As Excel.Worksheet, newWorksheet As Excel.Worksheet)
        Dim SourceRange As Excel.Range = baseWorksheet.UsedRange.Rows(1)
        Dim DestRange As Excel.Range = newWorksheet.Range("A1")
        SourceRange.Copy(DestRange)
        For I As Integer = 1 To 12
            Dim RDest As Excel.Range = DestRange.Cells(1, I)
            Dim RSource As Excel.Range = SourceRange.Cells(1, I)
            RDest.ColumnWidth = RSource.ColumnWidth
        Next
    End Sub
    Private Sub DumpData(Worksheet As Excel.Worksheet, DataTable As Dictionary(Of String, List(Of BookLine)))
        Dim CurrentLine As Integer = 1
        Dim StartRange As Excel.Range = Worksheet.Range("A3")
        For Each LineList As List(Of BookLine) In DataTable.Values
            For Each Line As BookLine In LineList
                StartRange.Cells(CurrentLine, 1).Value2 = Line.A_Cptegen
                StartRange.Cells(CurrentLine, 2).Value2 = Line.B_Rubrique
                StartRange.Cells(CurrentLine, 3).Value2 = Line.C_NumeroFlux
                StartRange.Cells(CurrentLine, 4).Value2 = Line.D_Nom
                StartRange.Cells(CurrentLine, 5).Value2 = Line.E_Libelle
                StartRange.Cells(CurrentLine, 6).Value2 = Line.F_MntEngHTR
                StartRange.Cells(CurrentLine, 7).Value2 = Line.G_MontantPa
                StartRange.Cells(CurrentLine, 8).Value2 = Line.H_Rapprochmt
                StartRange.Cells(CurrentLine, 9).Value2 = Line.I_RefFactF
                StartRange.Cells(CurrentLine, 10).Value2 = Line.J_DatePce
                StartRange.Cells(CurrentLine, 11).Value2 = GetDateCompteAsText(Line)
                StartRange.Cells(CurrentLine, 12).Value2 = Line.L_NumPiece
                CurrentLine += 1
                Globals.ThisAddIn.NextStep()
                Globals.ThisAddIn.NextStep()
                Globals.ThisAddIn.NextStep()
                Globals.ThisAddIn.NextStep()
                Globals.ThisAddIn.NextStep()
            Next
        Next
    End Sub

    Public Shared Function GetDateCompteAsText(Line As BookLine) As String
        If Line.K_DCompt IsNot Nothing Then
            Return Line.K_DCompt.AsText
        Else
            Return ""
        End If
    End Function

    Private Sub FeedTable(BaseWorksheet As Excel.Worksheet)
        Dim FullRange As Excel.Range = BaseWorksheet.UsedRange.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown)
        For RowNum As Integer = 1 To NumberOfLines
            Dim Line As Excel.Range = FullRange.Item(RowNum, 3)
            Dim Key As String = Line.Value2
            Dim Data = New List(Of BookLine)
            If FullTable.ContainsKey(Key) Then
                FullTable.TryGetValue(Key, Data)
            Else
                FullTable.Add(Key, Data)
            End If

            Dim NewLine As BookLine = ReadLine(FullRange, RowNum)
            Data.Add(NewLine)
            If NewLine.K_DCompt IsNot Nothing Then
                SheetYear = Math.Max(SheetYear, NewLine.K_DCompt.Year)
            End If
            Globals.ThisAddIn.NextStep()
        Next
    End Sub

    Public Shared Function ReadLine(FullRange As Range, RowNum As Integer) As BookLine
        Return New BookLine With {
            .A_Cptegen = FullRange.Cells(RowNum, 1).Value2,
            .B_Rubrique = FullRange.Cells(RowNum, 2).Value2,
            .C_NumeroFlux = FullRange.Cells(RowNum, 3).Value2,
            .D_Nom = FullRange.Cells(RowNum, 4).Value2,
            .E_Libelle = FullRange.Cells(RowNum, 5).Value2,
            .F_MntEngHTR = GetNumber(FullRange, RowNum, 6),
            .G_MontantPa = GetNumber(FullRange, RowNum, 7),
            .H_Rapprochmt = FullRange.Cells(RowNum, 8).Value2,
            .I_RefFactF = FullRange.Cells(RowNum, 9).Value2,
            .J_DatePce = FullRange.Cells(RowNum, 10).Value2,
            .K_DCompt = GetDateCompte(FullRange.Cells(RowNum, 11).Value2),
            .L_NumPiece = FullRange.Cells(RowNum, 12).Value2,
            .M_Comment = "",
            .N_From = ""
        }
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
