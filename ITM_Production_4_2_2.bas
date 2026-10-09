Attribute VB_Name = "ITM_Production"

Option Explicit

Private Const DEBUG_MODE As Boolean = True

Private Const DEBUG_AMBIGUITY As Boolean = True

Private Const HIDE_SUPPORTING_SHEETS As Boolean = True

Private MethodCounts As Object
Private MethodVolumes As Object

Private NextMatchID As Long
Private MatchIDs As Object
Private NextCandidateID As Long
Private NextContradictionID As Long
Private TransactionStatus As Object
Private CandidateLookup As Object
Private TransactionRowLookup As Object

Private TransferPatterns As Variant
Private TransferFlags() As Boolean

Private ConfirmationNumbers() As String
Private NarrativeIDs() As String
Private ReferencedAccounts() As String
Private EmbeddedAccounts() As String

Private ConfIndex As Object
Private NarrativeIndex As Object
Private RefAcctIndex As Object
Private AmtDateIndex As Object
Private TransferToSuffixes As Object
Private AccountIndex As Object

Private AmbiguousMethodCounts As Object
Private AmbiguousClusters As Object
Private ClusterSizes As Object
Private ClusterValues As Object
Private ClusterOriginMethods As Object
Private AmbiguousEdges As Object
Private SingleSuffixes As Object

Private CurrentRunID As String
Private CurrentVersion As String
Private RunStartTime As Double

Private Const STATUS_PENDING As String = "Pending"
Private Const STATUS_MATCHED As String = "Matched"
Private Const STATUS_AMBIGUOUS As String = "Ambiguous"
Private Const STATUS_UNMATCHED As String = "Unmatched"
Private Const STATUS_CONTRADICTED As String = "Contradicted"


' ACCOUNT STATISTICS ARRAY INDEXES - Used for Tables on Summmary

Private Const ACCT_MATCHED_IN_COUNT As Long = 0
Private Const ACCT_MATCHED_OUT_COUNT As Long = 1

Private Const ACCT_AMBIGUOUS_IN_COUNT As Long = 2
Private Const ACCT_AMBIGUOUS_OUT_COUNT As Long = 3

Private Const ACCT_UNMATCHED_IN_COUNT As Long = 4
Private Const ACCT_UNMATCHED_OUT_COUNT As Long = 5

Private Const ACCT_MATCHED_IN_VOLUME As Long = 6
Private Const ACCT_MATCHED_OUT_VOLUME As Long = 7

Private Const ACCT_AMBIGUOUS_IN_VOLUME As Long = 8
Private Const ACCT_AMBIGUOUS_OUT_VOLUME As Long = 9

Private Const ACCT_UNMATCHED_IN_VOLUME As Long = 10
Private Const ACCT_UNMATCHED_OUT_VOLUME As Long = 11

Private Const ACCT_SOURCE_DICT As Long = 12
Private Const ACCT_DESTINATION_DICT As Long = 13

Private Const ACCT_LARGEST_RECEIVED As Long = 14
Private Const ACCT_LARGEST_RECEIVED_STATUS As Long = 15

Private Const ACCT_LARGEST_SENT As Long = 16
Private Const ACCT_LARGEST_SENT_STATUS As Long = 17

Private Const SUMMARY_PANEL_FIRST_ROW As Long = 39
Private Const SUMMARY_PANEL_TABLE_ROW As Long = 42
Private Const SUMMARY_PANEL_FIRST_COL As Long = 1
Private Const SUMMARY_PANEL_LAST_COL As Long = 15


Private Const UNMATCH_NO_CANDIDATE As String = _
    "No Candidate Found"

Private Const UNMATCH_COUNTERPARTY_MISSING As String = _
    "Counterparty Missing"

Private Const UNMATCH_DATE_MISMATCH As String = _
    "Date Mismatch"

Private Const UNMATCH_AMOUNT_MISMATCH As String = _
    "Amount Mismatch"

Private Const UNMATCH_UNSUPPORTED_TYPE As String = _
    "Unsupported Transfer Type"
    
Private Const UNMATCH_CONTRADICTION_EXISTS As String = _
    "Candidate Rejected By Contradiction"
    

'=========================================================
' ITM Internal Transfer Matching
'
' PURPOSE
' -------
' Finds internal transfers between accounts and seperates unmatched
'
' DESIGN GOALS
' ------------
' 1. No PERSONAL.XLSB dependency.
' 2. No Code-column dependency.
' 3. Runs against any active worksheet.
' 4. Prevents false positives.
' 5. Easily expandable - lol
'
'
' A transfer is matched only if exactly one
' corresponding transfer exists.
'
' If multiple candidates exist:
'
'       Ambiguous
'
' If no candidate exists:
'
'       Unmatched
'
'==========================================================
' PROCESSING LOGIC
'
' 1. Read and normalize the source transaction data.
' 2. Attempt matches using the strongest evidence first:
'       Confirmation numbers
'       Narrative relationships
'       Referenced or embedded accounts
'       Transfer account suffixes
'       Amount and date
' 3. Send unresolved candidate relationships to ambiguity
'    analysis for graph-based resolution, followed by 2ndary
'    matching.
' 4. Classify all transactions as matched, ambiguous, or
'    unmatched.
' 5. Build technical report and analytics from the final
'    transaction outcomes.
'
' Matching passes must preserve earlier confirmed matches.
'==========================================================


Public Sub RunITM_Controller()

    Dim selectedCell As Range

    Dim selectedHostWb As Workbook
    Dim selectedSourceWs As Worksheet

    Dim hostWb As Workbook
    Dim sourceWs As Worksheet
    Dim ws As Worksheet

    Dim currentStage As String

    Dim errorNumber As Long
    Dim errorDescription As String
    Dim errorSource As String

    Dim originalScreenUpdating As Boolean
    Dim originalEnableEvents As Boolean
    Dim originalDisplayAlerts As Boolean
    Dim originalCalculation As XlCalculation
    
    Dim colAcct As Long
    Dim colCodeDesc As Long
    Dim colDate As Long
    Dim colAmount As Long
    Dim colDescription As Long
    
    Dim lastRow As Long
    Dim lastCol As Long
    Dim data As Variant
    
    Dim matchedCount As Long
    Dim ambiguousCount As Long
    Dim unmatchedCount As Long
    
    Dim matched() As Boolean
    Dim ambiguous() As Boolean
    Dim pendingAmbiguous() As Boolean
    Dim ambiguityPairs As Collection


    '---Capture current Excel state---


    originalScreenUpdating = _
        Application.ScreenUpdating

    originalEnableEvents = _
        Application.EnableEvents

    originalDisplayAlerts = _
        Application.DisplayAlerts

    originalCalculation = _
        Application.Calculation

    On Error GoTo ErrorHandler


    '---Select worksheet---


    currentStage = _
        "Selecting transaction source worksheet"
    
    Set selectedCell = _
        SelectITMSourceCell()
    
    If selectedCell Is Nothing Then
    
        Debug.Print _
            Format$(Now, "hh:mm:ss") & _
            " | Source selection cancelled."
    
        GoTo CleanExit
    
    End If
    

    Set selectedSourceWs = _
        selectedCell.Worksheet

    Set selectedHostWb = _
        selectedSourceWs.Parent

    If selectedHostWb Is ThisWorkbook Then

        MsgBox _
            "Select a worksheet in the transaction workbook, " & _
            "not a worksheet in this workbook.", _
            vbExclamation, _
            "ITM"

        GoTo CleanExit

    End If


    '---Designate environment---


    currentStage = _
        "Initializing run context"

    InitializeRunContext _
        selectedHostWb, _
        selectedSourceWs

    Set sourceWs = _
        modRunContext.sourceWs

    Set hostWb = _
        modRunContext.hostWb

    Set ws = sourceWs

       
    Debug.Print String(70, "=")
    Debug.Print "RUNNING ITM CONTROLLER"
    Debug.Print "Code workbook:   " & ThisWorkbook.Name
    Debug.Print "Code path:       " & ThisWorkbook.FullName
    Debug.Print "Host workbook:   " & hostWb.Name
    Debug.Print "Source worksheet:" & sourceWs.Name
    Debug.Print "Source parent:   " & sourceWs.Parent.Name
    Debug.Print "Context valid:   " & _
        CStr(sourceWs.Parent Is hostWb)
    Debug.Print String(70, "=")
    
    
    
    currentStage = _
        "Normalizing source worksheet"
        
    NormalizeSourceWorksheet ws

    
    CurrentVersion = "4.2.2"

    CurrentRunID = _
        Format(Now, "yyyymmdd_hhnnss")
    
    RunStartTime = Timer
    
    
    currentStage = _
        "Validating source headers"
    
   
    colAcct = FindColumnByAliases(ws, _
        Array("Customer Account", "Account", "Acct"))

    colCodeDesc = FindColumnByAliases(ws, _
        Array("Code Description", _
              "Transaction Description", _
              "Tran Description"))

    colDate = FindColumnByAliases(ws, _
        Array("Date", "Transaction Date"))

    colAmount = FindColumnByAliases(ws, _
        Array("Total Value", _
              "Non-Cash Value", _
              "Amount"))

    colDescription = FindColumnByAliases(ws, _
        Array("Description", "Transaction Detail"))

    If colAcct = 0 _
    Or colCodeDesc = 0 _
    Or colDate = 0 _
    Or colAmount = 0 _
    Or colDescription = 0 Then

        MsgBox _
        "Required columns not found.", _
        vbCritical

        Exit Sub

    End If
    
    
    currentStage = _
        "Creating output worksheets"

    CreateOutputSheets hostWb
    
    InitializeStatistics
    
    TransferPatterns = LoadTransferPatterns()
    
    
    currentStage = _
        "Data Read and Indexing"

    
    lastRow = _
        ws.Cells(ws.rows.Count, colAcct).End(xlUp).Row
        
    lastCol = _
        ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
   
    data = ws.Range( _
        ws.Cells(1, 1), _
        ws.Cells(lastRow, lastCol) _
    ).Value
    
    '--Initial data read--
    
    BuildIndexes _
        data, _
        lastRow, _
        colAcct, _
        colCodeDesc, _
        colDescription, _
        colDate, _
        colAmount
        
    Set TransactionRowLookup = Nothing
    
    '--Data storeage--
    
    currentStage = _
        "Store transactions to local memory"
    
    PersistTransactionsWarehouse _
        hostWb, _
        data, _
        lastRow, _
        colAcct, _
        colDate, _
        colAmount, _
        colCodeDesc, _
        colDescription
        
        
    BuildTransactionRowLookup hostWb
    
    ReDim matched(2 To lastRow)
    ReDim ambiguous(2 To lastRow)
    ReDim pendingAmbiguous(2 To lastRow)
    Set ambiguityPairs = New Collection
    
    
    '--Main logic passes--
    
    DebugLog String(59, "=")
    DebugLog "START MAIN PASS"
    
    currentStage = _
        "Main Pass Confirmation Numbers"
    
    PassConfirmationNumbers _
        hostWb, _
        ws, _
        data, _
        lastRow, _
        colAcct, _
        colCodeDesc, _
        colDate, _
        colAmount, _
        colDescription, _
        matched, _
        ambiguous, _
        ambiguityPairs
        
    currentStage = _
        "Main Pass Narrative Pairs"
        
      PassNarrativePairs _
        hostWb, _
        ws, _
        data, _
        lastRow, _
        colAcct, _
        colCodeDesc, _
        colDate, _
        colAmount, _
        colDescription, _
        matched, _
        ambiguous, _
        ambiguityPairs
        
        
    currentStage = _
        "Main Pass Referenced Accounts"
    
      PassReferencedAccounts _
        hostWb, _
        ws, _
        data, _
        lastRow, _
        colAcct, _
        colCodeDesc, _
        colDate, _
        colAmount, _
        colDescription, _
        matched, _
        ambiguous, _
        ambiguityPairs
        
           
    currentStage = _
        "Main Pass Embedded Accounts"
        
      PassEmbeddedAccounts _
        hostWb, _
        ws, _
        data, _
        lastRow, _
        colAcct, _
        colCodeDesc, _
        colDate, _
        colAmount, _
        colDescription, _
        matched, _
        ambiguous, _
        ambiguityPairs

    
    currentStage = _
        "Main Pass Suffixes"
        
      PassTransferSuffixes _
        hostWb, _
        ws, _
        data, _
        lastRow, _
        colAcct, _
        colCodeDesc, _
        colDate, _
        colAmount, _
        colDescription, _
        matched, _
        ambiguous, _
        ambiguityPairs
        
        
    currentStage = _
        "Main Pass Amount + Date"
        
      PassAmountDate _
        hostWb, _
        ws, _
        data, _
        lastRow, _
        colAcct, _
        colCodeDesc, _
        colDate, _
        colAmount, _
        colDescription, _
        matched, _
        ambiguous, _
        pendingAmbiguous, _
        ambiguityPairs
        
    DebugLog "END MAIN PASS"
    DebugLog String(59, "=")
        
        
    currentStage = _
        "Resolving ambiguous activity"
    
      ResolveAmbiguities _
        hostWb, _
        ws, _
        data, _
        lastRow, _
        colAcct, _
        colCodeDesc, _
        colDate, _
        colAmount, _
        colDescription, _
        matched, _
        ambiguous, _
        ambiguityPairs
        
        
    currentStage = _
        "Unmatched Pass"
    
      PassUnmatched _
        hostWb, _
        ws, _
        data, _
        lastRow, _
        colAcct, _
        colCodeDesc, _
        colDate, _
        colAmount, _
        colDescription, _
        matched, _
        ambiguous, _
        ambiguityPairs
        
        
    currentStage = _
        "Flush Transactions"
      
      FlushTransactionStatuses _
        hostWb, _
        data, _
        colAcct, _
        colDate, _
        colAmount
        
        
    currentStage = _
        "Finalizing Candidates"
        
    FinalizeCandidateStatuses hostWb
    
    matchedCount = GetMatchedCount()
    ambiguousCount = GetAmbiguousCount()
    unmatchedCount = GetUnmatchedCount()
    
    
    currentStage = _
        "Finalizing reports"
        
    DebugLog "START Finalize Run and Build Reports"
    
    BuildReports hostWb
    
    
    FinalizeRun _
        hostWb, _
        ws, _
        lastRow, _
        matchedCount, _
        ambiguousCount, _
        unmatchedCount
        
       
    FormatWorksheets hostWb
    
    RemoveUnusedOutputSheets hostWb
    
    If HIDE_SUPPORTING_SHEETS Then 'toggle data sheets from global
    
        HideSupportingSheets hostWb
    
    End If
    
        'Return to the worksheet where the run began.
    If sourceWs.Visible = xlSheetVisible Then

        sourceWs.Activate

    Else

        Debug.Print _
            "FinalizeRun: Source worksheet could not be restored " & _
            "because it is not visible."

    End If
    
    DebugLog "END Finalize Run and Build Reports"
        
    
    MsgBox _
    "Transfer analysis complete.", _
    vbInformation
    
    
CleanExit:

    On Error Resume Next

    Application.ScreenUpdating = _
        originalScreenUpdating

    Application.EnableEvents = _
        originalEnableEvents

    Application.DisplayAlerts = _
        originalDisplayAlerts

    Application.Calculation = _
        originalCalculation

    Application.StatusBar = False
    Application.Cursor = xlDefault

    If RunContextIsInitialized Then
        ClearRunContext
    End If

    On Error GoTo 0

    Exit Sub

ErrorHandler:

    errorNumber = Err.Number
    errorDescription = Err.Description
    errorSource = Err.source

    Debug.Print String$(70, "!")
    Debug.Print "ITM ERROR"
    Debug.Print "Entry procedure: RunITM_Controller"
    Debug.Print "Stage:           " & currentStage
    Debug.Print "Error source:    " & errorSource
    Debug.Print "Error number:    " & CStr(errorNumber)
    Debug.Print "Description:     " & errorDescription
    Debug.Print String$(70, "!")

    MsgBox _
        "ITM could not complete the analysis." & _
        vbCrLf & vbCrLf & _
        "Stage: " & currentStage & _
        vbCrLf & _
        "Error " & CStr(errorNumber) & ": " & _
        errorDescription & _
        vbCrLf & vbCrLf & _
        "If the error continues, provide the stage " & _
        "and error number to ITM support.", _
        vbExclamation, _
        "ITM Transfer Analysis"

    Resume CleanExit

End Sub

'---Initial data read and index build---

Private Sub BuildIndexes( _
    data As Variant, _
    lastRow As Long, _
    colAcct As Long, _
    colCodeDesc As Long, _
    colDescription As Long, _
    colDate As Long, _
    colAmount As Long)

    Dim i As Long
    Dim key As String
    
    DebugLog "START BuildIndexes"

    Set ConfIndex = CreateObject("Scripting.Dictionary")
    Set NarrativeIndex = CreateObject("Scripting.Dictionary")
    Set RefAcctIndex = CreateObject("Scripting.Dictionary")
    Set AmtDateIndex = CreateObject("Scripting.Dictionary")
    Set AccountIndex = CreateObject("Scripting.Dictionary")
    Set AmbiguousEdges = CreateObject("Scripting.Dictionary")
    Set TransferToSuffixes = CreateObject("Scripting.Dictionary")
    Set SingleSuffixes = CreateObject("Scripting.Dictionary")

    ReDim TransferFlags(2 To lastRow)

    ReDim ConfirmationNumbers(2 To lastRow)
    ReDim NarrativeIDs(2 To lastRow)
    ReDim ReferencedAccounts(2 To lastRow)
    ReDim EmbeddedAccounts(2 To lastRow)
    
    
    For i = 2 To lastRow

        TransferFlags(i) = _
            IsTransferRecord( _
                CStr(data(i, colCodeDesc)), _
                CStr(data(i, colDescription)))

        If Not TransferFlags(i) Then

            GoTo NextI

        End If
        
        AddToIndex _
            AccountIndex, _
            Trim(CStr(data(i, colAcct))), i
                
        ConfirmationNumbers(i) = _
            GetConfirmationNumber( _
                CStr(data(i, colDescription)))
        
        NarrativeIDs(i) = _
            GetTransferNarrativeID( _
                CStr(data(i, colDescription)))
        
        ReferencedAccounts(i) = _
            GetReferencedAccount( _
                CStr(data(i, colDescription)))
        
        EmbeddedAccounts(i) = _
            GetEmbeddedAccount( _
                CStr(data(i, colDescription)))
                
        Dim fromSuffix As String
        Dim toSuffix As String
        
        If GetTransferSuffixes( _
            CStr(data(i, colDescription)), _
            fromSuffix, _
            toSuffix) Then
        
        TransferToSuffixes(i) = _
            Replace(Trim(fromSuffix), " ", "") & _
            "|" & _
            Replace(Trim(toSuffix), " ", "")
                
        
        End If
        
        '---Suffix processing, used in first and 2nd matching---
        
        Dim singleSuffix As String

        singleSuffix = _
            GetSingleSuffix( _
                CStr(data(i, colDescription)))
        
        If singleSuffix <> "" Then
        
            SingleSuffixes(i) = _
                Trim(singleSuffix)
        
        End If
        
        If singleSuffix <> "" Then
        
            SingleSuffixes(i) = _
                Trim(singleSuffix)
        
        
        End If
        
        If GetTransferSuffixes( _
            CStr(data(i, colDescription)), _
            fromSuffix, _
            toSuffix) Then
        
            TransferToSuffixes(i) = _
                Replace(Trim$(fromSuffix), " ", "") & _
                "|" & _
                Replace(Trim$(toSuffix), " ", "")
        
        
        End If

        '--------------------------
        ' Confirmation Number
        '--------------------------
        key = ConfirmationNumbers(i)

        If key <> "" Then
            AddToIndex ConfIndex, key, i
        End If

        '--------------------------
        ' Narrative ID
        '--------------------------
        key = NarrativeIDs(i)

        If key <> "" Then
            AddToIndex NarrativeIndex, key, i
        End If

        '--------------------------
        ' Referenced Account
        '--------------------------
        key = ReferencedAccounts(i)

        If key <> "" Then
            AddToIndex RefAcctIndex, key, i
        End If

        '--------------------------
        ' Amount + Date
        '--------------------------
        key = _
            Format(data(i, colDate), "yyyymmdd") _
            & "|" _
            & Format(Abs(CDbl(data(i, colAmount))), "0.00")

        AddToIndex AmtDateIndex, key, i
        
       

NextI:
    Next i
    

    DebugLog "END BuildIndexes"
    

'Debug.Print _
'    "UBound(data,1) =", _
'    UBound(data, 1)
'
'Debug.Print _
'    "Indexed Rows:", _
'    UBound(TransferFlags)
'
'Debug.Print _
'    "Last Row:", _
'    lastRow
    
End Sub

Private Sub BuildReports( _
    ByVal hostWb As Workbook)

    BuildMatchedReport hostWb
    BuildAmbiguousPairsReport hostWb
    BuildUnmatchedReport hostWb
    
    BuildPrimaryHubAnalysis hostWb
    BuildTransferAnalysis hostWb
    BuildTransferNetworkAnalysis hostWb
    

End Sub


Private Sub NormalizeSourceWorksheet( _
    ByVal ws As Worksheet)
         
    On Error Resume Next

    If ws.FilterMode Then

        ws.ShowAllData

    End If

    On Error GoTo 0

End Sub


Private Sub HideSupportingSheets( _
    ByVal hostWb As Workbook)

    Dim veryHiddenSheets As Variant
    Dim hiddenSheets As Variant
    Dim sheetName As Variant

    '--Hide data sheets--
    
    veryHiddenSheets = Array( _
        "Transfer_Transactions", _
        "Transfer_Transaction_Status", _
        "Transfer_Ambiguities", _
        "Transfer_Ambiguity_Members", _
        "Transfer_Relationships", _
        "Transfer_Candidates", _
        "Transfer_Contradictions", _
        "Transfer_Unmatched_Reasons", _
        "Transfer_AutoMatch_Candidates", _
        "Investigation_Group_Metrics", _
        "Transfer_Resolved_Clusters", _
        "Investigation_Group_Summary", _
        "Cluster_Resolution_Preview", _
        "Cluster_Analysis")
        

    '--Show on button press--
    
    hiddenSheets = Array( _
        "ITM_Guide", _
        "ITM_Metadata", _
        "Primary_Hub_Analysis", _
        "Investigation_Groups", _
        "Transfer_Network_Analysis")


    '--very hidden loop--
    
    For Each sheetName In veryHiddenSheets

        If WorksheetExists( _
            hostWb, _
            CStr(sheetName)) Then

            hostWb.Worksheets( _
                CStr(sheetName)).Visible = _
                    xlSheetVeryHidden

        Else

            Debug.Print _
                "HideSupportingSheets: " & _
                "Sheet not found and not hidden: " & _
                CStr(sheetName)

        End If

    Next sheetName

        
    '--show-able loop--
    
    For Each sheetName In hiddenSheets

        If WorksheetExists( _
            hostWb, _
            CStr(sheetName)) Then

            hostWb.Worksheets( _
                CStr(sheetName)).Visible = _
                    xlSheetHidden

        Else

            Debug.Print _
                "HideSupportingSheets: " & _
                "Sheet not found and not hidden: " & _
                CStr(sheetName)

        End If

    Next sheetName
    
    
    If WorksheetExists( _
        hostWb, _
        "Primary_Hub_Analysis") Then
    
        hostWb.Worksheets( _
            "Primary_Hub_Analysis").Visible = _
                xlSheetHidden
    
    End If
    
    If GetTotalInvestigationGroups() <> 0 Then
    
        If WorksheetExists( _
            hostWb, _
            "Investigation_Groups") Then
    
            hostWb.Worksheets( _
                "Investigation_Groups").Visible = _
                    xlSheetHidden
    
        End If
    
    End If

End Sub


Private Sub DeleteIfSheetExists( _
    sheetName As String)

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = hostWb.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then Exit Sub

    Application.DisplayAlerts = False

    ws.Delete

    Application.DisplayAlerts = True

End Sub

Private Sub RemoveUnusedOutputSheets( _
    ByVal hostWb As Workbook)

    If GetStatusCount(STATUS_MATCHED) = 0 Then

        DeleteIfSheetExists _
            "Matched_Transfers"

        DeleteIfSheetExists _
            "Transfer_Network_Analysis"

    End If

    If GetStatusCount(STATUS_UNMATCHED) = 0 _
    And GetStatusCount(STATUS_AMBIGUOUS) = 0 Then
    
        DeleteIfSheetExists _
            "Unmatched_Transfers"
    
    End If

    If GetTotalInvestigationGroups() = 0 Then

        DeleteIfSheetExists _
            "Investigation_Groups"

    End If

End Sub


Private Sub FormatWorksheets( _
    ByVal hostWb As Workbook)

    Dim autoFitSheets As Variant
    Dim sheetName As Variant

    'Autofit all sheets
    autoFitSheets = Array( _
        "ITM_Metadata", _
        "Investigation_Groups", _
        "Transfer_Ambiguities", _
        "Transfer_Ambiguity_Members", _
        "Transfer_Candidates", _
        "Transfer_Contradictions", _
        "Primary_Hub_Analysis", _
        "Transfer_Network_Analysis", _
        "Transfer_Transaction_Status", _
        "Transfer_Transactions", _
        "Transfer_Unmatched_Reasons", _
        "Unmatched_Transfers", _
        "Matched_Transfers", _
        "Cluster_Resolution_Preview", _
        "Transfer_Summary", _
        "Transfer_Relationships")
    
    For Each sheetName In autoFitSheets
    
        hostWb.Worksheets( _
            CStr(sheetName)).Cells.EntireColumn.AutoFit
    
    Next sheetName
    
    

    With hostWb.Worksheets("Unmatched_Transfers")

        .Columns("A").numberFormat = "@"
        .Columns("B").numberFormat = "dd-mmm-yy"
        .Columns("D").numberFormat = "$#,##0.00;($#,##0.00)"

        .Columns("G").ColumnWidth = 60    'Investigator Notes
        .Columns("M:O").Hidden = True     'Source/Candidate rows and Investigation Key

    End With

    With hostWb.Worksheets("Matched_Transfers")
    
        .Columns("M:N").Hidden = True
        .Columns("B:C").numberFormat = "@"
        .Columns("D").numberFormat = "dd-mmm-yy"
        .Columns("E").numberFormat = "$#,##0.00"
    
        If .AutoFilterMode Then .AutoFilterMode = False
        .Range("A1").CurrentRegion.AutoFilter
    
    End With

    
    With hostWb.Worksheets("Transfer_Relationships")

        .Columns("B").numberFormat = "dd-mmm-yy"
        .Columns("C:D").numberFormat = "@"

    End With

    With hostWb.Worksheets("Transfer_Candidates")

        .Columns("E:F").numberFormat = "@"
        .Columns("G").numberFormat = "dd-mmm-yy"
        .Columns("H").numberFormat = "$#,##0.00;($#,##0.00)"

    End With

    With hostWb.Worksheets("Transfer_Transaction_Status")

        .Columns("C").numberFormat = "@"
        .Columns("D").numberFormat = "dd-mmm-yy"
        .Columns("E").numberFormat = "$#,##0.00;($#,##0.00)"

    End With
    
    With hostWb.Worksheets("Transfer_Transactions")
    
        .Columns("E").numberFormat = "$#,##0.00;($#,##0.00)"
        
    End With
    
    With hostWb.Worksheets("Transfer_Unmatched_Reasons")
        
        .Columns("D").numberFormat = "dd-mmm-yy"
        .Columns("E").numberFormat = "$#,##0.00;($#,##0.00)"
        
    End With
    
    With hostWb.Worksheets("Transfer_Ambiguity_Members")
        
        .Columns("E").numberFormat = "dd-mmm-yy"
        .Columns("F").numberFormat = "$#,##0.00;($#,##0.00)"
        
    End With
    
    With hostWb.Worksheets("Transfer_Summary")
    
        .Range(.Cells(5, 7), .Cells(7, 8)).HorizontalAlignment = xlLeft
    
    End With
    
    With hostWb.Worksheets("Primary_Hub_Analysis") 'not sure why the gridlines keep dissapearing
    
    ActiveWindow.DisplayGridlines = True
    
    End With
    


End Sub


Private Sub FinalizeRun( _
    ByVal hostWb As Workbook, _
    ByVal sourceWs As Worksheet, _
    ByVal lastRow As Long, _
    ByVal matchedCount As Long, _
    ByVal ambiguousCount As Long, _
    ByVal unmatchedCount As Long)

    Dim ws As Worksheet

    Dim r As Long
    Dim runtimeSeconds As Double
    Dim totalTransfers As Long

    
    Set ws = _
        hostWb.Worksheets("ITM_Metadata")


    totalTransfers = _
        matchedCount + _
        ambiguousCount + _
        unmatchedCount

    runtimeSeconds = _
        Round(Timer - RunStartTime, 2)


    ws.Cells.Clear

    r = 1


    With ws.Range( _
        ws.Cells(r, 1), _
        ws.Cells(r, 2))

        .Merge

        .Value = _
            "TRANSFER ANALYSIS RUN METADATA"

        .Font.Bold = _
            True

        .Font.size = _
            14

        .Font.Color = _
            vbWhite

        .Interior.Color = _
            RGB(31, 78, 121)

        .HorizontalAlignment = _
            xlLeft

    End With

    ws.rows(r).RowHeight = _
        24

    r = r + 2


    WriteMetadataSectionHeader _
        ws, _
        r, _
        "RUN INFORMATION"

    r = r + 1

    WriteMetadataItem _
        ws, r, _
        "Run ID", _
        CurrentRunID

    WriteMetadataItem _
        ws, r, _
        "Version", _
        CurrentVersion

    WriteMetadataItem _
        ws, r, _
        "Execution Timestamp", _
        Now, _
        "m/d/yyyy h:mm:ss AM/PM"


    WriteMetadataItem _
        ws, r, _
        "Workbook Name", _
        hostWb.Name

    WriteMetadataItem _
        ws, r, _
        "Source Worksheet", _
        sourceWs.Name

    WriteMetadataItem _
        ws, r, _
        "Runtime Seconds", _
        runtimeSeconds, _
        "0.00"

    r = r + 1


    WriteMetadataSectionHeader _
        ws, _
        r, _
        "TRANSACTION RESULTS"

    r = r + 1

    WriteMetadataItem _
        ws, r, _
        "Rows Processed", _
        lastRow - 1, _
        "#,##0"

    WriteMetadataItem _
        ws, r, _
        "Transfer Transactions", _
        totalTransfers, _
        "#,##0"

    WriteMetadataItem _
        ws, r, _
        "Matched Count", _
        matchedCount, _
        "#,##0"

    WriteMetadataItem _
        ws, r, _
        "Matched %", _
        SafePct( _
            matchedCount, _
            totalTransfers), _
        "0.00%"

    WriteMetadataItem _
        ws, r, _
        "Ambiguous Count", _
        ambiguousCount, _
        "#,##0"

    WriteMetadataItem _
        ws, r, _
        "Ambiguous %", _
        SafePct( _
            ambiguousCount, _
            totalTransfers), _
        "0.00%"

    WriteMetadataItem _
        ws, r, _
        "Unmatched Count", _
        unmatchedCount, _
        "#,##0"

    WriteMetadataItem _
        ws, r, _
        "Unmatched %", _
        SafePct( _
            unmatchedCount, _
            totalTransfers), _
        "0.00%"

    r = r + 1


    WriteMetadataSectionHeader _
        ws, _
        r, _
        "TRANSACTION VOLUME"

    r = r + 1

    WriteMetadataItem _
        ws, r, _
        "Matched Volume", _
        GetMatchedVolume(), _
        "$#,##0.00;$#,##0.00;-"

    WriteMetadataItem _
        ws, r, _
        "Ambiguous Volume", _
        GetAmbiguousExposure(), _
        "$#,##0.00;$#,##0.00;-"

    WriteMetadataItem _
        ws, r, _
        "Unmatched Volume", _
        GetUnmatchedVolume(), _
        "$#,##0.00;$#,##0.00;-"

    r = r + 1


    WriteMetadataSectionHeader _
        ws, _
        r, _
        "CLUSTER AND CANDIDATE STATISTICS"

    r = r + 1

    WriteMetadataItem _
        ws, r, _
        "Ambiguous Clusters", _
        GetAmbiguousClusterCount(), _
        "#,##0"

    WriteMetadataItem _
        ws, r, _
        "Largest Cluster Size", _
        GetLargestClusterSize(), _
        "#,##0"

    WriteMetadataItem _
        ws, r, _
        "Largest Cluster Exposure", _
        GetLargestClusterExposure(), _
        "$#,##0.00;$#,##0.00;-"

    WriteMetadataItem _
        ws, r, _
        "Candidates Generated", _
        NextCandidateID - 1, _
        "#,##0"

    r = r + 1


    WriteMetadataSectionHeader _
        ws, _
        r, _
        "MATCH METHODS"

    r = r + 1

    WriteMetadataItem _
        ws, r, _
        "Confirmation Number Matches", _
        GetMethodCount( _
            "Confirmation Number"), _
        "#,##0"

    WriteMetadataItem _
        ws, r, _
        "Narrative Pair Matches", _
        GetMethodCount( _
            "Narrative Pair"), _
        "#,##0"

    WriteMetadataItem _
        ws, r, _
        "Referenced Account Matches", _
        GetMethodCount( _
            "Referenced Account"), _
        "#,##0"

    WriteMetadataItem _
        ws, r, _
        "Embedded Account Matches", _
        GetMethodCount( _
            "Embedded Account"), _
        "#,##0"

    WriteMetadataItem _
        ws, r, _
        "Transfer Suffix Matches", _
        GetMethodCount( _
            "Transfer Suffix"), _
        "#,##0"

    WriteMetadataItem _
        ws, r, _
        "Branch Transfer Matches", _
        GetMethodCount( _
            "Branch Transfer"), _
        "#,##0"

    WriteMetadataItem _
        ws, r, _
        "Amount and Date Matches", _
        GetMethodCount( _
            "Amount + Date"), _
        "#,##0"

    r = r + 1



    WriteMetadataSectionHeader _
        ws, _
        r, _
        "AMBIGUITY METHODS"

    r = r + 1

    WriteMetadataItem _
        ws, r, _
        "Confirmation Number Ambiguities", _
        GetAmbiguousMethodCount( _
            "Confirmation Number"), _
        "#,##0"

    WriteMetadataItem _
        ws, r, _
        "Narrative Pair Ambiguities", _
        GetAmbiguousMethodCount( _
            "Narrative Pair"), _
        "#,##0"

    WriteMetadataItem _
        ws, r, _
        "Referenced Account Ambiguities", _
        GetAmbiguousMethodCount( _
            "Referenced Account"), _
        "#,##0"

    WriteMetadataItem _
        ws, r, _
        "Embedded Account Ambiguities", _
        GetAmbiguousMethodCount( _
            "Embedded Account"), _
        "#,##0"

    WriteMetadataItem _
        ws, r, _
        "Transfer Suffix Ambiguities", _
        GetAmbiguousMethodCount( _
            "Transfer Suffix"), _
        "#,##0"

    WriteMetadataItem _
        ws, r, _
        "Branch Transfer Ambiguities", _
        GetAmbiguousMethodCount( _
            "Branch Transfer"), _
        "#,##0"

    WriteMetadataItem _
        ws, r, _
        "Amount and Date Ambiguities", _
        GetAmbiguousMethodCount( _
            "Amount + Date"), _
        "#,##0"


    With ws

        .Columns(1).ColumnWidth = _
            38

        .Columns(2).ColumnWidth = _
            30

        .Columns(1).HorizontalAlignment = _
            xlLeft

        .Columns(2).HorizontalAlignment = _
            xlLeft

        .Columns(1).VerticalAlignment = _
            xlCenter

        .Columns(2).VerticalAlignment = _
            xlCenter

        .rows.RowHeight = _
            18

        .Range( _
            .Cells(1, 1), _
            .Cells(r - 1, 2)).Borders( _
                xlEdgeLeft).LineStyle = _
                    xlContinuous

        .Range( _
            .Cells(1, 1), _
            .Cells(r - 1, 2)).Borders( _
                xlEdgeRight).LineStyle = _
                    xlContinuous

    End With

End Sub

Private Sub WriteMetadataSectionHeader( _
    ws As Worksheet, _
    rowNum As Long, _
    sectionTitle As String)

    With ws.Range( _
        ws.Cells(rowNum, 1), _
        ws.Cells(rowNum, 2))

        .Merge

        .Value = _
            sectionTitle

        .Font.Bold = _
            True

        .Font.Color = _
            RGB(31, 78, 121)

        .Interior.Color = _
            RGB(221, 235, 247)

        .HorizontalAlignment = _
            xlLeft

        .Borders(xlEdgeBottom).LineStyle = _
            xlContinuous

        .Borders(xlEdgeBottom).Color = _
            RGB(180, 198, 231)

    End With

End Sub

Private Sub WriteMetadataItem( _
    ws As Worksheet, _
    ByRef rowNum As Long, _
    metricName As String, _
    metricValue As Variant, _
    Optional numberFormat As String = "")

    ws.Cells(rowNum, 1).Value = _
        metricName

    ws.Cells(rowNum, 2).Value = _
        metricValue

    If numberFormat <> "" Then

        ws.Cells(rowNum, 2).numberFormat = _
            numberFormat

    End If

    rowNum = _
        rowNum + 1

End Sub


'---Determine if items are transfers or not---

Private Function IsTransferRecord( _
    codeDesc As String, _
    fullDescription As String) As Boolean

    Dim txt As String
    txt = LCase(codeDesc & " " & fullDescription)
    txt = Replace(txt, "-", " ")
    
    Do While InStr(txt, "  ") > 0
        txt = Replace(txt, "  ", " ")
    Loop

    If IsEmpty(TransferPatterns) Then
        TransferPatterns = LoadTransferPatterns()
    End If

    Dim i As Long

    For i = LBound(TransferPatterns) _
        To UBound(TransferPatterns)

        If InStr(txt, _
            LCase(TransferPatterns(i))) > 0 Then

            IsTransferRecord = True
            Exit Function

        End If

    Next i

End Function

'=========================================================
' ***  Transfer Definitions  ***
'=========================================================
Private Function LoadTransferPatterns() As Variant

    Dim s As String

    s = "transfer to|transfer from|telephone transfer|telephone principal|"
    s = s & "sweep to|sweep from|loan payment|loan advance|"
    s = s & "ira transfer|ira reverse|ira death|"
    s = s & "investment sweep|intrasweep|internet transfer|"
    s = s & "interest transfer|in person transfer|"
    s = s & "hsa transfer|hsa direct transfer|hsa death distribution|"
    s = s & "heloc check or line of credit cash advance|"
    s = s & "death transfer|death distribution transfer|"
    s = s & "closing entry funds transfer|closing entry credit transfer|"
    s = s & "buseyescrow transfer|big change|"
    s = s & "automatic|additional payment|3rd party sweeps"

    LoadTransferPatterns = Split(s, "|")

End Function

'=====================================================================
' Transfer types which can match based on Transaction Description text
'=====================================================================

Private Function LoadNarrativePrefixes() As Variant

    LoadNarrativePrefixes = Array( _
        "telephone transfer credit", _
        "telephone transfer debit", _
        "transfer to dda", _
        "transfer from dda", _
        "transfer to savings", _
        "transfer from savings", _
        "transfer to sav", _
        "transfer from sav", _
        "transfer to loan", _
        "transfer from loan")

End Function

'---Probably unnecessary---

Private Function FindColumnByAliases( _
    ws As Worksheet, _
    aliases As Variant) As Long

    Dim lastCol As Long

    lastCol = ws.Cells(1, _
        ws.Columns.Count).End(xlToLeft).Column

    Dim c As Long
    Dim a As Variant

    For c = 1 To lastCol

        For Each a In aliases

            If LCase(Trim(ws.Cells(1, c).Value)) = _
               LCase(Trim(a)) Then

                FindColumnByAliases = c
                Exit Function

            End If

        Next a

    Next c

End Function


Private Sub CreateOutputSheets( _
    ByVal hostWb As Workbook)

    Dim sheetNames As Variant
    Dim i As Long
    
    sheetNames = Array( _
        "Matched_Transfers", _
        "Unmatched_Transfers", _
        "Transfer_Summary", _
        "Transfer_Network_Analysis", _
        "Investigation_Groups", _
        "Transfer_Relationships", _
        "Primary_Hub_Analysis", _
        "ITM_Guide", _
        "ITM_Metadata", _
        "Transfer_Contradictions", _
        "Transfer_Candidates", _
        "Transfer_Transaction_Status", _
        "Transfer_Ambiguities", _
        "Transfer_Ambiguity_Members", _
        "Transfer_Unmatched_Reasons", _
        "Transfer_Transactions", _
        "Cluster_Analysis", _
        "Cluster_Resolution_Preview", _
        "Investigation_Group_Summary", _
        "Transfer_Resolved_Clusters", _
        "Investigation_Group_Metrics", _
        "Transfer_AutoMatch_Candidates")
    
    For i = LBound(sheetNames) To UBound(sheetNames)
    
        CreateSheet _
            hostWb, _
            CStr(sheetNames(i))
    
    Next i
        
    
    WriteTransactionWarehouseHeaders _
        hostWb.Worksheets("Transfer_Transactions")
        
    WriteMatchedHeaders _
        hostWb.Worksheets("Matched_Transfers")
    
    WriteUnmatchedHeaders _
        hostWb.Worksheets("Unmatched_Transfers")
    
    WriteTransferRelationshipHeaders _
        hostWb.Worksheets("Transfer_Relationships")
    
    WriteContradictionHeaders _
        hostWb.Worksheets("Transfer_Contradictions")
    
    WriteCandidateHeaders _
        hostWb.Worksheets("Transfer_Candidates")
    
    WriteTransactionStatusHeaders _
        hostWb.Worksheets("Transfer_Transaction_Status")
    
    WriteAmbiguityHeaders _
        hostWb.Worksheets("Transfer_Ambiguities")
    
    WriteAmbiguityMemberHeaders _
        hostWb.Worksheets("Transfer_Ambiguity_Members")
    
    WriteUnmatchedReasonHeaders _
        hostWb.Worksheets("Transfer_Unmatched_Reasons")
    
    WriteMetadataHeaders _
        hostWb.Worksheets("ITM_Metadata")
    
    WriteResolutionPreviewHeaders _
        hostWb.Worksheets("Cluster_Resolution_Preview")
    
    WriteInvestigationGroupSummaryHeaders _
        hostWb.Worksheets("Investigation_Group_Summary")
    
    WriteResolvedClusterHeaders _
        hostWb.Worksheets("Transfer_Resolved_Clusters")
    
    WriteAutoMatchCandidateHeaders _
        hostWb.Worksheets("Transfer_AutoMatch_Candidates")
    
    WriteInvestigationGroupMetricsHeaders _
        hostWb.Worksheets("Investigation_Group_Metrics")
            
    
End Sub

Private Sub CreateSheet( _
    ByVal hostWb As Workbook, _
    ByVal sheetName As String)

    Dim ws As Worksheet
    Dim displayAlertsOriginal As Boolean
    
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo ErrorHandler

    If hostWb Is Nothing Then

        Err.Raise vbObjectError + 1020, _
                  "CreateSheet", _
                  "The host workbook is not available."

    End If

    If Len(Trim$(sheetName)) = 0 Then

        Err.Raise vbObjectError + 1021, _
                  "CreateSheet", _
                  "A blank worksheet name was supplied."

    End If

    displayAlertsOriginal = _
        Application.DisplayAlerts


    If WorksheetExists(hostWb, sheetName) Then

        Set ws = _
            hostWb.Worksheets(sheetName)

    End If

    If Not ws Is Nothing Then

        ws.Visible = xlSheetVisible

        Application.DisplayAlerts = False

        ws.Delete

        Application.DisplayAlerts = _
            displayAlertsOriginal

        Set ws = Nothing

    End If

  
    ' Create the new output sheet in the original workbook
 
    Set ws = hostWb.Worksheets.Add( _
        After:=hostWb.Worksheets( _
            hostWb.Worksheets.Count))

    ws.Name = sheetName

'    Debug.Print "CreateSheet: " & _
'                hostWb.Name & "!" & ws.Name

CleanExit:

    Application.DisplayAlerts = _
        displayAlertsOriginal

    Set ws = Nothing
    Exit Sub

ErrorHandler:


    errorNumber = Err.Number
    errorDescription = Err.Description

    Application.DisplayAlerts = _
        displayAlertsOriginal

    Debug.Print "CreateSheet ERROR"
    Debug.Print "  Host workbook = " & hostWb.Name
    Debug.Print "  Sheet name    = " & sheetName
    Debug.Print "  Error number  = " & CStr(errorNumber)
    Debug.Print "  Description   = " & errorDescription

    Err.Raise errorNumber, _
              "CreateSheet", _
              "Could not create worksheet '" & _
              sheetName & "' in workbook '" & _
              hostWb.Name & "'." & vbCrLf & _
              errorDescription

End Sub

'========
'Temp helper
'========
Public Function WorksheetExists( _
    ByVal hostWb As Workbook, _
    ByVal sheetName As String) As Boolean

    Dim ws As Worksheet

    WorksheetExists = False

    If hostWb Is Nothing Then
        Exit Function
    End If

    If Len(Trim$(sheetName)) = 0 Then
        Exit Function
    End If

    For Each ws In hostWb.Worksheets

        If StrComp( _
            ws.Name, _
            sheetName, _
            vbTextCompare) = 0 Then

            WorksheetExists = True
            Exit Function

        End If

    Next ws

End Function

'======================================================
' Extract Confirmation Number
'
' Example:
'
' Transfer to DDA Transfer from XXX2117 to XXX36 38: Conf #:1333192
'
' Returns:
' 1333192
'======================================================
Private Function GetConfirmationNumber( _
    txt As String) As String

    Dim RE As Object
    Dim matches As Object

    Set RE = CreateObject("VBScript.RegExp")

    RE.pattern = "Conf\s*#:\s*(\d+)"
    RE.IgnoreCase = True
    RE.Global = False

    If RE.Test(txt) Then

        Set matches = RE.Execute(txt)

        GetConfirmationNumber = _
            matches(0).SubMatches(0)

    End If

End Function

'=========================================================
' Extract Narrative
'
' Examples:
' Telephone Transfer Debit TELEPHONE TRANSFER PER: BRADLEY BERGER
'
' Returns:
' TELEPHONE TRANSFER PER: BRADLEY BERGER
'=========================================================

Private Function GetTransferNarrativeID( _
    txt As String) As String

    Dim prefixes As Variant
    Dim p As Variant

    txt = LCase(Trim(txt))

    prefixes = LoadNarrativePrefixes()

    For Each p In prefixes

        If Left(txt, Len(p)) = p Then

            GetTransferNarrativeID = _
                Trim(Mid(txt, Len(p) + 1))

            Exit Function

        End If

    Next p

End Function


'=========================================================
' Extract Referenced Account Number
'
' Examples:
' Investment Sweep From DDA Acct No. 9201968252-D
'
' Returns:
' 9201968252
'=========================================================
Private Function GetReferencedAccount( _
    txt As String) As String

    Dim RE As Object
    Dim matches As Object

    Set RE = CreateObject("VBScript.RegExp")

    RE.pattern = _
        "Acct\s*No\.?\s*(\d+)"

    RE.IgnoreCase = True
    RE.Global = False

    If RE.Test(txt) Then

        Set matches = RE.Execute(txt)

        GetReferencedAccount = _
            matches(0).SubMatches(0)

    End If

End Function

'---Match Writing, do not change---

Private Sub WriteMatched( _
    ByVal hostWb As Workbook, _
    ByRef data As Variant, _
    ByVal rowA As Long, _
    ByVal rowB As Long, _
    ByVal colAcct As Long, _
    ByVal colCodeDesc As Long, _
    ByVal colDate As Long, _
    ByVal colAmount As Long, _
    ByVal colDescription As Long, _
    ByVal matchMethod As String, _
    Optional ByVal clusterID As String = "", _
    Optional ByVal clusterShape As String = "", _
    Optional ByVal investigationGroup As String = "")

    Dim outWs As Worksheet

    Set outWs = _
        hostWb.Worksheets("Matched_Transfers")

    Debug.Assert outWs.Parent Is hostWb

    Dim r As Long
    r = outWs.Cells(outWs.rows.Count, 1).End(xlUp).Row + 1
    
    Dim matchID As String
    
    Dim creditRow As Long
    Dim debitRow As Long
    
    
    If CDbl(data(rowA, colAmount)) > 0 Then
    
        creditRow = rowA
        debitRow = rowB
    
    Else
    
        creditRow = rowB
        debitRow = rowA
    
    End If
    

    matchID = GetNextMatchID()
    
    RegisterTransactionStatus _
        rowA, _
        STATUS_MATCHED, _
        matchID
        
        
    RegisterTransactionStatus _
        rowB, _
        STATUS_MATCHED, _
        matchID

    
    MatchIDs(CStr(rowA)) = matchID
    MatchIDs(CStr(rowB)) = matchID

   
    WriteTransferRelationship _
        hostWb, _
        matchID, _
        data, _
        rowA, _
        rowB, _
        colAcct, _
        colDate, _
        colAmount, _
        matchMethod, _
        clusterID, _
        clusterShape, _
        investigationGroup
        
    UpdateCandidateResolution _
        hostWb, _
        rowA, _
        rowB, _
        STATUS_MATCHED, _
        matchID

    outWs.Cells(r, 1).Value = matchID

    outWs.Cells(r, 2).Value = _
        "'" & CStr(data(creditRow, colAcct))
    
    outWs.Cells(r, 3).Value = _
        "'" & CStr(data(debitRow, colAcct))
    
    outWs.Cells(r, 4).Value = _
        data(creditRow, colDate)
    
    outWs.Cells(r, 5).Value = _
        Abs(data(creditRow, colAmount))
    
    outWs.Cells(r, 6).Value = _
        data(creditRow, colDescription)
    
    outWs.Cells(r, 7).Value = _
        data(debitRow, colDescription)
    
    outWs.Cells(r, 8).Value = _
        matchMethod
    
    outWs.Cells(r, 9).Value = _
        GetConfidenceTier(matchMethod)
    
    outWs.Cells(r, 10).Value = clusterID
    
    outWs.Cells(r, 11).Value = clusterShape
    
    outWs.Cells(r, 12).Value = investigationGroup
    
    outWs.Cells(r, 13).Value = rowA
    
    outWs.Cells(r, 14).Value = rowB


    Call RecordMatchStatistic( _
        matchMethod, _
        CDbl(data(rowA, colAmount)))
        
    outWs.Columns("B:C").numberFormat = "@"

    outWs.Cells(r, 2).Value = _
        CStr(data(creditRow, colAcct))
    
    outWs.Cells(r, 3).Value = _
        CStr(data(debitRow, colAcct))
        
    outWs.rows(1).AutoFilter

End Sub

Private Sub WriteTransferRelationship( _
    ByVal hostWb As Workbook, _
    ByVal matchID As String, _
    ByRef data As Variant, _
    ByVal rowA As Long, _
    ByVal rowB As Long, _
    ByVal colAcct As Long, _
    ByVal colDate As Long, _
    ByVal colAmount As Long, _
    ByVal matchMethod As String, _
    ByVal clusterID As String, _
    ByVal clusterShape As String, _
    ByVal investigationGroup As String)

    Dim ws As Worksheet
    Dim r As Long
    
    Dim debitAcct As String
    Dim creditAcct As String
    Dim TransferAmt As Double
    
    
    Set ws = hostWb.Worksheets("Transfer_Relationships")

    r = ws.Cells(ws.rows.Count, 1).End(xlUp).Row + 1
    
    
    TransferAmt = Abs(CDbl(data(rowA, colAmount)))

    If CDbl(data(rowA, colAmount)) < 0 Then

        debitAcct = _
            CStr(data(rowA, colAcct))

        creditAcct = _
            CStr(data(rowB, colAcct))

    Else

        debitAcct = _
            CStr(data(rowB, colAcct))

        creditAcct = _
            CStr(data(rowA, colAcct))

    End If
    

    
    ws.Cells(r, 1).Value = matchID

    ws.Cells(r, 2).Value = _
        data(rowA, colDate)

    ws.Cells(r, 3).Value = "'" & CStr(debitAcct)
    ws.Cells(r, 4).Value = "'" & CStr(creditAcct)

    ws.Cells(r, 5).Value = TransferAmt

    ws.Cells(r, 6).Value = matchMethod

    ws.Cells(r, 7).Value = _
        GetConfidenceTier(matchMethod)

    ws.Cells(r, 8).Value = rowA
    ws.Cells(r, 9).Value = rowB
    
    ws.Cells(r, 10).Value = Now
    
    
    ws.Cells(r, 11).Value = clusterID

    ws.Cells(r, 12).Value = clusterShape
    
    ws.Cells(r, 13).Value = investigationGroup
    
   

End Sub



Private Sub WriteTransactionStatus( _
    ByVal hostWb As Workbook, _
    ByVal rowNum As Long, _
    ByVal acct As String, _
    ByVal txnDate As Variant, _
    ByVal amount As Variant, _
    ByVal statusValue As String, _
    Optional ByVal relatedID As String = "")

    Dim ws As Worksheet
    Dim r As Long

    Set ws = _
        hostWb.Worksheets("Transfer_Transaction_Status")
        

    r = ws.Cells(ws.rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(r, 1).Value = CurrentRunID

    ws.Cells(r, 2).Value = rowNum

    ws.Cells(r, 3).Value = "'" & acct
    ws.Cells(r, 4).Value = txnDate
    ws.Cells(r, 5).Value = amount

    ws.Cells(r, 6).Value = statusValue

    ws.Cells(r, 7).Value = relatedID

    ws.Cells(r, 8).Value = Now

End Sub

Private Sub RegisterTransactionStatus( _
    rowNum As Long, _
    statusValue As String, _
    Optional relatedID As String = "")

    Dim currentStatus As String

    If TransactionStatus.Exists(CStr(rowNum)) Then
        currentStatus = _
            TransactionStatus(CStr(rowNum))(0)
            
    If StatusPriority(currentStatus) >= _
           StatusPriority(statusValue) Then
            Exit Sub
        End If
    End If

    TransactionStatus(CStr(rowNum)) = _
        Array(statusValue, relatedID)

End Sub

Private Function StatusPriority( _
    ByVal statusValue As String) As Long

    Select Case statusValue

        Case STATUS_MATCHED
            StatusPriority = 5

        Case STATUS_AMBIGUOUS
            StatusPriority = 4

        Case STATUS_UNMATCHED
            StatusPriority = 3

        Case STATUS_CONTRADICTED
            StatusPriority = 2

        Case STATUS_PENDING
            StatusPriority = 1

        Case Else
            StatusPriority = 0

            Debug.Print _
                "StatusPriority: Unknown status [" & _
                statusValue & "]"

    End Select

End Function

Private Sub FlushTransactionStatuses( _
    ByVal hostWb As Workbook, _
    ByRef data As Variant, _
    ByVal colAcct As Long, _
    ByVal colDate As Long, _
    ByVal colAmount As Long)

    Dim rowKey As Variant
    Dim statusData As Variant
    Dim finalStatus As String
    Dim relatedID As String

    For Each rowKey In TransactionStatus.Keys
    
        statusData = TransactionStatus(rowKey)
    
        finalStatus = CStr(statusData(0))
        relatedID = CStr(statusData(1))
    
        WriteTransactionStatus _
            hostWb, _
            CLng(rowKey), _
            CStr(data(CLng(rowKey), colAcct)), _
            data(CLng(rowKey), colDate), _
            data(CLng(rowKey), colAmount), _
            finalStatus, _
            relatedID
    
    Next rowKey

End Sub



Private Sub WriteUnmatchedReasonHeaders(ws As Worksheet)

    ws.Cells(1, 1).Value = "Run ID"
    ws.Cells(1, 2).Value = "Row Number"
    ws.Cells(1, 3).Value = "Account"
    ws.Cells(1, 4).Value = "Date"
    ws.Cells(1, 5).Value = "Amount"
    ws.Cells(1, 6).Value = "Reason"
    ws.Cells(1, 7).Value = "Details"
    ws.Cells(1, 8).Value = "Created Timestamp"

    ws.rows(1).Font.Bold = True
    ws.rows(1).AutoFilter

End Sub

Private Sub WriteUnmatchedReason( _
    ByVal hostWb As Workbook, _
    ByVal rowNum As Long, _
    ByVal acct As String, _
    ByVal txnDate As Variant, _
    ByVal amount As Double, _
    ByVal reasonText As String, _
    Optional ByVal details As String = "")

    Dim ws As Worksheet
    Dim r As Long

    Set ws = hostWb.Worksheets("Transfer_Unmatched_Reasons")

    r = ws.Cells(ws.rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(r, 1).Value = CurrentRunID
    ws.Cells(r, 2).Value = rowNum
    ws.Cells(r, 3).Value = "'" & acct
    ws.Cells(r, 4).Value = txnDate
    ws.Cells(r, 5).Value = amount
    ws.Cells(r, 6).Value = reasonText
    ws.Cells(r, 7).Value = details
    ws.Cells(r, 8).Value = Now

End Sub

Private Function DetermineUnmatchedReason( _
    data As Variant, _
    rowNum As Long, _
    lastRow As Long, _
    colAcct As Long, _
    colDate As Long, _
    colAmount As Long, _
    colDescription As Long) As String

    If HasUnresolvedAccountEvidence( _
        data, _
        rowNum, _
        lastRow, _
        colAcct, _
        colDescription) Then

        DetermineUnmatchedReason = _
            UNMATCH_COUNTERPARTY_MISSING

        Exit Function

    End If

    If HasContradiction(rowNum) Then

        DetermineUnmatchedReason = _
            UNMATCH_CONTRADICTION_EXISTS 'Unmatching items that don't belong together

        Exit Function

    End If

    DetermineUnmatchedReason = _
        UNMATCH_NO_CANDIDATE

End Function

Private Function HasContradiction( _
    rowNum As Long) As Boolean

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = hostWb.Worksheets("Transfer_Contradictions")

    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow

        If ws.Cells(r, 3).Value = rowNum _
        Or ws.Cells(r, 4).Value = rowNum Then

            HasContradiction = True
            Exit Function

        End If

    Next r

End Function

Private Function GetContradictionReasonForRow( _
    rowNum As Long) As String

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = hostWb.Worksheets("Transfer_Contradictions")

    lastRow = _
        ws.Cells(ws.rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow

        If ws.Cells(r, 3).Value = rowNum _
        Or ws.Cells(r, 4).Value = rowNum Then

            GetContradictionReasonForRow = _
                CStr(ws.Cells(r, 9).Value)

            Exit Function

        End If

    Next r

End Function



Private Sub WriteContradiction( _
    ByVal hostWb As Workbook, _
    ByVal sourceRow As Long, _
    ByVal candidateRow As Long, _
    ByVal sourceAcct As String, _
    ByVal candidateAcct As String, _
    ByVal txnDate As Variant, _
    ByVal amount As Double, _
    ByVal contradictionType As String, _
    ByVal matchMethod As String)

    Dim ws As Worksheet
    Dim r As Long

    Set ws = hostWb.Worksheets("Transfer_Contradictions")

    r = ws.Cells(ws.rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(r, 1).Value = CurrentRunID
    ws.Cells(r, 2).Value = GetNextContradictionID()

    ws.Cells(r, 3).Value = sourceRow
    ws.Cells(r, 4).Value = candidateRow

    ws.Cells(r, 5).Value = "'" & sourceAcct
    ws.Cells(r, 6).Value = "'" & candidateAcct

    ws.Cells(r, 7).Value = txnDate
    ws.Cells(r, 8).Value = amount

    ws.Cells(r, 9).Value = contradictionType
    ws.Cells(r, 10).Value = matchMethod

    ws.Cells(r, 11).Value = Now
    
    RegisterTransactionStatus _
        sourceRow, _
        STATUS_CONTRADICTED
    
    RegisterTransactionStatus _
        candidateRow, _
        STATUS_CONTRADICTED
        
    UpdateCandidateResolution _
        hostWb, _
        sourceRow, _
        candidateRow, _
        STATUS_CONTRADICTED

End Sub

Private Sub WriteCandidate( _
    ByVal hostWb As Workbook, _
    ByVal sourceRow As Long, _
    ByVal candidateRow As Long, _
    ByVal sourceAcct As String, _
    ByVal candidateAcct As String, _
    ByVal txnDate As Variant, _
    ByVal amount As Double, _
    ByVal discoveredMethod As String)

    Dim ws As Worksheet
    Dim r As Long
    Dim CandidateID As String
    Dim CandidateKey As String

    Set ws = hostWb.Worksheets("Transfer_Candidates")

    r = ws.Cells(ws.rows.Count, 1).End(xlUp).Row + 1

    CandidateID = GetNextCandidateID()

    CandidateKey = _
        CStr(sourceRow) & "|" & _
        CStr(candidateRow)

    If Not CandidateLookup.Exists(CandidateKey) Then
        CandidateLookup.Add CandidateKey, r
    End If

    ws.Cells(r, 1).Value = CurrentRunID
    ws.Cells(r, 2).Value = CandidateID
    ws.Cells(r, 3).Value = sourceRow
    ws.Cells(r, 4).Value = candidateRow
    ws.Cells(r, 5).Value = "'" & sourceAcct
    ws.Cells(r, 6).Value = "'" & candidateAcct
    ws.Cells(r, 7).Value = txnDate
    ws.Cells(r, 8).Value = amount
    ws.Cells(r, 9).Value = discoveredMethod
    ws.Cells(r, 10).Value = STATUS_PENDING
    ws.Cells(r, 11).Value = ""
    ws.Cells(r, 12).Value = ""
    ws.Cells(r, 13).Value = Now

End Sub

Private Sub UpdateCandidateResolution( _
    ByVal hostWb As Workbook, _
    ByVal sourceRow As Long, _
    ByVal candidateRow As Long, _
    ByVal finalStatus As String, _
    Optional ByVal resolutionID As String = "")

    Dim ws As Worksheet
    Dim CandidateKey As String
    Dim reverseKey As String
    Dim r As Long

    Set ws = hostWb.Worksheets("Transfer_Candidates")

    CandidateKey = _
        CStr(sourceRow) & "|" & _
        CStr(candidateRow)

    reverseKey = _
        CStr(candidateRow) & "|" & _
        CStr(sourceRow)

    If CandidateLookup.Exists(CandidateKey) Then

        r = CandidateLookup(CandidateKey)

    ElseIf CandidateLookup.Exists(reverseKey) Then

        r = CandidateLookup(reverseKey)

    Else

        Exit Sub

    End If

    ws.Cells(r, 10).Value = finalStatus
    ws.Cells(r, 11).Value = resolutionID
    ws.Cells(r, 12).Value = Now

End Sub

Private Sub FinalizeCandidateStatuses( _
    ByVal hostWb As Workbook)

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = hostWb.Worksheets("Transfer_Candidates")

    lastRow = _
        ws.Cells(ws.rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow

        If ws.Cells(r, 10).Value = STATUS_PENDING Then

            ws.Cells(r, 10).Value = "Superseded"

            ws.Cells(r, 12).Value = Now

        End If

    Next r

End Sub

'=================================
' Executive Summary Page - Renamed Transfer_Summary
'=================================

Private Sub BuildTransferAnalysis( _
    ByVal hostWb As Workbook)

    Dim ws As Worksheet
    Dim r As Long

    Set ws = hostWb.Worksheets("Transfer_Summary")

    ws.Cells.Clear

    r = 1

    r = WriteExecutiveSummary(hostWb, ws, r)

    r = r + 4

    r = ws.Cells(ws.rows.Count, 1).End(xlUp).Row + 4

    Call WriteTopUnmatchedTransfers(hostWb, ws, 3, 5)
    Call WriteTopAmbiguousActivity(ws, 10, 5)
    Call WriteTopResolvedActivity(ws, 17, 5)
    Call WriteTopMatchedTransfers(hostWb, ws, 24, 5)
    
    
    'Verafin Search
    ws.Hyperlinks.Add _
    Anchor:=ws.Cells(31, 5), _
    Address:="https://us147.verafin.com/afc#/legacy?uri=%23table%2Fsearch%2Ftransaction", _
    TextToDisplay:="Open Verafin Transaction Search"
    
    
    '---Sheet and table buttons---
        
    Call AddNavigationButton( _
        ws, _
        "btnGuide", _
        "ITM Guide", _
        ws.Cells(31, 8), _
        "OpenGuide")

    Call AddNavigationButton( _
        ws, _
        "btnMetadata", _
        "ITM Metadata", _
        ws.Cells(33, 8), _
        "OpenMetadata")

    Call AddNavigationButton( _
        ws, _
        "btnAccountFinancial", _
        "Financial", _
        ws.Cells(14, 2), _
        "ShowSummaryAccountFinancial")

    Call AddNavigationButton( _
        ws, _
        "btnAccountTransfers", _
        "Transfers", _
        ws.Cells(15, 2), _
        "ShowSummaryAccountTransfers")

    Call AddNavigationButton( _
        ws, _
        "btnRelationships", _
        "Relationships", _
        ws.Cells(16, 2), _
        "ShowSummaryRelationships")

    Call AddNavigationButton( _
        ws, _
        "btnPrimaryHubAnalysis", _
        "Open", _
        ws.Cells(20, 2), _
        "OpenPrimaryHubAnalysis")

    Call AddNavigationButton( _
        ws, _
        "btnInvestigationGroups", _
        "Open", _
        ws.Cells(31, 2), _
        "OpenInvestigationGroups")

'    Call AddNavigationButton( _  No longer used, but may be re-implemented
'        ws, _
'        "btnGuide", _
'        "Open Guide", _
'        ws.Cells(31, 8), _
'        "OpenGuide")

    Call AddNavigationButton( _
        ws, _
        "btnMetadata", _
        "Open Metadata", _
        ws.Cells(31, 8), _
        "OpenMetadata")
        
    Call AddNavigationButton( _
        ws, _
        "btnOpenTransferNetwork", _
        "View All Tables", _
        ws.Cells(15, 3), _
        "OpenTransferNetworkAnalysis")

    
    ws.Columns("A:B").AutoFit
    ws.Columns("E:G").AutoFit

    
        
End Sub

Private Sub ClearExecutiveSummaryPanel( _
    ByVal ws As Worksheet)

    Dim tableIndex As Long
    Dim panelLastRow As Long
    Dim tableRange As Range
    Dim tableName As String
    Dim lastUsedCell As Range

    'Remove only Executive Summary panel tables.
    
    For tableIndex = ws.ListObjects.Count To 1 Step -1

        tableName = _
            ws.ListObjects(tableIndex).Name

        Select Case tableName

            Case "tblSummaryAccountFinancial", _
                 "tblSummaryAccountTransfer", _
                 "tblSummaryRelationshipAnalysis"

                Set tableRange = _
                    ws.ListObjects(tableIndex).Range

                ws.ListObjects(tableIndex).Unlist
                tableRange.Clear

        End Select

    Next tableIndex

    Set lastUsedCell = ws.Cells.Find( _
        What:="*", _
        After:=ws.Cells(1, 1), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False)

    If lastUsedCell Is Nothing Then

        panelLastRow = _
            SUMMARY_PANEL_FIRST_ROW

    Else

        panelLastRow = _
            lastUsedCell.Row

        If panelLastRow < _
                SUMMARY_PANEL_FIRST_ROW Then

            panelLastRow = _
                SUMMARY_PANEL_FIRST_ROW

        End If

    End If

    With ws.Range( _
        ws.Cells( _
            SUMMARY_PANEL_FIRST_ROW, _
            SUMMARY_PANEL_FIRST_COL), _
        ws.Cells( _
            panelLastRow, _
            SUMMARY_PANEL_LAST_COL))

        .UnMerge
        .Clear

    End With

End Sub

Private Sub WriteExecutiveSummaryPanelHeading( _
    ByVal ws As Worksheet, _
    ByVal headingText As String, _
    ByVal descriptionText As String)

    With ws.Range( _
        ws.Cells(SUMMARY_PANEL_FIRST_ROW, 1), _
        ws.Cells(SUMMARY_PANEL_FIRST_ROW, 15))

        .Merge
        .Value = headingText

        .Font.Bold = True
        .Font.size = 12

        .Interior.Color = _
            RGB(221, 235, 247)

        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter

    End With

    With ws.Range( _
        ws.Cells(SUMMARY_PANEL_FIRST_ROW + 1, 1), _
        ws.Cells(SUMMARY_PANEL_FIRST_ROW + 1, 15))

        .Merge
        .Value = descriptionText

        .Font.Italic = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter

    End With

End Sub

Private Sub FormatExecutiveSummaryPanelTable( _
    ByVal ws As Worksheet, _
    ByVal tableName As String)

    Dim outputTable As ListObject
    Dim checkTable As ListObject

    Set outputTable = Nothing

    For Each checkTable In ws.ListObjects

        If StrComp( _
                checkTable.Name, _
                tableName, _
                vbTextCompare) = 0 Then

            Set outputTable = checkTable
            Exit For

        End If

    Next checkTable

    If outputTable Is Nothing Then

        Debug.Print _
            "FormatExecutiveSummaryPanelTable: " & _
            "Table not found [" & tableName & "]"

        Exit Sub

    End If

    outputTable.TableStyle = _
        "TableStyleMedium26"

    outputTable.ShowTableStyleRowStripes = _
        True

    With outputTable.HeaderRowRange

        .WrapText = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 30

    End With

End Sub

'---Financial Table---

Public Sub ShowSummaryAccountFinancial()

    Dim hostWb As Workbook
    Dim ws As Worksheet
    Dim acctStats As Object

    If ActiveSheet Is Nothing Then

        MsgBox _
            "No active worksheet was found.", _
            vbExclamation, _
            "ITM"

        Exit Sub

    End If

    If Not TypeOf ActiveSheet Is Worksheet Then

        MsgBox _
            "The active sheet is not a worksheet.", _
            vbExclamation, _
            "ITM"

        Exit Sub

    End If

    Set ws = ActiveSheet
    Set hostWb = ws.Parent

    If hostWb Is ThisWorkbook Then

        MsgBox _
            "This button must be used from an ITM report " & _
            "in the transaction workbook.", _
            vbExclamation, _
            "ITM"

        Exit Sub

    End If

    If StrComp( _
            ws.Name, _
            "Transfer_Summary", _
            vbTextCompare) <> 0 Then

        MsgBox _
            "Open the Executive Summary before displaying " & _
            "the Account Financial Analysis.", _
            vbExclamation, _
            "ITM"

        Exit Sub

    End If

    Debug.Print _
        "ShowSummaryAccountFinancial host: " & _
        hostWb.Name

    ClearExecutiveSummaryPanel ws

    Set acctStats = _
        BuildAccountStatistics(hostWb)

    If acctStats Is Nothing Then Exit Sub
    If acctStats.Count = 0 Then Exit Sub

    WriteExecutiveSummaryPanelHeading _
        ws, _
        "ACCOUNT FINANCIAL ANALYSIS", _
        "Matched, ambiguous, unmatched, and total transfer volume"

    WriteAccountFinancialAnalysis _
        ws, _
        SUMMARY_PANEL_TABLE_ROW, _
        SUMMARY_PANEL_FIRST_COL, _
        acctStats, _
        "tblSummaryAccountFinancial"

    FormatExecutiveSummaryPanelTable _
        ws, _
        "tblSummaryAccountFinancial"
        
    DebugAccountStatistics _
        acctStats, _
        "2011014111"

End Sub

'---Transfers Table---

Public Sub ShowSummaryAccountTransfers()

    Dim hostWb As Workbook
    Dim ws As Worksheet
    Dim acctStats As Object

    If ActiveSheet Is Nothing Then Exit Sub
    If Not TypeOf ActiveSheet Is Worksheet Then Exit Sub

    Set ws = ActiveSheet
    Set hostWb = ws.Parent

    If hostWb Is ThisWorkbook Then

        MsgBox _
            "This button must be used from an ITM report " & _
            "in the transaction workbook.", _
            vbExclamation, _
            "ITM"

        Exit Sub

    End If

    If StrComp( _
            ws.Name, _
            "Transfer_Summary", _
            vbTextCompare) <> 0 Then

        MsgBox _
            "Open the Executive Summary before displaying " & _
            "the Account Transfer Analysis.", _
            vbExclamation, _
            "ITM"

        Exit Sub

    End If

    Debug.Print _
        "ShowSummaryAccountTransfers host: " & _
        hostWb.Name

    ClearExecutiveSummaryPanel ws

    Set acctStats = _
        BuildAccountStatistics(hostWb)

    If acctStats Is Nothing Then Exit Sub
    If acctStats.Count = 0 Then Exit Sub

    WriteExecutiveSummaryPanelHeading _
        ws, _
        "ACCOUNT TRANSFER ANALYSIS", _
        "Transfer counts and confirmed network relationships"

    WriteAccountTransferAnalysis _
        ws, _
        SUMMARY_PANEL_TABLE_ROW, _
        SUMMARY_PANEL_FIRST_COL, _
        acctStats, _
        "tblSummaryAccountTransfer"

    FormatExecutiveSummaryPanelTable _
        ws, _
        "tblSummaryAccountTransfer"

End Sub

'---Relationships Table---

Public Sub ShowSummaryRelationships()

    Dim hostWb As Workbook
    Dim ws As Worksheet
    Dim relWs As Worksheet

    Dim relationshipLastRow As Long

    Debug.Print String$(70, "=")
    Debug.Print "ShowSummaryRelationships: START"

    If Application.ActiveSheet Is Nothing Then

        MsgBox _
            "No active worksheet was found.", _
            vbExclamation, _
            "ITM"

        Exit Sub

    End If

    If Not TypeOf Application.ActiveSheet Is Worksheet Then

        MsgBox _
            "The active sheet is not a worksheet.", _
            vbExclamation, _
            "ITM"

        Exit Sub

    End If

    Set ws = Application.ActiveSheet
    Set hostWb = ws.Parent

    Debug.Print _
        "Derived worksheet: " & ws.Name

    Debug.Print _
        "Derived host workbook: " & hostWb.Name

    If hostWb Is ThisWorkbook Then

        MsgBox _
            "This button must be used from an ITM report " & _
            "in the transaction workbook.", _
            vbExclamation, _
            "ITM"

        Exit Sub

    End If

    If StrComp( _
            ws.Name, _
            "Transfer_Summary", _
            vbTextCompare) <> 0 Then

        MsgBox _
            "Open the Executive Summary before displaying " & _
            "Relationship Analysis.", _
            vbExclamation, _
            "ITM"

        Exit Sub

    End If

    ClearExecutiveSummaryPanel ws

    WriteExecutiveSummaryPanelHeading _
        ws, _
        "RELATIONSHIP ANALYSIS", _
        "Matched and confirmed transfer relationships"


    If Not WorksheetExists( _
            hostWb, _
            "Transfer_Relationships") Then

        MsgBox _
            "Transfer_Relationships sheet not found.", _
            vbExclamation, _
            "ITM"

        Exit Sub

    End If

    Set relWs = _
        hostWb.Worksheets( _
            "Transfer_Relationships")

    relationshipLastRow = _
        relWs.Cells( _
            relWs.rows.Count, _
            1).End(xlUp).Row

    Debug.Print _
        "Transfer_Relationships last row: " & _
        relationshipLastRow

    If relationshipLastRow <= 1 Then

        With ws.Cells( _
            SUMMARY_PANEL_TABLE_ROW, _
            SUMMARY_PANEL_FIRST_COL)

            .Value = _
                "No matched relationships available."

            .Font.Italic = True

        End With

        Debug.Print _
            "ShowSummaryRelationships: " & _
            "No relationship rows available."

        Exit Sub

    End If


    BuildRelationshipAnalysis _
        hostWb, _
        ws, _
        SUMMARY_PANEL_TABLE_ROW, _
        SUMMARY_PANEL_FIRST_COL, _
        "tblSummaryRelationshipAnalysis"

    FormatExecutiveSummaryPanelTable _
        ws, _
        "tblSummaryRelationshipAnalysis"

    Debug.Print _
        "ShowSummaryRelationships: END"

    Debug.Print String$(70, "=")

End Sub

Private Sub BuildTransferNetworkAnalysis( _
    ByVal hostWb As Workbook)

    Dim ws As Worksheet
    Dim acctStats As Object

    Dim r As Long
    Dim tableLastRow As Long

    Set ws = _
        hostWb.Worksheets( _
            "Transfer_Network_Analysis")


    Set acctStats = _
        BuildAccountStatistics(hostWb)


    ClearTransferNetworkOutput ws


    If acctStats Is Nothing Then

        With ws.Cells(1, 1)

            .Value = _
                "NO TRANSFER TRANSACTIONS FOUND"

            .Font.Bold = True

        End With

        Exit Sub

    End If

    If acctStats.Count = 0 Then

        With ws.Cells(1, 1)

            .Value = _
                "NO TRANSFER TRANSACTIONS FOUND"

            .Font.Bold = True

        End With

        Exit Sub

    End If


'    DebugAccountStatistics _
'        acctStats, _
'        "2011014111"


    r = 1

    With ws.Range( _
        ws.Cells(r, 1), _
        ws.Cells(r, 15))

        .Merge

        .Value = _
            "ACCOUNT FINANCIAL ANALYSIS"

        .Font.Bold = True
        .Font.size = 12

        .Interior.Color = _
            RGB(221, 235, 247)

    End With

    With ws.Range( _
        ws.Cells(r + 1, 1), _
        ws.Cells(r + 1, 15))
    
        .Merge
    
        .Value = _
            "Matched, ambiguous, and unmatched transfer activity"
    
        .Font.Italic = True
        .HorizontalAlignment = xlLeft
    
    End With

    r = r + 3

    WriteAccountFinancialAnalysis _
        ws, _
        r, _
        1, _
        acctStats, _
        "tblAccountFinancialAnalysis"

    tableLastRow = _
        GetNetworkTableLastRow( _
            ws, _
            "tblAccountFinancialAnalysis")

    If tableLastRow = 0 Then

        Debug.Print _
            "BuildTransferNetworkAnalysis: " & _
            "Financial table was not created."

        Exit Sub

    End If

    'Leave two blank rows after the financial table.
    r = tableLastRow + 3


    With ws.Range( _
        ws.Cells(r, 1), _
        ws.Cells(r, 10))

        .Merge

        .Value = _
            "ACCOUNT TRANSFER ANALYSIS"

        .Font.Bold = True
        .Font.size = 12

        .Interior.Color = _
            RGB(221, 235, 247)

    End With
    
    With ws.Range( _
        ws.Cells(r + 1, 1), _
        ws.Cells(r + 1, 10))
    
        .Merge
    
        .Value = _
            "Transfer counts and confirmed network relationships"
    
        .Font.Italic = True
        .HorizontalAlignment = xlLeft
    
    End With


    r = r + 3

    WriteAccountTransferAnalysis _
        ws, _
        r, _
        1, _
        acctStats, _
        "tblAccountTransferAnalysis"

    tableLastRow = _
        GetNetworkTableLastRow( _
            ws, _
            "tblAccountTransferAnalysis")

    If tableLastRow = 0 Then

        Debug.Print _
            "BuildTransferNetworkAnalysis: " & _
            "Transfer table was not created."

        Exit Sub

    End If

    'Leave two blank rows after the transfer table.
    r = tableLastRow + 3


    'Relationship Analysis still represents matched,
    'confirmed relationships only.
    
    If GetStatusCount(STATUS_MATCHED) > 0 Then

        With ws.Range( _
            ws.Cells(r, 1), _
            ws.Cells(r, 8))

            .Merge

            .Value = _
                "RELATIONSHIP ANALYSIS"

            .Font.Bold = True
            .Font.size = 12

            .Interior.Color = _
                RGB(221, 235, 247)

        End With

        ws.Cells(r + 1, 1).Value = _
            "Matched relationships only"

        r = r + 3

        BuildRelationshipAnalysis _
            hostWb, _
            ws, _
            r, _
            1, _
            "tblRelationshipAnalysis"

    Else

        With ws.Range( _
            ws.Cells(r, 1), _
            ws.Cells(r, 8))

            .Merge

            .Value = _
                "RELATIONSHIP ANALYSIS"

            .Font.Bold = True
            .Font.size = 12

            .Interior.Color = _
                RGB(221, 235, 247)

        End With

        With ws.Cells(r + 1, 1)

            .Value = _
                "No matched relationships available."

            .Font.Italic = True

        End With

    End If


    ApplyNetworkTableStyles ws

End Sub

Private Sub ApplyNetworkTableStyles( _
    ws As Worksheet)


    Dim tbl As ListObject

    For Each tbl In ws.ListObjects

        tbl.TableStyle = _
            "TableStyleMedium26"

        tbl.ShowTableStyleRowStripes = _
            True

        tbl.ShowTableStyleFirstColumn = _
            False

        tbl.ShowTableStyleLastColumn = _
            False

    Next tbl

End Sub

Private Sub ClearTransferNetworkOutput( _
    ByVal ws As Worksheet)

    Dim tableIndex As Long

    'Remove all existing Excel tables before clearing
    'the worksheet. Loop backward because the collection
    'shrinks each time a table is removed.  Out of my depths here
    
    For tableIndex = ws.ListObjects.Count To 1 Step -1

        ws.ListObjects(tableIndex).Unlist

    Next tableIndex

    ws.Cells.Clear

End Sub

Private Function GetNetworkTableLastRow( _
    ByVal ws As Worksheet, _
    ByVal tableName As String) As Long

    Dim checkTable As ListObject

    GetNetworkTableLastRow = 0

    For Each checkTable In ws.ListObjects

        If StrComp( _
                checkTable.Name, _
                tableName, _
                vbTextCompare) = 0 Then

            GetNetworkTableLastRow = _
                checkTable.Range.Row + _
                checkTable.Range.rows.Count - 1

            Exit Function

        End If

    Next checkTable

    Debug.Print _
        "GetNetworkTableLastRow: Table not found [" & _
        tableName & "]"

End Function

Private Function CountConfirmationCandidates( _
    data As Variant, _
    targetRow As Long, _
    colAcct As Long, _
    colDate As Long, _
    colAmount As Long, _
    colDescription As Long, _
    matched() As Boolean) As Long

    Dim targetConf As String
    
    Dim v As Variant
    Dim i As Long

    targetConf = ConfirmationNumbers(targetRow)

    If targetConf = "" Then Exit Function

    If Not ConfIndex.Exists(targetConf) Then Exit Function

    For Each v In ConfIndex(targetConf)

        i = CLng(v)

        If i = targetRow Then GoTo NextI
        If matched(i) Then GoTo NextI
        
        If IsSameAccountPair(data, targetRow, i, colAcct) Then GoTo NextI

        If Abs(CDbl(data(targetRow, colAmount))) <> _
           Abs(CDbl(data(i, colAmount))) Then GoTo NextI

        If CDbl(data(targetRow, colAmount)) * _
           CDbl(data(i, colAmount)) >= 0 Then GoTo NextI

        If data(targetRow, colDate) <> _
           data(i, colDate) Then GoTo NextI

        CountConfirmationCandidates = _
            CountConfirmationCandidates + 1
            
NextI:
    Next v

End Function


Private Sub BuildAccountAnalysis( _
    ByVal hostWb As Workbook, _
    ByVal ws As Worksheet, _
    ByVal startRow As Long)

    Dim relWs As Worksheet
    Set relWs = hostWb.Worksheets("Transfer_Relationships")

    Dim acctStats As Object
    Set acctStats = CreateObject("Scripting.Dictionary")

    Dim lastRow As Long
    Dim r As Long

    Dim debitAcct As String
    Dim creditAcct As String
    Dim amt As Double
    
    Dim acctRange As Range
    Dim acctTable As ListObject

    lastRow = relWs.Cells( _
        relWs.rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow

        debitAcct = _
            Trim(CStr(relWs.Cells(r, 3).Value))

        creditAcct = _
            Trim(CStr(relWs.Cells(r, 4).Value))

        amt = _
            CDbl(relWs.Cells(r, 5).Value)

        Call UpdateAccountStats( _
            acctStats, _
            debitAcct, _
            creditAcct, _
            amt, _
            False)

        Call UpdateAccountStats( _
            acctStats, _
            creditAcct, _
            debitAcct, _
            amt, _
            True)

    Next r

    ws.Cells(startRow, 1).Resize(1, 12).Value = Array( _
        "Account", _
        "Transfer Count", _
        "Transfers In", _
        "Transfers Out", _
        "Volume In", _
        "Volume Out", _
        "Gross Volume", _
        "Net Volume", _
        "Unique Sources", _
        "Unique Destinations", _
        "Largest Single Sent", _
        "Largest Single Received")

    ws.rows(startRow).Font.Bold = True

    Dim acct As Variant
    Dim outputRow As Long

    outputRow = startRow + 1
    
    ws.Range( _
        ws.Cells(startRow + 1, 1), _
        ws.Cells(startRow + acctStats.Count, 1) _
        ).numberFormat = "@"

    With ws.Cells(outputRow, 1)

        .numberFormat = "@"
        .Value2 = CStr(acct)
        .HorizontalAlignment = xlLeft
    
    End With
    
    
    For Each acct In acctStats.Keys
    
        ws.Cells(outputRow, 1).Value = acct
        ws.Cells(outputRow, 2).Value = acctStats(acct)(0)   'Transfer Count
        ws.Cells(outputRow, 3).Value = acctStats(acct)(1)   'Transfers In
        ws.Cells(outputRow, 4).Value = acctStats(acct)(2)   'Transfers Out
        ws.Cells(outputRow, 5).Value = acctStats(acct)(3)   'Volume In
        ws.Cells(outputRow, 6).Value = acctStats(acct)(4)   'Volume Out
        ws.Cells(outputRow, 7).Value = _
            acctStats(acct)(3) + acctStats(acct)(4)
        ws.Cells(outputRow, 8).Value = _
            acctStats(acct)(3) - acctStats(acct)(4)
        ws.Cells(outputRow, 9).Value = _
            acctStats(acct)(5).Count
        ws.Cells(outputRow, 10).Value = _
            acctStats(acct)(6).Count
        ws.Cells(outputRow, 11).Value = _
            acctStats(acct)(7)
        ws.Cells(outputRow, 12).Value = _
            acctStats(acct)(8)
    
        outputRow = outputRow + 1
    
    Next acct
    
  
    If acctStats.Count = 0 Then Exit Sub
    
    Set acctRange = ws.Range( _
        ws.Cells(startRow, 1), _
        ws.Cells(outputRow - 1, 12))
    
    Set acctTable = ws.ListObjects.Add( _
        xlSrcRange, _
        acctRange, _
        , xlYes)
    
    acctTable.Name = "tblAccountAnalysis"
    
    acctTable.ListColumns("Account").DataBodyRange.numberFormat = "@"
    
    acctTable.ListColumns("Volume In").DataBodyRange.numberFormat = _
    "$#,##0.00_);($#,##0.00)"
    
    acctTable.ListColumns("Volume Out").DataBodyRange.numberFormat = _
    "$#,##0.00_);($#,##0.00)"
    
    acctTable.ListColumns("Gross Volume").DataBodyRange.numberFormat = _
    "$#,##0.00_);($#,##0.00)"
    
    acctTable.ListColumns("Net Volume").DataBodyRange.numberFormat = _
    "$#,##0.00_);($#,##0.00)"
    
    acctTable.ListColumns("Largest Single Sent").DataBodyRange.numberFormat = _
    "$#,##0.00_);($#,##0.00)"
    
    acctTable.ListColumns("Largest Single Received").DataBodyRange.numberFormat = _
    "$#,##0.00_);($#,##0.00)"
    
    With acctTable.Sort
        .SortFields.Clear
    
        .SortFields.Add _
            key:=acctTable.ListColumns("Gross Volume").Range, _
            SortOn:=xlSortOnValues, _
            Order:=xlDescending
    
        .header = xlYes
        .Apply
    End With

End Sub

Private Sub UpdateAccountStats( _
    acctStats As Object, _
    acct As String, _
    counterparty As String, _
    amount As Double, _
    isIncoming As Boolean)

    Dim stats As Variant

    If acct = "" Then Exit Sub

    If Not acctStats.Exists(acct) Then

        acctStats.Add acct, Array( _
            0, _
            0, _
            0, _
            0#, _
            0#, _
            CreateObject("Scripting.Dictionary"), _
            CreateObject("Scripting.Dictionary"), _
            0#, _
            0#)

    End If

    stats = acctStats(acct)

    stats(0) = stats(0) + 1

    If isIncoming Then

        stats(1) = stats(1) + 1
        stats(3) = stats(3) + amount

        If amount > stats(8) Then
            stats(8) = amount
        End If

        If counterparty <> "" Then
            stats(5)(counterparty) = 1
        End If

    Else

        stats(2) = stats(2) + 1
        stats(4) = stats(4) + amount

        If amount > stats(7) Then
            stats(7) = amount
        End If

        If counterparty <> "" Then
            stats(6)(counterparty) = 1
        End If

    End If

    acctStats(acct) = stats

End Sub

Private Sub WriteAccountFinancialAnalysis( _
    ByVal ws As Worksheet, _
    ByVal startRow As Long, _
    ByVal startCol As Long, _
    ByVal acctStats As Object, _
    ByVal tableName As String)

    Dim acctTable As ListObject
    Dim existingTable As ListObject
    Dim existingRange As Range
    Dim tableRange As Range
    Dim checkTable As ListObject

    Dim acct As Variant
    Dim stats As Variant

    Dim outputRow As Long
    Dim lastCol As Long

    Dim matchedVolumeIn As Double
    Dim ambiguousVolumeIn As Double
    Dim unmatchedVolumeIn As Double
    Dim totalVolumeIn As Double

    Dim matchedVolumeOut As Double
    Dim ambiguousVolumeOut As Double
    Dim unmatchedVolumeOut As Double
    Dim totalVolumeOut As Double

    Dim grossVolume As Double
    Dim netVolume As Double

    If acctStats Is Nothing Then

        Debug.Print _
            "WriteAccountFinancialAnalysis: " & _
            "acctStats is Nothing."

        Exit Sub

    End If

    If acctStats.Count = 0 Then

        Debug.Print _
            "WriteAccountFinancialAnalysis: " & _
            "No account statistics available."

        Exit Sub

    End If

    'Remove the previous version of this table
    Set existingTable = Nothing

    For Each checkTable In ws.ListObjects

        If StrComp( _
                checkTable.Name, _
                tableName, _
                vbTextCompare) = 0 Then

            Set existingTable = checkTable
            Exit For

        End If

    Next checkTable

    If Not existingTable Is Nothing Then

        Set existingRange = _
            existingTable.Range

        existingTable.Unlist
        existingRange.Clear

    End If

    lastCol = startCol + 14


    ws.Cells(startRow, startCol).Resize(1, 15).Value = _
        Array( _
            "Account", _
            "Largest Received", _
            "Received Status", _
            "Largest Sent", _
            "Sent Status", _
            "Matched Volume In", _
            "Ambiguous Volume In", _
            "Unmatched Volume In", _
            "Total Volume In", _
            "Matched Volume Out", _
            "Ambiguous Volume Out", _
            "Unmatched Volume Out", _
            "Total Volume Out", _
            "Gross Volume", _
            "Net Volume")

    outputRow = startRow + 1


    For Each acct In acctStats.Keys

        If Len(Trim$(CStr(acct))) > 0 Then

            stats = acctStats(acct)

            matchedVolumeIn = _
                CDbl(stats( _
                    ACCT_MATCHED_IN_VOLUME))

            ambiguousVolumeIn = _
                CDbl(stats( _
                    ACCT_AMBIGUOUS_IN_VOLUME))

            unmatchedVolumeIn = _
                CDbl(stats( _
                    ACCT_UNMATCHED_IN_VOLUME))

            totalVolumeIn = _
                matchedVolumeIn + _
                ambiguousVolumeIn + _
                unmatchedVolumeIn

            matchedVolumeOut = _
                CDbl(stats( _
                    ACCT_MATCHED_OUT_VOLUME))

            ambiguousVolumeOut = _
                CDbl(stats( _
                    ACCT_AMBIGUOUS_OUT_VOLUME))

            unmatchedVolumeOut = _
                CDbl(stats( _
                    ACCT_UNMATCHED_OUT_VOLUME))

            totalVolumeOut = _
                matchedVolumeOut + _
                ambiguousVolumeOut + _
                unmatchedVolumeOut

            grossVolume = _
                totalVolumeIn + _
                totalVolumeOut

            netVolume = _
                totalVolumeIn - _
                totalVolumeOut

            With ws.Cells(outputRow, startCol)

                .numberFormat = "@"
                .Value2 = CStr(acct)
                .HorizontalAlignment = xlLeft

            End With

        ws.Cells(outputRow, startCol + 1).Value = _
            CDbl(stats( _
                ACCT_LARGEST_RECEIVED))
        
        ws.Cells(outputRow, startCol + 2).Value = _
            CStr(stats( _
                ACCT_LARGEST_RECEIVED_STATUS))
        
        ws.Cells(outputRow, startCol + 3).Value = _
            CDbl(stats( _
                ACCT_LARGEST_SENT))
        
        ws.Cells(outputRow, startCol + 4).Value = _
            CStr(stats( _
                ACCT_LARGEST_SENT_STATUS))
        
        ws.Cells(outputRow, startCol + 5).Value = _
            matchedVolumeIn
        
        ws.Cells(outputRow, startCol + 6).Value = _
            ambiguousVolumeIn
        
        ws.Cells(outputRow, startCol + 7).Value = _
            unmatchedVolumeIn
        
        ws.Cells(outputRow, startCol + 8).Value = _
            totalVolumeIn
        
        ws.Cells(outputRow, startCol + 9).Value = _
            matchedVolumeOut
        
        ws.Cells(outputRow, startCol + 10).Value = _
            ambiguousVolumeOut
        
        ws.Cells(outputRow, startCol + 11).Value = _
            unmatchedVolumeOut
        
        ws.Cells(outputRow, startCol + 12).Value = _
            totalVolumeOut
        
        ws.Cells(outputRow, startCol + 13).Value = _
            grossVolume
        
        ws.Cells(outputRow, startCol + 14).Value = _
            netVolume

            outputRow = outputRow + 1

        End If

    Next acct

    If outputRow = startRow + 1 Then
        Exit Sub
    End If


    Set tableRange = ws.Range( _
        ws.Cells(startRow, startCol), _
        ws.Cells(outputRow - 1, lastCol))

    Set acctTable = ws.ListObjects.Add( _
        xlSrcRange, _
        tableRange, _
        , _
        xlYes)

    acctTable.Name = tableName

    acctTable.ListColumns( _
        "Account").DataBodyRange.numberFormat = "@"


    ApplyFinancialColumnFormat _
        acctTable, _
        "Matched Volume In"

    ApplyFinancialColumnFormat _
        acctTable, _
        "Ambiguous Volume In"

    ApplyFinancialColumnFormat _
        acctTable, _
        "Unmatched Volume In"

    ApplyFinancialColumnFormat _
        acctTable, _
        "Total Volume In"

    ApplyFinancialColumnFormat _
        acctTable, _
        "Matched Volume Out"

    ApplyFinancialColumnFormat _
        acctTable, _
        "Ambiguous Volume Out"

    ApplyFinancialColumnFormat _
        acctTable, _
        "Unmatched Volume Out"

    ApplyFinancialColumnFormat _
        acctTable, _
        "Total Volume Out"

    ApplyFinancialColumnFormat _
        acctTable, _
        "Gross Volume"

    ApplyFinancialColumnFormat _
        acctTable, _
        "Net Volume"

    ApplyFinancialColumnFormat _
        acctTable, _
        "Largest Received"

    ApplyFinancialColumnFormat _
        acctTable, _
        "Largest Sent"


    With acctTable.Sort

        .SortFields.Clear

        .SortFields.Add _
            key:=acctTable.ListColumns( _
                "Gross Volume").Range, _
            SortOn:=xlSortOnValues, _
            Order:=xlDescending, _
            DataOption:=xlSortNormal

        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply

    End With

'    Debug.Print _
'        "WriteAccountFinancialAnalysis: " & _
'        acctTable.ListRows.Count & _
'        " accounts written."

End Sub

Private Sub ApplyFinancialColumnFormat( _
    ByVal outputTable As ListObject, _
    ByVal columnName As String)

    Dim dataRange As Range

    If outputTable Is Nothing Then Exit Sub

    On Error Resume Next

    Set dataRange = _
        outputTable.ListColumns( _
            columnName).DataBodyRange

    On Error GoTo 0

    If dataRange Is Nothing Then

        Debug.Print _
            "ApplyFinancialColumnFormat: " & _
            "Column not found or has no data [" & _
            columnName & "]"

        Exit Sub

    End If

    dataRange.numberFormat = _
        "$#,##0.00_);($#,##0.00)"

End Sub

Private Sub WriteAccountTransferAnalysis( _
    ByVal ws As Worksheet, _
    ByVal startRow As Long, _
    ByVal startCol As Long, _
    ByVal acctStats As Object, _
    ByVal tableName As String)

    Dim transferTable As ListObject
    Dim existingTable As ListObject
    Dim checkTable As ListObject

    Dim existingRange As Range
    Dim tableRange As Range

    Dim acct As Variant
    Dim stats As Variant

    Dim outputRow As Long
    Dim lastCol As Long

    Dim matchedIn As Long
    Dim matchedOut As Long

    Dim ambiguousIn As Long
    Dim ambiguousOut As Long

    Dim unmatchedIn As Long
    Dim unmatchedOut As Long

    Dim totalTransfers As Long


    If acctStats Is Nothing Then


        Exit Sub

    End If

    If acctStats.Count = 0 Then


        Exit Sub

    End If


    Set existingTable = Nothing

    For Each checkTable In ws.ListObjects

        If StrComp( _
                checkTable.Name, _
                tableName, _
                vbTextCompare) = 0 Then

            Set existingTable = checkTable
            Exit For

        End If

    Next checkTable

    If Not existingTable Is Nothing Then

        Set existingRange = _
            existingTable.Range

        existingTable.Unlist
        existingRange.Clear

    End If

    'Ten table columns, including Account.
    lastCol = startCol + 9


    ws.Cells(startRow, startCol).Resize(1, 10).Value = _
        Array( _
            "Account", _
            "Total Transfers", _
            "Matched In", _
            "Matched Out", _
            "Ambiguous In", _
            "Ambiguous Out", _
            "Unmatched In", _
            "Unmatched Out", _
            "Known Sources", _
            "Known Destinations")

    outputRow = startRow + 1


    For Each acct In acctStats.Keys

        If Len(Trim$(CStr(acct))) > 0 Then

            stats = acctStats(acct)

            matchedIn = _
                CLng(stats( _
                    ACCT_MATCHED_IN_COUNT))

            matchedOut = _
                CLng(stats( _
                    ACCT_MATCHED_OUT_COUNT))

            ambiguousIn = _
                CLng(stats( _
                    ACCT_AMBIGUOUS_IN_COUNT))

            ambiguousOut = _
                CLng(stats( _
                    ACCT_AMBIGUOUS_OUT_COUNT))

            unmatchedIn = _
                CLng(stats( _
                    ACCT_UNMATCHED_IN_COUNT))

            unmatchedOut = _
                CLng(stats( _
                    ACCT_UNMATCHED_OUT_COUNT))

            totalTransfers = _
                matchedIn + _
                matchedOut + _
                ambiguousIn + _
                ambiguousOut + _
                unmatchedIn + _
                unmatchedOut

            'Account numbers must remain text.
            With ws.Cells(outputRow, startCol)

                .numberFormat = "@"
                .Value2 = CStr(acct)
                .HorizontalAlignment = xlLeft

            End With

            ws.Cells(outputRow, startCol + 1).Value = _
                totalTransfers

            ws.Cells(outputRow, startCol + 2).Value = _
                matchedIn

            ws.Cells(outputRow, startCol + 3).Value = _
                matchedOut

            ws.Cells(outputRow, startCol + 4).Value = _
                ambiguousIn

            ws.Cells(outputRow, startCol + 5).Value = _
                ambiguousOut

            ws.Cells(outputRow, startCol + 6).Value = _
                unmatchedIn

            ws.Cells(outputRow, startCol + 7).Value = _
                unmatchedOut

            ws.Cells(outputRow, startCol + 8).Value = _
                stats(ACCT_SOURCE_DICT).Count

            ws.Cells(outputRow, startCol + 9).Value = _
                stats(ACCT_DESTINATION_DICT).Count

            outputRow = outputRow + 1

        End If

    Next acct

    If outputRow = startRow + 1 Then


        Exit Sub

    End If


    Set tableRange = ws.Range( _
        ws.Cells(startRow, startCol), _
        ws.Cells(outputRow - 1, lastCol))

    Set transferTable = ws.ListObjects.Add( _
        xlSrcRange, _
        tableRange, _
        , _
        xlYes)

    transferTable.Name = tableName

    transferTable.ListColumns( _
        "Account").DataBodyRange.numberFormat = "@"

    transferTable.ListColumns( _
        "Account").DataBodyRange.HorizontalAlignment = _
        xlLeft


    With transferTable.Sort

        .SortFields.Clear

        .SortFields.Add _
            key:=transferTable.ListColumns( _
                "Total Transfers").Range, _
            SortOn:=xlSortOnValues, _
            Order:=xlDescending, _
            DataOption:=xlSortNormal

        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply

    End With


End Sub

Private Function CreateEmptyAccountStats() As Variant

    '0 Matched In Count
    '1 Matched Out Count
    '2 Ambiguous In Count
    '3 Ambiguous Out Count
    '4 Unmatched In Count
    '5 Unmatched Out Count
    '6 Matched In Volume
    '7 Matched Out Voume
    '8 Ambiguous In Volume
    '9 Ambiguous Out Volume
    '10 Unmatched In Volume
    '11 Unmatched Out Volume
    '12 Known Sources
    '13 Known Destinations
    '14 Largest Received
    '15 Received Status
    '16 Largest Sent
    '17 Sent Status
    
    CreateEmptyAccountStats = Array( _
        0&, _
        0&, _
        0&, _
        0&, _
        0&, _
        0&, _
        0#, _
        0#, _
        0#, _
        0#, _
        0#, _
        0#, _
        CreateObject("Scripting.Dictionary"), _
        CreateObject("Scripting.Dictionary"), _
        0#, _
        vbNullString, _
        0#, _
        vbNullString)

End Function
Private Sub UpdateAccountTransactionStats( _
    ByVal acctStats As Object, _
    ByVal acct As String, _
    ByVal signedAmount As Double, _
    ByVal statusValue As String)

    Dim stats As Variant
    Dim transferAmount As Double
    Dim isIncoming As Boolean

    acct = Trim$(acct)
    statusValue = Trim$(statusValue)

    If Len(acct) = 0 Then Exit Sub

    If Not acctStats.Exists(acct) Then
        acctStats.Add _
            acct, _
            CreateEmptyAccountStats()
    End If

    stats = acctStats(acct)

    transferAmount = Abs(signedAmount)
    isIncoming = (signedAmount > 0)

    Select Case statusValue

        Case STATUS_MATCHED

            If isIncoming Then

                stats(ACCT_MATCHED_IN_COUNT) = _
                    stats(ACCT_MATCHED_IN_COUNT) + 1

                stats(ACCT_MATCHED_IN_VOLUME) = _
                    stats(ACCT_MATCHED_IN_VOLUME) + _
                    transferAmount

            Else

                stats(ACCT_MATCHED_OUT_COUNT) = _
                    stats(ACCT_MATCHED_OUT_COUNT) + 1

                stats(ACCT_MATCHED_OUT_VOLUME) = _
                    stats(ACCT_MATCHED_OUT_VOLUME) + _
                    transferAmount

            End If

        Case STATUS_AMBIGUOUS

            If isIncoming Then

                stats(ACCT_AMBIGUOUS_IN_COUNT) = _
                    stats(ACCT_AMBIGUOUS_IN_COUNT) + 1

                stats(ACCT_AMBIGUOUS_IN_VOLUME) = _
                    stats(ACCT_AMBIGUOUS_IN_VOLUME) + _
                    transferAmount

            Else

                stats(ACCT_AMBIGUOUS_OUT_COUNT) = _
                    stats(ACCT_AMBIGUOUS_OUT_COUNT) + 1

                stats(ACCT_AMBIGUOUS_OUT_VOLUME) = _
                    stats(ACCT_AMBIGUOUS_OUT_VOLUME) + _
                    transferAmount

            End If

        Case STATUS_UNMATCHED

            If isIncoming Then

                stats(ACCT_UNMATCHED_IN_COUNT) = _
                    stats(ACCT_UNMATCHED_IN_COUNT) + 1

                stats(ACCT_UNMATCHED_IN_VOLUME) = _
                    stats(ACCT_UNMATCHED_IN_VOLUME) + _
                    transferAmount

            Else

                stats(ACCT_UNMATCHED_OUT_COUNT) = _
                    stats(ACCT_UNMATCHED_OUT_COUNT) + 1

                stats(ACCT_UNMATCHED_OUT_VOLUME) = _
                    stats(ACCT_UNMATCHED_OUT_VOLUME) + _
                    transferAmount

            End If

        Case Else

            Debug.Print _
                "BuildAccountStatistics: Unknown status [" & _
                statusValue & "] for account [" & _
                acct & "]"

            Exit Sub

    End Select

    If isIncoming Then

        If transferAmount > _
                stats(ACCT_LARGEST_RECEIVED) Then

            stats(ACCT_LARGEST_RECEIVED) = _
                transferAmount

            stats(ACCT_LARGEST_RECEIVED_STATUS) = _
                statusValue

        ElseIf transferAmount = _
                stats(ACCT_LARGEST_RECEIVED) Then

            If GetAccountStatusPriority(statusValue) > _
                    GetAccountStatusPriority( _
                        CStr(stats( _
                            ACCT_LARGEST_RECEIVED_STATUS))) Then

                stats(ACCT_LARGEST_RECEIVED_STATUS) = _
                    statusValue

            End If

        End If

    ElseIf signedAmount < 0 Then

        If transferAmount > _
                stats(ACCT_LARGEST_SENT) Then

            stats(ACCT_LARGEST_SENT) = _
                transferAmount

            stats(ACCT_LARGEST_SENT_STATUS) = _
                statusValue

        ElseIf transferAmount = _
                stats(ACCT_LARGEST_SENT) Then

            If GetAccountStatusPriority(statusValue) > _
                    GetAccountStatusPriority( _
                        CStr(stats( _
                            ACCT_LARGEST_SENT_STATUS))) Then

                stats(ACCT_LARGEST_SENT_STATUS) = _
                    statusValue

            End If

        End If

    End If

    acctStats(acct) = stats

End Sub

Private Sub UpdateKnownCounterpartyStats( _
    ByVal acctStats As Object, _
    ByVal debitAcct As String, _
    ByVal creditAcct As String)

    Dim stats As Variant

    debitAcct = Trim$(debitAcct)
    creditAcct = Trim$(creditAcct)

    If Len(debitAcct) > 0 Then

        If Not acctStats.Exists(debitAcct) Then
            acctStats.Add _
                debitAcct, _
                CreateEmptyAccountStats()
        End If

        stats = acctStats(debitAcct)

        If Len(creditAcct) > 0 Then

            If Not stats( _
                    ACCT_DESTINATION_DICT).Exists( _
                    creditAcct) Then

                stats( _
                    ACCT_DESTINATION_DICT).Add _
                    creditAcct, _
                    True

            End If

        End If

        acctStats(debitAcct) = stats

    End If

    If Len(creditAcct) > 0 Then

        If Not acctStats.Exists(creditAcct) Then
            acctStats.Add _
                creditAcct, _
                CreateEmptyAccountStats()
        End If

        stats = acctStats(creditAcct)

        If Len(debitAcct) > 0 Then

            If Not stats( _
                    ACCT_SOURCE_DICT).Exists( _
                    debitAcct) Then

                stats( _
                    ACCT_SOURCE_DICT).Add _
                    debitAcct, _
                    True

            End If

        End If

        acctStats(creditAcct) = stats

    End If

End Sub

Private Function BuildAccountStatistics( _
    ByVal hostWb As Workbook) As Object

    Dim statusWs As Worksheet
    Dim relationshipWs As Worksheet

    Dim acctStats As Object

    Dim lastRow As Long
    Dim r As Long

    Dim acct As String
    Dim statusValue As String
    Dim amountValue As Double

    Dim debitAcct As String
    Dim creditAcct As String

    Set acctStats = _
        CreateObject("Scripting.Dictionary")

    acctStats.CompareMode = vbTextCompare

    Set statusWs = _
        hostWb.Worksheets( _
            "Transfer_Transaction_Status")

    Set relationshipWs = _
        hostWb.Worksheets( _
            "Transfer_Relationships")

'    Debug.Print String(70, "=")
'    Debug.Print "BuildAccountStatistics: START"
'    Debug.Print "Host workbook: " & hostWb.Name


    '---Transaction counts, volume, and largest---


    lastRow = statusWs.Cells( _
        statusWs.rows.Count, 1).End(xlUp).Row

'    Debug.Print _
'        "Status rows available: " & _
'        Application.Max(lastRow - 1, 0)

    For r = 2 To lastRow

        acct = Trim$( _
            CStr(statusWs.Cells(r, 3).Value2))

        statusValue = Trim$( _
            CStr(statusWs.Cells(r, 6).Value2))

        If Len(acct) = 0 Then

            Debug.Print _
                "Status row " & r & _
                ": Blank account skipped."

        ElseIf Not IsNumeric( _
                statusWs.Cells(r, 5).Value2) Then

            Debug.Print _
                "Status row " & r & _
                ": Non-numeric amount skipped for account [" & _
                acct & "]"

        Else

            amountValue = _
                CDbl(statusWs.Cells(r, 5).Value2)

            UpdateAccountTransactionStats _
                acctStats, _
                acct, _
                amountValue, _
                statusValue

        End If

    Next r
    

    ' ---Confirmed sources/destinations---

    lastRow = relationshipWs.Cells( _
        relationshipWs.rows.Count, 1).End(xlUp).Row

'    Debug.Print _
'        "Relationship rows available: " & _
'        Application.Max(lastRow - 1, 0)

    For r = 2 To lastRow

        debitAcct = Trim$( _
            CStr(relationshipWs.Cells(r, 3).Value2))

        creditAcct = Trim$( _
            CStr(relationshipWs.Cells(r, 4).Value2))

        If Len(debitAcct) = 0 And _
                Len(creditAcct) = 0 Then

'            Debug.Print _
'                "Relationship row " & r & _
'                ": Both accounts blank; row skipped."

        Else

            UpdateKnownCounterpartyStats _
                acctStats, _
                debitAcct, _
                creditAcct

        End If

    Next r

'    Debug.Print _
'        "Accounts built: " & acctStats.Count
'
'    Debug.Print "BuildAccountStatistics: END"
'    Debug.Print String(70, "=")

    Set BuildAccountStatistics = acctStats

End Function

Private Function GetAccountStatusPriority( _
    ByVal statusValue As String) As Long

    Select Case statusValue

        Case STATUS_UNMATCHED
            GetAccountStatusPriority = 3

        Case STATUS_AMBIGUOUS
            GetAccountStatusPriority = 2

        Case STATUS_MATCHED
            GetAccountStatusPriority = 1

        Case Else
            GetAccountStatusPriority = 0

    End Select

End Function

Private Sub BuildRelationshipAnalysis( _
    ByVal hostWb As Workbook, _
    ByVal ws As Worksheet, _
    ByVal startRow As Long, _
    ByVal startCol As Long, _
    ByVal tableName As String)

    Dim relWs As Worksheet
    Dim relStats As Object

    Dim existingTable As ListObject
    Dim checkTable As ListObject
    Dim existingRange As Range

    Dim lastRow As Long
    Dim r As Long

    Dim debitAcct As String
    Dim creditAcct As String
    Dim relKey As String

    Dim amt As Double
    Dim transferDate As Date
    Dim stats As Variant

    Dim k As Variant
    Dim outputRow As Long

    Dim relRange As Range
    Dim relTable As ListObject

    Set relWs = _
        hostWb.Worksheets( _
            "Transfer_Relationships")

    Set relStats = _
        CreateObject("Scripting.Dictionary")

    relStats.CompareMode = vbTextCompare


    Set existingTable = Nothing

    For Each checkTable In ws.ListObjects

        If StrComp( _
                checkTable.Name, _
                tableName, _
                vbTextCompare) = 0 Then

            Set existingTable = checkTable
            Exit For

        End If

    Next checkTable

    If Not existingTable Is Nothing Then

        Set existingRange = _
            existingTable.Range

        existingTable.Unlist
        existingRange.Clear

    End If


    lastRow = relWs.Cells( _
        relWs.rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow

        If IsDate(relWs.Cells(r, 2).Value) Then

            transferDate = _
                CDate(relWs.Cells(r, 2).Value)

        Else

            Debug.Print _
                "BuildRelationshipAnalysis: " & _
                "Invalid date skipped at row " & r

            GoTo NextRelationship

        End If

        debitAcct = _
            Trim$(CStr(relWs.Cells(r, 3).Value2))

        creditAcct = _
            Trim$(CStr(relWs.Cells(r, 4).Value2))

        If Len(debitAcct) = 0 Or _
                Len(creditAcct) = 0 Then

            Debug.Print _
                "BuildRelationshipAnalysis: " & _
                "Blank account skipped at row " & r

            GoTo NextRelationship

        End If

        If Not IsNumeric( _
                relWs.Cells(r, 5).Value2) Then

            Debug.Print _
                "BuildRelationshipAnalysis: " & _
                "Non-numeric amount skipped at row " & r

            GoTo NextRelationship

        End If

        relKey = _
            debitAcct & "|" & creditAcct

        If Not relStats.Exists(relKey) Then

            relStats.Add _
                relKey, _
                Array( _
                    debitAcct, _
                    creditAcct, _
                    0&, _
                    0#, _
                    0#, _
                    DateSerial(9999, 12, 31), _
                    DateSerial(1900, 1, 1))

        End If

        stats = relStats(relKey)

        amt = _
            CDbl(relWs.Cells(r, 5).Value2)

        stats(2) = stats(2) + 1
        stats(3) = stats(3) + amt

        If amt > stats(4) Then
            stats(4) = amt
        End If

        If transferDate < stats(5) Then
            stats(5) = transferDate
        End If

        If transferDate > stats(6) Then
            stats(6) = transferDate
        End If

        relStats(relKey) = stats

NextRelationship:

    Next r

    If relStats.Count = 0 Then

        Debug.Print _
            "BuildRelationshipAnalysis: " & _
            "No relationship statistics available."

        Exit Sub

    End If


    ws.Cells(startRow, startCol).Resize(1, 8).Value = _
        Array( _
            "Debit Account", _
            "Credit Account", _
            "Transfer Count", _
            "Transfer Volume", _
            "Average Transfer", _
            "Largest Transfer", _
            "First Transfer Date", _
            "Last Transfer Date")

    outputRow = startRow + 1


    For Each k In relStats.Keys

        stats = relStats(k)

        With ws.Cells(outputRow, startCol)

            .numberFormat = "@"
            .Value2 = CStr(stats(0))
            .HorizontalAlignment = xlLeft

        End With

        With ws.Cells(outputRow, startCol + 1)

            .numberFormat = "@"
            .Value2 = CStr(stats(1))
            .HorizontalAlignment = xlLeft

        End With

        ws.Cells(outputRow, startCol + 2).Value = _
            stats(2)

        ws.Cells(outputRow, startCol + 3).Value = _
            stats(3)

        If stats(2) > 0 Then

            ws.Cells(outputRow, startCol + 4).Value = _
                stats(3) / stats(2)

        Else

            ws.Cells(outputRow, startCol + 4).Value = _
                0#

        End If

        ws.Cells(outputRow, startCol + 5).Value = _
            stats(4)

        ws.Cells(outputRow, startCol + 6).Value = _
            stats(5)

        ws.Cells(outputRow, startCol + 7).Value = _
            stats(6)

        outputRow = outputRow + 1

    Next k


    Set relRange = ws.Range( _
        ws.Cells(startRow, startCol), _
        ws.Cells(outputRow - 1, startCol + 7))

    Set relTable = ws.ListObjects.Add( _
        xlSrcRange, _
        relRange, _
        , _
        xlYes)

    relTable.Name = tableName

    relTable.ListColumns( _
        "Debit Account").DataBodyRange.numberFormat = "@"

    relTable.ListColumns( _
        "Credit Account").DataBodyRange.numberFormat = "@"

    relTable.ListColumns( _
        "Transfer Volume").DataBodyRange.numberFormat = _
        "$#,##0.00_);($#,##0.00)"

    relTable.ListColumns( _
        "Average Transfer").DataBodyRange.numberFormat = _
        "$#,##0.00_);($#,##0.00)"

    relTable.ListColumns( _
        "Largest Transfer").DataBodyRange.numberFormat = _
        "$#,##0.00_);($#,##0.00)"

    relTable.ListColumns( _
        "First Transfer Date").DataBodyRange.numberFormat = _
        "dd-mmm-yy"

    relTable.ListColumns( _
        "Last Transfer Date").DataBodyRange.numberFormat = _
        "dd-mmm-yy"

    With relTable.Sort

        .SortFields.Clear

        .SortFields.Add _
            key:=relTable.ListColumns( _
                "Transfer Volume").Range, _
            SortOn:=xlSortOnValues, _
            Order:=xlDescending, _
            DataOption:=xlSortNormal

        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply

    End With

End Sub


'---Don't touch, working well---

Private Sub BuildMatchedReport( _
    ByVal hostWb As Workbook)

    Dim relWs As Worksheet
    Dim txnWs As Worksheet
    Dim outWs As Worksheet

    Dim lastRow As Long
    Dim r As Long
    Dim outRow As Long

    Dim sourceWarehouseRow As Long
    Dim candidateWarehouseRow As Long

    Dim sourceRow As Long
    Dim candidateRow As Long

    Dim debitAcct As String
    Dim creditAcct As String

    Set relWs = hostWb.Worksheets("Transfer_Relationships")
    Set txnWs = hostWb.Worksheets("Transfer_Transactions")
    Set outWs = hostWb.Worksheets("Matched_Transfers")

'    Debug.Print "BuildMatchedReport:"
'    Debug.Print "  Relationships = " & _
'                relWs.Parent.Name & "!" & relWs.Name
'    Debug.Print "  Transactions  = " & _
'                txnWs.Parent.Name & "!" & txnWs.Name
'    Debug.Print "  Output        = " & _
'                outWs.Parent.Name & "!" & outWs.Name

    Call WriteMatchedHeaders(outWs)

    lastRow = relWs.Cells( _
        relWs.rows.Count, 1).End(xlUp).Row

    outRow = 2

    For r = 2 To lastRow

        sourceRow = _
            CLng(relWs.Cells(r, 8).Value)

        candidateRow = _
            CLng(relWs.Cells(r, 9).Value)
            
'Debug.Print String(40, "-")
'Debug.Print "Relationship Row: " & r
'
'Debug.Print "Rel Col 8 Value = [" & _
'    relWs.Cells(r, 8).Text & "]"
'
'Debug.Print "Rel Col 8 Type  = " & _
'    TypeName(relWs.Cells(r, 8).Value)
'
'Debug.Print "sourceRow = [" & _
'    sourceRow & "]"
'
'Debug.Print "sourceRow Type = " & _
'    TypeName(sourceRow)
            
        sourceWarehouseRow = _
            FindTransactionRow( _
                hostWb, _
                sourceRow)
        
        candidateWarehouseRow = _
            FindTransactionRow( _
                hostWb, _
                candidateRow)
            
   
        debitAcct = relWs.Cells(r, 3).Value
        
        creditAcct = relWs.Cells(r, 4).Value
        

        
        If Trim(txnWs.Cells(sourceWarehouseRow, 3).Value) = debitAcct Then
        
            outWs.Cells(outRow, 6).Value = _
                txnWs.Cells(candidateWarehouseRow, 7).Value
        
            outWs.Cells(outRow, 7).Value = _
                txnWs.Cells(sourceWarehouseRow, 7).Value
        
        Else
        
            outWs.Cells(outRow, 6).Value = _
                txnWs.Cells(sourceWarehouseRow, 7).Value
        
            outWs.Cells(outRow, 7).Value = _
                txnWs.Cells(candidateWarehouseRow, 7).Value
        
        End If

        outWs.Cells(outRow, 1).Value = _
            relWs.Cells(r, 1).Value

        outWs.Cells(outRow, 2).Value = _
            creditAcct
        
        outWs.Cells(outRow, 3).Value = _
            debitAcct
            
        outWs.Cells(outRow, 4).Value = _
            relWs.Cells(r, 2).Value

        outWs.Cells(outRow, 5).Value = _
            relWs.Cells(r, 5).Value

        outWs.Cells(outRow, 8).Value = _
            relWs.Cells(r, 6).Value

        outWs.Cells(outRow, 9).Value = _
            relWs.Cells(r, 7).Value

        outWs.Cells(outRow, 10).Value = _
            GetFriendlyClusterShape( _
                CStr(relWs.Cells(r, 12).Value))

        outWs.Cells(outRow, 11).Value = _
            relWs.Cells(r, 13).Value
        
        outWs.Cells(outRow, 12).Value = _
            relWs.Cells(r, 11).Value
        
        outWs.Cells(outRow, 13).Value = _
            sourceRow
        
        outWs.Cells(outRow, 14).Value = _
            candidateRow
            

        outRow = outRow + 1

NextR:
    Next r

End Sub

Private Sub BuildUnmatchedReport( _
    ByVal hostWb As Workbook)
    
    Dim txnWs As Worksheet
    Dim statusWs As Worksheet
    Dim outWs As Worksheet

    Dim clusterLookup As Object
    Dim clusterData As Variant

    Dim lastRow As Long
    Dim r As Long
    Dim outRow As Long

    Dim rowNum As Long
    Dim txnRow As Long

    Dim acct As String
    Dim amount As Double
    Dim direction As String

    Dim statusValue As String
    Dim relatedID As String

    Dim clusterID As String
    Dim clusterShape As String
    Dim investigationGroup As String
    Dim investigationKey As String
    Dim clusterOutcome As String

    Dim investigatorNotes As String
    Dim unmatchedReasonDisplay As String
    
    Dim writtenRows As Object

    Set txnWs = hostWb.Worksheets("Transfer_Transactions")
    Set statusWs = hostWb.Worksheets("Transfer_Transaction_Status")
    Set outWs = hostWb.Worksheets("Unmatched_Transfers")
    
'Debug.Print "BuildUnmatchedReport:"
'Debug.Print "  Transactions = " & _
'            txnWs.Parent.Name & "!" & txnWs.Name
'Debug.Print "  Status       = " & _
'            statusWs.Parent.Name & "!" & statusWs.Name
'Debug.Print "  Output       = " & _
'            outWs.Parent.Name & "!" & outWs.Name

    Set clusterLookup = BuildClusterReportingLookup(hostWb)
        
    Set writtenRows = CreateObject("Scripting.Dictionary")
    
        
    outWs.rows("2:" & outWs.rows.Count).ClearContents

    WriteUnmatchedHeaders _
        outWs
        
    With outWs.Columns(1)

    .numberFormat = "@"
    .HorizontalAlignment = xlLeft

    End With

    outRow = 2

    lastRow = _
        statusWs.Cells( _
            statusWs.rows.Count, 2).End(xlUp).Row


For r = 2 To lastRow

    If Not IsNumeric( _
        statusWs.Cells(r, 2).Value) Then

        Debug.Print _
            "INVALID ROW NUMBER", _
            r, _
            statusWs.Cells(r, 2).Value

        GoTo NextR

    End If

    rowNum = _
        CLng(statusWs.Cells(r, 2).Value)

    statusValue = _
        UCase$(Trim$(CStr( _
            statusWs.Cells(r, 6).Value)))


    If Not StatusBelongsInUnmatchedReport( _
        statusValue) Then


        GoTo NextR

    End If

        rowNum = _
            CLng(statusWs.Cells(r, 2).Value)


        If TransactionStatus.Exists(CStr(rowNum)) Then

            If UCase$(Trim$(CStr( _
                TransactionStatus(CStr(rowNum))(0)))) = _
               UCase$(STATUS_MATCHED) Then

                GoTo NextR

            End If

        End If

        txnRow = _
            FindTransactionRow( _
            hostWb, _
            rowNum)
            

        If txnRow = 0 Then


            GoTo NextR

        End If


        acct = _
            Trim$(CStr(txnWs.Cells(txnRow, 3).Value))

        amount = _
            CDbl(txnWs.Cells(txnRow, 5).Value)

        If amount > 0 Then

            direction = _
                "Incoming"

        ElseIf amount < 0 Then

            direction = _
                "Outgoing"

        Else

            direction = _
                "Zero Amount"

        End If


        relatedID = _
            Trim$(CStr(statusWs.Cells(r, 7).Value))

        clusterID = ""
        clusterShape = ""
        investigationGroup = ""
        investigationKey = ""
        clusterOutcome = ""

        If relatedID <> "" Then

            clusterID = _
                relatedID

            If clusterLookup.Exists(clusterID) Then

                clusterData = _
                    clusterLookup(clusterID)

                ' BuildClusterReportingLookup positions:
                '
                ' 0 = Investigation Group
                ' 1 = Recommended Outcome
                ' 2 = Cluster Shape
                ' 9 = Investigation Key

                investigationGroup = _
                    CStr(clusterData(0))

                clusterOutcome = _
                    CStr(clusterData(1))

                clusterShape = _
                    CStr(clusterData(2))

                investigationKey = _
                    CStr(clusterData(9))

            Else

                Debug.Print _
                    "CLUSTER METADATA NOT FOUND", _
                    clusterID, _
                    rowNum

            End If

        End If

        
        investigatorNotes = ""
        
        unmatchedReasonDisplay = ""

        If statusValue = _
           UCase$(STATUS_UNMATCHED) Then
        
            unmatchedReasonDisplay = _
                "No final matched relationship"
        
        ElseIf statusValue = _
               UCase$(STATUS_AMBIGUOUS) Then
        
            unmatchedReasonDisplay = _
                "Transaction remains unresolved after Investigation Group analysis"
        
        End If
        
        Select Case UCase$(Trim$(clusterOutcome))

            Case "PARTIAL_MATCH"
        
                unmatchedReasonDisplay = _
                    "Residual transaction after partial Investigation Group resolution"
        
            Case "UNBALANCED"
        
                unmatchedReasonDisplay = _
                    "Residual transaction from an unbalanced transfer group"
        
            Case "AMBIGUOUS"
        
                unmatchedReasonDisplay = _
                    "Multiple valid counter-transaction relationships remain"
        
            Case "REVIEW"
        
                unmatchedReasonDisplay = _
                    "Transaction requires additional review"
        
        End Select

        
        outWs.Cells(outRow, 1).Value = _
            acct
        
        outWs.Cells(outRow, 1).numberFormat = _
            "@"
        
        outWs.Cells(outRow, 2).Value = _
            txnWs.Cells(txnRow, 4).Value
        
        outWs.Cells(outRow, 2).numberFormat = _
            "m/d/yyyy"
        
        outWs.Cells(outRow, 3).Value = _
            direction
        
        outWs.Cells(outRow, 4).Value = _
            amount
        
        outWs.Cells(outRow, 4).numberFormat = _
            "$#,##0.00;($#,##0.00)"
        
        outWs.Cells(outRow, 5).Value = _
            txnWs.Cells(txnRow, 6).Value
        
        outWs.Cells(outRow, 6).Value = _
            txnWs.Cells(txnRow, 7).Value
        
        
        With outWs.Cells(outRow, 7)
        
            .numberFormat = "@"
            .HorizontalAlignment = xlLeft
            .Value = investigatorNotes
        
        End With
        

        outWs.Cells(outRow, 8).Value = _
            unmatchedReasonDisplay
        
        
        outWs.Cells(outRow, 9).Value = _
            GetFriendlyClusterShape( _
                clusterShape)
        
        outWs.Cells(outRow, 10).Value = _
            GetFriendlyClusterOutcome( _
                clusterOutcome)
        
        outWs.Cells(outRow, 11).Value = _
            investigationGroup
        
        outWs.Cells(outRow, 12).Value = _
            clusterID
        
        
        outWs.Cells(outRow, 13).Value = _
            rowNum
        
        outWs.Cells(outRow, 14).Value = _
            investigationKey
        
        outRow = _
            outRow + 1
NextR:
    Next r


    If outWs.AutoFilterMode Then

        outWs.AutoFilterMode = _
            False

    End If

    outWs.rows(1).AutoFilter

    outWs.Columns("A:M").AutoFit
    
If outRow > 2 Then

    With outWs.Sort

        .SortFields.Clear

        ' Primary sort: Investigation Group

        .SortFields.Add _
            key:=outWs.Range( _
                outWs.Cells(2, 11), _
                outWs.Cells(outRow - 1, 11)), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending

        ' Secondary sort: Cluster ID

        .SortFields.Add _
            key:=outWs.Range( _
                outWs.Cells(2, 12), _
                outWs.Cells(outRow - 1, 12)), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending

        ' Third sort: Account

        .SortFields.Add _
            key:=outWs.Range( _
                outWs.Cells(2, 1), _
                outWs.Cells(outRow - 1, 1)), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending

        ' Fourth sort: Date

        .SortFields.Add _
            key:=outWs.Range( _
                outWs.Cells(2, 2), _
                outWs.Cells(outRow - 1, 2)), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending

        .SetRange outWs.Range( _
            outWs.Cells(1, 1), _
            outWs.Cells(outRow - 1, 14))

        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom

        .Apply

    End With

End If


End Sub

Private Function FindHeaderColumn( _
    ws As Worksheet, _
    headerText As String) As Long

    Dim lastCol As Long
    Dim colNum As Long

    lastCol = _
        ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For colNum = 1 To lastCol

        If UCase$(Trim$( _
            CStr(ws.Cells(1, colNum).Value))) = _
           UCase$(Trim$(headerText)) Then

            FindHeaderColumn = _
                colNum

            Exit Function

        End If

    Next colNum

    FindHeaderColumn = 0

End Function



Private Function StatusBelongsInUnmatchedReport( _
    statusValue As String) As Boolean

    statusValue = _
        UCase$(Trim$(statusValue))


    Select Case statusValue

        Case UCase$(Trim$(STATUS_UNMATCHED))

            StatusBelongsInUnmatchedReport = _
                True

        Case UCase$(Trim$(STATUS_AMBIGUOUS))

            StatusBelongsInUnmatchedReport = _
                True

        Case Else

            StatusBelongsInUnmatchedReport = _
                False

    End Select

End Function

'***Ambig resolution working well, DON'T TOUCH***

Private Sub BuildAmbiguousPairsReport( _
    ByVal hostWb As Workbook)

    Dim outWs As Worksheet
    Dim clusterID As Variant
    Dim members As Collection

    Dim metadata As Object
    Dim clusterDict As Object
    Dim groupDict As Object

    Dim outRow As Long
    
    Dim clusterList As Collection

    Dim investigationGroup As Variant
    Dim matchIDLookup As Object
    
    Dim transferPattern As String
    Dim outcome As String


    If hostWb Is Nothing Then
        Err.Raise vbObjectError + 1060, _
                  "BuildAmbiguousPairsReport", _
                  "The host workbook is Nothing."
    End If


    Set outWs = _
        hostWb.Worksheets("Investigation_Groups")


    Set metadata = _
        BuildClusterMetadataLookup(hostWb)


    Set clusterDict = _
        BuildFullClusterMemberLookup(hostWb)



    Set groupDict = _
        BuildInvestigationGroupLookup(hostWb)

        

    outWs.Cells.Clear
    outRow = 1

   
    For Each investigationGroup In _
        groupDict.Keys
    
        outRow = _
        WriteInvestigationHeader( _
            outWs, _
            outRow, _
            CStr(investigationGroup))
 
        Set clusterList = _
            groupDict(investigationGroup)
        
        Dim groupMembers As Collection

    Set groupMembers = _
        GetInvestigationMembers( _
            clusterList, _
            clusterDict)
            
   
    outRow = _
        WriteInvestigationReport( _
            hostWb, _
            outWs, _
            outRow, _
            CStr(investigationGroup), _
            clusterList, _
            groupMembers, _
            metadata, _
            matchIDLookup)

    
    Next investigationGroup

    With outWs
    
        .Columns("A:G").AutoFit
    
        .Columns("A:B").HorizontalAlignment = xlLeft

    End With

End Sub

Private Function GetInvestigationMembers( _
    clusterList As Collection, _
    clusterDict As Object) As Collection

    Dim members As New Collection

    Dim clusterID As Variant
    Dim rowNum As Variant

    Dim seen As Object

    Set seen = _
        CreateObject("Scripting.Dictionary")

    For Each clusterID In clusterList

        For Each rowNum In clusterDict(clusterID)

            If Not seen.Exists( _
                CStr(rowNum)) Then

                seen.Add _
                    CStr(rowNum), True

                members.Add rowNum

            End If

        Next rowNum

    Next clusterID

    Set GetInvestigationMembers = _
        members

End Function

Private Function GetTotalInvestigationGroups() _
    As Long

    With hostWb.Worksheets("Investigation_Group_Metrics")

        GetTotalInvestigationGroups = _
            .Cells(.rows.Count, 1).End(xlUp).Row - 1

    End With

End Function

Private Function GetResolvedGroupCount() _
    As Long

    Dim ws As Worksheet
    Dim r As Long
    Dim lastRow As Long

    Set ws = hostWb.Worksheets( _
        "Investigation_Group_Metrics")

    lastRow = _
        ws.Cells(ws.rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow

        If UCase( _
            CStr(ws.Cells(r, 2).Value)) = _
            "RESOLVED" Then

            GetResolvedGroupCount = _
                GetResolvedGroupCount + 1

        End If

    Next r

End Function

Private Function GetUnresolvedGroupCount() _
    As Long

    Dim ws As Worksheet
    Dim r As Long
    Dim lastRow As Long

    Set ws = hostWb.Worksheets( _
        "Investigation_Group_Metrics")

    lastRow = _
        ws.Cells(ws.rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow

        If UCase( _
            CStr(ws.Cells(r, 2).Value)) <> _
            "RESOLVED" Then

            GetUnresolvedGroupCount = _
                GetUnresolvedGroupCount + 1

        End If

    Next r

End Function

Private Function GetReviewGroupTransactionCount() As Long

    Dim ws As Worksheet
    Dim transactionRows As Object

    Dim lastRow As Long
    Dim r As Long

    Dim investigationGroup As String
    Dim sourceRow As Variant

    Dim colGroup As Long
    Dim colSourceRow As Long

    On Error Resume Next

    Set ws = _
        hostWb.Worksheets("Unmatched_Transfers")

    On Error GoTo 0


    If ws Is Nothing Then

        GetReviewGroupTransactionCount = 0
        Exit Function

    End If

    Set transactionRows = _
        CreateObject("Scripting.Dictionary")

    colGroup = _
        FindHeaderColumn( _
            ws, _
            "Investigation Group")

    colSourceRow = _
        FindHeaderColumn( _
            ws, _
            "Source Row")

    If colGroup = 0 _
    Or colSourceRow = 0 Then

        Debug.Print
        Debug.Print "REVIEW GROUP COUNT FAILED"
        Debug.Print "Investigation Group Column =", colGroup
        Debug.Print "Source Row Column =", colSourceRow

        Exit Function

    End If

    lastRow = _
        ws.Cells( _
            ws.rows.Count, _
            colSourceRow).End(xlUp).Row

    For r = 2 To lastRow

        investigationGroup = _
            Trim$(CStr( _
                ws.Cells(r, colGroup).Value))

        sourceRow = _
            ws.Cells(r, colSourceRow).Value

        If investigationGroup <> "" _
        And IsNumeric(sourceRow) Then

            transactionRows( _
                CStr(CLng(sourceRow))) = _
                True

        End If

    Next r

    GetReviewGroupTransactionCount = _
        transactionRows.Count

End Function

Private Function GetResolvedExposure() _
    As Double

    Dim ws As Worksheet
    Dim r As Long
    Dim lastRow As Long

    Set ws = hostWb.Worksheets( _
        "Investigation_Group_Metrics")

    lastRow = _
        ws.Cells(ws.rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow

        If UCase( _
            CStr(ws.Cells(r, 2).Value)) = _
            "RESOLVED" Then

            GetResolvedExposure = _
                GetResolvedExposure + _
                CDbl(ws.Cells(r, 6).Value)

        End If

    Next r

End Function

Private Function GetTransactionsRequiringReviewCount() As Long


    GetTransactionsRequiringReviewCount = _
        GetStatusCount(STATUS_UNMATCHED) + _
        GetStatusCount(STATUS_AMBIGUOUS)

End Function

Private Function GetResolvedTransactionCount() As Long


    GetResolvedTransactionCount = _
        GetStatusCount(STATUS_MATCHED)

End Function

Private Function GetResolvedGroupTransactionCount() As Long

    Dim ws As Worksheet
    Dim transactionRows As Object

    Dim lastRow As Long
    Dim r As Long

    Dim investigationGroup As String
    Dim sourceRow As Variant
    Dim candidateRow As Variant

    Dim colGroup As Long
    Dim colSourceRow As Long
    Dim colCandidateRow As Long

    Set ws = _
        hostWb.Worksheets("Transfer_Relationships")

    Set transactionRows = _
        CreateObject("Scripting.Dictionary")


    colGroup = _
        FindHeaderColumn( _
            ws, _
            "Investigation Group")

    colSourceRow = _
        FindHeaderColumn( _
            ws, _
            "Source Row")

    colCandidateRow = _
        FindHeaderColumn( _
            ws, _
            "Candidate Row")

    If colGroup = 0 _
    Or colSourceRow = 0 _
    Or colCandidateRow = 0 Then

        Debug.Print
        Debug.Print "GROUP TRANSACTION COUNT FAILED"
        Debug.Print "Investigation Group Column =", colGroup
        Debug.Print "Source Row Column =", colSourceRow
        Debug.Print "Candidate Row Column =", colCandidateRow

        Exit Function

    End If

    lastRow = _
        ws.Cells( _
            ws.rows.Count, _
            colSourceRow).End(xlUp).Row


    For r = 2 To lastRow

        investigationGroup = _
            Trim$(CStr( _
                ws.Cells(r, colGroup).Value))

        If investigationGroup <> "" Then

            sourceRow = _
                ws.Cells(r, colSourceRow).Value

            candidateRow = _
                ws.Cells(r, colCandidateRow).Value

            If IsNumeric(sourceRow) Then

                transactionRows( _
                    CStr(CLng(sourceRow))) = _
                    True

            End If

            If IsNumeric(candidateRow) Then

                transactionRows( _
                    CStr(CLng(candidateRow))) = _
                    True

            End If

        End If

    Next r

    GetResolvedGroupTransactionCount = _
        transactionRows.Count

End Function

Private Function GetUnresolvedExposure() _
    As Double

    Dim ws As Worksheet
    Dim r As Long
    Dim lastRow As Long

    Set ws = hostWb.Worksheets( _
        "Investigation_Group_Metrics")

    lastRow = _
        ws.Cells(ws.rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow

        If UCase( _
            CStr(ws.Cells(r, 2).Value)) <> _
            "RESOLVED" Then

            GetUnresolvedExposure = _
                GetUnresolvedExposure + _
                CDbl(ws.Cells(r, 6).Value)

        End If

    Next r

End Function

Private Function GetTotalExposure() _
    As Double

    GetTotalExposure = _
        GetResolvedExposure + _
        GetUnresolvedExposure

End Function

Private Function WriteInvestigationReport( _
    ByVal hostWb As Workbook, _
    ByVal ws As Worksheet, _
    ByVal reportRow As Long, _
    ByVal investigationGroup As String, _
    ByVal clusterList As Collection, _
    ByVal members As Collection, _
    ByVal metadata As Object, _
    ByVal matchIDLookup As Object) As Long

    Dim acctDict As Object
    Dim rowNum As Variant
    Dim clusterID As Variant
    Dim acctKey As Variant

    Dim txnRow As Long

    Dim account As String
    Dim exposure As Double
    
    Dim patternText As String
    Dim pattern As String
    
    Dim groupStatus As String
        
    
    Set acctDict = _
        CreateObject("Scripting.Dictionary")
        
    Set matchIDLookup = _
        BuildMatchIDBySourceRow()


    ws.Cells(reportRow, 1).Value = _
        "Investigation Group:"

    ws.Cells(reportRow, 2).Value = _
        investigationGroup

    ws.Cells(reportRow, 1).Font.Bold = True
    ws.Cells(reportRow, 2).Font.Bold = True
    ws.Cells(reportRow, 1).Font.size = 12
    ws.Cells(reportRow, 2).Font.size = 12
    
    With ws.Range( _
        ws.Cells(reportRow, 1), _
        ws.Cells(reportRow, 8))
    
        .Interior.Color = RGB(217, 225, 242)
        .Font.Bold = True
    
    End With

    reportRow = reportRow + 2


    ws.Cells(reportRow, 1).Value = _
        "Ambiguous Groups:"
    
    ws.Cells(reportRow, 1).Font.Bold = True
    
    Dim groupText As String
    
    For Each clusterID In clusterList
    
        If groupText <> "" Then
    
            groupText = _
                groupText & ", "
    
        End If
    
        groupText = _
            groupText & CStr(clusterID)
    
    Next clusterID
    
    ws.Cells(reportRow, 2).Value = _
        groupText
    
    reportRow = reportRow + 2
    
    
    For Each clusterID In clusterList
    
        If metadata.Exists(CStr(clusterID)) Then
    
            pattern = _
                CStr(metadata(clusterID)(1))
    
            If InStr( _
                1, _
                patternText, _
                pattern, _
                vbTextCompare) = 0 Then
    
                If patternText <> "" Then
    
                    patternText = _
                        patternText & ", "
    
                End If
    
                patternText = _
                    patternText & pattern
    
            End If
    
        End If
    
    Next clusterID


    For Each rowNum In members

        txnRow = _
            FindTransactionRow( _
                hostWb, _
                CLng(rowNum))

        If txnRow > 0 Then

            account = _
                CStr(hostWb.Worksheets( _
                    "Transfer_Transactions") _
                    .Cells(txnRow, 3).Value)

            If Not acctDict.Exists(account) Then

                acctDict.Add _
                    account, True

            End If

            exposure = exposure + _
                Abs(CDbl( _
                    hostWb.Worksheets( _
                        "Transfer_Transactions") _
                        .Cells(txnRow, 5).Value))

        End If

    Next rowNum


    groupStatus = "RESOLVED"
    
        For Each clusterID In clusterList
        
            If metadata.Exists(CStr(clusterID)) Then
        
                Select Case UCase(metadata(clusterID)(2))
                
                    Case "REVIEW"
                
                        groupStatus = "REVIEW"
                        Exit For
                
                    Case "UNBALANCED"
                
                        If groupStatus = "RESOLVED" Then
                
                            groupStatus = "UNBALANCED"
                
                        End If
                
                    Case "AMBIGUOUS"
                
                        If groupStatus = "RESOLVED" Then
                
                            groupStatus = "AMBIGUOUS"
                
                        End If
                
                    Case "PARTIAL_MATCH"
                
                        If groupStatus = "RESOLVED" Then
                
                            groupStatus = "PARTIALLY RESOLVED"
                
                        End If
                
                End Select
        
            End If
    
    Next clusterID
    
    ws.Cells(reportRow, 1).Value = "Status:"
    
    ws.Cells(reportRow, 2).Value = groupStatus
    
    ws.Cells(reportRow, 2).Font.Bold = True
    
        Select Case groupStatus
        
            Case "RESOLVED"
        
                ws.Cells(reportRow, 2).Font.Color = _
                    RGB(0, 128, 0)
        
            Case "AMBIGUOUS", "PARTIALLY RESOLVED"
        
                ws.Cells(reportRow, 2).Font.Color = _
                    RGB(192, 128, 0)
                    
        
            Case "REVIEW", "UNBALANCED"
        
                ws.Cells(reportRow, 2).Font.Color = _
                    RGB(192, 0, 0)
        
        End Select
    
    reportRow = reportRow + 2

    ws.Cells(reportRow, 1).Value = _
        "Transaction Count:"

    ws.Cells(reportRow, 2).Value = _
        members.Count

    reportRow = reportRow + 1

    ws.Cells(reportRow, 1).Value = _
        "Accounts Involved:"

    ws.Cells(reportRow, 2).Value = _
        acctDict.Count

    reportRow = reportRow + 1

    ws.Cells(reportRow, 1).Value = _
        "Total Exposure:"

    ws.Cells(reportRow, 2).Value = _
        exposure

    ws.Cells(reportRow, 2).numberFormat = _
        "$#,##0.00;($#,##0.00)"

    reportRow = reportRow + 2


    ws.Cells(reportRow, 1).Value = _
        "Accounts:"

    ws.Cells(reportRow, 1).Font.Bold = True

    reportRow = reportRow + 1

    For Each acctKey In acctDict.Keys
    
        ws.Cells(reportRow, 2).numberFormat = "@"

        ws.Cells(reportRow, 2).Value = _
            acctKey

        reportRow = reportRow + 1
        
 Next acctKey
 
    reportRow = reportRow + 1
    
    ws.Cells(reportRow, 1).Value = _
        "Transfer Pattern(s):"
    
    ws.Cells(reportRow, 1).Font.Bold = True
    
    ws.Cells(reportRow, 2).Value = _
        patternText
    
    reportRow = reportRow + 2


    ws.Cells(reportRow, 1).Value = _
        "Supporting Transactions"

    ws.Cells(reportRow, 1).Font.Bold = True

    ws.Cells(reportRow, 1).Interior.Color = _
        RGB(217, 225, 242)

    reportRow = reportRow + 1

    reportRow = _
        WriteSupportingTransactions( _
            hostWb, _
            ws, _
            reportRow, _
            members, _
            matchIDLookup)


    reportRow = reportRow + 1

    With ws.Range( _
        ws.Cells(reportRow, 1), _
        ws.Cells(reportRow, 8))

        .Interior.Color = _
            RGB(217, 225, 242)

    End With

    WriteInvestigationReport = _
        reportRow + 1

End Function

 

Private Sub ApplyClusterFormatting( _
    ws As Worksheet, _
    firstRow As Long, _
    lastRow As Long, _
    clusterIndex As Long)

    Dim fillColor As Long

    If clusterIndex Mod 2 = 0 Then
        fillColor = RGB(242, 242, 242)
    Else
        fillColor = RGB(255, 255, 255)
    End If

    ws.Range( _
        ws.Cells(firstRow, 1), _
        ws.Cells(lastRow, 6)).Interior.Color = fillColor

End Sub


'---No longer used due to Investigation Groups, afraid to remove it---

Private Function WriteClusterReport( _
    ByVal hostWb As Workbook, _
    ByVal ws As Worksheet, _
    ByVal reportRow As Long, _
    ByVal clusterID As String, _
    ByVal investigationGroup As String, _
    ByVal transferPattern As String, _
    ByVal outcome As String, _
    ByVal members As Collection) As Long

    Dim clusterSize As Long
    Dim clusterExposure As Double

    Dim txnDate As Variant
    Dim transferAmount As Double

    Dim relDict As Object
    Dim warningText As String
    
    Dim acctDict As Object
    Dim rowNum As Variant
    Dim acct As String
    
    Dim firstTxnRow As Long
    

    GetClusterMetrics _
        clusterID, _
        clusterSize, _
        clusterExposure
        
 
    Set relDict = BuildRelationshipsFromEdges( _
        hostWb, _
        members)
    
   
    Set relDict = CollapseRelationships(relDict)
    

    If members.Count > 0 Then
    
        firstTxnRow = FindTransactionRow( _
            hostWb, _
            CLng(members(1)))
    
        If firstTxnRow > 0 Then
    
            txnDate = hostWb.Worksheets("Transfer_Transactions") _
                        .Cells(firstTxnRow, 4).Value
    
            transferAmount = Abs(CDbl( _
                        hostWb.Worksheets("Transfer_Transactions") _
                        .Cells(firstTxnRow, 5).Value))
    
        End If
    
    End If


    ws.Cells(reportRow, 1).Font.Bold = True
    ws.Cells(reportRow, 1).Font.size = 12
    
    reportRow = reportRow + 1
    
    ws.Cells(reportRow, 1).Value = _
        "Ambiguous Group:"
    
    ws.Cells(reportRow, 2).Value = _
        clusterID
    
    reportRow = reportRow + 1
    
    ws.Cells(reportRow, 1).Value = _
        "Transfer Pattern:"
    
    ws.Cells(reportRow, 2).Value = _
        transferPattern
    
    reportRow = reportRow + 1
    
    ws.Cells(reportRow, 1).Value = _
        "Status:"

    ws.Cells(reportRow, 1).Value = "Status:"

    ws.Cells(reportRow, 2).Value = _
        outcome
    
    Select Case UCase(outcome)
    
        Case "UNBALANCED"
    
            ws.Cells(reportRow, 2).Font.Color = _
                RGB(192, 0, 0)
    
        Case "REVIEW"
    
            ws.Cells(reportRow, 2).Font.Color = _
                RGB(192, 0, 0)
    
        Case Else
    
            ws.Cells(reportRow, 2).Font.Color = _
                RGB(0, 128, 0)
    
    End Select
    
    ws.Cells(reportRow, 2).Font.Bold = True

    reportRow = reportRow + 2


    ws.Cells(reportRow, 1).Value = "Date:"

    With ws.Cells(reportRow, 2)

        .Value = txnDate
        .numberFormat = "m/d/yyyy"
        .HorizontalAlignment = xlLeft

    End With

    reportRow = reportRow + 1

    ws.Cells(reportRow, 1).Value = "Transfer Amount:"

    With ws.Cells(reportRow, 2)

        .Value = transferAmount
        .numberFormat = "$#,##0.00;($#,##0.00)"
        .HorizontalAlignment = xlLeft

    End With

    reportRow = reportRow + 1

    ws.Cells(reportRow, 1).Value = _
        "Transaction Count:"
        
    With ws.Cells(reportRow, 2)
    
        .Value = clusterSize
        .HorizontalAlignment = xlLeft
    
    End With
    

    reportRow = reportRow + 1

    ws.Cells(reportRow, 1).Value = _
        "Total Exposure:"

    With ws.Cells(reportRow, 2)

        .Value = clusterExposure
        .numberFormat = "$#,##0.00;($#,##0.00)"
        .HorizontalAlignment = xlLeft

    End With

    reportRow = reportRow + 1

    
    Set acctDict = _
        CreateObject("Scripting.Dictionary")
    
    For Each rowNum In members
    
        acct = _
            hostWb.Worksheets("Transfer_Transactions") _
                .Cells( _
                    FindTransactionRow( _
                        hostWb, _
                        CLng(rowNum)), _
                        3).Value
    
        If Not acctDict.Exists(acct) Then
    
            acctDict.Add acct, True
    
        End If
    
    Next rowNum


    Dim matchIDLookup As Object
    
    Set matchIDLookup = _
        BuildMatchIDBySourceRow()
    
    ws.Cells(reportRow, 1).Value = _
        "Supporting Transactions"
    
    ws.Cells(reportRow, 1).Font.Bold = _
        True
    
    ws.Cells(reportRow, 1).Interior.Color = _
        RGB(217, 225, 242)
    
    reportRow = _
        reportRow + 1
    
    reportRow = _
        WriteSupportingTransactions( _
            hostWb, _
            ws, _
            reportRow, _
            members, _
            matchIDLookup)



    If warningText <> "" Then

        reportRow = reportRow + 1

        reportRow = WriteWarnings( _
                        ws, _
                        reportRow, _
                        warningText)

    End If


    reportRow = reportRow + 1

    With ws.Range( _
        ws.Cells(reportRow, 1), _
        ws.Cells(reportRow, 6))

        .Interior.Color = RGB(217, 225, 242)

    End With

    WriteClusterReport = reportRow + 1

End Function

Private Function BuildInvestigationGroupLookup( _
    ByVal hostWb As Workbook) As Object

    Dim ws As Worksheet

    Dim groups As Object
    Dim clusters As Collection

    Dim lastRow As Long
    Dim r As Long

    Dim clusterID As String
    Dim investigationGroup As String

    Set ws = _
        hostWb.Worksheets("Cluster_Analysis")

    Set groups = _
        CreateObject("Scripting.Dictionary")

    lastRow = _
        ws.Cells(ws.rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow

        clusterID = _
            Trim$(CStr(ws.Cells(r, 1).Value2))

        investigationGroup = _
            Trim$(CStr(ws.Cells(r, 19).Value2))


        '---Do not create dictionary entries for blank IDs---

        If Len(clusterID) > 0 And _
           Len(investigationGroup) > 0 Then

            If Not groups.Exists(investigationGroup) Then

                Set clusters = New Collection

                groups.Add _
                    investigationGroup, _
                    clusters

            Else

                Set clusters = _
                    groups(investigationGroup)

            End If

            clusters.Add clusterID

        Else

            Debug.Print _
                "BuildInvestigationGroupLookup: " & _
                "Skipped Cluster_Analysis row " & CStr(r) & _
                "; Cluster ID = '" & clusterID & "'" & _
                "; Investigation Group = '" & _
                investigationGroup & "'"

        End If

    Next r


    Set BuildInvestigationGroupLookup = _
        groups

End Function

Private Sub BuildInvestigationGroupMetrics( _
    ByVal hostWb As Workbook)

    Dim analysisWs As Worksheet
    Dim outWs As Worksheet

    Dim groupLookup As Object
    Dim clusterData As Variant

    Dim lastRow As Long
    Dim r As Long
    Dim outRow As Long

    Dim investigationGroup As String

    Set analysisWs = hostWb.Worksheets("Cluster_Analysis")
    Set outWs = hostWb.Worksheets("Investigation_Group_Metrics")
    

    outWs.rows("2:" & outWs.rows.Count).ClearContents

    Set groupLookup = _
        CreateObject("Scripting.Dictionary")

    lastRow = _
        analysisWs.Cells( _
            analysisWs.rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow
    
        investigationGroup = _
            CStr(analysisWs.Cells(r, 19).Value)
    
        If Not groupLookup.Exists( _
            investigationGroup) Then
    
            groupLookup.Add _
                investigationGroup, _
                Array( _
                    "RESOLVED", _
                    0, _
                    0, _
                    0, _
                    0#)
    
        End If
    
        clusterData = _
            groupLookup(investigationGroup)
    
        clusterData(1) = _
            clusterData(1) + 1
    
        clusterData(2) = _
            clusterData(2) + _
            CLng(analysisWs.Cells(r, 2).Value)
    
        If CLng(analysisWs.Cells(r, 15).Value) > _
           clusterData(3) Then
    
            clusterData(3) = _
                CLng(analysisWs.Cells(r, 15).Value)
    
        End If
    
        clusterData(4) = _
            clusterData(4) + _
            CDbl(analysisWs.Cells(r, 17).Value)
    
        Select Case _
            UCase(CStr(analysisWs.Cells(r, 14).Value))
        
            Case "REVIEW"
        
                clusterData(0) = "REVIEW"
        
            Case "UNBALANCED"
        
                If clusterData(0) = "RESOLVED" Then
        
                    clusterData(0) = "UNBALANCED"
        
                End If
        
            Case "AMBIGUOUS"
        
                If clusterData(0) = "RESOLVED" Then
        
                    clusterData(0) = "AMBIGUOUS"
        
                End If
        
            Case "PARTIAL_MATCH"
        
                If clusterData(0) = "RESOLVED" Then
        
                    clusterData(0) = "PARTIAL"
        
                End If
        
        End Select
    
        groupLookup(investigationGroup) = _
            clusterData
    
   
    Next r
    
    outRow = 2

    Dim k As Variant

    For Each k In groupLookup.Keys

        clusterData = _
            groupLookup(k)

        outWs.Cells(outRow, 1).Value = k
        outWs.Cells(outRow, 2).Value = clusterData(0)
        outWs.Cells(outRow, 3).Value = clusterData(1)
        outWs.Cells(outRow, 4).Value = clusterData(2)
        outWs.Cells(outRow, 5).Value = clusterData(3)
        outWs.Cells(outRow, 6).Value = clusterData(4)

        outRow = outRow + 1

    Next k

    outWs.Columns.AutoFit

End Sub

Private Function WriteInvestigationHeader( _
    ws As Worksheet, _
    reportRow As Long, _
    investigationGroup As String) As Long

    ws.Cells(reportRow, 1).Font.Bold = True
    ws.Cells(reportRow, 2).Font.Bold = True

    WriteInvestigationHeader = _
        reportRow

End Function
Private Sub WriteInvestigationGroupMetricsHeaders( _
    ws As Worksheet)

    ws.Cells(1, 1).Value = "Investigation Group"
    ws.Cells(1, 2).Value = "Status"
    ws.Cells(1, 3).Value = "Cluster Count"
    ws.Cells(1, 4).Value = "Transaction Count"
    ws.Cells(1, 5).Value = "Account Count"
    ws.Cells(1, 6).Value = "Exposure"

    ws.rows(1).Font.Bold = True

End Sub

Private Sub WriteResolutionPreviewHeaders( _
    ws As Worksheet)

    ws.Cells(1, 1).Value = "Cluster ID"
    ws.Cells(1, 2).Value = "Investigation Key"
    ws.Cells(1, 3).Value = "Investigation Group"
    ws.Cells(1, 4).Value = "Cluster Shape"
    ws.Cells(1, 5).Value = "Outcome"
    
    ws.Cells(1, 6).Value = "Debit Row"
    ws.Cells(1, 7).Value = "Credit Row"
    
    ws.Cells(1, 8).Value = "Debit Account"
    ws.Cells(1, 9).Value = "Credit Account"
    
    ws.Cells(1, 10).Value = "Amount"
    
    ws.rows(1).Font.Bold = True


End Sub




Private Sub WriteResolutionPreview( _
    ByVal hostWb As Workbook, _
    ByVal clusterID As String, _
    ByVal investigationKey As String, _
    ByVal investigationGroup As String, _
    ByVal clusterShape As String, _
    ByVal outcome As String, _
    ByVal debitRow As Long, _
    ByVal creditRow As Long, _
    ByRef data As Variant, _
    ByVal colAcct As Long, _
    ByVal colAmount As Long)

    Dim ws As Worksheet
    Dim r As Long

    Set ws = _
        hostWb.Worksheets("Cluster_Resolution_Preview")

    r = ws.Cells( _
        ws.rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(r, 1).Value = clusterID
    ws.Cells(r, 2).Value = investigationKey
    ws.Cells(r, 3).Value = investigationGroup
    ws.Cells(r, 4).Value = clusterShape
    ws.Cells(r, 5).Value = outcome

    ws.Cells(r, 6).Value = debitRow
    ws.Cells(r, 7).Value = creditRow

    ws.Cells(r, 8).Value = _
        CStr(data(debitRow, colAcct))

    ws.Cells(r, 9).Value = _
        CStr(data(creditRow, colAcct))

    ws.Cells(r, 10).Value = _
        Abs(CDbl(data(debitRow, colAmount)))

End Sub

Private Sub PreviewResolvedCluster( _
    ByVal hostWb As Workbook, _
    ByVal clusterID As String, _
    ByVal investigationKey As String, _
    ByVal investigationGroup As String, _
    ByVal clusterShape As String, _
    ByVal recommendedOutcome As String, _
    ByVal members As Collection, _
    ByVal ambiguityPairs As Collection, _
    ByRef data As Variant, _
    ByVal colAcct As Long, _
    ByVal colAmount As Long)

    '========================================================
    ' Pass the final resolved-cluster information to the
    ' preview writer.
    '
    ' Investigation Key and Investigation Group are separate
    ' metadata fields and must both be carried forward.
    '========================================================

    WriteClusterPairingsToPreview _
        hostWb, _
        clusterID, _
        investigationKey, _
        investigationGroup, _
        clusterShape, _
        recommendedOutcome, _
        members, _
        ambiguityPairs, _
        data, _
        colAcct, _
        colAmount

End Sub

Private Sub WriteRelationshipMap( _
    ws As Worksheet, _
    startRow As Long, _
    relDict As Object)

    Dim relationship As Variant

    For Each relationship In relDict.Keys

        With ws.Cells(startRow, 1)

            .numberFormat = "@"
            .HorizontalAlignment = xlLeft
            .Value = relationship

        End With

        With ws.Cells(startRow, 2)

            .Value = relDict(relationship)
            .HorizontalAlignment = xlLeft

        End With

        startRow = startRow + 1

    Next relationship

End Sub

Private Function WriteSupportingTransactions( _
    ByVal hostWb As Workbook, _
    ByVal ws As Worksheet, _
    ByVal startRow As Long, _
    ByVal members As Collection, _
    ByVal matchIDLookup As Object) As Long


    Dim txnWs As Worksheet

    Dim rowNum As Variant
    Dim txnRow As Long

    Dim outRow As Long
    
    Set matchIDLookup = _
        BuildMatchIDBySourceRow()

    Set txnWs = _
        hostWb.Worksheets("Transfer_Transactions")

    outRow = _
        startRow


    ws.Cells(outRow, 1).Value = _
        "Account"

    ws.Cells(outRow, 2).Value = _
        "Code Description"

    ws.Cells(outRow, 3).Value = _
        "Date"

    ws.Cells(outRow, 4).Value = _
        "Amount"

    ws.Cells(outRow, 5).Value = _
        "Description"

    ws.Cells(outRow, 6).Value = _
        "Match ID"

    ws.rows(outRow).Font.Bold = _
        True

    outRow = _
        outRow + 1


    For Each rowNum In members

        txnRow = _
            FindTransactionRow( _
                hostWb, _
                CLng(rowNum))

        If txnRow > 0 Then

            ' Account

            With ws.Cells(outRow, 1)

                .numberFormat = "@"
                .HorizontalAlignment = xlLeft

                .Value = _
                    CStr(txnWs.Cells( _
                        txnRow, 3).Value)

            End With

            ' Code description

            ws.Cells(outRow, 2).Value = _
                txnWs.Cells(txnRow, 6).Value

            
            ' Transaction date

            With ws.Cells(outRow, 3)

                .Value = _
                    txnWs.Cells(txnRow, 4).Value

                .numberFormat = _
                    "m/d/yyyy"

            End With
            

            ' Transaction amount

            With ws.Cells(outRow, 4)

                .Value = _
                    txnWs.Cells(txnRow, 5).Value

                .numberFormat = _
                    "$#,##0.00;($#,##0.00)"

            End With


            ' Transaction description

            ws.Cells(outRow, 5).Value = _
                txnWs.Cells(txnRow, 7).Value


            ' Final Match ID

            With ws.Cells(outRow, 6)

                .numberFormat = "@"
                .HorizontalAlignment = xlLeft

                If matchIDLookup.Exists( _
                    CStr(CLng(rowNum))) Then

                    .Value = _
                        CStr(matchIDLookup( _
                            CStr(CLng(rowNum))))

                Else

                    .Value = ""

                End If

            End With

            outRow = _
                outRow + 1

        Else

            Debug.Print _
                "SUPPORTING TRANSACTION NOT FOUND", _
                rowNum

        End If

    Next rowNum


    If outRow > startRow + 1 Then

        With ws.Sort

            .SortFields.Clear

            .SortFields.Add _
                key:=ws.Range( _
                    ws.Cells(startRow + 1, 1), _
                    ws.Cells(outRow - 1, 1)), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending

            .SetRange ws.Range( _
                ws.Cells(startRow, 1), _
                ws.Cells(outRow - 1, 6))

            .header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom

            .Apply

        End With

    End If

    WriteSupportingTransactions = _
        outRow

End Function

Private Function BuildMatchIDBySourceRow() As Object

    Dim results As Object
    Dim relWs As Worksheet

    Dim lastRow As Long
    Dim r As Long

    Dim colMatchID As Long
    Dim colSourceRow As Long
    Dim colCandidateRow As Long

    Dim matchID As String
    Dim sourceRow As Long
    Dim candidateRow As Long

    Set results = _
        CreateObject("Scripting.Dictionary")

    Set relWs = _
        hostWb.Worksheets("Transfer_Relationships")

    colMatchID = _
        FindHeaderColumn( _
            relWs, _
            "Match ID")

    colSourceRow = _
        FindHeaderColumn( _
            relWs, _
            "Source Row")

    colCandidateRow = _
        FindHeaderColumn( _
            relWs, _
            "Candidate Row")

    If colMatchID = 0 _
    Or colSourceRow = 0 _
    Or colCandidateRow = 0 Then

        Set BuildMatchIDBySourceRow = _
            results

        Exit Function

    End If

    lastRow = _
        relWs.Cells( _
            relWs.rows.Count, _
            colMatchID).End(xlUp).Row


    For r = 2 To lastRow

        matchID = _
            Trim$(CStr( _
                relWs.Cells(r, colMatchID).Value))

        If matchID <> "" Then

            If IsNumeric( _
                relWs.Cells(r, colSourceRow).Value) Then

                sourceRow = _
                    CLng(relWs.Cells( _
                        r, colSourceRow).Value)

                results(CStr(sourceRow)) = _
                    matchID

            End If

            If IsNumeric( _
                relWs.Cells( _
                    r, colCandidateRow).Value) Then

                candidateRow = _
                    CLng(relWs.Cells( _
                        r, colCandidateRow).Value)

                results(CStr(candidateRow)) = _
                    matchID

            End If

        End If

    Next r

    Set BuildMatchIDBySourceRow = _
        results

End Function




Private Sub FormatAmbiguousPairsReport(ws As Worksheet)

    ws.Columns("A:F").AutoFit

    ws.Columns(5).numberFormat = _
        "$#,##0.00;($#,##0.00)"

    ws.rows.HorizontalAlignment = xlLeft

    ws.Activate

    ws.Range("A2").Select

    ActiveWindow.FreezePanes = True

End Sub


Private Function BuildAmbiguousStatusLookup() As Object

    Dim ws As Worksheet
    Dim dict As Object

    Dim lastRow As Long
    Dim r As Long

    Dim rowNumber As Long
    Dim clusterID As String

    Set dict = CreateObject("Scripting.Dictionary")
    Set ws = hostWb.Worksheets("Transfer_Transaction_Status")

    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow

        If UCase$(Trim$(ws.Cells(r, 6).Value)) = _
            UCase$(STATUS_AMBIGUOUS) Then

            rowNumber = CLng(ws.Cells(r, 2).Value)

            clusterID = _
                CStr(ws.Cells(r, 7).Value)

            dict(CStr(rowNumber)) = clusterID

        End If

    Next r

    Set BuildAmbiguousStatusLookup = dict

End Function



Private Sub BuildTransactionRowLookup( _
    ByVal hostWb As Workbook)

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim rowNumber As Long

    Set ws = hostWb.Worksheets("Transfer_Transactions")

    Set TransactionRowLookup = _
        CreateObject("Scripting.Dictionary")

    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow

        rowNumber = CLng(ws.Cells(r, 2).Value)

        TransactionRowLookup(rowNumber) = r

    Next r

End Sub

Private Function CountReferencedCandidates( _
    data As Variant, _
    targetRow As Long, _
    lastRow As Long, _
    colAcct As Long, _
    colDate As Long, _
    colAmount As Long, _
    colDescription As Long, _
    matched() As Boolean) As Long

    Dim v As Variant
    Dim i As Long

    Dim sourceAcct As String
    Dim sourceRef As String
    Dim sourceAmt As Double
    Dim sourceDate As Variant

    Dim candidateAcct As String
    Dim candidateRef As String
    Dim candidateAmt As Double

    sourceAcct = Trim(CStr(data(targetRow, colAcct)))

    sourceRef = ReferencedAccounts(targetRow)

    If sourceRef = "" Then Exit Function

    sourceAmt = CDbl(data(targetRow, colAmount))
    sourceDate = data(targetRow, colDate)

    If Not RefAcctIndex.Exists(sourceAcct) Then Exit Function

        For Each v In RefAcctIndex(sourceAcct)
        
            i = CLng(v)

        If i = targetRow Then GoTo NextI
        If matched(i) Then GoTo NextI
        
        If IsSameAccountPair(data, targetRow, i, colAcct) Then GoTo NextI

        candidateAcct = _
            Trim(CStr(data(i, colAcct)))

        candidateRef = ReferencedAccounts(i)

        candidateAmt = _
            CDbl(data(i, colAmount))

        If sourceRef <> candidateAcct Then GoTo NextI

        If candidateRef <> sourceAcct Then GoTo NextI

        If Abs(sourceAmt) <> _
           Abs(candidateAmt) Then GoTo NextI

        If sourceAmt * candidateAmt >= 0 Then GoTo NextI

        If sourceDate <> _
           data(i, colDate) Then GoTo NextI

        CountReferencedCandidates = _
            CountReferencedCandidates + 1

NextI:
    Next v

End Function


Private Sub WriteMatchedHeaders( _
    ws As Worksheet)


    ws.Cells(1, 1).Value = _
        "Match ID"

    ws.Cells(1, 2).Value = _
        "Credited Account"

    ws.Cells(1, 3).Value = _
        "Debited Account"

    ws.Cells(1, 4).Value = _
        "Date"

    ws.Cells(1, 5).Value = _
        "Amount"

    ws.Cells(1, 6).Value = _
        "Credited Description"

    ws.Cells(1, 7).Value = _
        "Debited Description"

    ws.Cells(1, 8).Value = _
        "Match Method"

    ws.Cells(1, 9).Value = _
        "Confidence Tier"


    ws.Cells(1, 10).Value = _
        "Transfer Pattern"

    ws.Cells(1, 11).Value = _
        "Investigation Group"

    ws.Cells(1, 12).Value = _
        "Cluster ID"


    ws.Cells(1, 13).Value = _
        "Source Row"

    ws.Cells(1, 14).Value = _
        "Candidate Row"

    With ws.rows(1)

        .Font.Bold = True

        .Interior.Color = _
            RGB(217, 225, 242)

    End With

    If ws.AutoFilterMode Then

        ws.AutoFilterMode = _
            False

    End If

    ws.rows(1).AutoFilter

End Sub



Private Sub WriteUnmatchedHeaders( _
    ws As Worksheet)


    ws.Cells(1, 1).Value = _
        "Account"

    ws.Cells(1, 2).Value = _
        "Date"

    ws.Cells(1, 3).Value = _
        "Direction"

    ws.Cells(1, 4).Value = _
        "Amount"

    ws.Cells(1, 5).Value = _
        "Transaction Type"

    ws.Cells(1, 6).Value = _
        "Description"

    ws.Cells(1, 7).Value = _
        "Investigator Notes"

    ws.Cells(1, 8).Value = _
        "Unmatched Reason"

    ws.Cells(1, 9).Value = _
        "Transfer Pattern"

    ws.Cells(1, 10).Value = _
        "Resolution Status"

    ws.Cells(1, 11).Value = _
        "Investigation Group"

    ws.Cells(1, 12).Value = _
        "Cluster ID"

    ws.Cells(1, 13).Value = _
        "Source Row"

    ws.Cells(1, 14).Value = _
        "Investigation Key"

    With ws.rows(1)

        .Font.Bold = True

        .Interior.Color = _
            RGB(217, 225, 242)

    End With

    If ws.AutoFilterMode Then

        ws.AutoFilterMode = _
            False

    End If

    ws.rows(1).AutoFilter

End Sub

'---Not really needed, but looks nice---

Private Function GetFriendlyClusterShape( _
    clusterShape As String) As String

    Select Case UCase$(Trim$(clusterShape))

        Case "ONE_TO_ONE"

            GetFriendlyClusterShape = _
                "One-to-One"

        Case "ONE_TO_ONE_REPEATED"

            GetFriendlyClusterShape = _
                "Repeated Transfers"

        Case "MANY_TO_ONE"

            GetFriendlyClusterShape = _
                "Many-to-One"

        Case "ONE_TO_MANY"

            GetFriendlyClusterShape = _
                "One-to-Many"

        Case "MANY_TO_MANY"

            GetFriendlyClusterShape = _
                "Many-to-Many"

        Case ""

            GetFriendlyClusterShape = _
                ""

        Case Else

            GetFriendlyClusterShape = _
                clusterShape

    End Select

End Function

Private Function GetFriendlyClusterOutcome( _
    clusterOutcome As String) As String

    Select Case UCase$(Trim$(clusterOutcome))

        Case "MATCHED"

            GetFriendlyClusterOutcome = _
                "Matched"

        Case "PARTIAL_MATCH"

            GetFriendlyClusterOutcome = _
                "Partially Matched"

        Case "UNBALANCED"

            GetFriendlyClusterOutcome = _
                "Unbalanced"

        Case "AMBIGUOUS"

            GetFriendlyClusterOutcome = _
                "Ambiguous"

        Case "REVIEW"

            GetFriendlyClusterOutcome = _
                "Review Required"

        Case ""

            GetFriendlyClusterOutcome = _
                ""

        Case Else

            GetFriendlyClusterOutcome = _
                clusterOutcome

    End Select

End Function


Private Sub WriteAmbiguityHeaders(ws As Worksheet)

    ws.Cells(1, 1).Value = "Run ID"
    ws.Cells(1, 2).Value = "Cluster ID"
    ws.Cells(1, 3).Value = "Cluster Size"
    ws.Cells(1, 4).Value = "Cluster Exposure"
    ws.Cells(1, 5).Value = "Created Timestamp"

    ws.rows(1).Font.Bold = True
    ws.rows(1).AutoFilter

End Sub

Private Sub WriteTransactionWarehouseHeaders(ws As Worksheet)

    ws.Cells(1, 1).Value = "Run ID"
    ws.Cells(1, 2).Value = "Row Number"
    ws.Cells(1, 3).Value = "Account"
    ws.Cells(1, 4).Value = "Date"
    ws.Cells(1, 5).Value = "Amount"
    ws.Cells(1, 6).Value = "Code Description"
    ws.Cells(1, 7).Value = "Description"
    ws.Cells(1, 8).Value = "Transfer Flag"
    ws.Cells(1, 9).Value = "Created Timestamp"

    ws.rows(1).Font.Bold = True
    ws.rows(1).AutoFilter

End Sub

Private Function WriteExecutiveSummary( _
    ByVal hostWb As Workbook, _
    ByVal ws As Worksheet, _
    ByVal startRow As Long) As Long
    
    Dim wsRel As Worksheet

    Dim r As Long
    
    Dim transferSectionStart As Long
    Dim transferSectionEnd As Long
    
    Dim investigationSectionStart As Long
    Dim investigationSectionEnd As Long
    
    Dim acctStats As Object
    Dim relStats As Object

    Dim hubAcct As String
    Dim hubScore As Double
    
    Set wsRel = _
        hostWb.Worksheets("Transfer_Relationships")

    r = startRow
    
    '---EXECUTIVE SUMMARY---
   
    With ws.Range( _
    ws.Cells(startRow, 1), _
    ws.Cells(startRow, 14))

    .Merge
    .Value = "EXECUTIVE SUMMARY"

    .Font.Bold = True
    .Font.size = 12

    .Interior.Color = RGB(221, 235, 247)

    End With

    r = r + 2
    
    transferSectionStart = r  'formatting outline
    
    ws.Cells(r, 1).Value = "TRANSFER STATISTICS"
    ws.Cells(r, 1).Font.Bold = True
    
    r = r + 1

    ws.Cells(r, 1).Value = "Total Transfer Transactions"
    ws.Cells(r, 2).Value = GetTotalTransferTransactions()
    
        r = r + 2
    
    ws.Cells(r, 1).Value = _
        "Transfers Requiring Review"
    
    ws.Cells(r, 2).Value = _
        GetTransfersRequiringReviewCount()
    
 
        With ws.Range( _
            ws.Cells(r, 1), _
            ws.Cells(r, 2))
    
   
            .Font.Bold = _
                True
    
        End With
    
    
    r = r + 1
    
    ws.Cells(r, 1).Value = _
        "  No Corresponding Transfer Identified"
    
    ws.Cells(r, 2).Value = _
        GetStatusCount(STATUS_UNMATCHED)
        
        If ws.Cells(r, 2).Value = 0 Then
        
            ws.Cells(r, 1).Value = _
            "  No External Unmatched Found!"
    
            With ws.Range( _
                ws.Cells(r, 1), _
                ws.Cells(r, 2))
        
                .Interior.Color = _
                    RGB(0, 176, 80)
        
                .Font.Color = _
                    RGB(255, 255, 255)
        
                .Font.Bold = _
                    True
        
            End With
        
       End If
    
    r = r + 1
    
    ws.Cells(r, 1).Value = _
        "  Unresolved Within Transfer Groups"
    
    ws.Cells(r, 2).Value = _
        GetStatusCount(STATUS_AMBIGUOUS)
    
    
    r = r + 1
    
    ws.Cells(r, 1).Value = _
        "Review Volume"
    
    ws.Cells(r, 2).Value = _
        GetTransfersRequiringReviewVolume()
    
    ws.Cells(r, 2).numberFormat = _
        "$#,##0.00;($#,##0.00)"

    r = r + 2

    ws.Cells(r, 1).Value = "Matched Transfers"
    ws.Cells(r, 2).Value = GetStatusCount(STATUS_MATCHED)
    
    If GetStatusCount(STATUS_MATCHED) = 0 Then
            With ws.Range(ws.Cells(r, 1), ws.Cells(r, 2))
                .Interior.Color = RGB(255, 0, 0)
                .Font.Color = RGB(255, 255, 255)
                .Font.Bold = True
            End With
    End If

    r = r + 1

    ws.Cells(r, 1).Value = "Matched Volume"
    ws.Cells(r, 2).Value = GetMatchedVolume()
    ws.Cells(r, 2).numberFormat = _
        "$#,##0.00;($#,##0.00)"
        
    r = r + 3

    ws.Cells(r, 1).Value = _
        "Analysis Tables"
        

    r = r + 3
        
    ws.Cells(r, 1).Value = _
        "Primary Hub Account"
        
        With ws.Cells(r, 2)
        
            .numberFormat = "@"
        
            .Value2 = _
                GetPrimaryHubAccount(ws)
        
            .HorizontalAlignment = _
                xlLeft
        
        End With
        
    r = r + 2

    ws.Cells(r, 1).Value = _
        "Primary Hub Analysis"

    transferSectionEnd = r
   
    r = r + 1


    ' INVESTIGATION SUMMARY

    
    r = r + 2
    
    investigationSectionStart = r
    
    ws.Cells(r, 1).Value = _
        "INVESTIGATION SUMMARY"
    
    ws.Cells(r, 1).Font.Bold = _
        True
    
    r = r + 1
    
    ws.Cells(r, 1).Value = _
        "Total Investigation Groups"
    
    ws.Cells(r, 2).Value = _
        GetTotalInvestigationGroups()
    
    r = r + 1
    
    ws.Cells(r, 1).Value = _
        "Total Exposure"
    
    ws.Cells(r, 2).Value = _
        GetTotalExposure()
    
    ws.Cells(r, 2).numberFormat = _
        "$#,##0.00;($#,##0.00)"
    

 
    r = r + 2
    
    ws.Cells(r, 1).Value = _
        "Unresolved Groups"
    
    ws.Cells(r, 2).Value = _
        GetUnresolvedGroupCount()
    
    r = r + 1
    
    ws.Cells(r, 1).Value = _
        "  Unresolved Unmatched Transactions"
    
    ws.Cells(r, 2).Value = _
        GetReviewGroupTransactionCount()
    
    r = r + 1
    
    ws.Cells(r, 1).Value = _
        "Unresolved Exposure"
    
    ws.Cells(r, 2).Value = _
        GetUnresolvedExposure()
    
    ws.Cells(r, 2).numberFormat = _
        "$#,##0.00;($#,##0.00)"
        
    r = r + 2

    ws.Cells(r, 1).Value = _
        "Investigation Groups"
    
    r = r + 2
    
    ws.Cells(r, 1).Value = _
        "Resolved Groups"
    
    ws.Cells(r, 2).Value = _
        GetResolvedGroupCount()
    
   
    r = r + 1
    
    ws.Cells(r, 1).Value = _
        "  Resolved Transactions"
    
    ws.Cells(r, 2).Value = _
        GetResolvedGroupTransactionCount()
    
    r = r + 1
    
    ws.Cells(r, 1).Value = _
        "Resolved Exposure"
    
    ws.Cells(r, 2).Value = _
        GetResolvedExposure()
    
    ws.Cells(r, 2).numberFormat = _
        "$#,##0.00;($#,##0.00)"

    investigationSectionEnd = r

    WriteExecutiveSummary = r
    
    With ws.Range( _
        ws.Cells(transferSectionStart, 1), _
        ws.Cells(transferSectionEnd, 2))
    
        .BorderAround _
            xlContinuous, _
            xlThin
    
    End With
        
    With ws.Range( _
        ws.Cells(investigationSectionStart, 1), _
        ws.Cells(investigationSectionEnd, 2))
    
        .BorderAround _
            xlContinuous, _
            xlThin
    
    End With
   
    

End Function


Private Sub WriteAmbiguityCluster( _
    ByVal hostWb As Workbook, _
    ByVal clusterID As String, _
    ByVal clusterSize As Long, _
    ByVal clusterExposure As Double)

    Dim ws As Worksheet
    Dim r As Long

    Set ws = _
        hostWb.Worksheets("Transfer_Ambiguities")

    r = ws.Cells(ws.rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(r, 1).Value = CurrentRunID
    ws.Cells(r, 2).Value = clusterID
    ws.Cells(r, 3).Value = clusterSize
    ws.Cells(r, 4).Value = clusterExposure
    ws.Cells(r, 5).Value = Now

End Sub

Private Sub WriteAmbiguityMember( _
    ByVal hostWb As Workbook, _
    ByVal clusterID As String, _
    ByVal rowNum As Long, _
    ByVal acct As String, _
    ByVal txnDate As Variant, _
    ByVal amount As Variant)

    Dim ws As Worksheet
    Dim r As Long

    Set ws = hostWb.Worksheets("Transfer_Ambiguity_Members")

    r = ws.Cells(ws.rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(r, 1).Value = CurrentRunID
    ws.Cells(r, 2).Value = clusterID
    ws.Cells(r, 3).Value = rowNum
    ws.Cells(r, 4).Value = "'" & acct
    ws.Cells(r, 5).Value = txnDate
    ws.Cells(r, 6).Value = amount
    ws.Cells(r, 7).Value = Now

End Sub

Private Sub WriteAmbiguityMemberHeaders(ws As Worksheet)

    ws.Cells(1, 1).Value = "Run ID"
    ws.Cells(1, 2).Value = "Cluster ID"
    ws.Cells(1, 3).Value = "Row Number"
    ws.Cells(1, 4).Value = "Account"
    ws.Cells(1, 5).Value = "Date"
    ws.Cells(1, 6).Value = "Amount"
    ws.Cells(1, 7).Value = "Created Timestamp"

    ws.rows(1).Font.Bold = True
    ws.rows(1).AutoFilter

End Sub

Private Sub PersistAmbiguityWarehouse( _
    ByVal hostWb As Workbook, _
    ByRef data As Variant, _
    ByVal colAcct As Long, _
    ByVal colDate As Long, _
    ByVal colAmount As Long)

    Dim rowNum As Variant
    Dim clusterID As String
    Dim processed As Object

    Set processed = _
        CreateObject("Scripting.Dictionary")

    For Each rowNum In AmbiguousClusters.Keys

        clusterID = _
            AmbiguousClusters(rowNum)

        If Not processed.Exists(clusterID) Then

            processed.Add clusterID, 1

            WriteAmbiguityCluster _
                hostWb, _
                clusterID, _
                ClusterSizes(clusterID), _
                ClusterValues(clusterID)

        End If

        WriteAmbiguityMember _
            hostWb, _
            clusterID, _
            CLng(rowNum), _
            CStr(data(CLng(rowNum), colAcct)), _
            data(CLng(rowNum), colDate), _
            data(CLng(rowNum), colAmount)

    Next rowNum

End Sub

Private Sub PersistTransactionsWarehouse( _
    ByVal hostWb As Workbook, _
    ByRef data As Variant, _
    ByVal lastRow As Long, _
    ByVal colAcct As Long, _
    ByVal colDate As Long, _
    ByVal colAmount As Long, _
    ByVal colCodeDesc As Long, _
    ByVal colDescription As Long)

    Dim ws As Worksheet
    Dim r As Long
    Dim i As Long

    Set ws = hostWb.Worksheets("Transfer_Transactions")

    r = 2

    For i = 2 To lastRow

        ws.Cells(r, 1).Value = CurrentRunID
        ws.Cells(r, 2).Value = i
        ws.Cells(r, 3).Value = "'" & _
            CStr(data(i, colAcct))
        ws.Cells(r, 4).Value = _
            data(i, colDate)
        ws.Cells(r, 5).Value = _
            data(i, colAmount)
        ws.Cells(r, 6).Value = _
            data(i, colCodeDesc)
        ws.Cells(r, 7).Value = _
            data(i, colDescription)
        ws.Cells(r, 8).Value = _
            TransferFlags(i)
        ws.Cells(r, 9).Value = _
            Now

        r = r + 1

    Next i

End Sub

Private Sub WriteMetadataHeaders(ws As Worksheet)

ws.Cells(1, 1).Value = "Run ID"
    ws.Cells(1, 2).Value = "Version"
    ws.Cells(1, 3).Value = "Execution Timestamp"

ws.Cells(1, 4).Value = "Workbook Name"
    ws.Cells(1, 5).Value = "Source Worksheet"

ws.Cells(1, 6).Value = "Rows Processed"
    ws.Cells(1, 7).Value = "Transfer Transactions"

ws.Cells(1, 8).Value = "Matched Count"
    ws.Cells(1, 9).Value = "Matched %"

ws.Cells(1, 10).Value = "Ambiguous Count"
    ws.Cells(1, 11).Value = "Ambiguous %"

ws.Cells(1, 12).Value = "Unmatched Count"
    ws.Cells(1, 13).Value = "Unmatched %"

ws.Cells(1, 14).Value = "Matched Volume"
    ws.Cells(1, 15).Value = "Ambiguous Volume"
    ws.Cells(1, 16).Value = "Unmatched Volume"

ws.Cells(1, 17).Value = "Ambiguous Clusters"
    ws.Cells(1, 18).Value = "Largest Cluster Size"
    ws.Cells(1, 19).Value = "Largest Cluster Exposure"

ws.Cells(1, 20).Value = "Candidates Generated"

ws.Cells(1, 21).Value = "Confirmation Matches"
    ws.Cells(1, 22).Value = "Narrative Matches"
    ws.Cells(1, 23).Value = "Referenced Account Matches"
    ws.Cells(1, 24).Value = "Embedded Account Matches"
    ws.Cells(1, 25).Value = "Transfer Suffix Matches"
    ws.Cells(1, 26).Value = "Branch Transfer Matches"
    ws.Cells(1, 27).Value = "Amount-Date Matches"

ws.Cells(1, 28).Value = "Confirmation Ambiguities"
    ws.Cells(1, 29).Value = "Narrative Ambiguities"
    ws.Cells(1, 30).Value = "Referenced Account Ambiguities"
    ws.Cells(1, 31).Value = "Embedded Account Ambiguities"
    ws.Cells(1, 32).Value = "Transfer Suffix Ambiguities"
    ws.Cells(1, 33).Value = "Branch Transfer Ambiguities"
    ws.Cells(1, 34).Value = "Amount-Date Ambiguities"

ws.Cells(1, 35).Value = "Runtime Seconds"

ws.rows(1).Font.Bold = True
    ws.rows(1).AutoFilter

End Sub


Private Sub WriteContradictionHeaders(ws As Worksheet)

    ws.Cells(1, 1).Value = "Run ID"
    ws.Cells(1, 2).Value = "Contradiction ID"
    ws.Cells(1, 3).Value = "Source Row"
    ws.Cells(1, 4).Value = "Candidate Row"
    ws.Cells(1, 5).Value = "Source Account"
    ws.Cells(1, 6).Value = "Candidate Account"
    ws.Cells(1, 7).Value = "Date"
    ws.Cells(1, 8).Value = "Amount"
    ws.Cells(1, 9).Value = "Contradiction Type"
    ws.Cells(1, 10).Value = "Rejected Match Method"
    ws.Cells(1, 11).Value = "Created Timestamp"

    ws.rows(1).Font.Bold = True
    ws.rows(1).AutoFilter

End Sub

Private Sub WriteCandidateHeaders(ws As Worksheet)

    ws.Cells(1, 1).Value = "Run ID"
    ws.Cells(1, 2).Value = "Candidate ID"
    ws.Cells(1, 3).Value = "Source Row"
    ws.Cells(1, 4).Value = "Candidate Row"
    ws.Cells(1, 5).Value = "Source Account"
    ws.Cells(1, 6).Value = "Candidate Account"
    ws.Cells(1, 7).Value = "Date"
    ws.Cells(1, 8).Value = "Amount"
    ws.Cells(1, 9).Value = "Discovered By Method"
    ws.Cells(1, 10).Value = "Candidate Status"
    ws.Cells(1, 11).Value = "Resolution ID"
    ws.Cells(1, 12).Value = "Resolution Timestamp"
    ws.Cells(1, 13).Value = "Created Timestamp"

    ws.rows(1).Font.Bold = True
    ws.rows(1).AutoFilter

End Sub

Private Sub WriteTransactionStatusHeaders(ws As Worksheet)

    ws.Cells(1, 1).Value = "Run ID"
    ws.Cells(1, 2).Value = "Row Number"
    ws.Cells(1, 3).Value = "Account"
    ws.Cells(1, 4).Value = "Date"
    ws.Cells(1, 5).Value = "Amount"
    ws.Cells(1, 6).Value = "Status"
    ws.Cells(1, 7).Value = "Related ID"
    ws.Cells(1, 8).Value = "Created Timestamp"

    ws.rows(1).Font.Bold = True
    ws.rows(1).AutoFilter

End Sub



Private Sub PassConfirmationNumbers( _
    ByVal hostWb As Workbook, _
    ByVal ws As Worksheet, _
    ByRef data As Variant, _
    ByVal lastRow As Long, _
    ByVal colAcct As Long, _
    ByVal colCodeDesc As Long, _
    ByVal colDate As Long, _
    ByVal colAmount As Long, _
    ByVal colDescription As Long, _
    ByRef matched() As Boolean, _
    ByRef ambiguous() As Boolean, _
    ByVal ambiguityPairs As Collection)
    
    DebugLog "START PassConfirmationNumbers"
    

    
    Dim i As Long
    Dim j As Long
    Dim confI As String
    Dim candidates As Collection
    Dim v As Variant
    Dim reverseCount As Long
    
    For i = 2 To lastRow
    
        If Not TransferFlags(i) Then GoTo NextI
        If matched(i) Then GoTo NextI
    
        confI = ConfirmationNumbers(i)
    
        If confI = "" Then GoTo NextI
    
        If Not ConfIndex.Exists(confI) Then GoTo NextI
    
        Set candidates = New Collection
    
        For Each v In ConfIndex(confI)
    
            j = CLng(v)
    
            If j <= i Then GoTo NextJ
    
            If matched(j) Then GoTo NextJ
            
            If IsSameAccountPair(data, i, j, colAcct) Then GoTo NextJ
    
            If Abs(CDbl(data(i, colAmount))) <> _
               Abs(CDbl(data(j, colAmount))) Then GoTo NextJ
    
            If CDbl(data(i, colAmount)) * _
               CDbl(data(j, colAmount)) >= 0 Then GoTo NextJ
    
            If data(i, colDate) <> _
               data(j, colDate) Then GoTo NextJ
               
            WriteCandidate _
                hostWb, _
                i, _
                j, _
                CStr(data(i, colAcct)), _
                CStr(data(j, colAcct)), _
                data(i, colDate), _
                Abs(CDbl(data(i, colAmount))), _
                "Confirmation Number"
    
            candidates.Add j
    
NextJ:
        Next v
                        
    
        Select Case candidates.Count
    
            Case 1
    
                reverseCount = _
                    CountConfirmationCandidates( _
                        data, _
                        CLng(candidates(1)), _
                        colAcct, _
                        colDate, _
                        colAmount, _
                        colDescription, _
                        matched)
    
                If reverseCount = 1 Then
    
                    WriteMatched _
                        hostWb, _
                        data, _
                        i, _
                        CLng(candidates(1)), _
                        colAcct, _
                        colCodeDesc, _
                        colDate, _
                        colAmount, _
                        colDescription, _
                        "Confirmation Number"

    
                    matched(i) = True
                    matched(CLng(candidates(1))) = True
    
                Else
    
                    AddAmbiguity _
                        ambiguityPairs, _
                        i, _
                        CLng(candidates(1)), _
                        "Confirmation Number"
    
    
                End If
    
        End Select
    
NextI:
    Next i
    
   
    DebugLog "END PassConfirmationNumbers"

End Sub

Private Sub PassNarrativePairs( _
    ByVal hostWb As Workbook, _
    ByVal ws As Worksheet, _
    ByRef data As Variant, _
    ByVal lastRow As Long, _
    ByVal colAcct As Long, _
    ByVal colCodeDesc As Long, _
    ByVal colDate As Long, _
    ByVal colAmount As Long, _
    ByVal colDescription As Long, _
    ByRef matched() As Boolean, _
    ByRef ambiguous() As Boolean, _
    ByVal ambiguityPairs As Collection)
    
    DebugLog "START PassNarrativePairs"

    Dim i As Long
    Dim j As Long
    Dim narrativeI As String
    Dim candidates As Collection
    Dim v As Variant
    Dim reverseCount As Long


    For i = 2 To lastRow

        If Not TransferFlags(i) Then GoTo NextI

        If matched(i) Then GoTo NextI
        
        narrativeI = NarrativeIDs(i)
          
            
        Set candidates = New Collection

        If narrativeI = "" Then GoTo NextI
        
        Set candidates = New Collection
        
        If Not NarrativeIndex.Exists(narrativeI) Then GoTo NextI
        
        For Each v In NarrativeIndex(narrativeI)

            j = CLng(v)
        
            If j <= i Then GoTo NextJ
        
            If matched(j) Then GoTo NextJ
            
            If IsSameAccountPair(data, i, j, colAcct) Then GoTo NextJ
        
            If Abs(CDbl(data(i, colAmount))) <> _
               Abs(CDbl(data(j, colAmount))) Then GoTo NextJ
        
            If CDbl(data(i, colAmount)) * _
               CDbl(data(j, colAmount)) >= 0 Then GoTo NextJ
        
            If data(i, colDate) <> _
               data(j, colDate) Then GoTo NextJ
               
            WriteCandidate _
                hostWb, _
                i, _
                j, _
                CStr(data(i, colAcct)), _
                CStr(data(j, colAcct)), _
                data(i, colDate), _
                Abs(CDbl(data(i, colAmount))), _
                "Narrative Pair"
        
            candidates.Add j

NextJ:
Next v
              
        Select Case candidates.Count
        
            Case 1
        
                reverseCount = CountNarrativeCandidates( _
                    data, _
                    CLng(candidates(1)), _
                    colAcct, _
                    colDate, _
                    colAmount, _
                    colDescription, _
                    matched)
                
                If reverseCount = 1 Then
                
                    WriteMatched _
                        hostWb, _
                        data, _
                        i, _
                        CLng(candidates(1)), _
                        colAcct, _
                        colCodeDesc, _
                        colDate, _
                        colAmount, _
                        colDescription, _
                        "Narrative Pair"

                
                    matched(i) = True
                    matched(CLng(candidates(1))) = True
                
                Else
                
                    AddAmbiguity _
                        ambiguityPairs, _
                        i, _
                        CLng(candidates(1)), _
                        "Narrative Pair"
                
                End If
        
            Case Is > 1
        
                For Each v In candidates

                    AddAmbiguity _
                        ambiguityPairs, _
                        i, _
                        CLng(v), _
                        "Narrative Pair"
                
                Next v

End Select

NextI:
    Next i

DebugLog "END PassNarrativePairs"

End Sub

Private Function CountNarrativeCandidates( _
    data As Variant, _
    targetRow As Long, _
    colAcct As Long, _
    colDate As Long, _
    colAmount As Long, _
    colDescription As Long, _
    matched() As Boolean) As Long

    Dim targetID As String
    Dim targetAmt As Double
    Dim targetDate As Variant

    Dim v As Variant
    Dim i As Long

    targetID = NarrativeIDs(targetRow)

    If targetID = "" Then Exit Function

    If Not NarrativeIndex.Exists(targetID) Then Exit Function

    targetAmt = CDbl(data(targetRow, colAmount))
    targetDate = data(targetRow, colDate)

    For Each v In NarrativeIndex(targetID)

        i = CLng(v)

        If i = targetRow Then GoTo NextI
        If matched(i) Then GoTo NextI
        
        If IsSameAccountPair(data, targetRow, i, colAcct) Then GoTo NextI

        If Abs(targetAmt) <> _
           Abs(CDbl(data(i, colAmount))) Then GoTo NextI

        If targetAmt * CDbl(data(i, colAmount)) >= 0 Then GoTo NextI

        If targetDate <> data(i, colDate) Then GoTo NextI

        CountNarrativeCandidates = _
            CountNarrativeCandidates + 1

NextI:
    Next v

End Function



Private Sub PassReferencedAccounts( _
    ByVal hostWb As Workbook, _
    ByVal ws As Worksheet, _
    ByRef data As Variant, _
    ByVal lastRow As Long, _
    ByVal colAcct As Long, _
    ByVal colCodeDesc As Long, _
    ByVal colDate As Long, _
    ByVal colAmount As Long, _
    ByVal colDescription As Long, _
    ByRef matched() As Boolean, _
    ByRef ambiguous() As Boolean, _
    ByVal ambiguityPairs As Collection)
    
    DebugLog "START PassReferencedAccounts"

    Dim i As Long
    Dim j As Long
    Dim v As Variant
    
    Dim reverseCount As Long

    Dim sourceAcct As String
    Dim sourceRef As String

    Dim candidateAcct As String
    Dim candidateRef As String

    Dim sourceAmt As Double
    Dim candidateAmt As Double

    Dim sourceDate As Variant
    
    
    For i = 2 To lastRow
    
        If Not TransferFlags(i) Then

        
            GoTo NextI
        
        End If
   
        If matched(i) Then GoTo NextI
        
        sourceRef = ReferencedAccounts(i)

        If sourceRef = "" Then GoTo NextI

        sourceAcct = Trim(CStr(data(i, colAcct)))

        sourceAmt = CDbl(data(i, colAmount))

        sourceDate = data(i, colDate)
        
        Dim candidates As Collection
        Set candidates = New Collection

        If Not RefAcctIndex.Exists(sourceAcct) Then GoTo NextI
        
        For Each v In RefAcctIndex(sourceAcct)

            j = CLng(v)
        
            If j <= i Then GoTo NextJ
        
            If matched(j) Then GoTo NextJ
            
            If IsSameAccountPair(data, i, j, colAcct) Then GoTo NextJ
        
            candidateRef = ReferencedAccounts(j)
        
            candidateAcct = _
                Trim(CStr(data(j, colAcct)))
        
            candidateAmt = _
                CDbl(data(j, colAmount))
        
            If sourceRef <> candidateAcct Then GoTo NextJ
        
            If candidateRef <> sourceAcct Then GoTo NextJ
        
            If Abs(sourceAmt) <> _
               Abs(candidateAmt) Then GoTo NextJ
        
            If sourceAmt * candidateAmt >= 0 Then GoTo NextJ
        
            If sourceDate <> data(j, colDate) Then GoTo NextJ
            
            WriteCandidate _
                hostWb, _
                i, _
                j, _
                CStr(data(i, colAcct)), _
                CStr(data(j, colAcct)), _
                data(i, colDate), _
                Abs(CDbl(data(i, colAmount))), _
                "Referenced Account"
        
            candidates.Add j

NextJ:
Next v

         
        Select Case candidates.Count

    Case 1

        reverseCount = _
            CountReferencedCandidates( _
            data, _
            CLng(candidates(1)), _
            lastRow, _
            colAcct, _
            colDate, _
            colAmount, _
            colDescription, _
            matched)
                
        If reverseCount = 1 Then
        
        WriteMatched _
            hostWb, _
            data, _
            i, _
            CLng(candidates(1)), _
            colAcct, _
            colCodeDesc, _
            colDate, _
            colAmount, _
            colDescription, _
            "Referenced Account"
            


        matched(i) = True
        matched(CLng(candidates(1))) = True

    Else
    
        AddAmbiguity _
            ambiguityPairs, _
            i, _
            CLng(candidates(1)), _
            "Referenced Account"
            
    End If

End Select
NextI:
    Next i
    
DebugLog "END PassReferencedAccounts"

End Sub


Private Sub PassEmbeddedAccounts( _
    ByVal hostWb As Workbook, _
    ByVal ws As Worksheet, _
    ByRef data As Variant, _
    ByVal lastRow As Long, _
    ByVal colAcct As Long, _
    ByVal colCodeDesc As Long, _
    ByVal colDate As Long, _
    ByVal colAmount As Long, _
    ByVal colDescription As Long, _
    ByRef matched() As Boolean, _
    ByRef ambiguous() As Boolean, _
    ByVal ambiguityPairs As Collection)
    
    DebugLog "START PassEmbeddedAccounts"
    
    Dim i As Long
    Dim j As Long
    Dim v As Variant

    Dim sourceAmt As Double
    Dim candidateAmt As Double

    Dim sourceDate As Variant
    
    Dim sourceAcct As String
    Dim candidateAcct As String
    
    Dim sourceEmbedded As String
    Dim candidateEmbedded As String
    


    For i = 2 To lastRow
            
        If matched(i) Then GoTo NextI
             
        If Not TransferFlags(i) Then

            GoTo NextI
        
        End If
            
                If InStr(1, _
                    CStr(data(i, colCodeDesc)), _
                    "Internet Transfer", _
                    vbTextCompare) = 0 Then
                
                    GoTo NextI
                
         End If

   
        sourceEmbedded = EmbeddedAccounts(i)
                
        If sourceEmbedded = "" Then GoTo NextI
        
        sourceAcct = _
            Trim(CStr(data(i, colAcct)))


        sourceAmt = CDbl( _
            data(i, colAmount))

        sourceDate = data(i, colDate)
            
        Dim candidates As Collection
        Set candidates = New Collection

        For j = i + 1 To lastRow
        
            If matched(j) Then GoTo NextJ
            
            If IsSameAccountPair(data, i, j, colAcct) Then GoTo NextJ

            If Not TransferFlags(j) Then

                GoTo NextJ
            End If
              
                    If InStr(1, _
                        CStr(data(j, colCodeDesc)), _
                        "Internet Transfer", _
                        vbTextCompare) = 0 Then
                    
                        GoTo NextJ
                    
             End If

      

            candidateAcct = _
                Trim(CStr(data(j, colAcct)))
                
            candidateEmbedded = EmbeddedAccounts(j)
                    
            candidateAmt = _
                CDbl(data(j, colAmount))
            
            If sourceEmbedded <> candidateAcct Then _
                GoTo NextJ
            
            If candidateEmbedded <> sourceAcct Then _
                GoTo NextJ

            If Abs(sourceAmt) <> _
               Abs(candidateAmt) Then _
                GoTo NextJ

            If sourceAmt * candidateAmt >= 0 Then _
                GoTo NextJ

            If sourceDate <> _
               data(j, colDate) Then _
                GoTo NextJ
                
            WriteCandidate _
                hostWb, _
                i, _
                j, _
                CStr(data(i, colAcct)), _
                CStr(data(j, colAcct)), _
                data(i, colDate), _
                Abs(CDbl(data(i, colAmount))), _
                "Embedded Account"
                
            candidates.Add j

NextJ:
        Next j


Select Case candidates.Count


        Case 1
        
            Dim reverseCount As Long
        
            reverseCount = CountEmbeddedCandidates( _
                data, _
                CLng(candidates(1)), _
                lastRow, _
                colAcct, _
                colCodeDesc, _
                colDate, _
                colAmount, _
                colDescription)
        
            If reverseCount = 1 Then
        
                WriteMatched _
                    hostWb, _
                    data, _
                    i, _
                    candidates(1), _
                    colAcct, _
                    colCodeDesc, _
                    colDate, _
                    colAmount, _
                    colDescription, _
                    "Embedded Account"
                    
    
                matched(i) = True
                matched(candidates(1)) = True
        
            Else

                AddAmbiguity _
                    ambiguityPairs, _
                    i, _
                    CLng(candidates(1)), _
                    "Embedded Account"
            
            End If


            Case Is > 1
            
                          
                For Each v In candidates
                
                    AddAmbiguity _
                        ambiguityPairs, _
                        i, _
                        CLng(v), _
                        "Embedded Account"
                
                Next v

End Select

NextI:
    Next i
    
DebugLog "END PassEmbeddedAccounts"

End Sub



Private Sub PassTransferSuffixes( _
    ByVal hostWb As Workbook, _
    ByVal ws As Worksheet, _
    ByRef data As Variant, _
    ByVal lastRow As Long, _
    ByVal colAcct As Long, _
    ByVal colCodeDesc As Long, _
    ByVal colDate As Long, _
    ByVal colAmount As Long, _
    ByVal colDescription As Long, _
    ByRef matched() As Boolean, _
    ByRef ambiguous() As Boolean, _
    ByVal ambiguityPairs As Collection)
    
    DebugLog "START PassTransferSuffixes"
    
    Dim i As Long
    Dim j As Long
    Dim v As Variant

    Dim fromSuffix As String
    Dim toSuffix As String

    Dim sourceAcct As String
    Dim candidateAcct As String

    Dim sourceAmt As Double
    Dim candidateAmt As Double

    Dim sourceDate As Variant
    
    
    
    For i = 2 To lastRow
    
        If matched(i) Then GoTo NextI
        
        If Not TransferFlags(i) Then

            GoTo NextI

        End If


        If Not GetTransferSuffixes( _
            CStr(data(i, colDescription)), _
            fromSuffix, _
            toSuffix) Then
            
            GoTo NextI

        End If
        
    
        fromSuffix = Replace(fromSuffix, " ", "")
        toSuffix = Replace(toSuffix, " ", "")
        
        sourceAcct = _
            Trim(CStr(data(i, colAcct)))

        sourceAmt = _
            CDbl(data(i, colAmount))

        sourceDate = data(i, colDate)
        
        Dim candidates As Collection
        Set candidates = New Collection

        For j = i + 1 To lastRow

            If matched(j) Then GoTo NextJ
            
            If IsSameAccountPair(data, i, j, colAcct) Then GoTo NextJ
            
            If Not TransferFlags(j) Then

                GoTo NextJ

            End If

            candidateAcct = _
                Trim(CStr(data(j, colAcct)))

            candidateAmt = _
                CDbl(data(j, colAmount))

            If Abs(sourceAmt) <> _
               Abs(candidateAmt) Then GoTo NextJ

            If sourceAmt * candidateAmt >= 0 Then _
                GoTo NextJ

            If sourceDate <> _
               data(j, colDate) Then GoTo NextJ

            If sourceAmt > 0 Then

                If AccountEndsWith(sourceAcct, toSuffix) _
                And AccountEndsWith(candidateAcct, fromSuffix) Then

                    candidates.Add j

                End If

            Else

                If AccountEndsWith(sourceAcct, fromSuffix) _
                And AccountEndsWith(candidateAcct, toSuffix) Then
                
                WriteCandidate _
                    hostWb, _
                    i, _
                    j, _
                    CStr(data(i, colAcct)), _
                    CStr(data(j, colAcct)), _
                    data(i, colDate), _
                    Abs(CDbl(data(i, colAmount))), _
                    "Transfer Suffix"

                    candidates.Add j
                   

                End If

            End If

NextJ:
        Next j
        
        Select Case candidates.Count

            Case 1
        
                Dim reverseCount As Long
        
                reverseCount = _
                    CountSuffixCandidates( _
                        data, _
                        CLng(candidates(1)), _
                        lastRow, _
                        colAcct, _
                        colDate, _
                        colAmount, _
                        colDescription, _
                        matched)
        
                If reverseCount = 1 Then
        
                    WriteMatched _
                        hostWb, _
                        data, _
                        i, _
                        CLng(candidates(1)), _
                        colAcct, _
                        colCodeDesc, _
                        colDate, _
                        colAmount, _
                        colDescription, _
                        "Transfer Suffix"
                        
        
                    matched(i) = True
                    matched(CLng(candidates(1))) = True
        
                Else
        
                    AddAmbiguity _
                        ambiguityPairs, _
                        i, _
                        CLng(candidates(1)), _
                        "Transfer Suffix"
        
                End If

End Select
NextI:
    Next i
    
    DebugLog "END PassTransferSuffixes"

End Sub

    
'---Leftovers, working well, DO NOT TOUCH---

Private Sub PassAmountDate( _
    ByVal hostWb As Workbook, _
    ByVal ws As Worksheet, _
    ByRef data As Variant, _
    ByVal lastRow As Long, _
    ByVal colAcct As Long, _
    ByVal colCodeDesc As Long, _
    ByVal colDate As Long, _
    ByVal colAmount As Long, _
    ByVal colDescription As Long, _
    ByRef matched() As Boolean, _
    ByRef ambiguous() As Boolean, _
    ByRef pendingAmbiguous() As Boolean, _
    ByVal ambiguityPairs As Collection)
    
    DebugLog "START PassAmountDate"

    Dim i As Long
    Dim j As Long
    Dim v As Variant

    Dim sourceAmt As Double
    Dim candidateAmt As Double

    Dim sourceDate As Variant

    Dim candidateCount As Long
    Dim candidateRow As Long
    Dim reverseCount As Long
    
    Dim passAmbiguous() As Boolean
    ReDim passAmbiguous(2 To lastRow)
    

    For i = 2 To lastRow
    
        If matched(i) Then GoTo NextI
        
        If pendingAmbiguous(i) Then GoTo NextI
        
        If Not TransferFlags(i) Then GoTo NextI
            
        sourceAmt = CDbl(data(i, colAmount))
        sourceDate = data(i, colDate)
        
        candidateCount = 0
        
                
        Dim candidates As Collection
        Set candidates = New Collection
        
        Dim amountDateKey As String

        amountDateKey = _
            Format(sourceDate, "yyyymmdd") & "|" & _
            Format(Abs(sourceAmt), "0.00")
            
            
        If Not AmtDateIndex.Exists(amountDateKey) Then _
            GoTo NextI

    For Each v In AmtDateIndex(amountDateKey)

        j = CLng(v)
        
            If i = j Then
            
                GoTo NextJ
                
            End If
        
            If matched(j) Then
            
                GoTo NextJ
    
            End If
                
           
            If IsSameAccountPair(data, i, j, colAcct) Then

                GoTo NextJ

            End If
        
        candidateAmt = CDbl(data(j, colAmount))
        
            If sourceAmt * candidateAmt >= 0 Then
        
                GoTo NextJ
            
            End If
            
            If HasUnresolvedAccountEvidence( _
                data, _
                i, _
                lastRow, _
                colAcct, _
                colDescription) Then _
                GoTo NextJ
        
'            If HasUnresolvedAccountEvidence( _
'                data, _
'                j, _
'                lastRow, _
'                colAcct, _
'                colDescription) Then _
'                GoTo NextJ

    
               
            Dim contradictionReason As String
            
            contradictionReason = _
                GetContradictionReason( _
                    data, _
                    i, _
                    j, _
                    colAcct, _
                    colDescription)
            
            If contradictionReason <> "" Then
            
                WriteContradiction _
                    hostWb, _
                    i, _
                    j, _
                    CStr(data(i, colAcct)), _
                    CStr(data(j, colAcct)), _
                    data(i, colDate), _
                    Abs(CDbl(data(i, colAmount))), _
                    contradictionReason, _
                    "Amount + Date"
            
                GoTo NextJ
            
            End If
            
            candidateCount = candidateCount + 1
            
            candidates.Add j
            
            candidateRow = j
            
            WriteCandidate _
                hostWb, _
                i, _
                j, _
                CStr(data(i, colAcct)), _
                CStr(data(j, colAcct)), _
                data(i, colDate), _
                Abs(CDbl(data(i, colAmount))), _
                "Amount + Date"
                
            candidateRow = j

NextJ:
Next v
        
    
        Select Case candidateCount

            Case 1

                reverseCount = _
                    CountAmountDateCandidates( _
                        data, _
                        candidateRow, _
                        lastRow, _
                        colAcct, _
                        colCodeDesc, _
                        colDate, _
                        colAmount, _
                        colDescription, _
                        matched)

    

            If passAmbiguous(i) _
            Or passAmbiguous(candidateRow) Then
            
                AddAmbiguity _
                    ambiguityPairs, _
                    i, _
                    candidateRow, _
                    "Amount + Date"
            
                GoTo NextI
            
            End If
                            
                If reverseCount = 1 Then

           
                    WriteMatched _
                        hostWb, _
                        data, _
                        i, _
                        candidateRow, _
                        colAcct, _
                        colCodeDesc, _
                        colDate, _
                        colAmount, _
                        colDescription, _
                        "Amount + Date"
                        
            
                    matched(i) = True
                    matched(candidateRow) = True
            
                Else

                    AddAmbiguity _
                        ambiguityPairs, _
                        i, _
                        candidateRow, _
                        "Amount + Date"
                
                End If
                
                       
            Case Is > 1
                      
            
                passAmbiguous(i) = True
                pendingAmbiguous(i) = True
            
                For Each v In candidates
                
                    passAmbiguous(CLng(v)) = True
                    pendingAmbiguous(CLng(v)) = True


                    AddAmbiguity _
                        ambiguityPairs, _
                        i, _
                        CLng(v), _
                        "Amount + Date"

                Next v
                
        End Select

NextI:
    Next i
    
    DebugLog "END PassAmountDate"

End Sub


Private Sub PassUnmatched( _
    ByVal hostWb As Workbook, _
    ByVal ws As Worksheet, _
    ByRef data As Variant, _
    ByVal lastRow As Long, _
    ByVal colAcct As Long, _
    ByVal colCodeDesc As Long, _
    ByVal colDate As Long, _
    ByVal colAmount As Long, _
    ByVal colDescription As Long, _
    ByRef matched() As Boolean, _
    ByRef ambiguous() As Boolean, _
    ByVal ambiguityPairs As Collection)


    DebugLog "START PassUnmatched"

    Dim i As Long
    Dim unmatchedReason As String
    Dim unmatchedDetails As String

    For i = 2 To lastRow

        If matched(i) Then GoTo NextI

        If ambiguous(i) Then GoTo NextI

        If Not TransferFlags(i) Then GoTo NextI

        unmatchedDetails = ""

        unmatchedReason = _
            DetermineUnmatchedReason( _
                data, _
                i, _
                lastRow, _
                colAcct, _
                colDate, _
                colAmount, _
                colDescription)

        If unmatchedReason = _
           UNMATCH_CONTRADICTION_EXISTS Then

            unmatchedDetails = _
                GetContradictionReasonForRow(i)

        End If

        RegisterTransactionStatus _
            i, _
            STATUS_UNMATCHED

        WriteUnmatchedReason _
            hostWb, _
            i, _
            CStr(data(i, colAcct)), _
            data(i, colDate), _
            CDbl(data(i, colAmount)), _
            unmatchedReason, _
            unmatchedDetails

NextI:
    Next i

    DebugLog "END PassUnmatched"

End Sub


Private Sub AddToIndex( _
    idx As Object, _
    key As String, _
    rowNum As Long)

    Dim rows As Collection

    If Not idx.Exists(key) Then
        Set rows = New Collection
        idx.Add key, rows
    End If

    idx(key).Add rowNum

End Sub

'============================================
'Same account transfer exclusion ---IMPORTANT
'============================================

Private Function IsSameAccountPair( _
    data As Variant, _
    rowA As Long, _
    rowB As Long, _
    colAcct As Long) As Boolean

    IsSameAccountPair = _
        Trim$(CStr(data(rowA, colAcct))) = _
        Trim$(CStr(data(rowB, colAcct)))

End Function


'=========================================================
' Extract Account Suffix from Description
'
' Examples:
' Transfer From DDA X8252
' Transfer To DDA X2489
'
' Returns:
' 8252
' 2489
'=========================================================


Private Function GetTransferSuffixes( _
    txt As String, _
    ByRef fromSuffix As String, _
    ByRef toSuffix As String) As Boolean

    Dim RE As Object
    Dim matches As Object

    Set RE = CreateObject("VBScript.RegExp")

    RE.pattern = _
        "Transfer\s+from\s+X+(\d+)\s+to\s+X+([\d ]+)"

    RE.IgnoreCase = True
    RE.Global = False

    If RE.Test(txt) Then

        Set matches = RE.Execute(txt)

        fromSuffix = matches(0).SubMatches(0)
        toSuffix = matches(0).SubMatches(1)
        
        GetTransferSuffixes = True

    End If
    

End Function

Private Function AccountEndsWith( _
    acct As String, _
    suffix As String) As Boolean

    acct = Trim(acct)

    If Len(acct) >= Len(suffix) Then

        AccountEndsWith = _
            (Right(acct, Len(suffix)) = suffix)

    End If

End Function

'=========================================================
' Counterevidence for matching
'
' Examples:
' Suffix contains X4567, counter account does end 4567
' Referenced account does not equal counter account
' Narrative descriptions do not match
' Embedded description accounts do not match counter account
'
' Returns:
' No match
'=========================================================

Private Function GetContradictionReason( _
    data As Variant, _
    rowA As Long, _
    rowB As Long, _
    colAcct As Long, _
    colDescription As Long) As String

    Dim acctA As String
    Dim acctB As String

    acctA = Trim(CStr(data(rowA, colAcct)))
    acctB = Trim(CStr(data(rowB, colAcct)))

 
    ' Transfer Suffix Contradiction

    Dim fromSuffix As String
    Dim toSuffix As String

    If GetTransferSuffixes( _
        CStr(data(rowA, colDescription)), _
        fromSuffix, _
        toSuffix) Then

        fromSuffix = Replace(fromSuffix, " ", "")
        toSuffix = Replace(toSuffix, " ", "")

        If Not AccountEndsWith(acctA, fromSuffix) _
        And Not AccountEndsWith(acctA, toSuffix) Then

            GetContradictionReason = _
                "Suffix Contradiction"
            
            Exit Function

        End If

        If Not AccountEndsWith(acctB, fromSuffix) _
        And Not AccountEndsWith(acctB, toSuffix) Then

            GetContradictionReason = _
                "Suffix Contradiction"
            
            Exit Function

        End If

    End If


    ' Referenced Account Contradiction


    Dim refAcct As String

    refAcct = ReferencedAccounts(rowA)

    If refAcct <> "" Then

        If acctB <> refAcct _
        And acctA <> refAcct Then

            GetContradictionReason = _
                "Referenced Account Contradiction"
            
            Exit Function

        End If

    End If

    refAcct = ReferencedAccounts(rowB)

    If refAcct <> "" Then

        If acctA <> refAcct _
        And acctB <> refAcct Then

            GetContradictionReason = _
                "Referenced Account Contradiction"
            
            Exit Function

        End If

    End If
    

    ' Narrative Description Contradiction

    Dim narrativeA As String
    Dim narrativeB As String
    
    narrativeA = NarrativeIDs(rowA)
    
    narrativeB = NarrativeIDs(rowB)
    
    If narrativeA <> "" _
    And narrativeB <> "" Then
    
        If narrativeA <> narrativeB Then
    
            GetContradictionReason = _
                "Narrative Contradiction"
            
            Exit Function
    
        End If
    
    End If


    ' Embedded Account Contradiction


    Dim embeddedAcct As String

    embeddedAcct = EmbeddedAccounts(rowA)

    If embeddedAcct <> "" Then

        If acctA <> embeddedAcct _
        And acctB <> embeddedAcct Then

            GetContradictionReason = _
                "Embedded Account Contradiction"
            
            Exit Function

        End If

    End If

    embeddedAcct = EmbeddedAccounts(rowB)

    If embeddedAcct <> "" Then

        If acctA <> embeddedAcct _
        And acctB <> embeddedAcct Then

            GetContradictionReason = _
                "Embedded Account Contradiction"
            
            Exit Function

        End If

    End If
    
GetContradictionReason = ""

End Function

'=========================================================
' Extract embedded account number from Internet Transfer
' This is specifically for CrossFirst pre-conversion transfers
'
' Examples:
'
' FUNDS TRANSFER TO DEP 2011676661 FROM
' FUNDS TRANSFER FROM DEP 2013448206 TO
'
' Returns:
' 2011676661
' 2013448206
'=========================================================

Private Function GetEmbeddedAccount( _
    txt As String) As String

    Dim RE As Object
    Dim matches As Object
    

    Set RE = CreateObject("VBScript.RegExp")

    RE.IgnoreCase = True
    RE.Global = False

    RE.pattern = _
        "TRANSFER\s+(?:TO|FROM|FRM)\s+DEP\s+(\d+)"

    If RE.Test(txt) Then

        Set matches = RE.Execute(txt)

        GetEmbeddedAccount = _
        matches(0).SubMatches(0)

    End If

End Function



Private Function CountEmbeddedCandidates( _
    data As Variant, _
    targetRow As Long, _
    lastRow As Long, _
    colAcct As Long, _
    colCodeDesc As Long, _
    colDate As Long, _
    colAmount As Long, _
    colDescription As Long) As Long

    Dim i As Long

    Dim sourceAcct As String
    Dim sourceEmbedded As String
    Dim sourceAmt As Double
    Dim sourceDate As Variant

    Dim candidateAcct As String
    Dim candidateEmbedded As String
    Dim candidateAmt As Double

    candidateAcct = Trim(CStr(data(targetRow, colAcct)))
    
    candidateEmbedded = EmbeddedAccounts(targetRow)

    candidateAmt = CDbl(data(targetRow, colAmount))

    For i = 2 To lastRow

        If i = targetRow Then GoTo NextI
        
        If IsSameAccountPair(data, targetRow, i, colAcct) Then GoTo NextI

        If Not TransferFlags(i) Then

            GoTo NextI
        End If

        If InStr(1, _
            CStr(data(i, colCodeDesc)), _
            "Internet Transfer", _
            vbTextCompare) = 0 Then
            GoTo NextI
        End If

        sourceAcct = Trim(CStr(data(i, colAcct)))

        sourceEmbedded = EmbeddedAccounts(i)

        sourceAmt = CDbl(data(i, colAmount))

        sourceDate = data(i, colDate)

        If sourceEmbedded <> candidateAcct Then GoTo NextI
        If candidateEmbedded <> sourceAcct Then GoTo NextI

        If Abs(sourceAmt) <> Abs(candidateAmt) Then GoTo NextI
        If sourceAmt * candidateAmt >= 0 Then GoTo NextI

        If sourceDate <> data(targetRow, colDate) Then GoTo NextI
        

        CountEmbeddedCandidates = CountEmbeddedCandidates + 1

NextI:
    Next i

End Function

Private Function CountAmountDateCandidates( _
    data As Variant, _
    targetRow As Long, _
    lastRow As Long, _
    colAcct As Long, _
    colCodeDesc As Long, _
    colDate As Long, _
    colAmount As Long, _
    colDescription As Long, _
    matched() As Boolean) As Long

    Dim v As Variant
    Dim i As Long
    
    Dim targetAmt As Double
    Dim candidateAmt As Double

    Dim targetDate As Variant

    targetAmt = CDbl(data(targetRow, colAmount))
    targetDate = data(targetRow, colDate)

    Dim amountDateKey As String
        
    amountDateKey = _
        Format(targetDate, "yyyymmdd") & "|" & _
        Format(Abs(targetAmt), "0.00")
    
    If Not AmtDateIndex.Exists(amountDateKey) Then Exit Function
    
    For Each v In AmtDateIndex(amountDateKey)
    
        i = CLng(v)

        If i = targetRow Then GoTo NextI
        '"If matched(i) Then GoTo NextI" was a problem, DO NOT CHANGE, it finally works
        
        If IsSameAccountPair(data, targetRow, i, colAcct) Then GoTo NextI

        If Not TransferFlags(i) Then
        
            GoTo NextI
        End If

        candidateAmt = CDbl(data(i, colAmount))

        If Abs(targetAmt) <> Abs(candidateAmt) Then _
            GoTo NextI

        If targetAmt * candidateAmt >= 0 Then _
            GoTo NextI

        If targetDate <> data(i, colDate) Then _
            GoTo NextI

        If GetContradictionReason( _
            data, _
            targetRow, _
            i, _
            colAcct, _
            colDescription) <> "" Then
        
            GoTo NextI
        
        End If
        
        If HasUnresolvedAccountEvidence( _
            data, _
            targetRow, _
            lastRow, _
            colAcct, _
            colDescription) Then
        
            GoTo NextI
        
        End If
        
        If HasUnresolvedAccountEvidence( _
            data, _
            i, _
            lastRow, _
            colAcct, _
            colDescription) Then
        
            GoTo NextI
        
        End If
        
        CountAmountDateCandidates = _
            CountAmountDateCandidates + 1

NextI:
    Next v

End Function

Private Function CountSuffixCandidates( _
    data As Variant, _
    targetRow As Long, _
    lastRow As Long, _
    colAcct As Long, _
    colDate As Long, _
    colAmount As Long, _
    colDescription As Long, _
    matched() As Boolean) As Long

    Dim i As Long

    Dim sourceAcct As String
    Dim candidateAcct As String

    Dim sourceAmt As Double
    Dim candidateAmt As Double

    Dim sourceDate As Variant

    Dim fromSuffix As String
    Dim toSuffix As String

    If Not GetTransferSuffixes( _
        CStr(data(targetRow, colDescription)), _
        fromSuffix, _
        toSuffix) Then Exit Function

    fromSuffix = Replace(fromSuffix, " ", "")
    toSuffix = Replace(toSuffix, " ", "")

    sourceAcct = Trim(CStr(data(targetRow, colAcct)))
    sourceAmt = CDbl(data(targetRow, colAmount))
    sourceDate = data(targetRow, colDate)

    For i = 2 To lastRow

        If i = targetRow Then GoTo NextI
        If matched(i) Then GoTo NextI
        
        If IsSameAccountPair(data, targetRow, i, colAcct) Then GoTo NextI

        candidateAcct = _
            Trim(CStr(data(i, colAcct)))

        candidateAmt = _
            CDbl(data(i, colAmount))

        If Abs(sourceAmt) <> _
           Abs(candidateAmt) Then GoTo NextI

        If sourceAmt * candidateAmt >= 0 Then GoTo NextI

        If sourceDate <> _
           data(i, colDate) Then GoTo NextI

        If sourceAmt > 0 Then

            If AccountEndsWith(sourceAcct, toSuffix) _
            And AccountEndsWith(candidateAcct, fromSuffix) Then

                CountSuffixCandidates = _
                    CountSuffixCandidates + 1

            End If

        Else

            If AccountEndsWith(sourceAcct, fromSuffix) _
            And AccountEndsWith(candidateAcct, toSuffix) Then

                CountSuffixCandidates = _
                    CountSuffixCandidates + 1

            End If

        End If

NextI:
    Next i

End Function

Private Function HasUnresolvedAccountEvidence( _
    data As Variant, _
    rowNum As Long, _
    lastRow As Long, _
    colAcct As Long, _
    colDescription As Long) As Boolean

    Dim acctToFind As String
    Dim i As Long

    Dim fromSuffix As String
    Dim toSuffix As String

    ' Transfer Suffix Evidence

    If GetTransferSuffixes( _
        CStr(data(rowNum, colDescription)), _
        fromSuffix, _
        toSuffix) Then

        If Not SuffixExistsInData( _
            data, _
            lastRow, _
            colAcct, _
            fromSuffix) _
        Or _
           Not SuffixExistsInData( _
            data, _
            lastRow, _
            colAcct, _
            toSuffix) Then
            
            HasUnresolvedAccountEvidence = True
            Exit Function

        End If

    End If


    ' Referenced Account Evidence

    acctToFind = GetReferencedAccount( _
        CStr(data(rowNum, colDescription)))

    If acctToFind <> "" Then

        If Not AccountExistsInData( _
            data, _
            lastRow, _
            colAcct, _
            acctToFind) Then
            
            HasUnresolvedAccountEvidence = True
            Exit Function

        End If

    End If


    ' Embedded Account Evidence

    acctToFind = GetEmbeddedAccount( _
        CStr(data(rowNum, colDescription)))

    If acctToFind <> "" Then

        If Not AccountExistsInData( _
            data, _
            lastRow, _
            colAcct, _
            acctToFind) Then
            
            HasUnresolvedAccountEvidence = True
            Exit Function

        End If

    End If

End Function



' Determine if account is in spreadsheet

Private Function AccountExistsInData( _
    data As Variant, _
    lastRow As Long, _
    colAcct As Long, _
    acctNumber As String) As Boolean

    Dim i As Long

    For i = 2 To lastRow

        If Trim(CStr(data(i, colAcct))) = _
           Trim(acctNumber) Then

            AccountExistsInData = True
            Exit Function

        End If

    Next i

End Function


' Determine if account suffix is in spreadsheet

Private Function SuffixExistsInData( _
    data As Variant, _
    lastRow As Long, _
    colAcct As Long, _
    suffix As String) As Boolean

    Dim i As Long

    suffix = Replace(suffix, " ", "")

    For i = 2 To lastRow

        If AccountEndsWith( _
            Trim(CStr(data(i, colAcct))), _
            suffix) Then

            SuffixExistsInData = True
            Exit Function

        End If

    Next i

End Function

Private Function SuffixMatchesAccount( _
    debitRow As Long, _
    creditAcct As String) As Boolean

    Dim suffix As String

    If Not TransferToSuffixes.Exists(debitRow) Then

        SuffixMatchesAccount = True
        Exit Function

    End If

    suffix = _
        CStr(TransferToSuffixes(debitRow))

    SuffixMatchesAccount = _
        (Right(creditAcct, Len(suffix)) = suffix)

End Function


'===============
'Debug Helpers
'===============

Private Sub DebugLog(msg As String)

    If DEBUG_MODE Then
        Debug.Print Format(Now, "hh:mm:ss") & " | " & msg
    End If
    
End Sub

Private Sub DebugAmbiguous(msg As String)

    If DEBUG_AMBIGUITY Then
    
        Debug.Print Format(Now, "hh:mm:ss") & " | " & msg
        
    End If
    
End Sub

Private Function BuildClusterMemberLookup() As Object

    Dim ws As Worksheet

    Dim clusters As Object
    Dim members As Collection

    Dim lastRow As Long
    Dim r As Long

    Dim clusterID As String
    Dim rowNumber As Long

    Set ws = hostWb.Worksheets("Transfer_Ambiguity_Members")

    Set clusters = CreateObject("Scripting.Dictionary")

    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow

        clusterID = _
            CStr(ws.Cells(r, 2).Value)

        rowNumber = _
            CLng(ws.Cells(r, 3).Value)

        If Not clusters.Exists(clusterID) Then

            Set members = New Collection

            clusters.Add _
                clusterID, _
                members

        End If

        clusters(clusterID).Add rowNumber

    Next r

    Set BuildClusterMemberLookup = clusters

End Function



'Ambiguous Clustering - Working correctly, do not change!!


Private Sub BuildAmbiguousClusters( _
    ambiguityPairs As Collection)

    Dim graph As Object
    Dim visited As Object

    Dim pairText As Variant
    Dim node As Variant

    Dim parts() As String

    Dim rowA As Long
    Dim rowB As Long

    Dim clusterID As Long

    Set graph = CreateObject("Scripting.Dictionary")
    Set visited = CreateObject("Scripting.Dictionary")
    Set AmbiguousClusters = CreateObject("Scripting.Dictionary")
    Set ClusterSizes = CreateObject("Scripting.Dictionary")

    DebugAmbiguous "START BuildAmbiguousClusters"
    ' Build Graph

    For Each pairText In ambiguityPairs
    
    
        parts = Split(CStr(pairText), "|")
        
        rowA = CLng(parts(0))
        rowB = CLng(parts(1))

        AddGraphEdge graph, rowA, rowB
        AddGraphEdge graph, rowB, rowA
        


    Next pairText


    ' Find Clusters

    clusterID = 1

    For Each node In graph.Keys
    
        If Not visited.Exists(node) Then

            ExploreCluster _
                graph, _
                visited, _
                CLng(node), _
                clusterID
                

            clusterID = clusterID + 1

        End If

    Next node
    
    DebugAmbiguous "END BuildAmbiguousClusters"

End Sub
    

'Ambiguous Graphing Helper

Private Sub AddGraphEdge( _
    graph As Object, _
    nodeA As Long, _
    nodeB As Long)

    If Not graph.Exists(nodeA) Then

        graph.Add nodeA, _
            CreateObject("Scripting.Dictionary")

    End If

    graph(nodeA)(nodeB) = 1

End Sub

Private Function BuildRelationshipsFromEdges( _
    ByVal hostWb As Workbook, _
    ByVal members As Collection) As Object

    Dim relDict As Object
    Dim memberRows As Object

    Dim edgeKey As Variant
    Dim parts() As String

    Dim rowA As Long
    Dim rowB As Long

    Dim txnRowA As Long
    Dim txnRowB As Long

    Dim acctA As String
    Dim acctB As String

    Dim relKey As String

    Dim txnWs As Worksheet
    Dim rowNum As Variant

    Set txnWs = _
        hostWb.Worksheets("Transfer_Transactions")

    Set relDict = CreateObject("Scripting.Dictionary")
    Set memberRows = CreateObject("Scripting.Dictionary")

    ' Build lookup of rows belonging to this cluster
    For Each rowNum In members
    
        memberRows(CStr(rowNum)) = True

    Next rowNum

    ' Process only edges that belong to this cluster
    For Each edgeKey In AmbiguousEdges.Keys

        parts = Split(CStr(edgeKey), "|")

        rowA = CLng(parts(0))
        rowB = CLng(parts(1))

        If memberRows.Exists(CStr(rowA)) _
        And memberRows.Exists(CStr(rowB)) Then

            txnRowA = FindTransactionRow( _
                hostWb, _
                rowA)
            txnRowB = FindTransactionRow( _
                hostWb, _
                rowB)

            If txnRowA > 0 _
            And txnRowB > 0 Then

                acctA = CStr(txnWs.Cells(txnRowA, 3).Value)
                acctB = CStr(txnWs.Cells(txnRowB, 3).Value)
                

                ' Normalize order so A|B and B|A collapse, no idea what I'm doing here anymore
                
                If StrComp(acctA, acctB, vbTextCompare) < 0 Then

                    relKey = acctA & "|" & acctB

                Else

                    relKey = acctB & "|" & acctA

                End If

                If Not relDict.Exists(relKey) Then

                    relDict.Add relKey, 0

                End If

                relDict(relKey) = _
                    CLng(relDict(relKey)) + 1

            End If
            

        End If
        
Next edgeKey

    Set BuildRelationshipsFromEdges = relDict
    
    For Each edgeKey In AmbiguousEdges.Keys

    parts = Split(edgeKey, "|")


Next edgeKey

End Function


'DFS Traversal - no idea what I'm doing at this point

Private Sub ExploreCluster( _
    graph As Object, _
    visited As Object, _
    startNode As Long, _
    clusterID As Long)

    Dim stack As Collection

    Dim node As Long
    Dim nbr As Variant
    
    Dim clusterName As String
    Dim size As Long

    clusterName = _
        "C" & Format(clusterID, "000000")

    Set stack = New Collection

    stack.Add startNode

    Do While stack.Count > 0

        node = stack(stack.Count)

        stack.Remove stack.Count

        If visited.Exists(node) Then GoTo NextNode

        visited(node) = True

                       
            AmbiguousClusters(CStr(node)) = _
                clusterName
                
            size = size + 1

        For Each nbr In graph(node).Keys

            If Not visited.Exists(nbr) Then

                stack.Add CLng(nbr)

            End If

        Next nbr
        
        ClusterSizes(clusterName) = size
        
NextNode:

    Loop

End Sub


'Ambiguous resolution - Do not touch!

Private Sub ResolveAmbiguities( _
    ByVal hostWb As Workbook, _
    ByVal ws As Worksheet, _
    ByRef data As Variant, _
    ByVal lastRow As Long, _
    ByVal colAcct As Long, _
    ByVal colCodeDesc As Long, _
    ByVal colDate As Long, _
    ByVal colAmount As Long, _
    ByVal colDescription As Long, _
    ByRef matched() As Boolean, _
    ByRef ambiguous() As Boolean, _
    ByVal ambiguityPairs As Collection)

    Dim pairText As Variant
    Dim parts() As String

    Dim rowA As Long
    Dim rowB As Long

    Dim matchMethod As String

    Dim lowRow As Long
    Dim highRow As Long
    Dim pairKey As String

    Dim seenPairs As Object
    Dim pairMethods As Object

    Dim key As Variant
    Dim clusterID As String
    
    DebugLog "START ResolveAmbiguities"
    DebugAmbiguous "START ResolveAmbiguities"

    Set seenPairs = CreateObject("Scripting.Dictionary")
    Set pairMethods = CreateObject("Scripting.Dictionary")

    For Each pairText In ambiguityPairs

        parts = Split(CStr(pairText), "|")

        rowA = CLng(parts(0))
        rowB = CLng(parts(1))

        If UBound(parts) >= 2 Then
            matchMethod = parts(2)
        Else
            matchMethod = ""
        End If

        lowRow = rowA
        highRow = rowB

        If lowRow > highRow Then
            lowRow = rowB
            highRow = rowA
        End If

        pairKey = CStr(lowRow) & "|" & CStr(highRow)

        If Not seenPairs.Exists(pairKey) Then

            seenPairs.Add pairKey, pairKey

            pairMethods.Add _
                pairKey, _
                CreateObject("Scripting.Dictionary")

        End If

        If matchMethod <> "" Then

            RecordAmbiguityMethod matchMethod

            If Not pairMethods(pairKey).Exists(matchMethod) Then

                pairMethods(pairKey).Add _
                    matchMethod, _
                    1

            Else

                pairMethods(pairKey)(matchMethod) = _
                    pairMethods(pairKey)(matchMethod) + 1

            End If

        End If

    Next pairText

    Dim methodText As String
    Dim methodKey As Variant

    Call BuildAmbiguousClusters(ambiguityPairs)
    
    Call CalculateClusterValues(data, colAmount)

    PersistAmbiguityWarehouse _
        hostWb, _
        data, _
        colAcct, _
        colDate, _
        colAmount
        

    AnalyzeClusters _
        hostWb, _
        data, _
        colAcct, _
        colDate, _
        colAmount, _
        colDescription, _
        ambiguityPairs
        
       

    ResolveRepeatedTransferMatches _
        hostWb, _
        data, _
        colAcct, _
        colCodeDesc, _
        colDate, _
        colAmount, _
        colDescription
        

    For Each key In seenPairs.Keys

        parts = Split(CStr(key), "|")
    
        rowA = CLng(parts(0))
        rowB = CLng(parts(1))
    
        ambiguous(rowA) = True
        ambiguous(rowB) = True
    
        If AmbiguousClusters.Exists(CStr(rowA)) Then
    
            clusterID = AmbiguousClusters(CStr(rowA))
            
    
            RegisterTransactionStatus _
                rowA, _
                STATUS_AMBIGUOUS, _
                clusterID
    
            RegisterTransactionStatus _
                rowB, _
                STATUS_AMBIGUOUS, _
                clusterID
    
        Else
    
    
        End If
    
    Next key

    For Each key In seenPairs.Keys

        parts = Split(CStr(key), "|")

        rowA = CLng(parts(0))
        rowB = CLng(parts(1))

        methodText = ""

        For Each methodKey In pairMethods(key).Keys

            If methodText <> "" Then
                methodText = methodText & "; "
            End If

            methodText = methodText & _
                methodKey & _
                " (" & _
                pairMethods(key)(methodKey) & _
                ")"

        Next methodKey
        
            If AmbiguousClusters.Exists(CStr(rowA)) Then
            
                clusterID = _
                    AmbiguousClusters(CStr(rowA))
                    
           
                ClusterOriginMethods(clusterID) = _
                    methodText
            
            End If
            
    Next key
   
   DebugAmbiguous "END ResolveAmbiguities"
   DebugLog "END ResolveAmbiguities"
   DebugLog String(59, "=")

End Sub

Private Function BuildFullClusterMemberLookup( _
    ByVal hostWb As Workbook) As Object

    Dim clusters As Object
    Dim rowNum As Variant
    Dim clusterID As String

    Dim members As Collection

    Set clusters = _
        CreateObject("Scripting.Dictionary")

    If AmbiguousClusters Is Nothing Then

        Err.Raise vbObjectError + 1080, _
                  "BuildFullClusterMemberLookup", _
                  "AmbiguousClusters has not been initialized."

    End If

    For Each rowNum In AmbiguousClusters.Keys

        clusterID = _
            Trim$(CStr(AmbiguousClusters(rowNum)))

        If Len(clusterID) > 0 Then

            If Not clusters.Exists(clusterID) Then

                Set members = New Collection

                clusters.Add _
                    clusterID, _
                    members

            Else

                Set members = clusters(clusterID)

            End If

            members.Add CLng(rowNum)

        Else

            Debug.Print _
                "BuildFullClusterMemberLookup: " & _
                "Blank Cluster ID for source row " & _
                CStr(rowNum)

        End If

    Next rowNum


    Set BuildFullClusterMemberLookup = _
        clusters

End Function


'Cluster Value Calculation


Private Sub CalculateClusterValues( _
    data As Variant, _
    colAmount As Long)

    Dim rowNum As Variant
    Dim clusterID As String

    Set ClusterValues = _
        CreateObject("Scripting.Dictionary")
        
    DebugAmbiguous "START CalculateClusterValues"

    For Each rowNum In AmbiguousClusters.Keys

        clusterID = _
            AmbiguousClusters(rowNum)

        If Not ClusterValues.Exists(clusterID) Then

            ClusterValues.Add clusterID, 0#

        End If

        ClusterValues(clusterID) = _
            ClusterValues(clusterID) + _
            Abs(CDbl(data(CLng(rowNum), colAmount)))

    Next rowNum
    
    DebugAmbiguous "END CalculateClusterValues"

End Sub


'Largest Cluster Calculation - Size


Private Function GetLargestClusterSizeID() _
    As String

    Dim k As Variant
    Dim maxSize As Long

    For Each k In ClusterSizes.Keys

        If ClusterSizes(k) > maxSize Then

            maxSize = ClusterSizes(k)

            GetLargestClusterSizeID = k

        End If

    Next k

End Function


'Largest Cluster Calculation - Value


Private Function GetLargestClusterValueID() _
    As String

    Dim k As Variant
    Dim maxValue As Double

    For Each k In ClusterValues.Keys

        If ClusterValues(k) > maxValue Then

            maxValue = ClusterValues(k)

            GetLargestClusterValueID = k

        End If

    Next k

End Function

Private Function GetClusterSize( _
    clusterID As String) As Long

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = hostWb.Worksheets("Transfer_Ambiguities")

    lastRow = _
        ws.Cells(ws.rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow

        If ws.Cells(r, 2).Value = clusterID Then

            GetClusterSize = _
                CLng(ws.Cells(r, 3).Value)

            Exit Function

        End If

    Next r

End Function

Private Function GetClusterExposure( _
    clusterID As String) As Double

    Dim ws As Worksheet
    Dim r As Long
    Dim lastRow As Long

    Set ws = hostWb.Worksheets("Transfer_Ambiguities")

    lastRow = _
        ws.Cells(ws.rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow

        If ws.Cells(r, 2).Value = clusterID Then

            GetClusterExposure = _
                CDbl(ws.Cells(r, 4).Value)

            Exit Function

        End If

    Next r

End Function

Private Function GetLargestClusterExposure() As Double

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim maxExposure As Double

    Set ws = hostWb.Worksheets("Transfer_Ambiguities")

    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow

        If CDbl(ws.Cells(r, 4).Value) > maxExposure Then
            maxExposure = CDbl(ws.Cells(r, 4).Value)
        End If

    Next r

    GetLargestClusterExposure = maxExposure

End Function


Private Function GetAmbiguousExposure() As Double

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = hostWb.Worksheets("Transfer_Ambiguities")

    lastRow = ws.Cells( _
        ws.rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow

        GetAmbiguousExposure = _
            GetAmbiguousExposure + _
            CDbl(ws.Cells(r, 4).Value)

    Next r

End Function

Private Function GetLargestClusterByExposure() As String

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Dim maxExposure As Double
    Dim clusterID As String

    Set ws = hostWb.Worksheets("Transfer_Ambiguities")

    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow

        If CDbl(ws.Cells(r, 4).Value) > maxExposure Then

            maxExposure = CDbl(ws.Cells(r, 4).Value)
            clusterID = CStr(ws.Cells(r, 2).Value)

        End If

    Next r

    GetLargestClusterByExposure = clusterID

End Function

Private Function GetLargestClusterSize() As Long

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim maxSize As Long

    Set ws = hostWb.Worksheets("Transfer_Ambiguities")

    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow

        If CLng(ws.Cells(r, 3).Value) > maxSize Then
            maxSize = CLng(ws.Cells(r, 3).Value)
        End If

    Next r

    GetLargestClusterSize = maxSize

End Function

Private Function GetLargestClusterAccounts() As Long

    Dim ws As Worksheet
    Dim clusterDict As Object
    Dim accountDict As Object

    Dim clusterID As String
    Dim acct As String

    Dim lastRow As Long
    Dim r As Long

    Dim key As Variant

    Set ws = hostWb.Worksheets("Transfer_Ambiguity_Members")
    Set clusterDict = CreateObject("Scripting.Dictionary")

    lastRow = _
        ws.Cells(ws.rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow

        clusterID = CStr(ws.Cells(r, 2).Value)
        acct = CStr(ws.Cells(r, 4).Value)

        If Not clusterDict.Exists(clusterID) Then

            Set accountDict = _
                CreateObject("Scripting.Dictionary")

            clusterDict.Add _
                clusterID, _
                accountDict

        End If

        clusterDict(clusterID)(acct) = 1

    Next r

    For Each key In clusterDict.Keys

        If clusterDict(key).Count > _
           GetLargestClusterAccounts Then

            GetLargestClusterAccounts = _
                clusterDict(key).Count

        End If

    Next key

End Function

Private Function GetAmbiguousInvestigationGroupCount() _
    As Long

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = hostWb.Worksheets( _
        "Investigation_Group_Metrics")

    lastRow = _
        ws.Cells(ws.rows.Count, 1) _
          .End(xlUp).Row

    For r = 2 To lastRow

        Select Case _
            UCase(CStr(ws.Cells(r, 2).Value))

            Case "REVIEW", _
                 "UNBALANCED"

                GetAmbiguousInvestigationGroupCount = _
                    GetAmbiguousInvestigationGroupCount + 1

        End Select

    Next r

End Function

Private Function GetAmbiguousTransferCount() _
    As Long

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = hostWb.Worksheets( _
        "Investigation_Group_Metrics")

    lastRow = _
        ws.Cells(ws.rows.Count, 1) _
          .End(xlUp).Row

    For r = 2 To lastRow

        Select Case _
            UCase(CStr(ws.Cells(r, 2).Value))

            Case "REVIEW", _
                 "UNBALANCED"

                GetAmbiguousTransferCount = _
                    GetAmbiguousTransferCount + _
                    CLng(ws.Cells(r, 4).Value)

        End Select

    Next r

End Function

Private Function _
    GetAmbiguousInvestigationExposure() _
    As Double

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = hostWb.Worksheets( _
        "Investigation_Group_Metrics")

    lastRow = _
        ws.Cells(ws.rows.Count, 1) _
          .End(xlUp).Row

    For r = 2 To lastRow

        Select Case _
            UCase(CStr(ws.Cells(r, 2).Value))

            Case "REVIEW", _
                 "UNBALANCED"

                GetAmbiguousInvestigationExposure = _
                    GetAmbiguousInvestigationExposure + _
                    CDbl(ws.Cells(r, 6).Value)

        End Select

    Next r

End Function

Private Function _
    GetLargestInvestigationGroupExposure() _
    As Double

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = hostWb.Worksheets( _
        "Investigation_Group_Metrics")

    lastRow = _
        ws.Cells(ws.rows.Count, 1) _
          .End(xlUp).Row

    For r = 2 To lastRow

        If UCase( _
            CStr(ws.Cells(r, 2).Value)) <> _
            "RESOLVED" Then

            If CDbl(ws.Cells(r, 6).Value) > _
               GetLargestInvestigationGroupExposure Then

                GetLargestInvestigationGroupExposure = _
                    CDbl(ws.Cells(r, 6).Value)

            End If

        End If

    Next r

End Function

Private Function _
    GetLargestResolvedGroupExposure() _
    As Double

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = hostWb.Worksheets( _
        "Investigation_Group_Metrics")

    lastRow = _
        ws.Cells(ws.rows.Count, 1) _
          .End(xlUp).Row

    For r = 2 To lastRow

        If UCase( _
            CStr(ws.Cells(r, 2).Value)) = _
            "RESOLVED" Then

            If CDbl(ws.Cells(r, 6).Value) > _
               GetLargestResolvedGroupExposure Then

                GetLargestResolvedGroupExposure = _
                    CDbl(ws.Cells(r, 6).Value)

            End If

        End If

    Next r

End Function

Private Function _
    GetLargestInvestigationAccounts() _
    As Long

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = hostWb.Worksheets( _
        "Investigation_Group_Metrics")

    lastRow = _
        ws.Cells(ws.rows.Count, 1) _
          .End(xlUp).Row

    For r = 2 To lastRow

        If UCase( _
            CStr(ws.Cells(r, 2).Value)) <> _
            "RESOLVED" Then

            If CLng(ws.Cells(r, 5).Value) > _
               GetLargestInvestigationAccounts Then

                GetLargestInvestigationAccounts = _
                    CLng(ws.Cells(r, 5).Value)

            End If

        End If

    Next r

End Function

Private Function _
    GetLargestResolvedAccounts() _
    As Long

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = hostWb.Worksheets( _
        "Investigation_Group_Metrics")

    lastRow = _
        ws.Cells(ws.rows.Count, 1) _
          .End(xlUp).Row

    For r = 2 To lastRow

        If UCase( _
            CStr(ws.Cells(r, 2).Value)) = _
            "RESOLVED" Then

            If CLng(ws.Cells(r, 5).Value) > _
               GetLargestResolvedAccounts Then

                GetLargestResolvedAccounts = _
                    CLng(ws.Cells(r, 5).Value)

            End If

        End If

    Next r

End Function

Private Function _
    GetLargestInvestigationAccountsID() _
    As String

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Dim maxCount As Long

    Set ws = hostWb.Worksheets( _
        "Investigation_Group_Metrics")

    lastRow = _
        ws.Cells(ws.rows.Count, 1) _
          .End(xlUp).Row

    For r = 2 To lastRow

        If UCase( _
            CStr(ws.Cells(r, 2).Value)) <> _
            "RESOLVED" Then

            If CLng(ws.Cells(r, 5).Value) > _
               maxCount Then

                maxCount = _
                    CLng(ws.Cells(r, 5).Value)

                GetLargestInvestigationAccountsID = _
                    CStr(ws.Cells(r, 1).Value)

            End If

        End If

    Next r

End Function

Private Function _
    GetLargestResolvedAccountsID() _
    As String

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Dim maxCount As Long

    Set ws = hostWb.Worksheets( _
        "Investigation_Group_Metrics")

    lastRow = _
        ws.Cells(ws.rows.Count, 1) _
          .End(xlUp).Row

    For r = 2 To lastRow

        If UCase( _
            CStr(ws.Cells(r, 2).Value)) = _
            "RESOLVED" Then

            If CLng(ws.Cells(r, 5).Value) > _
               maxCount Then

                maxCount = _
                    CLng(ws.Cells(r, 5).Value)

                GetLargestResolvedAccountsID = _
                    CStr(ws.Cells(r, 1).Value)

            End If

        End If

    Next r

End Function

Private Function _
    GetLargestInvestigationGroupSize() _
    As Long

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = hostWb.Worksheets( _
        "Investigation_Group_Metrics")

    lastRow = _
        ws.Cells(ws.rows.Count, 1) _
          .End(xlUp).Row

    For r = 2 To lastRow

        If UCase( _
            CStr(ws.Cells(r, 2).Value)) <> _
            "RESOLVED" Then

            If CLng(ws.Cells(r, 4).Value) > _
               GetLargestInvestigationGroupSize Then

                GetLargestInvestigationGroupSize = _
                    CLng(ws.Cells(r, 4).Value)

            End If

        End If

    Next r

End Function

Private Function _
    GetLargestResolvedGroupSize() _
    As Long

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = hostWb.Worksheets( _
        "Investigation_Group_Metrics")

    lastRow = _
        ws.Cells(ws.rows.Count, 1) _
          .End(xlUp).Row

    For r = 2 To lastRow

        If UCase( _
            CStr(ws.Cells(r, 2).Value)) = _
            "RESOLVED" Then

            If CLng(ws.Cells(r, 4).Value) > _
               GetLargestResolvedGroupSize Then

                GetLargestResolvedGroupSize = _
                    CLng(ws.Cells(r, 4).Value)

            End If

        End If

    Next r

End Function

Private Function _
    GetLargestInvestigationGroupSizeID() _
    As String

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Dim maxSize As Long

    Set ws = hostWb.Worksheets( _
        "Investigation_Group_Metrics")

    lastRow = _
        ws.Cells(ws.rows.Count, 1) _
          .End(xlUp).Row

    For r = 2 To lastRow

        If UCase( _
            CStr(ws.Cells(r, 2).Value)) <> _
            "RESOLVED" Then

            If CLng(ws.Cells(r, 4).Value) > _
               maxSize Then

                maxSize = _
                    CLng(ws.Cells(r, 4).Value)

                GetLargestInvestigationGroupSizeID = _
                    CStr(ws.Cells(r, 1).Value)

            End If

        End If

    Next r

End Function

Private Function _
    GetLargestResolvedGroupSizeID() _
    As String

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Dim maxSize As Long

    Set ws = hostWb.Worksheets( _
        "Investigation_Group_Metrics")

    lastRow = _
        ws.Cells(ws.rows.Count, 1) _
          .End(xlUp).Row

    For r = 2 To lastRow

        If UCase( _
            CStr(ws.Cells(r, 2).Value)) = _
            "RESOLVED" Then

            If CLng(ws.Cells(r, 4).Value) > _
               maxSize Then

                maxSize = _
                    CLng(ws.Cells(r, 4).Value)

                GetLargestResolvedGroupSizeID = _
                    CStr(ws.Cells(r, 1).Value)

            End If

        End If

    Next r

End Function



Private Function _
    GetLargestInvestigationGroupExposureID() _
    As String

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Dim maxExposure As Double

    Set ws = hostWb.Worksheets( _
        "Investigation_Group_Metrics")

    lastRow = _
        ws.Cells(ws.rows.Count, 1) _
          .End(xlUp).Row

    For r = 2 To lastRow

        If UCase( _
            CStr(ws.Cells(r, 2).Value)) <> _
            "RESOLVED" Then

            If CDbl(ws.Cells(r, 6).Value) > _
               maxExposure Then

                maxExposure = _
                    CDbl(ws.Cells(r, 6).Value)

                GetLargestInvestigationGroupExposureID = _
                    CStr(ws.Cells(r, 1).Value)

            End If

        End If

    Next r

End Function

Private Function _
    GetLargestResolvedGroupExposureID() _
    As String

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Dim maxExposure As Double

    Set ws = hostWb.Worksheets( _
        "Investigation_Group_Metrics")

    lastRow = _
        ws.Cells(ws.rows.Count, 1) _
          .End(xlUp).Row

    For r = 2 To lastRow

        If UCase( _
            CStr(ws.Cells(r, 2).Value)) = _
            "RESOLVED" Then

            If CDbl(ws.Cells(r, 6).Value) > _
               maxExposure Then

                maxExposure = _
                    CDbl(ws.Cells(r, 6).Value)

                GetLargestResolvedGroupExposureID = _
                    CStr(ws.Cells(r, 1).Value)

            End If

        End If

    Next r

End Function

Private Function _
    GetInvestigationGroupSize( _
        investigationGroup As String) _
    As Long

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = hostWb.Worksheets( _
        "Investigation_Group_Metrics")

    lastRow = _
        ws.Cells(ws.rows.Count, 1) _
          .End(xlUp).Row

    For r = 2 To lastRow

        If CStr(ws.Cells(r, 1).Value) = _
           investigationGroup Then

            GetInvestigationGroupSize = _
                CLng(ws.Cells(r, 4).Value)

            Exit Function

        End If

    Next r

End Function

Private Function _
    GetResolvedGroupSize( _
        resolvedGroup As String) _
    As Long

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = hostWb.Worksheets( _
        "Investigation_Group_Metrics")

    lastRow = _
        ws.Cells(ws.rows.Count, 1) _
          .End(xlUp).Row

    For r = 2 To lastRow

        If CStr(ws.Cells(r, 1).Value) = _
           resolvedGroup Then

            GetResolvedGroupSize = _
                CLng(ws.Cells(r, 4).Value)

            Exit Function

        End If

    Next r

End Function

Private Function _
    GetInvestigationGroupExposure( _
        investigationGroup As String) _
    As Double

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = hostWb.Worksheets( _
        "Investigation_Group_Metrics")

    lastRow = _
        ws.Cells(ws.rows.Count, 1) _
          .End(xlUp).Row

    For r = 2 To lastRow

        If CStr(ws.Cells(r, 1).Value) = _
           investigationGroup Then

            GetInvestigationGroupExposure = _
                CDbl(ws.Cells(r, 6).Value)

            Exit Function

        End If

    Next r

End Function

Private Function _
    GetResolvedGroupExposure( _
        resolvedGroup As String) _
    As Double

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = hostWb.Worksheets( _
        "Investigation_Group_Metrics")

    lastRow = _
        ws.Cells(ws.rows.Count, 1) _
          .End(xlUp).Row

    For r = 2 To lastRow

        If CStr(ws.Cells(r, 1).Value) = _
           resolvedGroup Then

            GetResolvedGroupExposure = _
                CDbl(ws.Cells(r, 6).Value)

            Exit Function

        End If

    Next r

End Function


Private Function BuildRelationshipCounts( _
    ByRef hostWb As Workbook, _
    ByRef members As Collection, _
    ByRef txnDate As Variant, _
    ByRef transferAmount As Double, _
    ByRef warningText As String) As Object

    Dim txnWs As Worksheet

    Dim rowNum As Variant
    Dim txnRow As Long

    Dim acct As String
    Dim amt As Double
    Dim txnDt As Variant

    Dim debitAccounts As Object
    Dim creditAccounts As Object

    Dim relDict As Object

    Dim debitAcct As Variant
    Dim creditAcct As Variant

    Dim supportCount As Long
    Dim key As String

    Dim debitTotal As Long
    Dim creditTotal As Long

    Dim dateDict As Object

    Set txnWs = hostWb.Worksheets("Transfer_Transactions")

    Set debitAccounts = CreateObject("Scripting.Dictionary")
    Set creditAccounts = CreateObject("Scripting.Dictionary")
    Set relDict = CreateObject("Scripting.Dictionary")
    Set dateDict = CreateObject("Scripting.Dictionary")

    For Each rowNum In members

        txnRow = FindTransactionRow( _
            hostWb, _
            CLng(rowNum))

        If txnRow > 0 Then

            acct = CStr(txnWs.Cells(txnRow, 3).Value)
            txnDt = txnWs.Cells(txnRow, 4).Value
            amt = CDbl(txnWs.Cells(txnRow, 5).Value)

            If IsEmpty(txnDate) Then
                txnDate = txnDt
            End If

            transferAmount = Abs(amt)

            If Not dateDict.Exists(CStr(txnDt)) Then
                dateDict.Add CStr(txnDt), True
            End If

            If amt > 0 Then

                creditTotal = creditTotal + 1

                If Not creditAccounts.Exists(acct) Then
                    creditAccounts.Add acct, 0
                End If

                creditAccounts(acct) = _
                    creditAccounts(acct) + 1

            Else

                debitTotal = debitTotal + 1

                If Not debitAccounts.Exists(acct) Then
                    debitAccounts.Add acct, 0
                End If

                debitAccounts(acct) = _
                    debitAccounts(acct) + 1

            End If

        End If

    Next rowNum



    If debitTotal <> creditTotal Then

        warningText = _
            warningText & _
            "Unbalanced credit/debit population" & vbCrLf & _
            "Credits: " & creditTotal & vbCrLf & _
            "Debits: " & debitTotal

    End If

    If dateDict.Count > 1 Then

        If warningText <> "" Then
            warningText = warningText & vbCrLf & vbCrLf
        End If

        warningText = warningText & _
            "Multiple dates detected within cluster"

    End If

    Set relDict = BuildRelationshipsFromEdges( _
        hostWb, _
        members)
    
    Set relDict = CollapseRelationships(relDict)

End Function


Private Sub GetClusterMetrics( _
    clusterID As String, _
    ByRef clusterSize As Long, _
    ByRef clusterExposure As Double)

    Dim ws As Worksheet
    Dim r As Long
    Dim lastRow As Long

    Set ws = hostWb.Worksheets("Transfer_Ambiguities")

    lastRow = ws.Cells(ws.rows.Count, 2).End(xlUp).Row

    For r = 2 To lastRow

        If CStr(ws.Cells(r, 2).Value) = clusterID Then

            clusterSize = CLng(ws.Cells(r, 3).Value)
            clusterExposure = CDbl(ws.Cells(r, 4).Value)

            Exit For

        End If

    Next r

End Sub

Private Function CollapseRelationships( _
    relDict As Object) As Object

    Dim outputDict As Object

    Dim key As Variant
    Dim parts() As String

    Set outputDict = CreateObject("Scripting.Dictionary")

    For Each key In relDict.Keys

        parts = Split(CStr(key), "|")

        outputDict( _
            parts(0) & " ? " & parts(1)) = _
            relDict(key)

    Next key
    
    Set CollapseRelationships = outputDict

End Function


Private Sub BuildInvestigationGroupSummary( _
    ByVal hostWb As Workbook)

    Dim srcWs As Worksheet
    Dim outWs As Worksheet

    Dim lastRow As Long
    Dim r As Long

    Dim groups As Object
    Dim groupData As Variant

    Dim groupID As String

    Set groups = CreateObject("Scripting.Dictionary")

    Set srcWs = hostWb.Worksheets("Cluster_Analysis")
    Set outWs = hostWb.Worksheets("Investigation_Group_Summary")

    outWs.rows("2:" & outWs.rows.Count).ClearContents

    lastRow = srcWs.Cells( _
        srcWs.rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow

        groupID = CStr(srcWs.Cells(r, 19).Value)

        If Not groups.Exists(groupID) Then

        ' 0 = Cluster Count
        ' 1 = Matched
        ' 2 = Unbalanced
        ' 3 = Review
        ' 4 = Investigation Key
        ' 5 = Cluster List
        
        groups.Add _
            groupID, _
            Array( _
                0, _
                0, _
                0, _
                0, _
                CStr(srcWs.Cells(r, 18).Value), _
                "" _
            )

        End If

        groupData = groups(groupID)

        groupData(0) = groupData(0) + 1

        Select Case _
            UCase(CStr(srcWs.Cells(r, 14).Value))

            Case "MATCHED"
                groupData(1) = groupData(1) + 1

            Case "UNBALANCED"
                groupData(2) = groupData(2) + 1

            Case Else
                groupData(3) = groupData(3) + 1

        End Select

        If groupData(5) <> "" Then

            groupData(5) = _
                groupData(5) & ", "

        End If

        groupData(5) = _
            groupData(5) & _
            CStr(srcWs.Cells(r, 1).Value)

        groups(groupID) = groupData

    Next r

    Dim outRow As Long
    Dim k As Variant

    outRow = 2

    For Each k In groups.Keys

        groupData = groups(k)

        outWs.Cells(outRow, 1).Value = k
        outWs.Cells(outRow, 2).Value = groupData(0)
        outWs.Cells(outRow, 3).Value = groupData(1)
        outWs.Cells(outRow, 4).Value = groupData(2)
        outWs.Cells(outRow, 5).Value = groupData(3)
        outWs.Cells(outRow, 6).Value = groupData(4)
        outWs.Cells(outRow, 7).Value = groupData(5)

        outRow = outRow + 1

    Next k

    outWs.Columns.AutoFit

End Sub


Private Sub WriteInvestigationGroupSummaryHeaders( _
    ws As Worksheet)

    ws.Cells(1, 1).Value = "Investigation Group"
    ws.Cells(1, 2).Value = "Cluster Count"
    ws.Cells(1, 3).Value = "Matched Clusters"
    ws.Cells(1, 4).Value = "Unbalanced Clusters"
    ws.Cells(1, 5).Value = "Review Clusters"
    ws.Cells(1, 6).Value = "Investigation Key"
    ws.Cells(1, 7).Value = "Cluster IDs"

    ws.rows(1).Font.Bold = True

End Sub

'---Don't touch!---

Private Sub AnalyzeClusters( _
    ByVal hostWb As Workbook, _
    ByRef data As Variant, _
    ByVal colAcct As Long, _
    ByVal colDate As Long, _
    ByVal colAmount As Long, _
    ByVal colDescription As Long, _
    ByVal ambiguityPairs As Collection)

    Dim ws As Worksheet
    Dim clusterDict As Object

    Dim clusterID As Variant
    Dim members As Collection

    Dim outRow As Long
    
    DebugAmbiguous "START AnalyzeClusters"

    Set ws = hostWb.Worksheets("Cluster_Analysis")

    ws.Cells.Clear

    ws.Cells(1, 1).Value = "Cluster ID"
    ws.Cells(1, 2).Value = "Members"
    ws.Cells(1, 3).Value = "Debits"
    ws.Cells(1, 4).Value = "Credits"
    ws.Cells(1, 5).Value = "Distinct Debit Accts"
    ws.Cells(1, 6).Value = "Distinct Credit Accts"
    ws.Cells(1, 7).Value = "Edge Count"
    ws.Cells(1, 8).Value = "Edge Density"
    ws.Cells(1, 9).Value = "Max Matching"
    ws.Cells(1, 10).Value = "Fully Matchable"
    ws.Cells(1, 11).Value = "Classification"
    ws.Cells(1, 12).Value = "Cluster Shape"
    ws.Cells(1, 13).Value = "Perfect Matchings"
    ws.Cells(1, 14).Value = "Recommended Outcome"
    ws.Cells(1, 15).Value = "Distinct Accounts"
    ws.Cells(1, 16).Value = "Account List"
    ws.Cells(1, 17).Value = "Cluster Exposure"
    ws.Cells(1, 18).Value = "Investigation Key"
    ws.Cells(1, 19).Value = "Investigation Group"
    

    ws.rows(1).Font.Bold = True

    Set clusterDict = BuildClusterMemberLookup()

    outRow = 2

    For Each clusterID In clusterDict.Keys

        Set members = clusterDict(clusterID)

        AnalyzeSingleCluster _
            hostWb, _
            ws, _
            outRow, _
            CStr(clusterID), _
            members, _
            data, _
            colAcct, _
            colDate, _
            colAmount, _
            ambiguityPairs

        outRow = outRow + 1

    Next clusterID
    
    
    AssignInvestigationGroups hostWb
    
    BuildInvestigationGroupSummary hostWb
    
    BuildInvestigationGroupMetrics hostWb
     
    BuildResolvedClusterReport hostWb
    
    BuildAutoMatchCandidates hostWb
    
    EnrichTransferRelationships hostWb
    
    
    ws.Columns.AutoFit
    
    DebugAmbiguous "END AnalyzeClusters"

End Sub


Private Function IsResolvedCluster( _
    clusterID As String) As Boolean

    Dim ws As Worksheet

    Dim lastRow As Long
    Dim r As Long

    Set ws = _
        hostWb.Worksheets("Transfer_Resolved_Clusters")

    lastRow = _
        ws.Cells(ws.rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow

        If CStr(ws.Cells(r, 1).Value) = _
           clusterID Then

        Select Case UCase( _
            CStr(ws.Cells(r, 3).Value))
        
            Case "MATCHED", "PARTIAL_MATCH"
        
                IsResolvedCluster = True
                Exit Function
        
        End Select

        End If

    Next r

End Function


Private Function IsAutoResolvableCluster( _
    clusterShape As String, _
    outcome As String, _
    fullyMatchable As String) As Boolean

    IsAutoResolvableCluster = False

    If UCase(Trim(outcome)) <> "MATCHED" Then Exit Function
    If UCase(Trim(fullyMatchable)) <> "YES" Then Exit Function

    IsAutoResolvableCluster = True

End Function


Private Function BuildClusterPairings( _
    members As Collection, _
    ambiguityPairs As Collection, _
    data As Variant, _
    colAmount As Long) As Collection

    Dim debits As Collection
    Dim credits As Collection

    Dim rowNum As Variant

    Set debits = New Collection
    Set credits = New Collection

    For Each rowNum In members

        If CDbl(data(CLng(rowNum), colAmount)) < 0 Then
            debits.Add CLng(rowNum)
        Else
            credits.Add CLng(rowNum)
        End If

    Next rowNum

    If debits.Count <> credits.Count Then
        Exit Function
    End If

    Dim edges As Object
    Set edges = CreateObject("Scripting.Dictionary")

    Dim pairText As Variant
    Dim parts() As String

    Dim rowA As Long
    Dim rowB As Long
    


    For Each pairText In ambiguityPairs

        parts = Split(CStr(pairText), "|")

        rowA = CLng(parts(0))
        rowB = CLng(parts(1))

        If CDbl(data(rowA, colAmount)) < 0 Then
            edges(CStr(rowA) & "|" & CStr(rowB)) = True
        Else
            edges(CStr(rowB) & "|" & CStr(rowA)) = True
        End If
        
    Next pairText
    
    Dim usedCredits As Object
    Set usedCredits = CreateObject("Scripting.Dictionary")

    Dim matches As Collection
    Set matches = New Collection

    If FindMatching( _
        1, _
        debits, _
        credits, _
        edges, _
        usedCredits, _
        matches) Then
        
        

    Set BuildClusterPairings = matches

    End If


End Function


Private Function BuildUniqueClusterPairings( _
    clusterID As String, _
    members As Collection, _
    ambiguityPairs As Collection, _
    data As Variant, _
    colAcct As Long, _
    colAmount As Long) As Collection

    Dim debits As Collection
    Dim credits As Collection

    Dim rowNum As Variant

    Set debits = New Collection
    Set credits = New Collection

    For Each rowNum In members

        If CDbl(data(CLng(rowNum), colAmount)) < 0 Then

            debits.Add CLng(rowNum)

        Else

            credits.Add CLng(rowNum)

        End If

    Next rowNum

    If debits.Count <> credits.Count Then
        Exit Function
    End If

    Dim edges As Object
    Set edges = CreateObject("Scripting.Dictionary")

    Dim pairText As Variant
    Dim parts() As String

    Dim rowA As Long
    Dim rowB As Long

    For Each pairText In ambiguityPairs

        parts = Split(CStr(pairText), "|")

        rowA = CLng(parts(0))
        rowB = CLng(parts(1))

    Dim debitRow As Long
    Dim creditRow As Long
    
    If CDbl(data(rowA, colAmount)) < 0 Then
    
        debitRow = rowA
        creditRow = rowB
    
    Else
    
        debitRow = rowB
        creditRow = rowA
    
    End If
    
If EdgeSupportedByEvidence( _
    debitRow, _
    creditRow, _
    CStr(data(debitRow, colAcct)), _
    CStr(data(creditRow, colAcct))) Then
    
        edges( _
            CStr(debitRow) & "|" & _
            CStr(creditRow)) = True
    
    
    End If
    
        Next pairText
    
        Dim usedCredits As Object
        Set usedCredits = _
            CreateObject("Scripting.Dictionary")
    
        Dim currentMatches As Collection
        Set currentMatches = New Collection
    
        Dim firstSolution As Collection
    
        Dim solutionCount As Long
        
    
    For Each rowNum In members


Next rowNum

    CountClusterSolutions _
        1, _
        debits, _
        credits, _
        edges, _
        usedCredits, _
        currentMatches, _
        solutionCount, _
        firstSolution
        



    If solutionCount = 1 Then
    

For Each pairText In firstSolution

    parts = Split(CStr(pairText), "|")


Next pairText

        Set BuildUniqueClusterPairings = _
            firstSolution

    End If
    
  


End Function

Private Function BuildPartialClusterPairings( _
    clusterID As String, _
    members As Collection, _
    ambiguityPairs As Collection, _
    data As Variant, _
    colAcct As Long, _
    colAmount As Long) As Collection

    Dim debits As Collection
    Dim credits As Collection

    Dim rowNum As Variant

    Set debits = New Collection
    Set credits = New Collection
    

    For Each rowNum In members

        If CDbl(data(CLng(rowNum), colAmount)) < 0 Then

            debits.Add CLng(rowNum)

        Else

            credits.Add CLng(rowNum)

        End If

    Next rowNum


    Dim edges As Object
    Set edges = CreateObject("Scripting.Dictionary")

    Dim pairText As Variant
    Dim parts() As String

    Dim rowA As Long
    Dim rowB As Long
    
    For Each pairText In ambiguityPairs

        parts = Split(CStr(pairText), "|")

        rowA = CLng(parts(0))
        rowB = CLng(parts(1))

    Dim debitRow As Long
    Dim creditRow As Long
    
    If CDbl(data(rowA, colAmount)) < 0 Then
    
        debitRow = rowA
        creditRow = rowB
    
    Else
    
        debitRow = rowB
        creditRow = rowA
    
    End If

    
If EdgeSupportedByEvidence( _
    debitRow, _
    creditRow, _
    CStr(data(debitRow, colAcct)), _
    CStr(data(creditRow, colAcct))) Then
    
        
        edges( _
            CStr(debitRow) & "|" & _
            CStr(creditRow)) = True
            
  
    End If
    
        Next pairText
        
        Dim usedCredits As Object
        Set usedCredits = _
            CreateObject("Scripting.Dictionary")
    
   
        Dim currentMatches As Collection
        Set currentMatches = New Collection
    
        Dim firstSolution As Collection
    
        Dim solutionCount As Long
        

Dim edgeKey As Variant
    
    CountClusterSolutions _
        1, _
        debits, _
        credits, _
        edges, _
        usedCredits, _
        currentMatches, _
        solutionCount, _
        firstSolution
        

    If Not firstSolution Is Nothing Then

    If firstSolution.Count > 0 Then  'check here if there are goofy clusting results
    
'Debug.Print
'Debug.Print "RETURNING PARTIAL SOLUTION"

'For Each pairText In firstSolution
'
'    Debug.Print pairText
'
'Next pairText

        Set BuildPartialClusterPairings = _
            firstSolution

    End If

End If
    
    
End Function
    

Private Sub CountClusterSolutions( _
    debitIndex As Long, _
    debits As Collection, _
    credits As Collection, _
    edges As Object, _
    usedCredits As Object, _
    currentMatches As Collection, _
    solutionCount As Long, _
    firstSolution As Collection)

    If solutionCount >= 2 Then Exit Sub

    If debitIndex > debits.Count Then

        solutionCount = solutionCount + 1

        If solutionCount = 1 Then

            Set firstSolution = _
                CopyPairCollection(currentMatches)

        End If


    If Not firstSolution Is Nothing Then
    
    
    
    End If

        Exit Sub

    End If

    Dim debitRow As Long
    Dim creditRow As Long
    Dim i As Long

    debitRow = CLng(debits(debitIndex))

    For i = 1 To credits.Count

        creditRow = CLng(credits(i))

        If Not usedCredits.Exists( _
            CStr(creditRow)) Then
            
            If edges.Exists( _
                CStr(debitRow) & "|" & _
                CStr(creditRow)) Then
                
                usedCredits( _
                    CStr(creditRow)) = True
                    
                currentMatches.Add _
                    CStr(debitRow) & "|" & _
                    CStr(creditRow)
                    
                CountClusterSolutions _
                    debitIndex + 1, _
                    debits, _
                    credits, _
                    edges, _
                    usedCredits, _
                    currentMatches, _
                    solutionCount, _
                    firstSolution
                    
                usedCredits.Remove _
                    CStr(creditRow)
                    
                currentMatches.Remove _
                    currentMatches.Count
                    
                If solutionCount >= 2 Then Exit Sub
                
            End If
            
        End If

    Next i

End Sub


Private Function CopyPairCollection( _
    source As Collection) As Collection

    Dim result As New Collection
    Dim item As Variant

    For Each item In source
        result.Add item
    Next item

    Set CopyPairCollection = result

End Function

Private Function FindEvidenceMatch( _
    debitRow As Long, _
    credits As Collection, _
    data As Variant, _
    colAcct As Long) As Long

    Dim targetAcct As String
    Dim creditRow As Variant

    If ReferencedAccounts(debitRow) <> "" Then

        targetAcct = _
            ReferencedAccounts(debitRow)

    ElseIf EmbeddedAccounts(debitRow) <> "" Then

        targetAcct = _
            EmbeddedAccounts(debitRow)

    Else

        Exit Function

    End If

    For Each creditRow In credits

        If Right( _
            CStr(data(CLng(creditRow), colAcct)), _
            Len(targetAcct)) = targetAcct Then

            FindEvidenceMatch = _
                CLng(creditRow)

            Exit Function

        End If

    
    Next creditRow
    

End Function


Private Function FindMatching( _
    debitIndex As Long, _
    debits As Collection, _
    credits As Collection, _
    edges As Object, _
    usedCredits As Object, _
    matches As Collection) As Boolean

    If debitIndex > debits.Count Then

        FindMatching = True
        Exit Function

    End If

    Dim debitRow As Long
    debitRow = CLng(debits(debitIndex))

    Dim i As Long
    Dim creditRow As Long
    
    '---Don't remember why this is commented out but things are working without it
    
'    Dim preferredCredit As Long
'
'    preferredCredit = _
'        FindEvidenceMatch( _
'            debitRow, _
'            credits, _
'            data, _
'            colAcct)
'
'    If preferredCredit > 0 Then
'
'        If Not usedCredits.Exists( _
'            CStr(preferredCredit)) Then
'
'            usedCredits(CStr(preferredCredit)) = True
'
'            matches.Add _
'                CStr(debitRow) & "|" & _
'                CStr(preferredCredit)
'
'            If FindMatching( _
'                debitIndex + 1, _
'                debits, _
'                credits, _
'                edges, _
'                usedCredits, _
'                matches) Then
'
'                FindMatching = True
'                Exit Function
'
'            End If
'
'            usedCredits.Remove _
'                CStr(preferredCredit)
'
'            matches.Remove matches.Count
'
'        End If
'
'    End If

    For i = 1 To credits.Count

        creditRow = CLng(credits(i))

        If Not usedCredits.Exists(CStr(creditRow)) Then

            If edges.Exists( _
                CStr(debitRow) & "|" & CStr(creditRow)) Then

                usedCredits(CStr(creditRow)) = True

                matches.Add _
                    CStr(debitRow) & "|" & _
                    CStr(creditRow)

                If FindMatching( _
                    debitIndex + 1, _
                    debits, _
                    credits, _
                    edges, _
                    usedCredits, _
                    matches) Then

                    FindMatching = True
                    Exit Function

                End If

                usedCredits.Remove _
                    CStr(creditRow)

                matches.Remove matches.Count

            End If

        End If

    Next i

End Function


Private Sub WriteClusterPairingsToPreview( _
    ByVal hostWb As Workbook, _
    ByVal clusterID As String, _
    ByVal investigationKey As String, _
    ByVal investigationGroup As String, _
    ByVal clusterShape As String, _
    ByVal outcome As String, _
    ByVal members As Collection, _
    ByVal ambiguityPairs As Collection, _
    ByRef data As Variant, _
    ByVal colAcct As Long, _
    ByVal colAmount As Long)


    Dim pairings As Collection

    If UCase(clusterShape) = _
       "ONE_TO_ONE_REPEATED" Then

        Set pairings = _
            BuildClusterPairings( _
                members, _
                ambiguityPairs, _
                data, _
                colAmount)

    Else

        Set pairings = _
            BuildUniqueClusterPairings( _
                clusterID, _
                members, _
                ambiguityPairs, _
                data, _
                colAcct, _
                colAmount)

    End If

    If pairings Is Nothing Then
        Exit Sub
    End If

    Dim pairText As Variant
    Dim parts() As String

    For Each pairText In pairings

        parts = Split(CStr(pairText), "|")
        
    WriteResolutionPreview _
        hostWb, _
        clusterID, _
        investigationKey, _
        investigationGroup, _
        clusterShape, _
        outcome, _
        CLng(parts(0)), _
        CLng(parts(1)), _
        data, _
        colAcct, _
        colAmount

    Next pairText

End Sub

Private Sub ResolveRepeatedTransferMatches( _
    ByVal hostWb As Workbook, _
    ByRef data As Variant, _
    ByVal colAcct As Long, _
    ByVal colCodeDesc As Long, _
    ByVal colDate As Long, _
    ByVal colAmount As Long, _
    ByVal colDescription As Long)

    Dim analysisWs As Worksheet
    Dim previewWs As Worksheet

    Dim eligibleClusters As Object

    Dim lastRow As Long
    Dim r As Long

    Dim clusterID As String

    Dim debitRow As Long
    Dim creditRow As Long
    
    Dim investigationGroup As String
    Dim clusterShape As String
    Dim outcome As String
    Dim originMethod As String
    
    DebugAmbiguous "START ResolveRepeatedTransferMatches"

    Set analysisWs = hostWb.Worksheets("Cluster_Analysis")
    Set previewWs = hostWb.Worksheets("Cluster_Resolution_Preview")

    Set eligibleClusters = _
        CreateObject("Scripting.Dictionary")
        
    Dim clusterMetadata As Object

    Set clusterMetadata = _
        BuildClusterMetadataLookup(hostWb)
        
    
    ' Build Eligible Cluster List


    lastRow = _
        analysisWs.Cells( _
            analysisWs.rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow

        If IsAutoResolvableCluster( _
            CStr(analysisWs.Cells(r, 12).Value), _
            CStr(analysisWs.Cells(r, 14).Value), _
            CStr(analysisWs.Cells(r, 10).Value)) Then

            eligibleClusters( _
                CStr(analysisWs.Cells(r, 1).Value)) = True

        End If

    Next r


    ' Resolve Preview Rows

    
    lastRow = _
        previewWs.Cells( _
            previewWs.rows.Count, 1).End(xlUp).Row
    
    For r = 2 To lastRow
    
        clusterID = _
            CStr(previewWs.Cells(r, 1).Value)
    
        Dim previewOutcome As String
    
        previewOutcome = _
            UCase$( _
                Trim$( _
                    CStr(previewWs.Cells(r, 5).Value)))
    
    
        ' Allow:
        '   MATCHED clusters
        '   PARTIAL_MATCH preview rows
     
        If Not eligibleClusters.Exists(clusterID) Then
    
    
            If previewOutcome <> _
               "PARTIAL_MATCH" Then
    
                GoTo NextRow
    
            End If
    
        End If
    
        If clusterMetadata.Exists(clusterID) Then
    
            investigationGroup = _
                clusterMetadata(clusterID)(0)
    
            clusterShape = _
                clusterMetadata(clusterID)(1)
    
            outcome = _
                clusterMetadata(clusterID)(2)
    
            originMethod = _
                clusterMetadata(clusterID)(3)
    
        End If
    
        debitRow = _
            CLng(previewWs.Cells(r, 6).Value)
    
        creditRow = _
            CLng(previewWs.Cells(r, 7).Value)
    
        If TransactionStatus.Exists( _
            CStr(debitRow)) Then
    
            If TransactionStatus( _
                CStr(debitRow))(0) = _
                STATUS_MATCHED Then
    
                GoTo NextRow
    
            End If
    
        End If
    
        If TransactionStatus.Exists( _
            CStr(creditRow)) Then
    
            If TransactionStatus( _
                CStr(creditRow))(0) = _
                STATUS_MATCHED Then
    
                GoTo NextRow
    
            End If
    
        End If
    
        Dim resolvedMethod As String
    
        resolvedMethod = _
            "Investigation Group Resolution"
    
        If clusterShape = "ONE_TO_ONE" Then
    
            If originMethod <> "" Then
    
                resolvedMethod = _
                    Split(originMethod, " (")(0)
    
            End If
    
        End If
        
    
        WriteMatched _
            hostWb, _
            data, _
            debitRow, _
            creditRow, _
            colAcct, _
            colCodeDesc, _
            colDate, _
            colAmount, _
            colDescription, _
            resolvedMethod, _
            clusterID, _
            clusterShape, _
            investigationGroup
            


NextRow:
Next r

    DebugAmbiguous "END ResolveRepeatedTransferMatches"

End Sub



Private Sub BuildAutoMatchCandidates( _
    ByVal hostWb As Workbook)

    Dim analysisWs As Worksheet
    Dim previewWs As Worksheet
    Dim outWs As Worksheet

    Dim clusterLookup As Object

    Dim lastRow As Long
    Dim r As Long

    Dim clusterID As String

    Set analysisWs = hostWb.Worksheets("Cluster_Analysis")
    Set previewWs = hostWb.Worksheets("Cluster_Resolution_Preview")
    Set outWs = hostWb.Worksheets("Transfer_AutoMatch_Candidates")

    outWs.rows("2:" & outWs.rows.Count).ClearContents

    Set clusterLookup = _
        CreateObject("Scripting.Dictionary")

    lastRow = _
        analysisWs.Cells( _
            analysisWs.rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow
    
        If IsAutoResolvableCluster( _
            CStr(analysisWs.Cells(r, 12).Value), _
            CStr(analysisWs.Cells(r, 14).Value), _
            CStr(analysisWs.Cells(r, 10).Value)) Then
    
            clusterLookup( _
                CStr(analysisWs.Cells(r, 1).Value)) = _
                CStr(analysisWs.Cells(r, 19).Value)
    
        End If
    
    Next r
        Dim outRow As Long
    
        outRow = 2

    lastRow = _
        previewWs.Cells( _
            previewWs.rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow

        clusterID = _
            CStr(previewWs.Cells(r, 1).Value)

        If clusterLookup.Exists(clusterID) Then

            outWs.Cells(outRow, 1).Value = _
                clusterID

            outWs.Cells(outRow, 2).Value = _
                clusterLookup(clusterID)

            outWs.Cells(outRow, 3).Value = _
                previewWs.Cells(r, 5).Value
            
            outWs.Cells(outRow, 4).Value = _
                previewWs.Cells(r, 6).Value
            
            outWs.Cells(outRow, 5).Value = _
                previewWs.Cells(r, 7).Value
            
            outWs.Cells(outRow, 6).Value = _
                previewWs.Cells(r, 8).Value
            
            outWs.Cells(outRow, 7).Value = _
                previewWs.Cells(r, 9).Value

            outRow = outRow + 1

        End If

    Next r

    outWs.Columns.AutoFit

End Sub

'---Adding Ambig resolution data to matched and unmatched items

Private Sub EnrichTransferRelationships( _
    ByVal hostWb As Workbook)

    Dim relWs As Worksheet

    Dim clusterMetadata As Object
    Dim clusterMembers As Object

    Dim lastRow As Long
    Dim r As Long

    Dim rowA As Long
    Dim rowB As Long

    Dim clusterID As String

    Set relWs = _
        hostWb.Worksheets("Transfer_Relationships")

    Set clusterMetadata = _
        BuildClusterMetadataLookup(hostWb)

    Set clusterMembers = _
        BuildClusterMemberLookup()

    lastRow = _
        relWs.Cells( _
            relWs.rows.Count, 1).End(xlUp).Row
            
    
    For r = 2 To lastRow
    
   
        ' already populated
        If Trim(CStr(relWs.Cells(r, 11).Value)) <> "" Then
            GoTo NextR
        End If

        rowA = _
            CLng(relWs.Cells(r, 8).Value)

        rowB = _
            CLng(relWs.Cells(r, 9).Value)

        clusterID = _
            FindRelationshipCluster( _
                rowA, _
                rowB, _
                clusterMembers)
                
              

        If clusterID <> "" Then
        

            If clusterMetadata.Exists(clusterID) Then

                If clusterMetadata(clusterID)(1) = _
                   "ONE_TO_ONE" Then
                
                
                    relWs.Cells(r, 11).Value = _
                        clusterID
                
                    relWs.Cells(r, 12).Value = _
                        clusterMetadata(clusterID)(1)
                
                    relWs.Cells(r, 13).Value = _
                        clusterMetadata(clusterID)(0)

                                             
                End If

            End If


        End If

NextR:
    Next r

End Sub

Private Function FindRelationshipCluster( _
    rowA As Long, _
    rowB As Long, _
    clusterMembers As Object) As String

    Dim clusterID As Variant
    Dim member As Variant

    Dim hasA As Boolean
    Dim hasB As Boolean

    For Each clusterID In clusterMembers.Keys

        hasA = False
        hasB = False

        For Each member In _
            clusterMembers(clusterID)

            If CLng(member) = rowA Then
                hasA = True
            End If

            If CLng(member) = rowB Then
                hasB = True
            End If

        Next member

        If hasA And hasB Then
        

            FindRelationshipCluster = _
                CStr(clusterID)
                

            Exit Function

        End If

    Next clusterID

End Function


Private Sub WriteAutoMatchCandidateHeaders( _
    ws As Worksheet)

    ws.Cells(1, 1).Value = "Cluster ID"
    ws.Cells(1, 2).Value = "Investigation Group"
    ws.Cells(1, 3).Value = "Debit Row"
    ws.Cells(1, 4).Value = "Credit Row"
    ws.Cells(1, 5).Value = "Debit Account"
    ws.Cells(1, 6).Value = "Credit Account"
    ws.Cells(1, 7).Value = "Amount"

    ws.rows(1).Font.Bold = True

End Sub


Private Sub WriteResolvedClusterHeaders( _
    ws As Worksheet)

    ws.Cells(1, 1).Value = "Cluster ID"
    ws.Cells(1, 2).Value = "Investigation Group"
    ws.Cells(1, 3).Value = "Outcome"
    ws.Cells(1, 4).Value = "Cluster Shape"
    ws.Cells(1, 5).Value = "Members"
    ws.Cells(1, 6).Value = "Distinct Accounts"
    ws.Cells(1, 7).Value = "Account List"
    ws.Cells(1, 8).Value = "Max Matching"
    ws.Cells(1, 9).Value = "Perfect Matchings"
    ws.Cells(1, 10).Value = "Fully Matchable"

    ws.rows(1).Font.Bold = True

End Sub


Private Sub WriteResolvedCluster( _
    ByVal hostWb As Workbook, _
    ByVal clusterID As String, _
    ByVal investigationGroup As String, _
    ByVal outcome As String, _
    ByVal clusterShape As String, _
    ByVal memberCount As Long, _
    ByVal distinctAccounts As Long, _
    ByVal accountList As String)

    Dim ws As Worksheet
    Dim r As Long

    Set ws = _
        hostWb.Worksheets("Transfer_Resolved_Clusters")

    r = ws.Cells( _
        ws.rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(r, 1).Value = clusterID
    ws.Cells(r, 2).Value = investigationGroup
    ws.Cells(r, 3).Value = outcome
    ws.Cells(r, 4).Value = clusterShape
    ws.Cells(r, 5).Value = memberCount
    ws.Cells(r, 6).Value = distinctAccounts
    ws.Cells(r, 7).Value = accountList

End Sub


Private Sub BuildResolvedClusterReport( _
    ByVal hostWb As Workbook)

    Dim srcWs As Worksheet
    Dim outWs As Worksheet

    Dim lastRow As Long
    Dim r As Long
    Dim outRow As Long

    Set srcWs = hostWb.Worksheets("Cluster_Analysis")
    Set outWs = hostWb.Worksheets("Transfer_Resolved_Clusters")

    outWs.rows("2:" & outWs.rows.Count).ClearContents

    lastRow = _
        srcWs.Cells(srcWs.rows.Count, 1).End(xlUp).Row

    outRow = 2

    For r = 2 To lastRow
    
    If Trim(CStr(srcWs.Cells(r, 14).Value)) <> "" Then

        outWs.Cells(outRow, 1).Value = _
            srcWs.Cells(r, 1).Value     ' Cluster ID

        outWs.Cells(outRow, 2).Value = _
            srcWs.Cells(r, 19).Value    ' Investigation Group

        outWs.Cells(outRow, 3).Value = _
            srcWs.Cells(r, 14).Value    ' Recommended Outcome

        outWs.Cells(outRow, 4).Value = _
            srcWs.Cells(r, 12).Value    ' Cluster Shape

        outWs.Cells(outRow, 5).Value = _
            srcWs.Cells(r, 2).Value     ' Members

        outWs.Cells(outRow, 6).Value = _
            srcWs.Cells(r, 15).Value    ' Distinct Accounts

        outWs.Cells(outRow, 7).Value = _
            srcWs.Cells(r, 16).Value    ' Account List
            
        outWs.Cells(outRow, 8).Value = _
            srcWs.Cells(r, 9)           ' Max Matching
            
        outWs.Cells(outRow, 9).Value = _
            srcWs.Cells(r, 13)          ' Perfect Matchings
            
        outWs.Cells(outRow, 10).Value = _
            srcWs.Cells(r, 10)          ' Fully Matchable

        outRow = outRow + 1
        
    End If

    Next r

    outWs.Columns.AutoFit

End Sub

Private Function BuildInvestigationKeyFromDict( _
    allAccounts As Object) As String

    Dim accounts() As String
    Dim i As Long
    Dim j As Long
    Dim temp As String

    ReDim accounts(0 To allAccounts.Count - 1)

    Dim k As Variant
    Dim idx As Long

    For Each k In allAccounts.Keys

        accounts(idx) = CStr(k)
        idx = idx + 1

    Next k

    For i = LBound(accounts) To UBound(accounts) - 1

        For j = i + 1 To UBound(accounts)

            If accounts(j) < accounts(i) Then

                temp = accounts(i)
                accounts(i) = accounts(j)
                accounts(j) = temp

            End If

        Next j

    Next i

    For i = LBound(accounts) To UBound(accounts)

        If BuildInvestigationKeyFromDict <> "" Then

            BuildInvestigationKeyFromDict = _
                BuildInvestigationKeyFromDict & "|"

        End If

        BuildInvestigationKeyFromDict = _
            BuildInvestigationKeyFromDict & _
            accounts(i)

    Next i

End Function


Private Function BuildInvestigationKey( _
    accountList As String, _
    clusterDate As Date, _
    clusterAmount As Double) As String

    Dim accounts() As String

    accounts = Split(accountList, ",")

    Dim i As Long
    Dim j As Long
    Dim temp As String

    For i = LBound(accounts) To UBound(accounts) - 1

        For j = i + 1 To UBound(accounts)

            If Trim(accounts(j)) < Trim(accounts(i)) Then

                temp = accounts(i)
                accounts(i) = accounts(j)
                accounts(j) = temp

            End If

        Next j

    Next i

    Dim keyText As String

    For i = LBound(accounts) To UBound(accounts)

        If keyText <> "" Then

            keyText = keyText & "|"

        End If

        keyText = keyText & Trim(accounts(i))

    Next i

    BuildInvestigationKey = _
        keyText & "|" & _
        Format(clusterDate, "yyyymmdd") & "|" & _
        Format(Abs(clusterAmount), "0.00")

End Function

Private Function GetMaximumMatchingSize( _
    members As Collection, _
    ambiguityPairs As Collection, _
    data As Variant, _
    colAmount As Long, _
    colAcct As Long) As Long

    Dim debitNodes As Object
    Dim creditNodes As Object
    Dim edges As Object

    Dim rowNum As Variant
    Dim pairText As Variant

    Dim parts() As String

    Dim rowA As Long
    Dim rowB As Long

    Dim debitRow As Long
    Dim creditRow As Long

    Dim memberLookup As Object

    Set debitNodes = CreateObject("Scripting.Dictionary")
    Set creditNodes = CreateObject("Scripting.Dictionary")
    Set edges = CreateObject("Scripting.Dictionary")
    Set memberLookup = CreateObject("Scripting.Dictionary")

   
   
    For Each rowNum In members
    
    
        memberLookup(CStr(rowNum)) = True

        If CDbl(data(CLng(rowNum), colAmount)) < 0 Then

            debitNodes(CStr(rowNum)) = True

        Else

            creditNodes(CStr(rowNum)) = True

        End If

    Next rowNum

    For Each pairText In ambiguityPairs

        parts = Split(CStr(pairText), "|")

        rowA = CLng(parts(0))
        rowB = CLng(parts(1))

        If Not memberLookup.Exists(CStr(rowA)) Then GoTo NextPair
        If Not memberLookup.Exists(CStr(rowB)) Then GoTo NextPair

        If CDbl(data(rowA, colAmount)) < 0 Then

            debitRow = rowA
            creditRow = rowB

        Else

            debitRow = rowB
            creditRow = rowA

        End If
        
        If Not edges.Exists(CStr(debitRow)) Then

            Set edges(CStr(debitRow)) = _
                CreateObject("Scripting.Dictionary")

        End If

    If EdgeSupportedByEvidence( _
        debitRow, _
        creditRow, _
        CStr(data(debitRow, colAcct)), _
        CStr(data(creditRow, colAcct))) Then

            edges(CStr(debitRow))(CStr(creditRow)) = True
            
        End If

NextPair:
    Next pairText

    Dim usedCredits As Object

    Set usedCredits = CreateObject("Scripting.Dictionary")

    Dim debitKey As Variant
    Dim creditKey As Variant

    For Each debitKey In edges.Keys

        For Each creditKey In edges(debitKey).Keys

            If Not usedCredits.Exists(CStr(creditKey)) Then

                usedCredits.Add _
                    CStr(creditKey), _
                    True

                GetMaximumMatchingSize = _
                    GetMaximumMatchingSize + 1

                Exit For

            End If

        Next creditKey

    Next debitKey
    

End Function

Private Sub AnalyzeResidualCluster( _
    residualMembers As Collection, _
    residualPairs As Collection, _
    data As Variant, _
    colAcct As Long, _
    colAmount As Long)

    Dim residualMaxMatching As Long
    Dim residualPerfectMatchings As Long

    residualMaxMatching = _
        GetMaximumMatchingSize( _
            residualMembers, _
            residualPairs, _
            data, _
            colAmount, _
            colAcct)

    residualPerfectMatchings = _
        CountPerfectMatchings( _
            residualMembers, _
            residualPairs, _
            data, _
            colAmount, _
            colAcct)


End Sub


Private Function GetClusterEdgeCount( _
    clusterID As String, _
    members As Collection, _
    ambiguityPairs As Collection) As Long

    Dim memberLookup As Object
    Dim uniqueEdges As Object

    Dim pairText As Variant
    Dim parts() As String

    Dim rowA As Long
    Dim rowB As Long

    Dim edgeKey As String

    Set memberLookup = CreateObject("Scripting.Dictionary")
    Set uniqueEdges = CreateObject("Scripting.Dictionary")

    Dim v As Variant

    For Each v In members

        memberLookup(CStr(v)) = True

    Next v

    For Each pairText In ambiguityPairs

        parts = Split(CStr(pairText), "|")

        rowA = CLng(parts(0))
        rowB = CLng(parts(1))

        If rowA > rowB Then

            edgeKey = _
                CStr(rowB) & "|" & CStr(rowA)

        Else

            edgeKey = _
                CStr(rowA) & "|" & CStr(rowB)

        End If

        If memberLookup.Exists(CStr(rowA)) _
        And memberLookup.Exists(CStr(rowB)) Then

            If Not uniqueEdges.Exists(edgeKey) Then

                uniqueEdges.Add edgeKey, True

            End If

        End If

    Next pairText

    GetClusterEdgeCount = uniqueEdges.Count

End Function


'DO NOT TOUCH UNLESS YOU HAVE A LOT OF TIME, this one sucks

Private Sub AnalyzeSingleCluster( _
    ByVal hostWb As Workbook, _
    ByVal ws As Worksheet, _
    ByRef outRow As Long, _
    ByVal clusterID As String, _
    ByVal members As Collection, _
    ByRef data As Variant, _
    ByVal colAcct As Long, _
    ByVal colDate As Long, _
    ByVal colAmount As Long, _
    ByVal ambiguityPairs As Collection)


    Dim rowNum As Variant
    Dim pairText As Variant
    Dim acctKey As Variant

    Dim parts() As String

    Dim debitRow As Long
    Dim creditRow As Long

    Dim acct As String
    Dim amt As Double

    '--------------------------------------------------------
    ' Original cluster statistics
    '--------------------------------------------------------

    Dim debitCount As Long
    Dim creditCount As Long

    Dim debitAccts As Object
    Dim creditAccts As Object
    Dim allAccounts As Object

    Dim distinctAccounts As Long
    Dim accountList As String

    Dim clusterExposure As Double
    Dim clusterDate As Date
    Dim clusterAmount As Double

    Dim originalClusterShape As String
    Dim clusterShape As String

    Dim investigationKey As String
    Dim investigationGroup As String

    '--------------------------------------------------------
    ' Working and residual cluster state
    '--------------------------------------------------------

    Dim workingMembers As Collection
    Dim workingPairs As Collection

    Dim evidencePairs As Collection
    Dim partialPairings As Collection

    Dim residualDebitAccts As Object
    Dim residualCreditAccts As Object
    Dim residualDebitCount As Long
    Dim residualCreditCount As Long
    Dim residualClusterShape As String

    Dim partialLookup As Object

    Dim previousMemberCount As Long
    Dim passCount As Long

    '--------------------------------------------------------
    ' Residual graph metrics
    '--------------------------------------------------------

    Dim edgeCount As Long
    Dim edgeDensity As Double

    Dim maxMatching As Long
    Dim perfectMatchings As Long
    Dim isUnique As Boolean

    Dim fullyMatchable As String
    Dim uniqueSolution As String

    Dim resolvedPairCount As Long
    Dim recommendedOutcome As String

    ' Verify whether another procedure normally populates this.
    Dim classification As String

    '========================================================
    ' INITIALIZE DICTIONARIES
    '========================================================

    Set allAccounts = _
        CreateObject("Scripting.Dictionary")

    Set debitAccts = _
        CreateObject("Scripting.Dictionary")

    Set creditAccts = _
        CreateObject("Scripting.Dictionary")

    Set residualDebitAccts = _
        CreateObject("Scripting.Dictionary")

    Set residualCreditAccts = _
        CreateObject("Scripting.Dictionary")

    Set partialLookup = _
        CreateObject("Scripting.Dictionary")

    ' workingMembers and workingPairs begin with the original
    ' cluster. Removal functions return new collections, so the
    ' original collections are not physically altered.

    Set workingMembers = members
    
    Set workingPairs = _
        FilterAmbiguityPairs( _
            ambiguityPairs, _
            workingMembers)


    '========================================================
    ' PHASE 1
    ' CALCULATE ORIGINAL CLUSTER STATISTICS
    '========================================================

    For Each rowNum In members

        acct = _
            Trim$(CStr(data(CLng(rowNum), colAcct)))

        If Not allAccounts.Exists(acct) Then

            allAccounts.Add _
                acct, _
                True

        End If

        amt = _
            CDbl(data(CLng(rowNum), colAmount))

        clusterExposure = _
            clusterExposure + Abs(amt)

        If amt < 0 Then

            debitCount = _
                debitCount + 1

            If Not debitAccts.Exists(acct) Then

                debitAccts.Add _
                    acct, _
                    True

            End If

        Else

            creditCount = _
                creditCount + 1

            If Not creditAccts.Exists(acct) Then

                creditAccts.Add _
                    acct, _
                    True

            End If

        End If

        ' Use the first member to establish the cluster date
        ' and amount used by the investigation key.

        If clusterDate = 0 Then

            clusterDate = _
                CDate(data(CLng(rowNum), colDate))

        End If

        If clusterAmount = 0 Then

            clusterAmount = _
                Abs(CDbl(data(CLng(rowNum), colAmount)))

        End If

    Next rowNum

    '--------------------------------------------------------
    ' Build original account list and shape
    '--------------------------------------------------------

    distinctAccounts = _
        allAccounts.Count

    accountList = ""

    For Each acctKey In allAccounts.Keys

        If accountList <> "" Then

            accountList = _
                accountList & ", "

        End If

        accountList = _
            accountList & CStr(acctKey)

    Next acctKey

    investigationKey = _
        BuildInvestigationKey( _
            accountList, _
            clusterDate, _
            clusterAmount)

    originalClusterShape = _
        GetClusterShape( _
            debitAccts.Count, _
            creditAccts.Count, _
            debitCount, _
            creditCount)

    ' Preserve the original cluster shape for reporting.
    clusterShape = _
        originalClusterShape

    '========================================================
    ' PHASE 2
    ' ITERATIVELY EXTRACT RESOLVABLE EVIDENCE PAIRS
    '========================================================

    passCount = 0

    Do

        passCount = _
            passCount + 1

        ' Emergency protection while the evidence loop is
        ' being regression tested.

        If passCount > 100 Then

            Debug.Print _
                "EMERGENCY LOOP EXIT", _
                clusterID

            Exit Do

        End If

        If workingMembers.Count = 0 Then

            Exit Do

        End If

        previousMemberCount = _
            workingMembers.Count

        Set evidencePairs = _
            ExtractResolvableEvidencePairs( _
                workingPairs, _
                data, _
                colAcct)

        ' No additional certainty was found.
        If evidencePairs.Count = 0 Then

            Exit Do

        End If

        If partialPairings Is Nothing Then

            Set partialPairings = _
                New Collection

        End If

        ' Add evidence pairings to the cumulative output
        ' collection. The dictionary prevents duplicate writes
        ' when the same row pair appeared under multiple methods.

        For Each pairText In evidencePairs

            If Trim$(CStr(pairText)) <> "" Then

                If Not partialLookup.Exists( _
                    CStr(pairText)) Then

                    partialPairings.Add _
                        CStr(pairText)

                    partialLookup.Add _
                        CStr(pairText), _
                        True

                End If

            Else

                Debug.Print _
                    "IGNORED BLANK EVIDENCE PAIR", _
                    clusterID

            End If

        Next pairText

        ' Remove resolved transaction rows from the working
        ' member collection.

        Set workingMembers = _
            RemoveResolvedMembers( _
                workingMembers, _
                evidencePairs)


        If workingMembers.Count = 0 Then

            Exit Do

        End If

        ' Remove ambiguity pairs involving transactions that
        ' have already been resolved.

        Set workingPairs = _
            FilterAmbiguityPairs( _
                workingPairs, _
                workingMembers)


        ' Evidence was returned, but no transaction was removed.
        ' Without this guard, the same evidence can be found
        ' repeatedly and Excel can enter an infinite loop.

        If workingMembers.Count >= previousMemberCount Then

            Debug.Print _
                "NO EVIDENCE PROGRESS - EXITING LOOP", _
                clusterID

            Exit Do

        End If

    Loop

        '========================================================
        ' PHASE 3
        ' REBUILD RESIDUAL ACCOUNT AND TRANSACTION COUNTS
        '========================================================
        
        Set residualDebitAccts = _
            CreateObject("Scripting.Dictionary")
        
        Set residualCreditAccts = _
            CreateObject("Scripting.Dictionary")
        
        residualDebitCount = 0
        residualCreditCount = 0
        
        For Each rowNum In workingMembers
        
            acct = _
                Trim$(CStr(data(CLng(rowNum), colAcct)))
        
            amt = _
                CDbl(data(CLng(rowNum), colAmount))
        
            If amt < 0 Then
        
                residualDebitCount = _
                    residualDebitCount + 1
        
                If Not residualDebitAccts.Exists(acct) Then
        
                    residualDebitAccts.Add _
                        acct, _
                        True
        
                End If
        
            Else
        
                residualCreditCount = _
                    residualCreditCount + 1
        
                If Not residualCreditAccts.Exists(acct) Then
        
                    residualCreditAccts.Add _
                        acct, _
                        True
        
                End If
        
            End If
        
        Next rowNum

    '--------------------------------------------------------
    ' Determine the shape of the residual cluster.
    '
    ' If no members remain, preserve the original shape for
    ' reporting, although the outcome will be MATCHED.
    '--------------------------------------------------------

    If workingMembers.Count > 0 Then

    residualClusterShape = _
        GetClusterShape( _
            residualDebitAccts.Count, _
            residualCreditAccts.Count, _
            residualDebitCount, _
            residualCreditCount)
        
    Else

        residualClusterShape = _
            originalClusterShape

    End If

    '========================================================
    ' PHASE 4
    ' CALCULATE RESIDUAL GRAPH METRICS
    '========================================================

    If workingMembers.Count > 0 Then

        edgeCount = _
            GetClusterEdgeCount( _
                clusterID, _
                workingMembers, _
                workingPairs)

        maxMatching = _
            GetMaximumMatchingSize( _
                workingMembers, _
                workingPairs, _
                data, _
                colAmount, _
                colAcct)

        perfectMatchings = _
            CountPerfectMatchings( _
                workingMembers, _
                workingPairs, _
                data, _
                colAmount, _
                colAcct)

        isUnique = _
            HasUniqueSolution( _
                workingMembers, _
                workingPairs, _
                data, _
                colAmount)

        ' Edge density now describes the residual cluster.
        ' Avoid division by zero when every member was resolved.

        edgeDensity = _
            edgeCount / workingMembers.Count

    Else

        edgeCount = 0
        maxMatching = 0
        perfectMatchings = 0
        isUnique = True
        edgeDensity = 0

    End If
    

    '--------------------------------------------------------
    ' Calculate fully-matchable only after maxMatching has
    ' been recalculated using the residual cluster.
    '--------------------------------------------------------

    If workingMembers.Count = 0 Then

        fullyMatchable = _
            "YES"

    ElseIf maxMatching * 2 = _
           workingMembers.Count Then

        fullyMatchable = _
            "YES"

    Else

        fullyMatchable = _
            "NO"

    End If

    If isUnique Then

        uniqueSolution = _
            "YES"

    Else

        uniqueSolution = _
            "NO"

    End If


    '========================================================
    ' PHASE 5
    ' DETERMINE ONE AUTHORITATIVE OUTCOME
    '========================================================

    resolvedPairCount = 0

    If Not partialPairings Is Nothing Then

        resolvedPairCount = _
            partialPairings.Count

    End If
    

    recommendedOutcome = _
        DetermineClusterOutcome( _
            residualClusterShape, _
            members.Count, _
            workingMembers.Count, _
            residualDebitCount, _
            residualCreditCount, _
            resolvedPairCount, _
            maxMatching, _
            perfectMatchings)


    '========================================================
    ' PHASE 6
    ' WRITE RESOLUTION OUTPUT EXACTLY ONCE
    '========================================================

    Select Case recommendedOutcome

        Case "PARTIAL_MATCH"

            ' Write only the evidence-supported pairs.
            ' Residual members remain unresolved.

            If Not partialPairings Is Nothing Then

                For Each pairText In partialPairings

                    parts = _
                        Split(CStr(pairText), "|")

                    If UBound(parts) >= 1 Then

                        debitRow = _
                            CLng(parts(0))

                        creditRow = _
                            CLng(parts(1))

                        WriteResolutionPreview _
                            hostWb, _
                            clusterID, _
                            investigationKey, _
                            investigationGroup, _
                            originalClusterShape, _
                            "PARTIAL_MATCH", _
                            debitRow, _
                            creditRow, _
                            data, _
                            colAcct, _
                            colAmount

                    Else

                        Debug.Print _
                            "BAD PARTIAL PAIR:", _
                            pairText

                    End If

                Next pairText

            End If

        Case "MATCHED"
        
            If resolvedPairCount > 0 _
            And workingMembers.Count = 0 Then
        
                ' Existing evidence-pair output logic.
        
                For Each pairText In partialPairings
        
                    parts = _
                        Split(CStr(pairText), "|")
        
                    If UBound(parts) >= 1 Then
        
                        debitRow = _
                            CLng(parts(0))
        
                        creditRow = _
                            CLng(parts(1))
        
                        WriteResolutionPreview _
                            hostWb, _
                            clusterID, _
                            investigationKey, _
                            investigationGroup, _
                            originalClusterShape, _
                            "MATCHED", _
                            debitRow, _
                            creditRow, _
                            data, _
                            colAcct, _
                            colAmount
        
                    End If
        
                Next pairText
        
            ElseIf residualClusterShape = _
                   "ONE_TO_ONE_REPEATED" Then
        
                PreviewRepeatedTransferCluster _
                    hostWb, _
                    clusterID, _
                    investigationKey, _
                    investigationGroup, _
                    residualClusterShape, _
                    recommendedOutcome, _
                    workingMembers, _
                    data, _
                    colAcct, _
                    colAmount
        
            ElseIf IsEquivalentAccountAmbiguity( _
                residualClusterShape, _
                workingMembers.Count, _
                residualDebitCount, _
                residualCreditCount, _
                maxMatching, _
                perfectMatchings) Then
        
                Dim equivalentPairings As Collection
        
                Set equivalentPairings = _
                    BuildEquivalentPairings( _
                        workingMembers, _
                        workingPairs, _
                        data, _
                        colAmount)
        
                If equivalentPairings.Count * 2 = _
                   workingMembers.Count Then
        
                    For Each pairText In equivalentPairings
        
                        parts = _
                            Split(CStr(pairText), "|")
        
                        debitRow = _
                            CLng(parts(0))
        
                        creditRow = _
                            CLng(parts(1))
        
                        WriteResolutionPreview _
                            hostWb, _
                            clusterID, _
                            investigationKey, _
                            investigationGroup, _
                            originalClusterShape, _
                            "MATCHED", _
                            debitRow, _
                            creditRow, _
                            data, _
                            colAcct, _
                            colAmount
        
                    Next pairText
        
                Else
        
                    Debug.Print _
                        "EQUIVALENT PAIRING INCOMPLETE", _
                        clusterID, _
                        equivalentPairings.Count, _
                        workingMembers.Count
        
                End If
        
            Else
        
            PreviewResolvedCluster _
                hostWb, _
                clusterID, _
                investigationKey, _
                investigationGroup, _
                residualClusterShape, _
                recommendedOutcome, _
                workingMembers, _
                workingPairs, _
                data, _
                colAcct, _
                colAmount
        
            End If

        Case "UNBALANCED"

            ' The residual cluster cannot consume every member.
            ' Remaining workingMembers should later be written
            ' to Unmatched_Transfers with cluster metadata.

        Case "AMBIGUOUS"

            ' Multiple meaningful residual solutions remain.
            ' Existing ambiguity reporting can handle these rows.

        Case "REVIEW"

            ' Unexpected or contradictory residual state.
            ' Preserve the cluster for manual review.

        Case Else

            Debug.Print _
                "UNKNOWN CLUSTER OUTCOME", _
                clusterID, _
                recommendedOutcome

    End Select

    '========================================================
    ' PHASE 7
    ' WRITE CLUSTER ANALYSIS ROW
    '========================================================

    ws.Cells(outRow, 1).Value = _
        clusterID

    ws.Cells(outRow, 2).Value = _
        members.Count

    ws.Cells(outRow, 3).Value = _
        debitCount

    ws.Cells(outRow, 4).Value = _
        creditCount

    ws.Cells(outRow, 5).Value = _
        debitAccts.Count

    ws.Cells(outRow, 6).Value = _
        creditAccts.Count

    ' These graph metrics describe the residual cluster after
    ' evidence-supported pairings have been removed.

    ws.Cells(outRow, 7).Value = _
        edgeCount

    ws.Cells(outRow, 8).Value = _
        Round(edgeDensity, 2)

    ws.Cells(outRow, 9).Value = _
        maxMatching

    ws.Cells(outRow, 10).Value = _
        fullyMatchable

    ws.Cells(outRow, 11).Value = _
        classification

    ' Preserve original shape in the existing report column.

    ws.Cells(outRow, 12).Value = _
        originalClusterShape

    ws.Cells(outRow, 13).Value = _
        perfectMatchings

    ws.Cells(outRow, 14).Value = _
        recommendedOutcome

    ws.Cells(outRow, 15).Value = _
        distinctAccounts

    ws.Cells(outRow, 16).Value = _
        accountList

    ws.Cells(outRow, 17).Value = _
        clusterExposure

    ws.Cells(outRow, 17).numberFormat = _
        "$#,##0.00;($#,##0.00)"

    ws.Cells(outRow, 18).Value = _
        investigationKey

End Sub

Private Function DetermineClusterOutcome( _
    clusterShape As String, _
    originalMemberCount As Long, _
    remainingMemberCount As Long, _
    remainingDebitCount As Long, _
    remainingCreditCount As Long, _
    resolvedPairCount As Long, _
    maxMatching As Long, _
    perfectMatchings As Long) As String

    '====================================================
    ' All original members were consumed.
    '====================================================

    If remainingMemberCount = 0 Then

        DetermineClusterOutcome = _
            "MATCHED"

        Exit Function

    End If

    '====================================================
    ' Some evidence matches resolved, but transactions
    ' remain for investigation.
    '====================================================

    If resolvedPairCount > 0 Then

        DetermineClusterOutcome = _
            "PARTIAL_MATCH"

        Exit Function

    End If

    '====================================================
    ' Imbalance means the residual cluster has
    ' unequal debit and credit transaction counts.
    '====================================================

    If remainingDebitCount <> remainingCreditCount Then

        DetermineClusterOutcome = _
            "UNBALANCED"

        Exit Function

    End If
    
    '====================================================
    ' Equivalent row-level ambiguity
    '
    ' Multiple matchings exist, but all matching
    ' variations produce the same account relationships.
    '====================================================
    
    If IsEquivalentAccountAmbiguity( _
        clusterShape, _
        remainingMemberCount, _
        remainingDebitCount, _
        remainingCreditCount, _
        maxMatching, _
        perfectMatchings) Then
    
        DetermineClusterOutcome = _
            "MATCHED"
    
        Exit Function
    
    End If

    '====================================================
    ' Repeated A-to-B transfers are functionally resolved
    ' if all residual members can be paired.
    '====================================================

    If clusterShape = "ONE_TO_ONE_REPEATED" Then

        If maxMatching * 2 = remainingMemberCount Then

            DetermineClusterOutcome = _
                "MATCHED"

            Exit Function

        End If

    End If

    '====================================================
    ' Standard one-to-one cluster.
    '====================================================

    If clusterShape = "ONE_TO_ONE" Then

        If maxMatching = 1 Then

            DetermineClusterOutcome = _
                "MATCHED"

            Exit Function

        End If

    End If

    '====================================================
    ' Exactly one complete graph solution exists.
    '====================================================

    If perfectMatchings = 1 _
    And maxMatching * 2 = remainingMemberCount Then

        DetermineClusterOutcome = _
            "MATCHED"

        Exit Function

    End If
    


    '====================================================
    ' Multiple meaningful complete solutions remain.
    '====================================================

    If perfectMatchings > 1 Then

        DetermineClusterOutcome = _
            "AMBIGUOUS"

        Exit Function

    End If

    '====================================================
    ' Equal sides but no complete solution means the graph
    ' is deficient or candidate relationships are missing.
    ' This is ambiguity, not structural imbalance.
    '====================================================

    If maxMatching * 2 < remainingMemberCount Then

        DetermineClusterOutcome = _
            "AMBIGUOUS"

        Exit Function

    End If

    '====================================================
    ' Unexpected resolution
    '====================================================

    DetermineClusterOutcome = _
        "REVIEW"

End Function

Private Function FilterAmbiguityPairs( _
    ambiguityPairs As Collection, _
    remainingMembers As Collection) As Collection

    '========================================================
    ' Return only ambiguity pairs where both transaction rows
    ' still belong to the current residual cluster.
    '
    ' The original pair text is preserved, including the
    ' discovery method:
    '
    '     rowA|rowB|Amount + Date
    '========================================================

    Dim results As New Collection

    Dim memberLookup As Object
    Dim seenPairs As Object

    Dim rowNum As Variant
    Dim pairText As Variant

    Dim parts() As String
    Dim rowA As Long
    Dim rowB As Long

    Dim pairKey As String

    Set memberLookup = _
        CreateObject("Scripting.Dictionary")

    Set seenPairs = _
        CreateObject("Scripting.Dictionary")


    ' Build lookup containing only residual transaction rows.

    For Each rowNum In remainingMembers

        memberLookup(CStr(CLng(rowNum))) = _
            True

    Next rowNum


    ' Keep pairs only when both rows remain in the cluster.

    For Each pairText In ambiguityPairs

        If Trim$(CStr(pairText)) <> "" Then

            parts = _
                Split(CStr(pairText), "|")

            If UBound(parts) >= 1 Then

                rowA = _
                    CLng(Trim$(parts(0)))

                rowB = _
                    CLng(Trim$(parts(1)))

                If memberLookup.Exists(CStr(rowA)) _
                And memberLookup.Exists(CStr(rowB)) Then

                    ' Preserve the method in the key.
                    '
                    ' The same row pair may legitimately appear
                    ' under Narrative Pair and Amount + Date.
                    ' Those are different candidate records.

                    pairKey = _
                        CStr(pairText)

                    If Not seenPairs.Exists(pairKey) Then

                        results.Add _
                            pairKey

                        seenPairs.Add _
                            pairKey, _
                            True

                    End If

                End If

            Else

                Debug.Print _
                    "BAD AMBIGUITY PAIR:", _
                    pairText

            End If

        Else

            Debug.Print _
                "BLANK AMBIGUITY PAIR IGNORED"

        End If

    Next pairText

    Set FilterAmbiguityPairs = _
        results

End Function

Private Function IsEquivalentAccountAmbiguity( _
    clusterShape As String, _
    remainingMemberCount As Long, _
    remainingDebitCount As Long, _
    remainingCreditCount As Long, _
    maxMatching As Long, _
    perfectMatchings As Long) As Boolean

    '====================================================
    ' Recognize balanced ONE_TO_MANY or MANY_TO_ONE
    ' clusters where multiple row-level assignments exist,
    ' but every assignment produces the same account-level
    ' relationships.
    '====================================================

    IsEquivalentAccountAmbiguity = _
        False

    ' There must be an equal number of debit and credit
    ' transactions before all rows can be consumed.

    If remainingDebitCount <> _
       remainingCreditCount Then

        Exit Function

    End If

    ' The graph must be capable of matching every residual
    ' transaction.

    If maxMatching * 2 <> _
       remainingMemberCount Then

        Exit Function

    End If

    ' This rule is specifically for multiple mathematically
    ' valid row-level solutions.

    If perfectMatchings <= 1 Then

        Exit Function

    End If

    ' ONE_TO_MANY and MANY_TO_ONE have only one distinct
    ' account on one side. Therefore varying the individual
    ' row assignments does not change the account-level
    ' outcome.

    Select Case UCase$(Trim$(clusterShape))

        Case "ONE_TO_MANY", _
             "MANY_TO_ONE"

            IsEquivalentAccountAmbiguity = _
                True

    End Select

End Function


Private Function BuildEquivalentPairings( _
    members As Collection, _
    ambiguityPairs As Collection, _
    data As Variant, _
    colAmount As Long) As Collection

    Dim results As New Collection

    Dim debitRows As New Collection
    Dim creditRows As New Collection

    Dim memberLookup As Object
    Dim usedDebits As Object
    Dim usedCredits As Object

    Dim rowNum As Variant
    Dim pairText As Variant
    Dim parts() As String

    Dim rowA As Long
    Dim rowB As Long

    Dim debitRow As Long
    Dim creditRow As Long
    Dim pairKey As String

    Set memberLookup = _
        CreateObject("Scripting.Dictionary")

    Set usedDebits = _
        CreateObject("Scripting.Dictionary")

    Set usedCredits = _
        CreateObject("Scripting.Dictionary")

    '====================================================
    ' BUILD RESIDUAL MEMBER LOOKUP
    '====================================================

    For Each rowNum In members

        memberLookup(CStr(rowNum)) = _
            True

    Next rowNum

    '====================================================
    ' Any valid matching is functionally equivalent for
    ' ONE_TO_MANY or MANY_TO_ONE clusters. We take the first
    ' available valid edge for each unused transaction.
    '====================================================

    For Each pairText In ambiguityPairs

        parts = _
            Split(CStr(pairText), "|")

        If UBound(parts) >= 1 Then

            rowA = _
                CLng(parts(0))

            rowB = _
                CLng(parts(1))

            If memberLookup.Exists(CStr(rowA)) _
            And memberLookup.Exists(CStr(rowB)) Then

                If CDbl(data(rowA, colAmount)) < 0 Then

                    debitRow = rowA
                    creditRow = rowB

                Else

                    debitRow = rowB
                    creditRow = rowA

                End If

                If Not usedDebits.Exists(CStr(debitRow)) _
                And Not usedCredits.Exists(CStr(creditRow)) Then

                    pairKey = _
                        CStr(debitRow) & "|" & _
                        CStr(creditRow)

                    results.Add _
                        pairKey

                    usedDebits.Add _
                        CStr(debitRow), _
                        True

                    usedCredits.Add _
                        CStr(creditRow), _
                        True

                End If

            End If

        End If

    Next pairText

    Set BuildEquivalentPairings = _
        results

End Function


Private Sub PreviewRepeatedTransferCluster( _
    ByVal hostWb As Workbook, _
    ByVal clusterID As String, _
    ByVal investigationKey As String, _
    ByVal investigationGroup As String, _
    ByVal clusterShape As String, _
    ByVal outcome As String, _
    ByVal members As Collection, _
    ByRef data As Variant, _
    ByVal colAcct As Long, _
    ByVal colAmount As Long)


    Dim debits As Collection
    Dim credits As Collection

    Dim rowNum As Variant
    
    
    Set debits = New Collection
    Set credits = New Collection

    For Each rowNum In members

        If CDbl(data(CLng(rowNum), colAmount)) < 0 Then

            debits.Add CLng(rowNum)

        Else

            credits.Add CLng(rowNum)

        End If

    Next rowNum

    If debits.Count <> credits.Count Then Exit Sub

    Dim i As Long

'    For i = 1 To debits.Count
'
'        WriteResolutionPreview _
'            clusterID, _
'            "", _
'            "ONE_TO_ONE_REPEATED", _
'            "MATCHED", _
'            CLng(debits(i)), _
'            CLng(credits(i)), _
'            data, _
'            colAcct, _
'            colAmount
'
'    Next i

    
    If members.Count = 2 Then
    
        clusterShape = "ONE_TO_ONE"
    
    Else
    
        clusterShape = "ONE_TO_ONE_REPEATED"
    
    End If
    
    For i = 1 To debits.Count
    
        WriteResolutionPreview _
            hostWb, _
            clusterID, _
            investigationKey, _
            investigationGroup, _
            clusterShape, _
            "MATCHED", _
            CLng(debits(i)), _
            CLng(credits(i)), _
            data, _
            colAcct, _
            colAmount
    
    Next i

End Sub

Private Function HasStrongResolutionEvidence( _
    members As Collection, _
    data As Variant, _
    colAcct As Long) As Boolean

    Dim rowNum As Variant
    Dim targetAcct As String

    Dim evidenceMatches As Long

    For Each rowNum In members

        targetAcct = ""

        If ReferencedAccounts(CLng(rowNum)) <> "" Then

            targetAcct = _
                ReferencedAccounts(CLng(rowNum))

        ElseIf EmbeddedAccounts(CLng(rowNum)) <> "" Then

            targetAcct = _
                EmbeddedAccounts(CLng(rowNum))

        End If

        If targetAcct <> "" Then

            If ClusterContainsAccount( _
                members, _
                data, _
                colAcct, _
                targetAcct) Then

                evidenceMatches = _
                    evidenceMatches + 1

            End If

        End If
        

    Next rowNum

    HasStrongResolutionEvidence = _
        (evidenceMatches > 0)

End Function

Private Function HasResolvableSuffixEvidence( _
    members As Collection, _
    data As Variant, _
    colAcct As Long, _
    colAmount As Long) As Boolean

    Dim rowNum As Variant

    Dim fromSuffix As String
    Dim toSuffix As String

    Dim creditCount As Long

    Dim acct As Variant

    For Each rowNum In members

        If CDbl(data(CLng(rowNum), colAmount)) < 0 Then

            fromSuffix = ""
            toSuffix = ""

            If GetTransferSuffixes( _
                CStr(data(CLng(rowNum))), _
                fromSuffix, _
                toSuffix) Then

                creditCount = 0

                For Each acct In GetClusterAccountCollection( _
                    members, _
                    data, _
                    colAcct)

                    If Right( _
                        CStr(acct), _
                        Len(toSuffix)) = toSuffix Then

                        creditCount = _
                            creditCount + 1

                    End If
                    


                Next acct

                If creditCount <> 1 Then

                    Exit Function

                End If

            End If

        End If

    Next rowNum

    HasResolvableSuffixEvidence = True

End Function

Private Function GetClusterAccountCollection( _
    members As Collection, _
    data As Variant, _
    colAcct As Long) As Collection

    Dim results As New Collection

    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")

    Dim rowNum As Variant
    Dim acct As String

    For Each rowNum In members

        acct = Trim( _
            CStr(data(CLng(rowNum), colAcct)))

        If acct <> "" Then

            If Not seen.Exists(acct) Then

                seen.Add acct, True

                results.Add acct

            End If

        End If

    Next rowNum

    Set GetClusterAccountCollection = results

End Function

Private Function ClusterContainsAccount( _
    members As Collection, _
    data As Variant, _
    colAcct As Long, _
    targetAcct As String) As Boolean

    Dim rowNum As Variant

    For Each rowNum In members

        If Right( _
            Trim(CStr(data(CLng(rowNum), colAcct))), _
            Len(targetAcct)) = targetAcct Then

            ClusterContainsAccount = True
            Exit Function

        End If

    Next rowNum

End Function


Private Sub AssignInvestigationGroups( _
    ByVal hostWb As Workbook)

    Dim ws As Worksheet

    Dim lastRow As Long
    Dim r As Long

    Dim groupLookup As Object

    Dim investigationKey As String

    Dim nextGroup As Long

    Set groupLookup = _
        CreateObject("Scripting.Dictionary")

    Set ws = _
        hostWb.Worksheets("Cluster_Analysis")

    lastRow = _
        ws.Cells(ws.rows.Count, 1).End(xlUp).Row

    nextGroup = 1

    For r = 2 To lastRow

        investigationKey = _
            CStr(ws.Cells(r, 18).Value)

        If Not groupLookup.Exists( _
                investigationKey) Then

            groupLookup.Add _
                investigationKey, _
                "G" & _
                Format(nextGroup, "000000")

            nextGroup = nextGroup + 1

        End If

        ws.Cells(r, 19).Value = _
            groupLookup(investigationKey)

    Next r

End Sub

Private Function CountPerfectMatchings( _
    members As Collection, _
    ambiguityPairs As Collection, _
    data As Variant, _
    colAmount As Long, _
    colAcct As Long) As Long

    Dim debits As Collection
    Dim credits As Collection

    Dim rowNum As Variant

    Set debits = New Collection
    Set credits = New Collection

    For Each rowNum In members

        If CDbl(data(CLng(rowNum), colAmount)) < 0 Then

            debits.Add CLng(rowNum)

        Else

            credits.Add CLng(rowNum)

        End If

    Next rowNum

    If debits.Count <> 2 Then
    
        CountPerfectMatchings = -1
        Exit Function
    
    End If
    
    If credits.Count <> 2 Then
    
        CountPerfectMatchings = -1
        Exit Function
    
    End If

    Dim D1 As Long
    Dim D2 As Long

    Dim C1 As Long
    Dim C2 As Long

    D1 = debits(1)
    D2 = debits(2)

    C1 = credits(1)
    C2 = credits(2)

    Dim E As Object
    Set E = CreateObject("Scripting.Dictionary")

    Dim pairText As Variant
    Dim parts() As String

    Dim rowA As Long
    Dim rowB As Long
    
    Dim debitRow As Long
    Dim creditRow As Long
    

    For Each pairText In ambiguityPairs

        parts = Split(CStr(pairText), "|")

        rowA = CLng(parts(0))
        rowB = CLng(parts(1))

        If CDbl(data(rowA, colAmount)) < 0 Then
        
            debitRow = rowA
            creditRow = rowB
        
        Else
        
            debitRow = rowB
            creditRow = rowA
        
        End If
        
If EdgeSupportedByEvidence( _
    debitRow, _
    creditRow, _
    CStr(data(debitRow, colAcct)), _
    CStr(data(creditRow, colAcct))) Then
        
            E(CStr(debitRow) & "|" & _
              CStr(creditRow)) = True
        
        End If

    Next pairText

    ' Solution 1
    If E.Exists(CStr(D1) & "|" & CStr(C1)) _
    And E.Exists(CStr(D2) & "|" & CStr(C2)) Then

        CountPerfectMatchings = _
            CountPerfectMatchings + 1

    End If

    ' Solution 2
    If E.Exists(CStr(D1) & "|" & CStr(C2)) _
    And E.Exists(CStr(D2) & "|" & CStr(C1)) Then

        CountPerfectMatchings = _
            CountPerfectMatchings + 1

    End If

End Function

Private Function HasUniqueSolution( _
    members As Collection, _
    ambiguityPairs As Collection, _
    data As Variant, _
    colAmount As Long) As Boolean
    
    
    
    End Function
    
Private Function GetClusterShape( _
    debitAcctCount As Long, _
    creditAcctCount As Long, _
    Optional debitTxnCount As Long = 0, _
    Optional creditTxnCount As Long = 0) As String

    '========================================================
    ' Classify a cluster using:
    '
    '   1. Number of distinct debit accounts
    '   2. Number of distinct credit accounts
    '   3. Number of debit transactions
    '   4. Number of credit transactions
    '
    ' Transaction counts are necessary to distinguish a true
    ' ONE_TO_ONE cluster from a repeated A-to-B cluster.
    '========================================================

    '--------------------------------------------------------
    ' Exactly one debit transaction and one credit transaction
    '--------------------------------------------------------

    If debitTxnCount = 1 _
    And creditTxnCount = 1 Then

        GetClusterShape = _
            "ONE_TO_ONE"

        Exit Function

    End If

    '--------------------------------------------------------
    ' Multiple transactions between one account on each side
    '
    ' Examples:
    '
    '   A -> B
    '   A -> B
    '
    ' or:
    '
    '   A -> B repeated three times
    '--------------------------------------------------------

    If debitAcctCount = 1 _
    And creditAcctCount = 1 _
    And debitTxnCount > 1 _
    And creditTxnCount > 1 Then

        GetClusterShape = _
            "ONE_TO_ONE_REPEATED"

        Exit Function

    End If

    '--------------------------------------------------------
    ' Multiple debit accounts transferring to one credit
    ' account
    '--------------------------------------------------------

    If debitAcctCount > 1 _
    And creditAcctCount = 1 Then

        GetClusterShape = _
            "MANY_TO_ONE"

        Exit Function

    End If

    '--------------------------------------------------------
    ' One debit account transferring to multiple credit
    ' accounts
    '--------------------------------------------------------

    If debitAcctCount = 1 _
    And creditAcctCount > 1 Then

        GetClusterShape = _
            "ONE_TO_MANY"

        Exit Function

    End If

    '--------------------------------------------------------
    ' Multiple accounts exist on both sides
    '--------------------------------------------------------

    If debitAcctCount > 1 _
    And creditAcctCount > 1 Then

        GetClusterShape = _
            "MANY_TO_MANY"

        Exit Function

    End If

    '--------------------------------------------------------
    ' Defensive fallback
    '
    ' This can occur if one side has no recognized accounts.
    '--------------------------------------------------------

    GetClusterShape = _
        "REVIEW"

End Function

Private Function ExtractResolvableEvidencePairs( _
    ambiguityPairs As Collection, _
    data As Variant, _
    colAcct As Long) As Collection

    Dim results As New Collection

    Dim pairText As Variant
    Dim parts() As String

    Dim debitRow As Long
    Dim creditRow As Long

    Dim debitAcct As String
    Dim creditAcct As String
    
    Dim seenPairs As Object
    Dim pairKey As String
    
    Dim debitCounts As Object
    Dim creditCounts As Object
    
    Dim candidatePairs As Collection
    Dim candidateText As Variant
    
    Set debitCounts = _
        CreateObject("Scripting.Dictionary")
    
    Set creditCounts = _
        CreateObject("Scripting.Dictionary")

    Set seenPairs = _
        CreateObject("Scripting.Dictionary")
        

    Set candidatePairs = New Collection
    
    
    For Each pairText In ambiguityPairs
    

        parts = Split(CStr(pairText), "|")

        If UBound(parts) >= 1 Then

        debitRow = CLng(parts(0))
        creditRow = CLng(parts(1))

        debitAcct = _
            CStr(data(debitRow, colAcct))

        creditAcct = _
            CStr(data(creditRow, colAcct))

        If EdgeSupportedByEvidence( _
            debitRow, _
            creditRow, _
            debitAcct, _
            creditAcct, _
            True) Then
            
                pairKey = _
                    CStr(debitRow) & "|" & _
                    CStr(creditRow)
            
                candidatePairs.Add pairKey
            
                If Not debitCounts.Exists( _
                    CStr(debitRow)) Then
            
                    debitCounts( _
                        CStr(debitRow)) = 0
            
                End If
            
                debitCounts( _
                    CStr(debitRow)) = _
                    debitCounts( _
                        CStr(debitRow)) + 1
            
                If Not creditCounts.Exists( _
                    CStr(creditRow)) Then
            
                    creditCounts( _
                        CStr(creditRow)) = 0
            
                End If
            
                creditCounts( _
                    CStr(creditRow)) = _
                    creditCounts( _
                        CStr(creditRow)) + 1
            
            End If
        
        End If
            
Next pairText
            

For Each candidateText In candidatePairs

    parts = Split( _
        CStr(candidateText), _
        "|")

    If UBound(parts) < 1 Then


        GoTo NextCandidate

    End If

    debitRow = _
        CLng(parts(0))

    creditRow = _
        CLng(parts(1))

    If debitCounts( _
        CStr(debitRow)) = 1 _
    And creditCounts( _
        CStr(creditRow)) = 1 Then

        results.Add candidateText

    End If

NextCandidate:

Next candidateText



    Set ExtractResolvableEvidencePairs = _
        results
        
       
        
End Function

Private Function RemoveResolvedMembers( _
    members As Collection, _
    resolvedPairs As Collection) As Collection

    Dim remaining As New Collection

    Dim resolvedRows As Object
    Set resolvedRows = _
        CreateObject("Scripting.Dictionary")

    Dim pairText As Variant
    Dim parts() As String

    Dim rowNum As Variant

    '--------------------------
    ' Mark resolved rows
    '--------------------------

    For Each pairText In resolvedPairs
    
        parts = Split(CStr(pairText), "|")
    
        If UBound(parts) >= 1 Then
    
            resolvedRows(CStr(parts(0))) = True
            resolvedRows(CStr(parts(1))) = True
    
        Else
    
    
        End If
    
    Next pairText

    '--------------------------
    ' Build residual collection
    '--------------------------

    For Each rowNum In members

        If Not resolvedRows.Exists( _
            CStr(rowNum)) Then

            remaining.Add rowNum

        End If

    Next rowNum

    Set RemoveResolvedMembers = _
        remaining

End Function


Private Function EdgeSupportedByEvidence( _
    debitRow As Long, _
    creditRow As Long, _
    debitAcct As String, _
    creditAcct As String, _
    Optional requireEvidence As Boolean = False) As Boolean

    Dim fromSuffix As String
    Dim toSuffix As String
    Dim evidence As String
    


    ' Transfer Suffix

    If GetEdgeSuffixEvidence( _
        debitRow, _
        creditRow, _
        fromSuffix, _
        toSuffix) Then
    
        Dim debitMatchesFrom As Boolean
        Dim creditMatchesTo As Boolean
    
        debitMatchesFrom = _
            AccountEndsWith( _
                debitAcct, _
                fromSuffix)
    
        creditMatchesTo = _
            AccountEndsWith( _
                creditAcct, _
                toSuffix)
    
    
        EdgeSupportedByEvidence = _
            debitMatchesFrom _
            And creditMatchesTo
    
    
        Exit Function
    
    End If



    ' Single Suffix Evidence
    
    
    Dim debitSuffix As String
    Dim creditSuffix As String
    
    If GetSingleSuffixEvidence( _
        debitRow, _
        creditRow, _
        debitSuffix, _
        creditSuffix) Then
    
        EdgeSupportedByEvidence = True
    
        If debitSuffix <> "" Then
    
            EdgeSupportedByEvidence = _
                EdgeSupportedByEvidence _
                And _
                AccountEndsWith( _
                    creditAcct, _
                    debitSuffix)
    
        End If
    
        If creditSuffix <> "" Then
    
            EdgeSupportedByEvidence = _
                EdgeSupportedByEvidence _
                And _
                AccountEndsWith( _
                    debitAcct, _
                    creditSuffix)
    
        End If
    
    
        Exit Function
    
    End If


    ' Credit-side Referenced Acct


    evidence = Trim$( _
        CStr(ReferencedAccounts(creditRow)))
    
    If evidence <> "" Then
    
    
        EdgeSupportedByEvidence = _
            (debitAcct = evidence)
            
        Exit Function
    
    End If
    
    
        ' Debit-side Referenced Acct
    
    evidence = Trim$( _
        CStr(ReferencedAccounts(debitRow)))
    
    If evidence <> "" Then
    
    
        EdgeSupportedByEvidence = _
            (creditAcct = evidence)
    
        Exit Function
    
    End If


    EdgeSupportedByEvidence = False
    
    If requireEvidence Then
    
        ' Used by ExtractResolvableEvidencePairs.
        ' No evidence means this is not an extractable pair.
    
        EdgeSupportedByEvidence = _
            False
    
    Else
    
        ' Used by the normal residual graph.
        ' No contradictory evidence means the candidate edge
        ' remains available to structural ambiguity analysis.
    
        EdgeSupportedByEvidence = _
            True
    
    End If


    End Function
    
    '---This may be re-enabled later, but is working well without,
    '   and will inject even more unnecessary complexity.---
''
''    '--------------------------
''    ' Credit-side Embedded Acct
''    '--------------------------
''
''    If Trim$( _
''        CStr(EmbeddedAccounts(creditRow))) <> "" Then
''
''        evidence = _
''            Trim$( _
''                CStr(EmbeddedAccounts(creditRow)))
''
''        Debug.Print _
''            "CREDIT EMBED", _
''            creditRow, _
''            evidence
''
''        EdgeSupportedByEvidence = _
''            (debitAcct = evidence)
''
''        Exit Function
''
''    End If
''
''    '--------------------------
''    ' Debit-side Embedded Acct
''    '--------------------------
''
''    If Trim$( _
''        CStr(EmbeddedAccounts(debitRow))) <> "" Then
''
''        evidence = _
''            Trim$( _
''                CStr(EmbeddedAccounts(debitRow)))
''
''        Debug.Print _
''            "DEBIT EMBED", _
''            debitRow, _
''            evidence
''
''        EdgeSupportedByEvidence = _
''            (creditAcct = evidence)
''
''        Exit Function
''
''    End If

'End Function


Private Function GetEdgeSuffixEvidence( _
    debitRow As Long, _
    creditRow As Long, _
    ByRef fromSuffix As String, _
    ByRef toSuffix As String) As Boolean

    Dim parts() As String
    Dim suffixText As String

    fromSuffix = ""
    toSuffix = ""


    If TransferToSuffixes.Exists(debitRow) Then

        suffixText = _
            CStr(TransferToSuffixes(debitRow))


    ElseIf TransferToSuffixes.Exists(creditRow) Then

        suffixText = _
            CStr(TransferToSuffixes(creditRow))

    Else

 
        Exit Function

    End If

    parts = _
        Split(suffixText, "|")

    If UBound(parts) <> 1 Then

        Debug.Print _
            "INVALID SUFFIX FORMAT =", _
            suffixText

        Exit Function

    End If

    fromSuffix = _
        Trim$(parts(0))

    toSuffix = _
        Trim$(parts(1))


    GetEdgeSuffixEvidence = _
        True

End Function

Private Function GetSingleSuffix( _
    txt As String) As String

    Dim RE As Object
    Dim matches As Object

    Set RE = CreateObject("VBScript.RegExp")

    RE.pattern = _
        "x(\d{4})"

    RE.IgnoreCase = True
    RE.Global = False

    If RE.Test(txt) Then

        Set matches = RE.Execute(txt)

        GetSingleSuffix = _
            matches(0).SubMatches(0)

    End If

End Function

Private Function GetSingleSuffixEvidence( _
    debitRow As Long, _
    creditRow As Long, _
    ByRef debitSuffix As String, _
    ByRef creditSuffix As String) As Boolean

    If SingleSuffixes.Exists(debitRow) Then

        debitSuffix = _
            CStr(SingleSuffixes(debitRow))

    End If

    If SingleSuffixes.Exists(creditRow) Then

        creditSuffix = _
            CStr(SingleSuffixes(creditRow))

    End If

    GetSingleSuffixEvidence = _
        (debitSuffix <> "" Or creditSuffix <> "")

End Function

Private Function ExtractEvidencePairings( _
    ambiguityPairs As Collection, _
    data As Variant, _
    colAcct As Long) As Collection

    Dim results As New Collection

    Dim pairText As Variant

    Dim parts() As String

    Dim debitRow As Long
    Dim creditRow As Long

    For Each pairText In ambiguityPairs

        parts = Split(CStr(pairText), "|")

        debitRow = CLng(parts(0))
        creditRow = CLng(parts(1))

        If EdgeSupportedByEvidence( _
            debitRow, _
            creditRow, _
            CStr(data(debitRow, colAcct)), _
            CStr(data(creditRow, colAcct))) Then

            results.Add pairText

        End If

    Next pairText

    Set ExtractEvidencePairings = results

End Function

Private Function ClusterContainsStrongEvidence( _
    members As Collection) As Boolean

    Dim rowNum As Variant
    

    For Each rowNum In members

            If TransferToSuffixes.Exists(CLng(rowNum)) Then
            

            If Trim(CStr( _
                TransferToSuffixes(CLng(rowNum)))) <> "" Then

                ClusterContainsStrongEvidence = True
                Exit Function

            End If

        End If

            If Trim(CStr(ReferencedAccounts(CLng(rowNum)))) <> "" Then
            
                ClusterContainsStrongEvidence = True
                Exit Function
            
            End If


            If Trim(CStr(EmbeddedAccounts(CLng(rowNum)))) <> "" Then
            
                ClusterContainsStrongEvidence = True
                Exit Function
            
            End If


    Next rowNum

End Function

Private Function EdgeSupportedBySuffix( _
    debitRow As Long, _
    creditAcct As String) As Boolean

    Dim suffix As String

    If Not TransferToSuffixes.Exists(debitRow) Then

        EdgeSupportedBySuffix = True
        Exit Function

    End If

    suffix = _
        Replace( _
            CStr(TransferToSuffixes(debitRow)), _
            " ", "")

    creditAcct = _
        Replace(creditAcct, " ", "")

    EdgeSupportedBySuffix = _
        (Right(creditAcct, Len(suffix)) = suffix)

End Function
    
    
Private Function GetDebitToSuffix( _
    rowNum As Long, _
    data As Variant)

    Dim fromSuffix As String
    Dim toSuffix As String

    If GetTransferSuffixes( _
        CStr(data(rowNum)), _
        fromSuffix, _
        toSuffix) Then

        GetDebitToSuffix = _
            Trim(toSuffix)

    End If

End Function

'hostWb refactoring completed
Private Function AccountMatchesSuffix( _
    acct As String, _
    suffix As String) As Boolean

    acct = _
        Replace(acct, " ", "")

    suffix = _
        Replace(suffix, " ", "")

    AccountMatchesSuffix = _
        (Right(acct, Len(suffix)) = suffix)

End Function

Private Sub BuildRelationshipSupport( _
    relDict As Object, _
    debits As Object, _
    credits As Object)

    Dim debitAcct As Variant
    Dim creditAcct As Variant

    Dim supportCount As Long
    Dim key As String

    For Each debitAcct In debits.Keys

        For Each creditAcct In credits.Keys

            supportCount = _
                WorksheetFunction.Min( _
                    CLng(debits(debitAcct)), _
                    CLng(credits(creditAcct)))

            If supportCount > 0 Then

                key = debitAcct & "|" & creditAcct

                relDict(key) = supportCount

            End If

        Next creditAcct

    Next debitAcct

End Sub


Private Function WriteWarnings( _
    ws As Worksheet, _
    reportRow As Long, _
    warningText As String) As Long

    ws.Cells(reportRow, 1).Value = "Warnings"
    ws.Cells(reportRow, 1).Font.Bold = True

    reportRow = reportRow + 1

    ws.Cells(reportRow, 1).Value = warningText

    WriteWarnings = reportRow + 1

End Function



Private Function GetLargestClusterAccountsID() As String

    Dim ws As Worksheet
    Dim clusterDict As Object
    Dim accountDict As Object

    Dim clusterID As String
    Dim acct As String

    Dim lastRow As Long
    Dim r As Long

    Dim key As Variant
    Dim largestCount As Long

    Set ws = hostWb.Worksheets("Transfer_Ambiguity_Members")
    Set clusterDict = CreateObject("Scripting.Dictionary")

    lastRow = _
        ws.Cells(ws.rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow

        clusterID = CStr(ws.Cells(r, 2).Value)
        acct = CStr(ws.Cells(r, 4).Value)

        If Not clusterDict.Exists(clusterID) Then

            Set accountDict = _
                CreateObject("Scripting.Dictionary")

            clusterDict.Add _
                clusterID, _
                accountDict

        End If

        clusterDict(clusterID)(acct) = 1

    Next r

    For Each key In clusterDict.Keys

        If clusterDict(key).Count > largestCount Then

            largestCount = _
                clusterDict(key).Count

            GetLargestClusterAccountsID = _
                CStr(key)

        End If

    Next key

End Function


Private Function GetLargestClusterBySize() As String

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Dim maxSize As Long
    Dim clusterID As String

    Set ws = hostWb.Worksheets("Transfer_Ambiguities")

    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow

        If CLng(ws.Cells(r, 3).Value) > maxSize Then

            maxSize = CLng(ws.Cells(r, 3).Value)

            clusterID = ws.Cells(r, 2).Value

        End If

    Next r

    GetLargestClusterBySize = clusterID

End Function


Private Sub AddAmbiguity( _
    ambiguityPairs As Collection, _
    rowA As Long, _
    rowB As Long, _
    matchMethod As String)

    Dim edgeKey As String

    ambiguityPairs.Add _
        CStr(rowA) & "|" & _
        CStr(rowB) & "|" & _
        matchMethod

    edgeKey = _
        CStr(rowA) & "|" & _
        CStr(rowB)

    If Not AmbiguousEdges.Exists(edgeKey) Then

        AmbiguousEdges.Add _
            edgeKey, _
            matchMethod

    End If

End Sub




Private Sub WriteClusterSummary( _
    ByVal hostWb As Workbook, _
    ByVal ws As Worksheet, _
    ByRef reportRow As Long, _
    ByVal clusterID As String, _
    ByVal members As Collection)

    Dim txnWs As Worksheet

    Dim rowNum As Variant
    Dim txnRow As Long

    Dim activityDate As Variant
    Dim totalVolume As Double

    Dim txnCount As Long

    Dim creditAccounts As Object
    Dim debitAccounts As Object

    Set txnWs = hostWb.Worksheets("Transfer_Transactions")

    Set creditAccounts = CreateObject("Scripting.Dictionary")
    Set debitAccounts = CreateObject("Scripting.Dictionary")

    txnCount = members.Count

    For Each rowNum In members

        txnRow = FindTransactionRow( _
            hostWb, _
            CLng(rowNum))

        If txnRow > 0 Then

            If IsEmpty(activityDate) Then
                activityDate = txnWs.Cells(txnRow, 4).Value
            End If

            totalVolume = totalVolume + _
                Abs(CDbl(txnWs.Cells(txnRow, 5).Value))

            If CDbl(txnWs.Cells(txnRow, 5).Value) > 0 Then

                If Not creditAccounts.Exists(CStr(txnWs.Cells(txnRow, 3).Value)) Then
                    creditAccounts.Add CStr(txnWs.Cells(txnRow, 3).Value), True
                End If

            Else

                If Not debitAccounts.Exists(CStr(txnWs.Cells(txnRow, 3).Value)) Then
                    debitAccounts.Add CStr(txnWs.Cells(txnRow, 3).Value), True
                End If

            End If

        End If

    Next rowNum

    With ws.Range(ws.Cells(reportRow, 1), ws.Cells(reportRow, 6))
        .Merge
        .Value = "Cluster " & clusterID
        .Font.Bold = True
        .Interior.Color = RGB(217, 225, 242)
    End With

    reportRow = reportRow + 1

    ws.Cells(reportRow, 1).Value = _
        "Date: " & Format(activityDate, "dd-mmm-yy")

    reportRow = reportRow + 1

    ws.Cells(reportRow, 1).Value = _
        "Transaction Count: " & txnCount

    ws.Cells(reportRow, 2).Value = _
        "Total Volume: " & Format(totalVolume, "$#,##0.00")

    reportRow = reportRow + 1

    ws.Cells(reportRow, 1).Value = _
        "Credit Accounts: " & creditAccounts.Count

    ws.Cells(reportRow, 2).Value = _
        "Debit Accounts: " & debitAccounts.Count

    ws.Cells(reportRow, 3).Value = _
        "Potential Relationships: " & _
        (creditAccounts.Count * debitAccounts.Count)
        
    

End Sub


Private Function BuildClusterMetadataLookup( _
    ByVal hostWb As Workbook) As Object

    Dim ws As Worksheet
    Dim dict As Object

    Dim lastRow As Long
    Dim r As Long

    Dim clusterID As String
    Dim investigationGroup As String
    Dim clusterShape As String
    Dim outcome As String
    Dim originMethod As String

    Set ws = _
        hostWb.Worksheets("Cluster_Analysis")

    Set dict = _
        CreateObject("Scripting.Dictionary")

    lastRow = _
        ws.Cells(ws.rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow

        clusterID = _
            CStr(ws.Cells(r, 1).Value)

        clusterShape = _
            CStr(ws.Cells(r, 12).Value)

        outcome = _
            CStr(ws.Cells(r, 14).Value)

        investigationGroup = _
            CStr(ws.Cells(r, 19).Value)

        If ClusterOriginMethods.Exists(clusterID) Then

            originMethod = _
                ClusterOriginMethods(clusterID)

        Else

            originMethod = ""

        End If

        dict(clusterID) = Array( _
            investigationGroup, _
            clusterShape, _
            outcome, _
            originMethod)

    Next r

    Set BuildClusterMetadataLookup = dict

End Function


Private Function BuildClusterReportingLookup( _
    ByVal hostWb As Workbook) As Object

    '========================================================
    ' Build cluster metadata once for use by:
    '
    '   BuildResolvedClusterReport
    '   BuildUnmatchedReport
    '
    ' Dictionary key:
    '   Cluster ID
    '
    ' Array:
    '   0 = Investigation Group
    '   1 = Recommended Outcome
    '   2 = Cluster Shape
    '   3 = Original Member Count
    '   4 = Distinct Account Count
    '   5 = Account List
    '   6 = Maximum Matching
    '   7 = Perfect Matchings
    '   8 = Fully Matchable
    '   9 = Investigation Key
    '========================================================

    Dim results As Object
    Dim srcWs As Worksheet

    Dim lastRow As Long
    Dim r As Long

    Dim clusterID As String

    Set results = _
        CreateObject("Scripting.Dictionary")

    Set srcWs = hostWb.Worksheets("Cluster_Analysis")

    lastRow = _
        srcWs.Cells( _
            srcWs.rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow

        clusterID = _
            Trim$(CStr(srcWs.Cells(r, 1).Value))

        If clusterID <> "" Then

            results(clusterID) = _
                Array( _
                    CStr(srcWs.Cells(r, 19).Value), _
                    CStr(srcWs.Cells(r, 14).Value), _
                    CStr(srcWs.Cells(r, 12).Value), _
                    CLng(srcWs.Cells(r, 2).Value), _
                    CLng(srcWs.Cells(r, 15).Value), _
                    CStr(srcWs.Cells(r, 16).Value), _
                    CLng(srcWs.Cells(r, 9).Value), _
                    CLng(srcWs.Cells(r, 13).Value), _
                    CStr(srcWs.Cells(r, 10).Value), _
                    CStr(srcWs.Cells(r, 18).Value))

        End If

    Next r

    Set BuildClusterReportingLookup = results

End Function


Private Function GetAmbiguousClusterCount() As Long

    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = hostWb.Worksheets("Transfer_Ambiguities")

    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row

    GetAmbiguousClusterCount = Application.Max(lastRow - 1, 0)

End Function




Private Function IsBranchDeposit( _
    codeDesc As String) As Boolean

    Select Case UCase(Trim(codeDesc))

        Case "10-D - POD - CREDIT/DEPOSIT"

            IsBranchDeposit = True

    End Select

End Function



Private Function GetConfidenceTier( _
    matchMethod As String) As String

    Select Case matchMethod

        Case "Confirmation Number", _
             "Narrative Pair", _
             "Referenced Account", _
             "Embedded Account"

            GetConfidenceTier = "High"

        Case "Transfer Suffix"

            GetConfidenceTier = "Medium"

        Case "Amount + Date"

            GetConfidenceTier = "Low"
            
        Case "Investigation Group Resolution"

            GetConfidenceTier = "System-Resolved"

        Case Else

            GetConfidenceTier = "Unknown"

    End Select

End Function


Private Sub InitializeStatistics()

    Set MethodCounts = _
        CreateObject("Scripting.Dictionary")

    Set MethodVolumes = _
        CreateObject("Scripting.Dictionary")
        
    Set AmbiguousMethodCounts = _
        CreateObject("Scripting.Dictionary")
        
    Set MatchIDs = _
        CreateObject("Scripting.Dictionary")
        
    Set TransactionStatus = _
        CreateObject("Scripting.Dictionary")
        
    Set CandidateLookup = _
        CreateObject("Scripting.Dictionary")
        
    Set ClusterOriginMethods = _
        CreateObject("Scripting.Dictionary")
        
        
    NextMatchID = 1
    NextContradictionID = 1
    NextCandidateID = 1

End Sub

Private Sub RecordMatchStatistic( _
    matchMethod As String, _
    amount As Double)

    If Not MethodCounts.Exists(matchMethod) Then

        MethodCounts.Add matchMethod, 0
        MethodVolumes.Add matchMethod, 0#

    End If

    MethodCounts(matchMethod) = _
        MethodCounts(matchMethod) + 1

    MethodVolumes(matchMethod) = _
        MethodVolumes(matchMethod) + Abs(amount)

End Sub

Private Sub RecordAmbiguityMethod( _
    matchMethod As String)

    If matchMethod = "" Then Exit Sub

    If Not AmbiguousMethodCounts.Exists(matchMethod) Then
        AmbiguousMethodCounts.Add matchMethod, 0
    End If

    AmbiguousMethodCounts(matchMethod) = _
        AmbiguousMethodCounts(matchMethod) + 1

End Sub

Private Function GetNextMatchID() As String

    GetNextMatchID = _
        "M" & Format(NextMatchID, "000000")

    NextMatchID = NextMatchID + 1

End Function

Private Function GetNextContradictionID() As String

    GetNextContradictionID = _
        "CN" & Format(NextContradictionID, "000000")

    NextContradictionID = _
        NextContradictionID + 1

End Function

Private Function GetNextCandidateID() As String

    GetNextCandidateID = _
        "CA" & Format(NextCandidateID, "000000")

    NextCandidateID = _
        NextCandidateID + 1

End Function

Private Function GetMatchID( _
    rowNum As Long) As String

    If MatchIDs.Exists(CStr(rowNum)) Then

        GetMatchID = _
            MatchIDs(CStr(rowNum))

    End If

End Function

Private Sub RegisterMatch( _
    rowA As Long, _
    rowB As Long, _
    matchID As String)

    MatchIDs(CStr(rowA)) = matchID
    MatchIDs(CStr(rowB)) = matchID

End Sub

Private Sub WriteTransferRelationshipHeaders( _
    ws As Worksheet)

    ws.Cells(1, 1).Value = "Match ID"
    ws.Cells(1, 2).Value = "Date"

    ws.Cells(1, 3).Value = "Debit Account"
    ws.Cells(1, 4).Value = "Credit Account"

    ws.Cells(1, 5).Value = "Amount"

    ws.Cells(1, 6).Value = "Match Method"
    ws.Cells(1, 7).Value = "Confidence"

    ws.Cells(1, 8).Value = "Source Row"
    ws.Cells(1, 9).Value = "Candidate Row"
    
    ws.Cells(1, 10).Value = "Created Timestamp"
    
    ws.Cells(1, 11).Value = "Cluster ID"
    
    ws.Cells(1, 12).Value = "Cluster Shape"
    
    ws.Cells(1, 13).Value = "Investigation Group"

    
    ws.rows(1).Font.Bold = True
    ws.rows(1).AutoFilter

End Sub


Private Function GetAccountRows( _
    acctNumber As String) As Collection

    If AccountIndex.Exists(Trim(acctNumber)) Then

        Set GetAccountRows = _
            AccountIndex(Trim(acctNumber))

    End If

End Function


Private Function FindTransactionRow( _
    ByVal hostWb As Workbook, _
    ByVal rowNumber As Long) As Long

    If TransactionRowLookup Is Nothing Then
        BuildTransactionRowLookup hostWb
    End If

    If TransactionRowLookup.Exists(rowNumber) Then

        FindTransactionRow = _
            TransactionRowLookup(rowNumber)

    Else

        FindTransactionRow = 0

    End If

End Function



Private Function GetAccountTransactionCount( _
    acctNumber As String) As Long

    If AccountIndex.Exists(Trim(acctNumber)) Then

        GetAccountTransactionCount = _
            AccountIndex(Trim(acctNumber)).Count

    End If

End Function

Private Function GetMatchedVolume() As Double

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = hostWb.Worksheets("Transfer_Relationships")

    lastRow = ws.Cells( _
        ws.rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow

        GetMatchedVolume = _
            GetMatchedVolume + _
            CDbl(ws.Cells(r, 5).Value)

    Next r

End Function

Private Function GetUnmatchedVolume() As Double

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = hostWb.Worksheets("Transfer_Transaction_Status")

    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow

        If ws.Cells(r, 6).Value = STATUS_UNMATCHED Then

            GetUnmatchedVolume = _
                GetUnmatchedVolume + _
                Abs(CDbl(ws.Cells(r, 5).Value))

        End If

    Next r

End Function

Private Function GetTransfersRequiringReviewVolume() As Double

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = hostWb.Worksheets("Transfer_Transaction_Status")

    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow

        If ws.Cells(r, 6).Value = STATUS_UNMATCHED _
        Or ws.Cells(r, 6).Value = STATUS_AMBIGUOUS Then

            GetTransfersRequiringReviewVolume = _
                GetTransfersRequiringReviewVolume + _
                Abs(CDbl(ws.Cells(r, 5).Value))

        End If

    Next r

End Function


Private Function GetTotalTransferTransactions() As Long

    GetTotalTransferTransactions = _
        GetStatusCount(STATUS_MATCHED) + _
        GetStatusCount(STATUS_AMBIGUOUS) + _
        GetStatusCount(STATUS_UNMATCHED)

End Function


Private Function GetVolumeByStatus( _
    targetStatus As String) As Double

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = hostWb.Worksheets("Transfer_Transactions")

    lastRow = _
        ws.Cells(ws.rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow

        If ws.Cells(r, 8).Value = targetStatus Then

            GetVolumeByStatus = _
                GetVolumeByStatus + _
                Abs(CDbl(ws.Cells(r, 5).Value))

        End If

    Next r

End Function


Private Function AccountExists( _
    acctNumber As String) As Boolean

    AccountExists = _
        AccountIndex.Exists(Trim(acctNumber))

End Function

Private Sub WriteTopMatchedTransfers( _
    ByVal hostWb As Workbook, _
    ByVal ws As Worksheet, _
    ByVal startRow As Long, _
    ByVal startCol As Long)

    Dim relWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim r As Long
    Dim sectionLastCol As Long

    Set relWs = hostWb.Worksheets("Transfer_Relationships")
    
    sectionLastCol = startCol + 3

    lastRow = relWs.Cells( _
        relWs.rows.Count, 1).End(xlUp).Row
    
    Dim matchCount As Long

        matchCount = lastRow - 1    'subtract header row
        
        If matchCount = 0 Then
        
        With ws.Range( _
            ws.Cells(startRow, startCol), _
            ws.Cells(startRow, sectionLastCol))
        
            .Merge
        
            .Value = _
                "TOP MATCHED TRANSFERS"
        
            .Font.Bold = True
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        
            With .Borders(xlEdgeBottom)
        
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = vbBlack
        
            End With
        
        End With
        
        With ws.Range( _
            ws.Cells(startRow + 1, startCol), _
            ws.Cells(startRow + 1, sectionLastCol))
        
            .Merge
        
            .Value = _
                "No matches found."
        
            .Interior.Color = _
                RGB(255, 0, 0)
        
            .Font.Color = _
                RGB(255, 255, 255)
        
            .Font.Bold = True
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        
            With .Borders(xlEdgeBottom)
        
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = vbBlack
        
            End With
        
        End With
        
        Exit Sub
        
        End If
    
    If lastRow <= 1 Then Exit Sub

    relWs.Sort.SortFields.Clear

    relWs.Sort.SortFields.Add _
        key:=relWs.Columns(5), _
        Order:=xlDescending
        

    With relWs.Sort
    
        .SetRange relWs.Range("A1:M" & lastRow)
    
        .header = xlYes
    
        .Apply
    
    End With

    With ws.Range( _
        ws.Cells(startRow, startCol), _
        ws.Cells(startRow, sectionLastCol))
    
        .Merge
    
        .Value = _
            "TOP MATCHED TRANSFERS"
    
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    
        With .Borders(xlEdgeBottom)
    
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = vbBlack
    
        End With
    
    End With

    ws.Cells(startRow + 1, startCol).Resize(1, 4).Value = _
        Array("From Account", _
              "To Account", _
              "Date", _
              "Amount")
              
    With ws.Range( _
        ws.Cells(startRow + 1, startCol), _
        ws.Cells(startRow + 1, sectionLastCol))
    
        .Font.Bold = True
    
        With .Borders(xlEdgeBottom)
    
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(180, 180, 180)
    
        End With
    
    End With

    r = startRow + 2

    For i = 2 To WorksheetFunction.Min(lastRow, 4)
    
        ws.Cells(r, startCol).Value = _
            "'" & CStr(relWs.Cells(i, 3).Value)
    
        ws.Cells(r, startCol + 1).Value = _
            "'" & CStr(relWs.Cells(i, 4).Value)
    
        ws.Cells(r, startCol + 2).Value = _
            relWs.Cells(i, 2).Value
    
        ws.Cells(r, startCol + 3).Value = _
            relWs.Cells(i, 5).Value
    
        r = r + 1
    
    Next i

    ws.Range( _
        ws.Cells(startRow + 2, startCol + 2), _
        ws.Cells(r - 1, startCol + 2) _
    ).numberFormat = "dd-mmm-yy"

    ws.Range( _
        ws.Cells(startRow + 2, startCol + 3), _
        ws.Cells(r - 1, startCol + 3) _
    ).numberFormat = "$#,##0.00"
    
    If r > startRow + 2 Then
    
        With ws.Range( _
            ws.Cells(r - 1, startCol), _
            ws.Cells(r - 1, sectionLastCol))
    
            With .Borders(xlEdgeBottom)
    
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = vbBlack
    
            End With
    
        End With
    
    End If
    
    ws.Range( _
        ws.Cells(startRow + 2, startCol + 2), _
        ws.Cells(r - 1, startCol + 2) _
    ).HorizontalAlignment = xlLeft
    
    ws.Range( _
        ws.Cells(startRow + 2, startCol + 3), _
        ws.Cells(r - 1, startCol + 3) _
    ).HorizontalAlignment = xlLeft

    
    
End Sub

Private Sub WriteTopUnmatchedTransfers( _
    ByVal hostWb As Workbook, _
    ByVal ws As Worksheet, _
    ByVal startRow As Long, _
    ByVal startCol As Long)

    Dim txnWs As Worksheet
    Dim statWs As Worksheet

    Dim lastRow As Long
    Dim r As Long
    Dim txnRow As Long

    Dim TopRows(1 To 3) As Long
    Dim TopValues(1 To 3) As Double
    
    Dim statusValue As String

    Dim amount As Double
    Dim rowNum As Long
    
    Dim i As Long
    Dim j As Long
    
    Const DESCRIPTION_COLUMN_SPAN As Long = 7
    
    Dim descriptionFirstCol As Long
    Dim descriptionLastCol As Long
    Dim sectionLastCol As Long

    Set txnWs = hostWb.Worksheets("Transfer_Transactions")
    Set statWs = hostWb.Worksheets("Transfer_Transaction_Status")
    
    descriptionFirstCol = startCol + 3

    descriptionLastCol = _
        descriptionFirstCol + _
        DESCRIPTION_COLUMN_SPAN - 1
    
    sectionLastCol = descriptionLastCol
   
    lastRow = statWs.Cells( _
        statWs.rows.Count, 1).End(xlUp).Row
        
       
    For r = 2 To lastRow
    
        statusValue = _
            Trim$(CStr(statWs.Cells(r, 6).Value2))
    
        Select Case statusValue
    
            Case STATUS_UNMATCHED, _
                 STATUS_AMBIGUOUS
    
                rowNum = _
                    CLng(statWs.Cells(r, 2).Value2)
    
                txnRow = _
                    FindTransactionRow( _
                        hostWb, _
                        rowNum)
    
                If txnRow = 0 Then GoTo NextR
    
                amount = _
                    Abs(CDbl( _
                        txnWs.Cells(txnRow, 5).Value2))
    
                For i = 1 To 3
    
                    If amount > TopValues(i) Then
    
                        For j = 3 To i + 1 Step -1
    
                            TopValues(j) = _
                                TopValues(j - 1)
    
                            TopRows(j) = _
                                TopRows(j - 1)
    
                        Next j
    
                        TopValues(i) = amount
                        TopRows(i) = txnRow
    
                        Exit For
    
                    End If
    
                Next i
    
        End Select
    
NextR:
    Next r
    
    ws.Cells(startRow, startCol).Value = _
        "TOP TRANSFERS REQUIRING REVIEW"

    ws.Cells(startRow, startCol).Font.Bold = True
    
        If TopRows(1) = 0 Then

            ws.Cells(startRow + 1, startCol).Value = _
                "No transfers requiring review detected."
                
            With ws.Cells(startRow + 1, startCol)
                .Value = "No transfers requiring review detected."
                .Interior.Color = RGB(0, 176, 80)
                .Font.Bold = True
                .Font.Color = RGB(255, 255, 255)
            End With
                
                Exit Sub
                
            End If

    ws.Cells(startRow + 1, startCol).Resize(1, 4).Value = _
        Array("Account", _
              "Date", _
              "Amount", _
              "Description")

        With ws.Range( _
            ws.Cells(startRow + 1, startCol), _
            ws.Cells(startRow + 1, sectionLastCol))
        
            .Font.Bold = True
        
            With .Borders(xlEdgeBottom)
        
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = RGB(180, 180, 180)
        
            End With
        
        End With
    
        r = startRow + 2

    For i = 1 To 3

    
        If TopRows(i) > 0 Then

            txnRow = TopRows(i)

            ws.Cells(r, startCol).Value = _
                "'" & CStr(txnWs.Cells(txnRow, 3).Value)

            ws.Cells(r, startCol + 1).Value = _
                txnWs.Cells(txnRow, 4).Value

            ws.Cells(r, startCol + 2).Value = _
                txnWs.Cells(txnRow, 5).Value

        With ws.Range( _
            ws.Cells(r, descriptionFirstCol), _
            ws.Cells(r, descriptionLastCol))
        
            .Merge
        
            .Value = _
                CStr(txnWs.Cells(txnRow, 7).Value)
        
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .WrapText = False
        
        End With


            r = r + 1

        End If

    Next i
    
        ws.Range( _
        ws.Cells(startRow + 2, startCol + 1), _
        ws.Cells(r - 1, startCol + 1) _
    ).numberFormat = "dd-mmm-yy"
    
        ws.Range( _
        ws.Cells(startRow + 2, startCol + 1), _
        ws.Cells(r - 1, startCol + 1) _
    ).HorizontalAlignment = xlLeft


    ws.Range( _
        ws.Cells(startRow + 2, startCol + 2), _
        ws.Cells(r - 1, startCol + 2) _
    ).numberFormat = "$#,##0.00;($#,##0.00)"
    
    With ws.Range( _
        ws.Cells(startRow, startCol), _
        ws.Cells(startRow, sectionLastCol))
    
        .Merge
    
        .Value = _
            "TOP TRANSFERS REQUIRING REVIEW"
    
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    
        With .Borders(xlEdgeBottom)
    
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = vbBlack
    
        End With
    
    End With

    If r > startRow + 2 Then
    
        With ws.Range( _
            ws.Cells(r - 1, startCol), _
            ws.Cells(r - 1, sectionLastCol))
    
            With .Borders(xlEdgeBottom)
    
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = vbBlack
    
            End With
    
        End With
    
    End If
    
End Sub
    
    

Private Sub WriteTopAmbiguousActivity( _
    ByVal ws As Worksheet, _
    ByVal startRow As Long, _
    ByVal startCol As Long)

    Dim largestSizeID As String
    Dim largestValueID As String
    Dim unresolvedCount As Long

    Dim sectionLastCol As Long

    sectionLastCol = startCol + 2

    'Remove merged ranges left by a prior run.
    With ws.Range( _
        ws.Cells(startRow, startCol), _
        ws.Cells(startRow + 4, sectionLastCol))

        .UnMerge

    End With

    unresolvedCount = _
        GetUnresolvedGroupCount()


    With ws.Range( _
        ws.Cells(startRow, startCol), _
        ws.Cells(startRow, sectionLastCol))

        .Merge

        .Value = _
            "TOP UNRESOLVED ACTIVITY"

        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter

        With .Borders(xlEdgeBottom)

            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = vbBlack

        End With

    End With


    If unresolvedCount = 0 Then

        With ws.Range( _
            ws.Cells(startRow + 1, startCol), _
            ws.Cells(startRow + 1, sectionLastCol))

            .Merge

            .Value = _
                "No unresolved investigative groups detected."

            .Font.Italic = True
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter

            With .Borders(xlEdgeBottom)

                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = vbBlack

            End With

        End With

        Exit Sub

    End If

    largestSizeID = _
        GetLargestInvestigationGroupSizeID()

    largestValueID = _
        GetLargestInvestigationGroupExposureID()


    ws.Cells(startRow + 1, startCol).Value = _
        "Metric"

    ws.Cells(startRow + 1, startCol + 1).Value = _
        "Value"

    ws.Cells(startRow + 1, startCol + 2).Value = _
        "Group"

    With ws.Range( _
        ws.Cells(startRow + 1, startCol), _
        ws.Cells(startRow + 1, sectionLastCol))

        .Font.Bold = True

        With .Borders(xlEdgeBottom)

            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(180, 180, 180)

        End With

    End With


    ws.Cells(startRow + 2, startCol).Value = _
        "Most Involved Accounts"

    ws.Cells(startRow + 2, startCol + 1).Value = _
        GetLargestInvestigationAccounts()

    ws.Cells(startRow + 2, startCol + 2).Value = _
        GetLargestInvestigationAccountsID()

    ws.Cells(startRow + 3, startCol).Value = _
        "Most Transactions"

    ws.Cells(startRow + 3, startCol + 1).Value = _
        GetInvestigationGroupSize(largestSizeID)

    ws.Cells(startRow + 3, startCol + 2).Value = _
        largestSizeID

    ws.Cells(startRow + 4, startCol).Value = _
        "Largest Exposure"

    ws.Cells(startRow + 4, startCol + 1).Value = _
        GetInvestigationGroupExposure(largestValueID)

    ws.Cells(startRow + 4, startCol + 2).Value = _
        largestValueID

    ws.Cells( _
        startRow + 4, _
        startCol + 1).numberFormat = _
        "$#,##0.00;($#,##0.00)"

    ws.Range( _
        ws.Cells(startRow + 2, startCol + 1), _
        ws.Cells(startRow + 4, startCol + 1) _
        ).HorizontalAlignment = xlLeft



    With ws.Range( _
        ws.Cells(startRow + 4, startCol), _
        ws.Cells(startRow + 4, sectionLastCol))

        With .Borders(xlEdgeBottom)

            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = vbBlack

        End With

    End With

End Sub

Private Sub WriteTopResolvedActivity( _
    ByVal ws As Worksheet, _
    ByVal startRow As Long, _
    ByVal startCol As Long)

    Dim clusterCount As Long
    Dim largestSizeID As String
    Dim largestValueID As String

    Dim sectionLastCol As Long

    sectionLastCol = startCol + 2

    'Remove merged ranges left by a prior run.
    With ws.Range( _
        ws.Cells(startRow, startCol), _
        ws.Cells(startRow + 4, sectionLastCol))

        .UnMerge

    End With

    clusterCount = _
        GetAmbiguousClusterCount()


    With ws.Range( _
        ws.Cells(startRow, startCol), _
        ws.Cells(startRow, sectionLastCol))

        .Merge

        .Value = _
            "TOP RESOLVED ACTIVITY"

        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter

        With .Borders(xlEdgeBottom)

            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = vbBlack

        End With

    End With


    If clusterCount = 0 Then

        With ws.Range( _
            ws.Cells(startRow + 1, startCol), _
            ws.Cells(startRow + 1, sectionLastCol))

            .Merge

            .Value = _
                "No resolved investigative groups detected."

            .Font.Italic = True
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter

            With .Borders(xlEdgeBottom)

                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = vbBlack

            End With

        End With

        Exit Sub

    End If

    largestSizeID = _
        GetLargestResolvedGroupSizeID()

    largestValueID = _
        GetLargestResolvedGroupExposureID()


    ws.Cells(startRow + 1, startCol).Value = _
        "Metric"

    ws.Cells(startRow + 1, startCol + 1).Value = _
        "Value"

    ws.Cells(startRow + 1, startCol + 2).Value = _
        "Group"

    With ws.Range( _
        ws.Cells(startRow + 1, startCol), _
        ws.Cells(startRow + 1, sectionLastCol))

        .Font.Bold = True

        With .Borders(xlEdgeBottom)

            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(180, 180, 180)

        End With

    End With



    ws.Cells(startRow + 2, startCol).Value = _
        "Most Involved Accounts"

    ws.Cells(startRow + 2, startCol + 1).Value = _
        GetLargestResolvedAccounts()

    ws.Cells(startRow + 2, startCol + 2).Value = _
        GetLargestResolvedAccountsID()

    ws.Cells(startRow + 3, startCol).Value = _
        "Most Transactions"

    ws.Cells(startRow + 3, startCol + 1).Value = _
        GetResolvedGroupSize(largestSizeID)

    ws.Cells(startRow + 3, startCol + 2).Value = _
        largestSizeID

    ws.Cells(startRow + 4, startCol).Value = _
        "Largest Exposure"

    ws.Cells(startRow + 4, startCol + 1).Value = _
        GetResolvedGroupExposure(largestValueID)

    ws.Cells(startRow + 4, startCol + 2).Value = _
        largestValueID

    ws.Cells( _
        startRow + 4, _
        startCol + 1).numberFormat = _
        "$#,##0.00;($#,##0.00)"

    ws.Range( _
        ws.Cells(startRow + 2, startCol + 1), _
        ws.Cells(startRow + 4, startCol + 1) _
        ).HorizontalAlignment = xlLeft


    With ws.Range( _
        ws.Cells(startRow + 4, startCol), _
        ws.Cells(startRow + 4, sectionLastCol))

        With .Borders(xlEdgeBottom)

            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = vbBlack

        End With

    End With

End Sub

Private Function GetMatchedCount() As Long

    With hostWb.Worksheets("Matched_Transfers")

        GetMatchedCount = _
            .Cells(.rows.Count, 1).End(xlUp).Row - 1

    End With

End Function

Private Function GetMethodCount( _
    methodName As String) As Long
    
    If MethodCounts.Exists(methodName) Then
    
    GetMethodCount = _
        MethodCounts(methodName)
    
    Else
    
    GetMethodCount = 0
    
    End If

End Function

Private Function GetAmbiguousMethodCount( _
    methodName As String) As Long

    If AmbiguousMethodCounts.Exists(methodName) Then

    GetAmbiguousMethodCount = _
        AmbiguousMethodCounts(methodName)

    Else

    GetAmbiguousMethodCount = 0

    End If
    
End Function

Private Function GetAmbiguousCount() As Long

    With hostWb.Worksheets("Transfer_Ambiguity_Members")

        GetAmbiguousCount = _
            .Cells(.rows.Count, 1).End(xlUp).Row - 1

    End With

End Function


Private Function GetUnmatchedCount() As Long

    With hostWb.Worksheets("Unmatched_Transfers")

        GetUnmatchedCount = _
            .Cells(.rows.Count, 1).End(xlUp).Row - 1

    End With

End Function

Private Function GetStatusCount( _
    statusValue As String) As Long

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = hostWb.Worksheets("Transfer_Transaction_Status")

    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow

        If ws.Cells(r, 6).Value = statusValue Then

            GetStatusCount = _
                GetStatusCount + 1

        End If

    Next r

End Function
Private Function GetTransfersRequiringReviewCount() As Long


    GetTransfersRequiringReviewCount = _
        GetStatusCount(STATUS_UNMATCHED) + _
        GetStatusCount(STATUS_AMBIGUOUS)

End Function


Private Sub BuildPrimaryHubAnalysis( _
    ByVal hostWb As Workbook)

    Dim ws As Worksheet
    Dim wsRel As Worksheet

    Dim r As Long
    Dim startRow As Long

    Dim acctStats As Object
    Dim relStats As Object

    Dim hubAcct As String
    Dim hubScore As Double

    Set ws = _
        hostWb.Worksheets("Primary_Hub_Analysis")

    Set wsRel = _
        hostWb.Worksheets("Transfer_Relationships")

    ws.Cells.Clear

    startRow = 1
    r = startRow

    Set acctStats = _
        CreateObject("Scripting.Dictionary")

    Set relStats = _
        CreateObject("Scripting.Dictionary")

    CalculateHubStatistics _
        wsRel, _
        acctStats, _
        relStats

    
    If acctStats.Count = 0 Then
        ws.Cells(startRow, 1).Value = _
                "PRIMARY HUB ANALYSIS"
                
        ws.Cells(startRow, 1).Font.Bold = True
        ws.Cells(startRow + 2, 1).Value = _
                "No matched transfer relationships found."
        ws.Cells(startRow + 2, 1).Font.Italic = True
        
        Exit Sub
        
    End If

    hubAcct = DeterminePrimaryHub(acctStats, hubScore)

    WritePrimaryHubSummary _
        ws, _
        startRow, _
        acctStats, _
        hubAcct, _
        hubScore
        
    WriteTopConnectedAccounts _
        ws, _
        startRow, _
        relStats, _
        hubAcct
        
    WriteVelocityMetrics _
        ws, _
        startRow, _
        acctStats, _
        relStats, _
        hubAcct
        
    WriteFlowCharacteristics _
        ws, _
        startRow, _
        acctStats, _
        hubAcct
        
    WriteTimelineAnalysis _
        ws, _
        startRow, _
        acctStats, _
        hubAcct


ActiveWindow.DisplayGridlines = True
        

End Sub


Private Sub CalculateHubStatistics( _
    ByVal wsRel As Worksheet, _
    ByRef acctStats As Object, _
    ByRef relStats As Object)

    Dim lastRow As Long
    Dim r As Long

    Dim acct As Object

    Dim txDate As Date
    Dim debitAcct As String
    Dim creditAcct As String
    Dim amt As Double
    Dim relKey As String

    lastRow = wsRel.Cells(wsRel.rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow

        If Len(wsRel.Cells(r, 3).Value) = 0 Then GoTo ContinueLoop
        If Len(wsRel.Cells(r, 4).Value) = 0 Then GoTo ContinueLoop

        txDate = wsRel.Cells(r, 2).Value
        debitAcct = Trim(CStr(wsRel.Cells(r, 3).Value))
        creditAcct = Trim(CStr(wsRel.Cells(r, 4).Value))
        amt = Abs(CDbl(wsRel.Cells(r, 5).Value))

        EnsureAccount acctStats, debitAcct, txDate
        EnsureAccount acctStats, creditAcct, txDate

        Set acct = acctStats(debitAcct)
        acct("TransferCount") = acct("TransferCount") + 1
        acct("TransferOut") = acct("TransferOut") + 1
        acct("VolumeOut") = acct("VolumeOut") + amt
        acct("Destinations")(creditAcct) = 1
        acct("Counterparties")(creditAcct) = 1
        AddDailyVolume acct, txDate, amt
        UpdateDates acct, txDate

        Set acct = acctStats(creditAcct)
        acct("TransferCount") = acct("TransferCount") + 1
        acct("TransferIn") = acct("TransferIn") + 1
        acct("VolumeIn") = acct("VolumeIn") + amt
        acct("Sources")(debitAcct) = 1
        acct("Counterparties")(debitAcct) = 1
        AddDailyVolume acct, txDate, amt
        UpdateDates acct, txDate

        relKey = debitAcct & "|" & creditAcct

        If Not relStats.Exists(relKey) Then
            relStats.Add relKey, Array(0&, 0#)
        End If

        Dim arr
        arr = relStats(relKey)
        arr(0) = arr(0) + 1
        arr(1) = arr(1) + amt
        relStats(relKey) = arr

ContinueLoop:
    Next r

End Sub

Private Sub EnsureAccount(ByRef acctStats As Object, _
                          ByVal acctNum As String, _
                          ByVal txDate As Date)

    Dim d As Object

    If acctStats.Exists(acctNum) Then Exit Sub

    Set d = CreateObject("Scripting.Dictionary")

    d.Add "TransferCount", 0&
    d.Add "TransferIn", 0&
    d.Add "TransferOut", 0&
    d.Add "VolumeIn", 0#
    d.Add "VolumeOut", 0#
    d.Add "FirstDate", txDate
    d.Add "LastDate", txDate

    Set d("Sources") = CreateObject("Scripting.Dictionary")
    Set d("Destinations") = CreateObject("Scripting.Dictionary")
    Set d("Counterparties") = CreateObject("Scripting.Dictionary")
    Set d("DailyVolume") = CreateObject("Scripting.Dictionary")

    acctStats.Add acctNum, d

End Sub


Private Sub AddDailyVolume(ByVal acct As Object, _
                           ByVal txDate As Date, _
                           ByVal amt As Double)

    Dim k As String

    k = Format$(txDate, "yyyymmdd")

    If Not acct("DailyVolume").Exists(k) Then
        acct("DailyVolume").Add k, 0#
    End If

    acct("DailyVolume")(k) = acct("DailyVolume")(k) + amt

End Sub

Private Sub UpdateDates(ByVal acct As Object, _
                        ByVal txDate As Date)

    If txDate < acct("FirstDate") Then acct("FirstDate") = txDate
    If txDate > acct("LastDate") Then acct("LastDate") = txDate

End Sub

Private Function DeterminePrimaryHub( _
    ByVal acctStats As Object, _
    ByRef hubScore As Double) As String

    Dim acct As Variant

    Dim maxVolume As Double
    Dim maxCentrality As Double
    Dim maxVelocity As Double
    Dim maxCounterparty As Double

    Dim grossVol As Double
    Dim centrality As Double
    Dim velocity As Double
    Dim counterparties As Double
    Dim activeDays As Double

    Dim score As Double
    Dim bestScore As Double

    For Each acct In acctStats.Keys

        grossVol = acctStats(acct)("VolumeIn") + acctStats(acct)("VolumeOut")
        centrality = acctStats(acct)("Sources").Count + acctStats(acct)("Destinations").Count
        counterparties = acctStats(acct)("Counterparties").Count

        activeDays = acctStats(acct)("LastDate") - acctStats(acct)("FirstDate") + 1
        If activeDays < 1 Then activeDays = 1

        velocity = grossVol / activeDays

        If grossVol > maxVolume Then maxVolume = grossVol
        If centrality > maxCentrality Then maxCentrality = centrality
        If velocity > maxVelocity Then maxVelocity = velocity
        If counterparties > maxCounterparty Then maxCounterparty = counterparties

    Next acct

    For Each acct In acctStats.Keys

        grossVol = acctStats(acct)("VolumeIn") + acctStats(acct)("VolumeOut")
        centrality = acctStats(acct)("Sources").Count + acctStats(acct)("Destinations").Count
        counterparties = acctStats(acct)("Counterparties").Count

        activeDays = acctStats(acct)("LastDate") - acctStats(acct)("FirstDate") + 1
        If activeDays < 1 Then activeDays = 1

        velocity = grossVol / activeDays
        
        If maxVolume = 0 Then
            maxVolume = 1
        End If
        
        If maxCentrality = 0 Then
            maxCentrality = 1
        End If
        
        If maxVelocity = 0 Then
            maxVelocity = 1
        End If
        
        If maxCounterparty = 0 Then
            maxCounterparty = 1
        End If


        score = _
            ((grossVol / maxVolume) * 40#) + _
            ((centrality / maxCentrality) * 20#) + _
            ((velocity / maxVelocity) * 20#) + _
            ((counterparties / maxCounterparty) * 20#)

        If score > bestScore Then
            bestScore = score
            DeterminePrimaryHub = acct
        End If

    Next acct

    hubScore = bestScore

End Function

Private Sub WritePrimaryHubSummary( _
    ByVal ws As Worksheet, _
    ByVal startRow As Long, _
    ByVal acctStats As Object, _
    ByVal hubAcct As String, _
    ByVal hubScore As Double)

    
    Dim summaryRow As Long
    Dim velocityRow As Long
    Dim flowRow As Long
    Dim timelineRow As Long
    
    Dim s As Object
    Dim grossVol As Double
    
    Dim networkVol As Double
    Dim dominance As Double
    
    summaryRow = startRow + 2
    velocityRow = startRow + 13
    flowRow = startRow + 13
    timelineRow = startRow + 26
    
        
    
    Set s = acctStats(hubAcct)
    
    grossVol = s("VolumeIn") + s("VolumeOut")
    
    networkVol = GetNetworkVolume(acctStats)
    
    If networkVol > 0 Then
    
        dominance = grossVol / networkVol
    
    Else
    
        dominance = 0
    
    End If
    

    summaryRow = startRow + 2
    
    With ws.Range( _
        ws.Cells(startRow, 1), _
        ws.Cells(startRow, 8))
    
        .Merge
        .Value = "PRIMARY HUB ANALYSIS"
    
        .Font.Bold = True
        .Font.size = 12
    
        .Interior.Color = RGB(221, 235, 247)
    
    End With
    
    ws.Cells(summaryRow, 1).Value = _
        "PRIMARY HUB ACCOUNT"
    ws.Cells(summaryRow, 2).Value = _
        "'" & hubAcct
    ws.Cells(summaryRow + 1, 1).Value = _
        "Hub Dominance"
    ws.Cells(summaryRow + 1, 2).Value = _
        dominance
    ws.Cells(summaryRow + 2, 1).Value = _
        "Hub Score"
    ws.Cells(summaryRow + 2, 2).Value = _
        Round(hubScore, 1) & " / 100"
    ws.Cells(summaryRow + 3, 1).Value = _
        "Total Transfers"
    ws.Cells(summaryRow + 3, 2).Value = _
        s("TransferCount")
    ws.Cells(summaryRow + 4, 1).Value = _
        "Total Volume"
    ws.Cells(summaryRow + 4, 2).Value = _
        grossVol
    ws.Cells(summaryRow + 5, 1).Value = _
        "Transfers In"
    ws.Cells(summaryRow + 5, 2).Value = _
        s("TransferIn")
    ws.Cells(summaryRow + 6, 1).Value = _
        "Volume In"
    ws.Cells(summaryRow + 6, 2).Value = _
        s("VolumeIn")
    ws.Cells(summaryRow + 7, 1).Value = _
        "Transfers Out"
    ws.Cells(summaryRow + 7, 2).Value = _
        s("TransferOut")
    ws.Cells(summaryRow + 8, 1).Value = _
        "Volume Out"
    ws.Cells(summaryRow + 8, 2).Value = _
        s("VolumeOut")
    ws.Cells(summaryRow + 9, 1).Value = _
        "Net Flow"
    ws.Cells(summaryRow + 9, 2).Value = _
        s("VolumeIn") - s("VolumeOut")
        
    ws.Cells(summaryRow + 1, 2).numberFormat = "0.00%"
    
    ws.Cells(summaryRow + 4, 2).numberFormat = "$#,##0.00;($#,##0.00)" 'Total Volume
    ws.Cells(summaryRow + 6, 2).numberFormat = "$#,##0.00;($#,##0.00)" 'Volume In
    ws.Cells(summaryRow + 8, 2).numberFormat = "$#,##0.00;($#,##0.00)" 'Volume Out
    ws.Cells(summaryRow + 9, 2).numberFormat = "$#,##0.00;($#,##0.00)" 'Net Flow
    
        
  
    ws.Range( _
        ws.Cells(summaryRow + 1, 2), _
        ws.Cells(summaryRow + 9, 2) _
    ).HorizontalAlignment = xlRight
    
    With ws.Range( _
    ws.Cells(summaryRow, 1), _
    ws.Cells(summaryRow + 9, 2))

    .Borders(xlEdgeLeft).LineStyle = xlContinuous
    .Borders(xlEdgeRight).LineStyle = xlContinuous
    .Borders(xlEdgeTop).LineStyle = xlContinuous
    .Borders(xlEdgeBottom).LineStyle = xlContinuous

    End With
    
    ws.Cells(summaryRow, 2).Value = _
        "'" & hubAcct

    ws.Cells(summaryRow, 2).Font.Bold = True
    
    ws.Cells(summaryRow, 1).Font.Bold = True
    ws.Cells(summaryRow, 2).Font.Bold = True
    ws.Cells(summaryRow + 1, 2).Font.Bold = True   'Hub Dominance
    ws.Cells(summaryRow + 2, 2).Font.Bold = True   'Hub Score



End Sub

Private Function GetPrimaryHubAccount( _
    sourceWs As Worksheet) As String

    Dim hostWb As Workbook
    Dim hubWs As Worksheet
    Dim labelCell As Range

    Set hostWb = _
        sourceWs.Parent


    On Error Resume Next

    Set hubWs = _
        hostWb.Worksheets( _
            "Primary_Hub_Analysis")

    On Error GoTo 0

    If hubWs Is Nothing Then

        GetPrimaryHubAccount = _
            "Not Available"

        Exit Function

    End If


    Set labelCell = _
        hubWs.Columns(1).Find( _
            What:="PRIMARY HUB ACCOUNT", _
            After:=hubWs.Cells(1, 1), _
            LookIn:=xlValues, _
            LookAt:=xlWhole, _
            SearchOrder:=xlByRows, _
            SearchDirection:=xlNext, _
            MatchCase:=False)

    If labelCell Is Nothing Then

        GetPrimaryHubAccount = _
            "Not Available"

        Exit Function

    End If


    GetPrimaryHubAccount = _
        Trim$(CStr( _
            labelCell.Offset(0, 1).Value2))

    If GetPrimaryHubAccount = "" Then

        GetPrimaryHubAccount = _
            "Not Available"

    End If

End Function


Private Sub WriteTopConnectedAccounts( _
    ByVal ws As Worksheet, _
    ByVal startRow As Long, _
    ByVal relStats As Object, _
    ByVal hubAcct As String)

    Dim cpStats As Object
    Dim k As Variant
    Dim parts() As String

    Dim counterparty As String

    Dim relData As Variant

    Dim cpVolume As Double
    Dim cpCount As Long

    Dim rowOut As Long

    Dim TopAcct(1 To 3) As String
    Dim TopCount(1 To 3) As Long
    Dim TopVolume(1 To 3) As Double

    Dim i As Long
    Dim j As Long

    Set cpStats = CreateObject("Scripting.Dictionary")

    For Each k In relStats.Keys

        parts = Split(k, "|")

        If UBound(parts) < 1 Then GoTo NextRelationship

        If parts(0) = hubAcct Then

            counterparty = parts(1)

        ElseIf parts(1) = hubAcct Then

            counterparty = parts(0)

        Else

            GoTo NextRelationship

        End If

        relData = relStats(k)

        cpCount = relData(0)
        cpVolume = relData(1)

        If Not cpStats.Exists(counterparty) Then

            cpStats.Add counterparty, _
                Array(0&, 0#)

        End If

        Dim arr As Variant

        arr = cpStats(counterparty)

        arr(0) = arr(0) + cpCount
        arr(1) = arr(1) + cpVolume

        cpStats(counterparty) = arr

NextRelationship:
    Next k

    For Each k In cpStats.Keys

        arr = cpStats(k)

        cpCount = arr(0)
        cpVolume = arr(1)

        For i = 1 To 3

            If cpVolume > TopVolume(i) Then

                For j = 3 To i + 1 Step -1

                    TopVolume(j) = TopVolume(j - 1)
                    TopCount(j) = TopCount(j - 1)
                    TopAcct(j) = TopAcct(j - 1)

                Next j

                TopVolume(i) = cpVolume
                TopCount(i) = cpCount
                TopAcct(i) = k

                Exit For

            End If

        Next i

    Next k

    Dim sectionRow As Long

    sectionRow = startRow + 2
    
    If cpStats.Count = 0 Then

        ws.Cells(sectionRow + 1, 6).Value = _
            "No connected accounts found."
    
        Exit Sub
    
    End If

    ws.Cells(sectionRow, 5).Value = _
        "TOP CONNECTED ACCOUNTS"

    ws.Cells(sectionRow, 5).Font.Bold = True

    ws.Cells(sectionRow + 1, 5).Value = _
        "Account"

    ws.Cells(sectionRow + 1, 6).Value = _
        "Transfers"

    ws.Cells(sectionRow + 1, 7).Value = _
        "Volume"

    ws.Range( _
        ws.Cells(sectionRow + 1, 5), _
        ws.Cells(sectionRow + 1, 7)).Font.Bold = True

    rowOut = sectionRow + 2

        For i = 1 To 3

        If Len(TopAcct(i)) > 0 Then

            ws.Cells(rowOut, 5).Value = _
                "'" & TopAcct(i)

            ws.Cells(rowOut, 6).Value = _
                TopCount(i)

            ws.Cells(rowOut, 7).Value = _
                TopVolume(i)

            rowOut = rowOut + 1

        End If

    Next i

    If rowOut > sectionRow + 2 Then

        ws.Range( _
            ws.Cells(sectionRow + 2, 7), _
            ws.Cells(rowOut - 1, 7) _
        ).numberFormat = _
            "$#,##0.00;($#,##0.00)"

    End If
    
    Dim LargestCounterparty As String
    Dim LargestCounterpartyVolume As Double
    
    LargestCounterparty = TopAcct(1)
    LargestCounterpartyVolume = TopVolume(1)
    
    With ws.Range( _
        ws.Cells(sectionRow, 5), _
        ws.Cells(sectionRow + 4, 7))
    
        .BorderAround xlContinuous, xlThin

    End With

End Sub


Private Sub WriteVelocityMetrics( _
    ByVal ws As Worksheet, _
    ByVal startRow As Long, _
    ByVal acctStats As Object, _
    ByVal relStats As Object, _
    ByVal hubAcct As String)

    Dim d As Double
    Dim s As Object
    Dim grossVol As Double

    Dim velocityRow As Long

    Dim largestVolume As Double
    Dim largestPct As Double

    Set s = acctStats(hubAcct)

    grossVol = s("VolumeIn") + s("VolumeOut")
    
    largestVolume = _
    GetLargestCounterpartyVolume( _
        relStats, _
        hubAcct)

    d = s("LastDate") - s("FirstDate") + 1

    If d < 1 Then d = 1

    velocityRow = startRow + 13

    ws.Cells(velocityRow, 1).Value = _
        "ACTIVITY VELOCITY"

    ws.Cells(velocityRow, 1).Font.Bold = True

    ws.Cells(velocityRow + 1, 1).Value = _
        "Date Range"

    ws.Cells(velocityRow + 1, 2).Value = _
        Format$(s("FirstDate"), "dd-mmm-yy") & _
        " - " & _
        Format$(s("LastDate"), "dd-mmm-yy")

    ws.Cells(velocityRow + 2, 1).Value = _
        "Transfers / Day"

    ws.Cells(velocityRow + 2, 2).Value = _
        s("TransferCount") / d

    ws.Cells(velocityRow + 3, 1).Value = _
        "Transfers / Week"

    ws.Cells(velocityRow + 3, 2).Value = _
        (s("TransferCount") / d) * 7

    ws.Cells(velocityRow + 4, 1).Value = _
        "Transfers / Month"

    ws.Cells(velocityRow + 4, 2).Value = _
        (s("TransferCount") / d) * 30.44

    ws.Cells(velocityRow + 5, 1).Value = _
        "Volume / Day"

    ws.Cells(velocityRow + 5, 2).Value = _
        grossVol / d

    ws.Cells(velocityRow + 6, 1).Value = _
        "Volume / Week"

    ws.Cells(velocityRow + 6, 2).Value = _
        (grossVol / d) * 7

    ws.Cells(velocityRow + 7, 1).Value = _
        "Volume / Month"

    ws.Cells(velocityRow + 7, 2).Value = _
        (grossVol / d) * 30.44

    largestVolume = GetLargestCounterpartyVolume( _
                        relStats, _
                        hubAcct)

    ws.Cells(velocityRow + 8, 1).Value = _
        "Largest Counterparty Volume"


    ws.Cells(velocityRow + 8, 2).Value = _
        largestVolume
        
    If grossVol > 0 Then

        largestPct = largestVolume / grossVol

    End If

    ws.Cells(velocityRow + 9, 1).Value = _
        "Largest Counterparty %"
    
    ws.Cells(velocityRow + 9, 2).Value = _
        largestPct
    
    ' Transfer rates
    ws.Range( _
        ws.Cells(velocityRow + 2, 2), _
        ws.Cells(velocityRow + 4, 2) _
    ).numberFormat = "0.00"
    
    ' Monetary metrics
    ws.Range( _
        ws.Cells(velocityRow + 5, 2), _
        ws.Cells(velocityRow + 8, 2) _
    ).numberFormat = _
        "$#,##0.00;($#,##0.00)"
    
    ' Percentages
    ws.Cells(velocityRow + 9, 2).numberFormat = _
        "0.00%"
        
    With ws.Range( _
        ws.Cells(velocityRow, 1), _
        ws.Cells(velocityRow + 9, 2))
    
        .BorderAround xlContinuous, xlThin
    
    End With



End Sub

Private Sub WriteFlowCharacteristics( _
    ByVal ws As Worksheet, _
    ByVal startRow As Long, _
    ByVal acctStats As Object, _
    ByVal hubAcct As String)

    Dim s As Object
    Dim grossVol As Double

    Dim flowRow As Long

    Set s = acctStats(hubAcct)

    grossVol = s("VolumeIn") + s("VolumeOut")

    flowRow = startRow + 8

    ws.Cells(flowRow, 5).Value = _
        "FLOW CHARACTERISTICS"

    ws.Cells(flowRow, 5).Font.Bold = True

    If grossVol > 0 Then

        ws.Cells(flowRow + 1, 5).Value = _
            "Inbound Volume %"

        ws.Cells(flowRow + 1, 6).Value = _
            s("VolumeIn") / grossVol

        ws.Cells(flowRow + 2, 5).Value = _
            "Outbound Volume %"

        ws.Cells(flowRow + 2, 6).Value = _
            s("VolumeOut") / grossVol

        ws.Cells(flowRow + 1, 6).numberFormat = _
            "0.00%"

        ws.Cells(flowRow + 2, 6).numberFormat = _
            "0.00%"

    Else

        ws.Cells(flowRow + 1, 6).Value = _
            "Inbound Volume %"

        ws.Cells(flowRow + 1, 6).Value = 0

        ws.Cells(flowRow + 2, 5).Value = _
            "Outbound Volume %"

        ws.Cells(flowRow + 2, 6).Value = 0

    End If

    ws.Cells(flowRow + 4, 5).Value = _
        "Source Accounts"

    ws.Cells(flowRow + 4, 6).Value = _
        s("Sources").Count

    ws.Cells(flowRow + 5, 5).Value = _
        "Destination Accounts"

    ws.Cells(flowRow + 5, 6).Value = _
        s("Destinations").Count
        
    With ws.Range( _
        ws.Cells(flowRow, 5), _
        ws.Cells(flowRow + 5, 6))
    
        .BorderAround xlContinuous, xlThin
    
    End With

End Sub

Private Sub WriteTimelineAnalysis( _
    ByVal ws As Worksheet, _
    ByVal startRow As Long, _
    ByVal acctStats As Object, _
    ByVal hubAcct As String)

    Dim s As Object
    Dim k As Variant

    Dim peakDate As String
    Dim peakDt As Date

    Dim peakVol As Double
    Dim grossVol As Double

    Dim timelineRow As Long

    Set s = acctStats(hubAcct)

    grossVol = s("VolumeIn") + s("VolumeOut")

    For Each k In s("DailyVolume").Keys

        If s("DailyVolume")(k) > peakVol Then

            peakVol = s("DailyVolume")(k)
            peakDate = CStr(k)

        End If

    Next k

    If Len(peakDate) = 8 Then

        peakDt = DateSerial( _
            CLng(Left$(peakDate, 4)), _
            CLng(Mid$(peakDate, 5, 2)), _
            CLng(Right$(peakDate, 2)))

    End If

    timelineRow = startRow + 15

    ws.Cells(timelineRow, 5).Value = _
        "TIMELINE ANALYSIS"

    ws.Cells(timelineRow, 5).Font.Bold = True

    ws.Cells(timelineRow + 1, 5).Value = _
        "First Transfer Date"

    ws.Cells(timelineRow + 1, 6).Value = _
        s("FirstDate")

    ws.Cells(timelineRow + 2, 5).Value = _
        "Last Transfer Date"

    ws.Cells(timelineRow + 2, 6).Value = _
        s("LastDate")

    ws.Cells(timelineRow + 3, 5).Value = _
        "Active Days"

    ws.Cells(timelineRow + 3, 6).Value = _
        s("LastDate") - s("FirstDate") + 1

    ws.Cells(timelineRow + 4, 5).Value = _
        "Peak Transfer Day"

    ws.Cells(timelineRow + 4, 6).Value = _
        peakDt

    ws.Cells(timelineRow + 5, 5).Value = _
        "Peak Day Volume"

    ws.Cells(timelineRow + 5, 6).Value = _
        peakVol

    ws.Cells(timelineRow + 6, 5).Value = _
        "Peak Day % Volume"

    If grossVol > 0 Then

        ws.Cells(timelineRow + 6, 6).Value = _
            peakVol / grossVol

    End If

    ws.Cells(timelineRow + 1, 6).numberFormat = _
        "dd-mmm-yy"

    ws.Cells(timelineRow + 2, 6).numberFormat = _
        "dd-mmm-yy"

    ws.Cells(timelineRow + 4, 6).numberFormat = _
        "dd-mmm-yy"

    ws.Cells(timelineRow + 5, 6).numberFormat = _
        "$#,##0.00;($#,##0.00)"

    ws.Cells(timelineRow + 6, 6).numberFormat = _
        "0.00%"
        
    With ws.Range( _
        ws.Cells(timelineRow, 5), _
        ws.Cells(timelineRow + 6, 6))
    
        .BorderAround xlContinuous, xlThin
    
    End With

End Sub

Private Function GetNetworkVolume( _
    ByVal acctStats As Object) As Double

    Dim acct As Variant
    Dim totalVol As Double

    For Each acct In acctStats.Keys

        totalVol = totalVol + _
            acctStats(acct)("VolumeIn") + _
            acctStats(acct)("VolumeOut")

    Next acct

    GetNetworkVolume = totalVol / 2

End Function

Private Function GetLargestCounterpartyVolume( _
        ByVal relStats As Object, _
        ByVal hubAcct As String) As Double
        
    Dim cpStats As Object
    Dim k As Variant
    Dim parts() As String
    Dim counterparty As String
    Dim relData As Variant
    Dim arr As Variant
    Dim largestVolume As Double
    
    Set cpStats = CreateObject("Scripting.Dictionary")
    
    For Each k In relStats.Keys
        parts = Split(k, "|")
    
    If UBound(parts) < 1 Then GoTo NextRelationship
    
    If parts(0) = hubAcct Then
        
        counterparty = parts(1)
    
    ElseIf parts(1) = hubAcct Then
    
        counterparty = parts(0)
    Else
    
    GoTo NextRelationship
    
    End If
    
    relData = relStats(k)
    
    If Not cpStats.Exists(counterparty) Then
        
        cpStats.Add counterparty, 0#
    
    End If
    
    cpStats(counterparty) = _
        cpStats(counterparty) + relData(1)
NextRelationship:
    Next k
    
    For Each k In cpStats.Keys
    
    If cpStats(k) > largestVolume Then
    
        largestVolume = cpStats(k)
    
    End If
    
Next k
    
    GetLargestCounterpartyVolume = largestVolume

End Function


'---This entire section will likely be removed---

Private Sub BuildITMGuide()

    Dim ws As Worksheet

    Set ws = hostWb.Worksheets("ITM_Guide")

    ws.Cells.Clear

    ws.Columns("A").ColumnWidth = 30
    ws.Columns("B").ColumnWidth = 120

    ws.Cells.WrapText = True

    '=================================
    ' Title
    '=================================

    With ws.Range("A1:B3")

        .Merge
        .Value = _
            "Internal Transfer Matching (ITM) v8" & vbCrLf & _
            "Internal Transfer Analysis and Investigation Guide"

        .Font.Bold = True
        .Font.size = 18

        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = vbWhite

    End With

    ws.Cells(5, 1).Value = "Purpose"

    ws.Cells(5, 1).Font.Bold = True
    ws.Cells(5, 1).Interior.Color = RGB(220, 230, 241)

    ws.Cells(5, 2).Value = _
        "The Internal Transfer Matching (ITM) Module identifies likely internal transfers and classifies activity as Matched, Ambiguous, or Unmatched. The module is intended to reduce manual review effort while maintaining investigator visibility into unresolved activity."

    '=================================
    ' Matching Process
    '=================================

    ws.Cells(8, 1).Value = "Matching Process"

    ws.Cells(8, 1).Font.Bold = True
    ws.Cells(8, 1).Interior.Color = RGB(220, 230, 241)

    ws.Cells(8, 2).Value = _
        "Transactions are processed using the following match priority:" & vbCrLf & vbCrLf & _
        "1. Confirmation Number" & vbCrLf & _
        "2. Narrative Pair" & vbCrLf & _
        "3. Referenced Account" & vbCrLf & _
        "4. Embedded Account" & vbCrLf & _
        "5. Transfer Suffix" & vbCrLf & _
        "6. Branch Transfer" & vbCrLf & _
        "7. Amount + Date"

    '=================================
    ' Reports
    '=================================

    ws.Cells(18, 1).Value = "Reports"

    ws.Cells(18, 1).Font.Bold = True
    ws.Cells(18, 1).Interior.Color = RGB(220, 230, 241)

    ws.Cells(18, 2).Value = _
        "Transfer_Summary - Executive summary of matching results." & vbCrLf & vbCrLf & _
        "Transfer_Network_Analysis - Account and relationship analysis." & vbCrLf & vbCrLf & _
        "Matched_Transfers - Confirmed transfer matches." & vbCrLf & vbCrLf & _
        "Investigation_Groups - Potential transfer relationships requiring analyst review." & vbCrLf & vbCrLf & _
        "Unmatched_Transfers - Transfer activity that could not be confidently matched."

    '=================================
    ' Calculation Reference
    '=================================

    ws.Cells(26, 1).Value = "Calculation Reference"

    ws.Cells(26, 1).Font.Bold = True
    ws.Cells(26, 1).Interior.Color = RGB(220, 230, 241)

    ws.Cells(26, 2).Value = _
        "Primary Hub" & vbCrLf & _
        "Definition: Account with the highest Gross Volume." & vbCrLf & vbCrLf & _
        "Gross Volume = Incoming Volume + Outgoing Volume" & vbCrLf & vbCrLf & _
        "Primary Hub = Account with Maximum Gross Volume"

    ws.Cells(34, 2).Value = _
        "Cluster Exposure" & vbCrLf & _
        "Definition: Total absolute dollar value contained within an ambiguity cluster." & vbCrLf & vbCrLf & _
        "Formula:" & vbCrLf & _
        "Cluster Exposure = Sum(ABS(Transaction Amount))"

    ws.Cells(42, 2).Value = _
        "Velocity" & vbCrLf & _
        "Definition: Transaction activity occurring during a specific time period." & vbCrLf & vbCrLf & _
        "Examples:" & vbCrLf & _
        "Transfers Per Day" & vbCrLf & _
        "Transfer Volume Per Day"

    ws.Cells(49, 2).Value = _
        "Relationship Count" & vbCrLf & _
        "Definition: Number of unique account relationships associated with an account."

    '=================================
    ' Interpretation Guide
    '=================================

    ws.Cells(55, 1).Value = "Understanding Results"

    ws.Cells(55, 1).Font.Bold = True
    ws.Cells(55, 1).Interior.Color = RGB(220, 230, 241)

    ws.Cells(55, 2).Value = _
        "Matched - Sufficient evidence exists to confidently pair transactions." & vbCrLf & vbCrLf & _
        "Ambiguous - Multiple potential relationships exist and analyst review is required." & vbCrLf & vbCrLf & _
        "Unmatched - No acceptable counterpart was identified." & vbCrLf & vbCrLf & _
        "Contradicted - Evidence exists indicating a proposed match is unlikely."

    '=================================
    ' Detected Activity
    '=================================

    ws.Cells(63, 1).Value = "Detected Activity"

    ws.Cells(63, 1).Font.Bold = True
    ws.Cells(63, 1).Interior.Color = RGB(220, 230, 241)

    ws.Cells(63, 2).Value = _
        "Internal account transfers" & vbCrLf & _
        "Member-to-member transfers" & vbCrLf & _
        "Transfers identified through confirmation numbers" & vbCrLf & _
        "Transfers identified through account references" & vbCrLf & _
        "Transfers identified through transaction narratives" & vbCrLf & _
        "Transfers identified through amount/date relationships"

    '=================================
    ' Limitations
    '=================================

    ws.Cells(71, 1).Value = "Known Limitations"

    ws.Cells(71, 1).Font.Bold = True
    ws.Cells(71, 1).Interior.Color = RGB(220, 230, 241)

    ws.Cells(71, 2).Value = _
        "Activity occurring outside the available dataset." & vbCrLf & _
        "Transfers lacking identifiable linking evidence." & vbCrLf & _
        "Transactions with incomplete source data." & vbCrLf & _
        "Activity intentionally structured to conceal transfer relationships." & vbCrLf & _
        "External transfers without sufficient transfer indicators."

    '=================================
    ' Architecture
    '=================================

    ws.Cells(79, 1).Value = "Warehouse Architecture"

    ws.Cells(79, 1).Font.Bold = True
    ws.Cells(79, 1).Interior.Color = RGB(220, 230, 241)

    ws.Cells(79, 2).Value = _
        "Warehouse Tables" & vbCrLf & vbCrLf & _
        "Transfer_Transactions" & vbCrLf & _
        "Transfer_Transaction_Status" & vbCrLf & _
        "Transfer_Ambiguities" & vbCrLf & _
        "Transfer_Ambiguity_Members" & vbCrLf & _
        "Transfer_Relationships" & vbCrLf & vbCrLf & _
        "Warehouse tables are considered the system source of truth."

    '=================================
    ' Metadata
    '=================================

    ws.Cells(91, 1).Value = "Metadata Definitions"

    ws.Cells(91, 1).Font.Bold = True
    ws.Cells(91, 1).Interior.Color = RGB(220, 230, 241)

    ws.Cells(91, 2).Value = _
        "Match Rate = Matched Transactions / Total Transfer Transactions" & vbCrLf & vbCrLf & _
        "Ambiguity Rate = Ambiguous Transactions / Total Transfer Transactions" & vbCrLf & vbCrLf & _
        "Largest Cluster Exposure = Largest dollar exposure associated with a single ambiguity cluster" & vbCrLf & vbCrLf & _
        "Candidates Generated = Total candidate transfer relationships evaluated during processing"

    '=================================
    ' Support
    '=================================

    ws.Cells(100, 1).Value = "Support"

    ws.Cells(100, 1).Font.Bold = True
    ws.Cells(100, 1).Interior.Color = RGB(220, 230, 241)

    ws.Cells(100, 2).Value = _
        "Questions, enhancement requests, or issues:" & vbCrLf & vbCrLf & _
        "Michael Zipprich" & vbCrLf & _
        "Financial Crimes Analyst I" & vbCrLf & vbCrLf & _
        "Email: michael.zipprich@busey.com"

    ws.rows.AutoFit

    ActiveWindow.DisplayGridlines = False

End Sub


Private Function SafePct( _
    ByVal numerator As Long, _
    ByVal denominator As Long) As Double

    If denominator = 0 Then
        SafePct = 0
    Else
        SafePct = numerator / denominator
    End If

End Function


Private Sub AddInfoButton( _
    ByVal hostWb As Workbook)

    Dim ws As Worksheet
    Dim shp As Shape
    Dim targetCell As Range

    Set ws = hostWb.Worksheets("Transfer_Summary")

    Set targetCell = ws.Cells(29, 2)

    On Error Resume Next
    ws.Shapes("btnGuide").Delete
    On Error GoTo 0

    Set shp = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        targetCell.Left, _
        targetCell.Top, _
        90, _
        18)

    With shp

        .Name = "btnGuide"

        .TextFrame.Characters.Text = _
            "Guide"

        .OnAction = _
            "'" & ThisWorkbook.Name & _
            "'!ToggleInformationSheets"

        'Light blue fill
        .Fill.ForeColor.RGB = RGB(221, 235, 247)

        'Black text
        .TextFrame.Characters.Font.Color = vbBlack

        .TextFrame.Characters.Font.Bold = True

        'Light border
        .Line.ForeColor.RGB = RGB(180, 198, 231)
        
        .Placement = xlMove
        
        .TextFrame.VerticalAlignment = _
            xlVAlignCenter

    End With

End Sub

Private Function SelectITMSourceCell() As Range

    Dim selectedRange As Range

    On Error Resume Next

    Set selectedRange = Application.InputBox( _
        Prompt:= _
            "Select any cell on the transaction source worksheet.", _
        Title:="Select ITM Source Worksheet", _
        Type:=8)

    Err.Clear

    On Error GoTo 0

    Set SelectITMSourceCell = selectedRange

End Function

'*****************************
'
'
'        DEBUGGING
'
' Call Trace("MATCH", MatchID)
' Call Trace("CLUSTER", clusterID)
' Call Trace("STATUS", rowNum, statusValue)
'
'*****************************


'-----------------------------------------
' Match ID - Call DumpMatchInfo("M000108")
'-----------------------------------------

Private Sub DumpMatchInfo( _
    matchID As String)

    Dim ws As Worksheet
    Dim r As Long

    Set ws = hostWb.Worksheets("Transfer_Relationships")

    For r = 2 To ws.Cells( _
        ws.rows.Count, 1).End(xlUp).Row

        If CStr(ws.Cells(r, 1).Value) = _
           matchID Then

            Debug.Print _
                "MATCH", _
                matchID, _
                "ROW", r, _
                "CLUSTER", ws.Cells(r, 11).Value, _
                "SHAPE", ws.Cells(r, 12).Value, _
                "GROUP", ws.Cells(r, 13).Value

            Exit Sub

        End If

    Next r

End Sub

'-----------------------------------------------------
' Transaction Status - Call DumpTransactionStatus(327)
'-----------------------------------------------------

Private Sub DumpTransactionStatus( _
    rowNum As Long)

    If TransactionStatus.Exists(CStr(rowNum)) Then

        Debug.Print _
            "STATUS", _
            rowNum, _
            TransactionStatus(CStr(rowNum))(0), _
            TransactionStatus(CStr(rowNum))(1)

    Else

        Debug.Print _
            "STATUS", _
            rowNum, _
            "NOT FOUND"

    End If

End Sub

'-------------------------------------------
' Dump Cluster - Call DumpCluster("C000029")
'-------------------------------------------
Private Sub DumpCluster( _
    clusterID As String)

    Dim ws As Worksheet
    Dim r As Long

    Set ws = hostWb.Worksheets("Cluster_Analysis")

    For r = 2 To ws.Cells( _
        ws.rows.Count, 1).End(xlUp).Row

        If CStr(ws.Cells(r, 1).Value) = _
           clusterID Then

            Debug.Print _
                "CLUSTER", _
                clusterID, _
                "SHAPE", _
                ws.Cells(r, 12).Value, _
                "MATCHABLE", _
                ws.Cells(r, 10).Value, _
                "OUTCOME", _
                ws.Cells(r, 14).Value

            Exit Sub

        End If

    Next r

End Sub




Private Sub Trace( _
    location As String, _
    ParamArray values())
    
'------------------------------------
' Generic Trace - Call Trace( _
'                   "WRITE MATCH", _
'                   clusterID, _
'                   debitRow, _
'                   creditRow)
'------------------------------------

    Dim i As Long
    Dim msg As String

    msg = location

    For i = LBound(values) To _
             UBound(values)

        msg = msg & " | " & _
              CStr(values(i))

    Next i

    Debug.Print msg

End Sub


Private Sub DebugAccountStatistics( _
    ByVal acctStats As Object, _
    ByVal accountNumber As String)
    
'===================
'Debug by account
' Debug lead found in accont statistics
'    DebugAccountStatistics _
'        acctStats, _
'        "2011014111"
'===================

    Dim stats As Variant

    Dim totalInCount As Long
    Dim totalOutCount As Long
    Dim totalCount As Long

    Dim totalInVolume As Double
    Dim totalOutVolume As Double
    Dim grossVolume As Double
    Dim netVolume As Double

    accountNumber = Trim$(accountNumber)

    Debug.Print String(70, "-")
    Debug.Print "ACCOUNT STATISTICS DEBUG"
    Debug.Print "Account: " & accountNumber

    If Not acctStats.Exists(accountNumber) Then

        Debug.Print "Account not found."
        Debug.Print String(70, "-")
        Exit Sub

    End If

    stats = acctStats(accountNumber)

    totalInCount = _
        stats(ACCT_MATCHED_IN_COUNT) + _
        stats(ACCT_AMBIGUOUS_IN_COUNT) + _
        stats(ACCT_UNMATCHED_IN_COUNT)

    totalOutCount = _
        stats(ACCT_MATCHED_OUT_COUNT) + _
        stats(ACCT_AMBIGUOUS_OUT_COUNT) + _
        stats(ACCT_UNMATCHED_OUT_COUNT)

    totalCount = _
        totalInCount + totalOutCount

    totalInVolume = _
        stats(ACCT_MATCHED_IN_VOLUME) + _
        stats(ACCT_AMBIGUOUS_IN_VOLUME) + _
        stats(ACCT_UNMATCHED_IN_VOLUME)

    totalOutVolume = _
        stats(ACCT_MATCHED_OUT_VOLUME) + _
        stats(ACCT_AMBIGUOUS_OUT_VOLUME) + _
        stats(ACCT_UNMATCHED_OUT_VOLUME)

    grossVolume = _
        totalInVolume + totalOutVolume

    netVolume = _
        totalInVolume - totalOutVolume

    Debug.Print "COUNTS"
    Debug.Print _
        "  Matched In:       " & _
        stats(ACCT_MATCHED_IN_COUNT)

    Debug.Print _
        "  Matched Out:      " & _
        stats(ACCT_MATCHED_OUT_COUNT)

    Debug.Print _
        "  Ambiguous In:     " & _
        stats(ACCT_AMBIGUOUS_IN_COUNT)

    Debug.Print _
        "  Ambiguous Out:    " & _
        stats(ACCT_AMBIGUOUS_OUT_COUNT)

    Debug.Print _
        "  Unmatched In:     " & _
        stats(ACCT_UNMATCHED_IN_COUNT)

    Debug.Print _
        "  Unmatched Out:    " & _
        stats(ACCT_UNMATCHED_OUT_COUNT)

    Debug.Print _
        "  Total Transfers:  " & totalCount

    Debug.Print "VOLUME"

    Debug.Print _
        "  Matched In:       " & _
        Format$( _
            stats(ACCT_MATCHED_IN_VOLUME), _
            "$#,##0.00")

    Debug.Print _
        "  Matched Out:      " & _
        Format$( _
            stats(ACCT_MATCHED_OUT_VOLUME), _
            "$#,##0.00")

    Debug.Print _
        "  Ambiguous In:     " & _
        Format$( _
            stats(ACCT_AMBIGUOUS_IN_VOLUME), _
            "$#,##0.00")

    Debug.Print _
        "  Ambiguous Out:    " & _
        Format$( _
            stats(ACCT_AMBIGUOUS_OUT_VOLUME), _
            "$#,##0.00")

    Debug.Print _
        "  Unmatched In:     " & _
        Format$( _
            stats(ACCT_UNMATCHED_IN_VOLUME), _
            "$#,##0.00")

    Debug.Print _
        "  Unmatched Out:    " & _
        Format$( _
            stats(ACCT_UNMATCHED_OUT_VOLUME), _
            "$#,##0.00")

    Debug.Print _
        "  Total In:         " & _
        Format$(totalInVolume, "$#,##0.00")

    Debug.Print _
        "  Total Out:        " & _
        Format$(totalOutVolume, "$#,##0.00")

    Debug.Print _
        "  Gross Volume:     " & _
        Format$(grossVolume, "$#,##0.00")

    Debug.Print _
        "  Net Volume:       " & _
        Format$(netVolume, "$#,##0.00")

    Debug.Print "NETWORK"

    Debug.Print _
        "  Known Sources:    " & _
        stats(ACCT_SOURCE_DICT).Count

    Debug.Print _
        "  Known Destinations: " & _
        stats(ACCT_DESTINATION_DICT).Count

    Debug.Print "LARGEST TRANSACTIONS"

    Debug.Print _
        "  Largest Received: " & _
        Format$( _
            stats(ACCT_LARGEST_RECEIVED), _
            "$#,##0.00")

    Debug.Print _
        "  Received Status:  " & _
        CStr(stats( _
            ACCT_LARGEST_RECEIVED_STATUS))

    Debug.Print _
        "  Largest Sent:     " & _
        Format$( _
            stats(ACCT_LARGEST_SENT), _
            "$#,##0.00")

    Debug.Print _
        "  Sent Status:      " & _
        CStr(stats( _
            ACCT_LARGEST_SENT_STATUS))

    Debug.Print String(70, "-")

End Sub

'===========================
' Ambiguity Debug
'===========================
'Private Sub DebugAmbiguityRow( _
'    ByVal stageName As String, _
'    ByRef data As Variant, _
'    ByVal rowNum As Long, _
'    ByVal colAcct As Long, _
'    ByVal colDate As Long, _
'    ByVal colAmount As Long, _
'    ByVal colDescription As Long, _
'    ByRef matched() As Boolean, _
'    ByRef ambiguous() As Boolean)
'
'    If Not DEBUG_AMBIGUITY Then Exit Sub
'
'    If rowNum < LBound(data, 1) Or _
'            rowNum > UBound(data, 1) Then
'
'        Debug.Print _
'            "AMBIGUITY DEBUG | " & stageName & _
'            " | Invalid row: " & rowNum
'
'        Exit Sub
'
'    End If
'
'    Debug.Print _
'        "AMBIGUITY DEBUG" & _
'        " | Stage=[" & stageName & "]" & _
'        " | Row=" & rowNum & _
'        " | Account=[" & CStr(data(rowNum, colAcct)) & "]" & _
'        " | Date=[" & CStr(data(rowNum, colDate)) & "]" & _
'        " | Amount=[" & CStr(data(rowNum, colAmount)) & "]" & _
'        " | Matched=" & CStr(matched(rowNum)) & _
'        " | Ambiguous=" & CStr(ambiguous(rowNum)) & _
'        " | Description=[" & _
'            Left$( _
'                CStr(data(rowNum, colDescription)), _
'                80) & "]"
'
'End Sub

Private Function CountTrueFlags( _
    ByRef flags() As Boolean, _
    ByVal firstRow As Long, _
    ByVal lastRow As Long) As Long

    Dim rowNum As Long
    Dim safeFirstRow As Long
    Dim safeLastRow As Long

    safeFirstRow = firstRow
    safeLastRow = lastRow

    If safeFirstRow < LBound(flags) Then
        safeFirstRow = LBound(flags)
    End If

    If safeLastRow > UBound(flags) Then
        safeLastRow = UBound(flags)
    End If

    If safeLastRow < safeFirstRow Then Exit Function

    For rowNum = safeFirstRow To safeLastRow

        If flags(rowNum) Then
            CountTrueFlags = CountTrueFlags + 1
        End If

    Next rowNum

End Function

Private Sub DebugAmbiguityCheckpoint( _
    ByVal procedureName As String, _
    ByVal checkpointName As String, _
    ByVal hostWb As Workbook, _
    ByVal ws As Worksheet, _
    ByRef data As Variant, _
    ByVal lastRow As Long, _
    ByVal colAcct As Long, _
    ByVal colCodeDesc As Long, _
    ByVal colDate As Long, _
    ByVal colAmount As Long, _
    ByVal colDescription As Long, _
    ByRef matched() As Boolean, _
    ByRef ambiguous() As Boolean, _
    ByVal ambiguityPairs As Collection, _
    Optional ByVal sourceRow As Long = 0, _
    Optional ByVal candidateRow As Long = 0, _
    Optional ByVal methodName As String = "", _
    Optional ByVal messageText As String = "")

    Dim pairCount As Long
    Dim matchedCount As Long
    Dim ambiguousCount As Long

    If Not DEBUG_AMBIGUITY Then Exit Sub

    matchedCount = _
        CountTrueFlags( _
            matched, _
            2, _
            lastRow)

    ambiguousCount = _
        CountTrueFlags( _
            ambiguous, _
            2, _
            lastRow)

    If ambiguityPairs Is Nothing Then
        pairCount = 0
    Else
        pairCount = ambiguityPairs.Count
    End If

    Debug.Print String$(78, "=")
    Debug.Print "AMBIGUITY CHECKPOINT"
    Debug.Print "Procedure:       " & procedureName
    Debug.Print "Checkpoint:      " & checkpointName

    If Not hostWb Is Nothing Then

        Debug.Print _
            "Host workbook:   " & hostWb.Name

    Else

        Debug.Print _
            "Host workbook:   [Not provided]"

    End If

    If Not ws Is Nothing Then

        Debug.Print _
            "Source sheet:    " & ws.Name

    Else

        Debug.Print _
            "Source sheet:    [Not provided]"

    End If

    Debug.Print "Last data row:   " & lastRow
    Debug.Print "Matched rows:    " & matchedCount
    Debug.Print "Ambiguous rows:  " & ambiguousCount
    Debug.Print "Ambiguity pairs: " & pairCount

    If Len(methodName) > 0 Then

        Debug.Print _
            "Match method:    " & methodName

    End If

    If Len(messageText) > 0 Then

        Debug.Print _
            "Message:         " & messageText

    End If

    If sourceRow > 0 Then

        DebugAmbiguityCheckpointRow _
            "Source", _
            data, _
            sourceRow, _
            lastRow, _
            colAcct, _
            colCodeDesc, _
            colDate, _
            colAmount, _
            colDescription, _
            matched, _
            ambiguous

    End If

    If candidateRow > 0 Then

        DebugAmbiguityCheckpointRow _
            "Candidate", _
            data, _
            candidateRow, _
            lastRow, _
            colAcct, _
            colCodeDesc, _
            colDate, _
            colAmount, _
            colDescription, _
            matched, _
            ambiguous

    End If

    Debug.Print String$(78, "=")

End Sub


Private Sub DebugAmbiguityCheckpointRow( _
    ByVal rowRole As String, _
    ByRef data As Variant, _
    ByVal rowNum As Long, _
    ByVal lastRow As Long, _
    ByVal colAcct As Long, _
    ByVal colCodeDesc As Long, _
    ByVal colDate As Long, _
    ByVal colAmount As Long, _
    ByVal colDescription As Long, _
    ByRef matched() As Boolean, _
    ByRef ambiguous() As Boolean)

    If Not DEBUG_AMBIGUITY Then Exit Sub

    If rowNum < LBound(data, 1) Or _
            rowNum > UBound(data, 1) Then

        Debug.Print _
            rowRole & " row:      " & rowNum & _
            " [Outside data-array bounds]"

        Exit Sub

    End If

    If rowNum < LBound(matched) Or _
            rowNum > UBound(matched) Then

        Debug.Print _
            rowRole & " row:      " & rowNum & _
            " [Outside matched-array bounds]"

        Exit Sub

    End If

    If rowNum < LBound(ambiguous) Or _
            rowNum > UBound(ambiguous) Then

        Debug.Print _
            rowRole & " row:      " & rowNum & _
            " [Outside ambiguous-array bounds]"

        Exit Sub

    End If

    Debug.Print String$(78, "-")

    Debug.Print _
        rowRole & " row:      " & rowNum

    Debug.Print _
        rowRole & " account:  [" & _
        CStr(data(rowNum, colAcct)) & "]"

    Debug.Print _
        rowRole & " code:     [" & _
        CStr(data(rowNum, colCodeDesc)) & "]"

    Debug.Print _
        rowRole & " date:     [" & _
        CStr(data(rowNum, colDate)) & "]"

    Debug.Print _
        rowRole & " amount:   [" & _
        CStr(data(rowNum, colAmount)) & "]"

    Debug.Print _
        rowRole & " matched:  " & _
        CStr(matched(rowNum))

    Debug.Print _
        rowRole & " ambiguous:" & _
        CStr(ambiguous(rowNum))

    Debug.Print _
        rowRole & " narrative:[" & _
        Left$( _
            CStr(data(rowNum, colDescription)), _
            100) & "]"

End Sub
