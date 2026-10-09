Attribute VB_Name = "modDashboardNavigation"
Option Explicit




Public Sub AddNavigationButton( _
    ByVal ws As Worksheet, _
    ByVal buttonName As String, _
    ByVal buttonText As String, _
    ByVal targetCell As Range, _
    ByVal macroName As String)

    Dim shp As Shape
    Dim shapeFound As Boolean
    Dim buttonWidth As Double
    Dim buttonHeight As Double
    
    For Each shp In ws.Shapes
    
        If StrComp( _
            shp.Name, _
            buttonName, _
            vbTextCompare) = 0 Then
    
            shp.Delete
    
            shapeFound = True
    
            Exit For
    
        End If
    
    Next shp

    buttonWidth = targetCell.Width
    buttonHeight = targetCell.Height

    If buttonWidth < 50 Then
        buttonWidth = 50
    End If

    If buttonHeight < 18 Then
        buttonHeight = 18
    End If

    Set shp = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        targetCell.Left, _
        targetCell.Top, _
        targetCell.Width, _
        targetCell.Height)
        
    With shp
    
        .Name = buttonName
    
        .TextFrame.Characters.Text = _
            buttonText
    
        .OnAction = _
            "'" & ThisWorkbook.Name & "'!" & _
            macroName
    
        'Light blue background.
        .Fill.ForeColor.RGB = _
            RGB(221, 235, 247)
    
        'Medium blue outline.
        .Line.ForeColor.RGB = _
            RGB(91, 155, 213)
    
        'Dark navy button text.
        With .TextFrame.Characters.Font
    
            .Color = _
                RGB(31, 78, 121)
    
            .Bold = True
            .size = 9
    
        End With
    
        'Center the text within the target cell.
        .TextFrame.HorizontalAlignment = _
            xlHAlignCenter
    
        .TextFrame.VerticalAlignment = _
            xlVAlignCenter
    
        .Placement = _
            xlMoveAndSize
    
    End With

End Sub

Private Sub OpenReportSheet( _
    ByVal sheetName As String)

    Dim hostWb As Workbook
    Dim targetWs As Worksheet

    'The active Executive Summary belongs to the host workbook.
    If ActiveSheet Is Nothing Then
        MsgBox _
            "No active worksheet was found.", _
            vbExclamation, _
            "ITM"
        Exit Sub
    End If

    Set hostWb = ActiveSheet.Parent

    If Not WorksheetExists(hostWb, sheetName) Then
        MsgBox _
            sheetName & " sheet not found.", _
            vbExclamation, _
            "ITM"
        Exit Sub
    End If

    Set targetWs = hostWb.Worksheets(sheetName)

    'This also handles sheets that are currently hidden.
    targetWs.Visible = xlSheetVisible
    targetWs.Activate

End Sub



Public Sub OpenInvestigationGroups()

    OpenReportSheet _
        "Investigation_Groups"

End Sub

Public Sub OpenTransferNetworkAnalysis()

    OpenReportSheet _
        "Transfer_Network_Analysis"

End Sub

Public Sub OpenPrimaryHubAnalysis()

    OpenReportSheet _
        "Primary_Hub_Analysis"

End Sub

Public Sub OpenGuide()

    OpenReportSheet _
        "ITM_Guide"

End Sub

Public Sub OpenMetadata()

    OpenReportSheet _
        "ITM_Metadata"

End Sub

