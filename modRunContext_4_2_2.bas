Attribute VB_Name = "modRunContext"
Option Explicit

Private mHostWb As Workbook
Private mSourceWs As Worksheet
Private mCodeWb As Workbook
Private mRunContextInitialized As Boolean

Public Sub InitializeRunContext( _
    ByVal selectedHostWb As Workbook, _
    ByVal selectedSourceWs As Worksheet)

    'Clear references left from a prior run.
    Set mSourceWs = Nothing
    Set mHostWb = Nothing
    Set mCodeWb = Nothing

    mRunContextInitialized = False


    ' Validate Workbook


    If selectedHostWb Is Nothing Then

        Err.Raise _
            vbObjectError + 1000, _
            "InitializeRunContext", _
            "No transaction workbook was provided."

    End If

    If selectedSourceWs Is Nothing Then

        Err.Raise _
            vbObjectError + 1001, _
            "InitializeRunContext", _
            "No transaction source worksheet was provided."

    End If

    If Not selectedSourceWs.Parent Is selectedHostWb Then

        Err.Raise _
            vbObjectError + 1002, _
            "InitializeRunContext", _
            "The selected source worksheet does not belong " & _
            "to the selected transaction workbook."

    End If


    ' Store environment


    Set mHostWb = selectedHostWb
    Set mSourceWs = selectedSourceWs
    Set mCodeWb = ThisWorkbook

    If mHostWb Is mCodeWb Then

        ClearRunContext

        Err.Raise _
            vbObjectError + 1003, _
            "InitializeRunContext", _
            "The workbook containing the ITM code cannot " & _
            "be used as the transaction workbook."

    End If

    mRunContextInitialized = True

    Debug.Print "ITM RUN CONTEXT INITIALIZED"

End Sub

Public Property Get hostWb() As Workbook

    ValidateRunContext "HostWb"

    Set hostWb = mHostWb

End Property


Public Property Get sourceWs() As Worksheet

    ValidateRunContext "SourceWs"

    Set sourceWs = mSourceWs

End Property


Public Property Get CodeWb() As Workbook

    ValidateRunContext "CodeWb"

    Set CodeWb = mCodeWb

End Property


Public Property Get RunContextIsInitialized() As Boolean

    RunContextIsInitialized = mRunContextInitialized

End Property

Private Sub ValidateRunContext( _
    ByVal requestedObject As String)

    If Not mRunContextInitialized Then
        Err.Raise vbObjectError + 1010, _
                  "ValidateRunContext", _
                  "The ITM run context has not been initialized. " & _
                  "Requested object: " & requestedObject
    End If

    If mHostWb Is Nothing Then
        Err.Raise vbObjectError + 1011, _
                  "ValidateRunContext", _
                  "The ITM host workbook is not available."
    End If

    If mSourceWs Is Nothing Then
        Err.Raise vbObjectError + 1012, _
                  "ValidateRunContext", _
                  "The ITM source worksheet is not available."
    End If

    If mCodeWb Is Nothing Then
        Err.Raise vbObjectError + 1013, _
                  "ValidateRunContext", _
                  "The ITM code workbook is not available."
    End If

End Sub


Public Sub ClearRunContext()

    Set mSourceWs = Nothing
    Set mHostWb = Nothing
    Set mCodeWb = Nothing

    mRunContextInitialized = False

    Debug.Print "ITM run context cleared."

End Sub

