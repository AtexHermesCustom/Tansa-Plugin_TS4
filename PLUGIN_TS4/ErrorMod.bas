Attribute VB_Name = "ErrorMod"
Option Explicit

'error handler for Tansa
Private m_oTSErrHandler As TS4C.errorHandler

Public Sub ErrH(lErrNum As Long, sErrDesc As String, sErrSrc As String, sProcName As String, bShowMsg As Boolean)
On Error GoTo EH
    'log error
    WriteLog "ERROR at " & sProcName & ". ErrNum: " & lErrNum & _
            ", Desc: " & sErrDesc & ", Source: " & sErrSrc

    If bShowMsg Then    'display error
        MsgBox "ERROR at procedure " & sProcName & vbNewLine & _
            "Err Number: " & lErrNum & vbNewLine & _
            "Description: " & sErrDesc, vbCritical
    End If
    Exit Sub
EH:
    MsgBox "ErrH ERROR: " & Err.Number & ", " & Err.Description & ", " & Err.Source, vbCritical
End Sub

'Tansa error handler
Public Property Get ErrHandler() As TS4C.errorHandler
    If m_oTSErrHandler Is Nothing Then
      Set m_oTSErrHandler = New TS4C.errorHandler
    End If
    
    Set ErrHandler = m_oTSErrHandler
End Property


