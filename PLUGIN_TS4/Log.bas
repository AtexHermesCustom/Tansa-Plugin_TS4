Attribute VB_Name = "Log"
Option Explicit

'constants
Private Const LOGFILENAME As String = "NewsroomTansaPlugin_YYYYMMDD.log"
Private Const LOGDATEFORMAT As String = "YYYYMMDD"
Public Const TIMESTAMPFORMAT As String = "yyyy-mm-dd HH:mm:ss"

Public g_sLogFile As String

Public Function InitLog() As Boolean
On Error GoTo EH
    If FolderExists(g_sLogPath) Then
        g_sLogFile = g_sLogPath & "\" & Replace$(LOGFILENAME, LOGDATEFORMAT, Format$(Now, LOGDATEFORMAT))
        InitLog = True
    Else
        MsgBox "ERROR: Log path " & g_sLogPath & " does not exist.", vbCritical
        InitLog = False
    End If
    Exit Function
EH:
    Err.Raise Err.Number, "InitLog:" & Err.Source, Err.Description
End Function

Public Sub WriteLog(sMsg As String)
On Error GoTo EH
    Dim iLogNum As Integer
    iLogNum = FreeFile
    
    If FileExists(g_sLogFile) Then
        Open g_sLogFile For Append As iLogNum
    Else
        Open g_sLogFile For Output As iLogNum
    End If
    
    Print #iLogNum, sMsg
    Close #iLogNum
    Exit Sub
EH:
    MsgBox "WriteLog ERROR: " & Err.Number & ", " & Err.Description & ", " & Err.Source, vbCritical
End Sub

Public Sub DeleteExpiredLogs(iLogRetentionDays)
On Error GoTo EH
    If g_bDebug Then WriteLog ">>>DeleteExpiredLogs"

    Dim sExpireDate As String
    Dim sLog As String
    
    If Trim$(g_sLogPath) <> "" Then
        sExpireDate = Format$(CDate(Now - iLogRetentionDays), LOGDATEFORMAT)
        
        sLog = Dir$(g_sLogPath & "\" & Replace$(LOGFILENAME, LOGDATEFORMAT, "*"))
        
        Do While Len(sLog)
            If UCase(Left$(sLog, 20)) = UCase(Left$(LOGFILENAME, 20)) And _
                UCase(Right$(sLog, 4)) = UCase(Right$(LOGFILENAME, 4)) And _
                Mid$(sLog, 21, 8) < sExpireDate Then
                
                'delete log file
                If g_bDebug Then WriteLog "Delete file: " & g_sLogPath & "\" & sLog
                Kill (g_sLogPath & "\" & sLog)
            End If

            sLog = Dir$
        Loop
    End If
    
    Exit Sub
EH:
    Err.Raise Err.Number, "DeleteExpiredLogs:" & Err.Source, Err.Description
End Sub
