VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Newsroom Tansa Plugin"
   ClientHeight    =   1395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   1560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer myTimer 
      Interval        =   200
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'**************************************************************************************
' NewsroomTansaPlugin
'
' Version History:
'
' 4.0.0.4 20150215 jpm: updated NRCOM functions to support Unicode
'   -use UTF-16 versions of the following functions: GetContentCom, Replace
'   -replace message to "MSG_TSNE_MESSAGE"
' 4.0.0.3 20110822 jpm: used Microsoft XML, v4.0 instead of Microsoft XML, v3.0
'   as suggested by Espen
' 4.0.0.2 20110819 jpm: removed reference to scrrun.dll by
'   changing the way files' and folders' existence is checked.
' 4.0.0.1 20110520 jpm: added support for Special Spaces:
'   -mapping between Newsroom space commands and special Unicode chars in Tansa
' 4.0.0.0 20100629 jpm: updated to work with Tansa 4
'
' 3.6.0.12 20090724 jpm: workaround for an NRCOM issue:
'   issue when text within notes are replaced and then the Show/Hide Commands menu function is called,
'   space after the corrected text is removed
' 3.6.0.11 20090213 jpm: get Newsroom path from INI file, not from xml config
'   since installer package can edit the INI file during installation, but not an xml config.
' 3.6.0.10 20071127 jpm:
'   default log path to %TEMP% to make sure user has write access to the folder
' 3.6.0.9 20071126 jpm:
'   just updated version number so new installer package can be released
' 3.6.0.8 20071121 jpm:
'   just updated version number so new package can be released
'   to go with nrxtansa.nrx v1.0.0.2 release (fix: run custom functions only when icon is clicked)
' 3.6.0.7 20071116 jpm:
'   just updated version number so new installer package can be released
' 3.6.0.6 20071115 jpm:
'   use DOMDocument30 (mxsml3.dll) instead of DOMDocument40 (mxsml4.dll)
'   msxml3.dll is included with Windows XP by default whereas msxml4.dll is not
' 3.6.0.5 20070920 jpm:
'   just updated version number so new installer package can be released
' 3.6.0.4 20070808 jpm:
'   new way of handling WEBHED's and notes - WEBHEAD tag outside of the note:
'       [WEBHED]<NO1>note content to be proofed<NO>[/WEBHED]
'   handling double spaces
'   handling notes for proofing that are concatenated with regular text.
'       e.g. <NO1>note<NO>regular text - by using NotesMarker read from config file
'   EscapeCommandChars option - whether to escape special commands entered by user or not
'   handling for spacecommands - e.g. <EM>, <EN>, <TH>
' 3.6.0.3 20070731 jpm:
'   just updated version number so new installer package can be released
' 3.6.0.2 20070707 jpm:
'   escape special chars (< > [ ] {} \) in the text corrections received from Tansa
'   load list of Newsroom invisible commands (WC1 WC, carton delimiters (eg. {TEXT}) and US commands)
'       from config file and remove them from text before proofing
'   call RemoveInvisibleCommands before HandleNotes-so any invisible commands will
'       not be substituted and retained before getting deleted
'   do not substitute escaped chars and esc chars in notes with the subst. char
'       -their pos need to be recorded
'   turn on/off appropriate Newsroom menu items during start and set them back
'       to their orig. values when done
' 3.6.0.1 20070606 jpm:
'   provided functionality to run Tansa proofing and hyphenation
'   through keyboard shortcuts
' 3.6.0.0:
'   initial release
'
'**************************************************************************************

Dim bShowTansaMenu As Boolean
Dim bRunTansaProofing As Boolean
Dim bRunTansaHyphenation As Boolean

Private Sub Form_Load()
On Error GoTo EH
    If App.PrevInstance Then End 'exit if another instance is already running

    Dim bInitOK As Boolean
    bInitOK = False
    
    If LoadConfig Then  'load config info
    
        WriteLog Format$(Now, TIMESTAMPFORMAT) & " *** START Newsroom Tansa Plugin ***"
        WriteLog "Info: Windows user=" & Environ("UserName") & _
            ", Workstation=" & Environ("ComputerName") & _
            ", Parent App " & PARENTAPPNAME & " Version=" & g_sParentAppVersion & _
            ", Plugin Version=" & App.Major & "." & App.Minor & ".0." & App.Revision

        If InitNRCOM Then   'initialize NRCOM
            If ConnectToTansa Then  'Connect to Tansa server
                bInitOK = True
                frmMain.Show    'load main form
            Else
                WriteLog "Error connecting to Tansa server"
                MsgBox "error connecting to Tansa server", vbCritical
            End If
        End If
        
    End If

    If bInitOK Then 'successful init
        With Me
            .Caption = MSG_CLIENTTITLE 'important to set caption for form window for msg purposes
            'plugin form "hide"
            .Height = 1
            .Width = 1
            
            HookForm .hwnd    'use custom windows handler for messages from Newsroom
        End With
    Else
        MsgBox "Newsroom Tansa Plugin was not successfully initialized." & vbNewLine & _
            "Its functions will not be available.", vbCritical
        End 'exit
    End If
    
    Exit Sub
EH:
    If Trim$(g_sLogFile) = "" Then
        MsgBox "Form_Load ERROR: " & Err.Number & ", " & Err.Description & ", " & Err.Source, vbCritical
    Else
        ErrH Err.Number, Err.Description, Err.Source, "Form_Load", True
    End If
    
    MsgBox "Newsroom Tansa Plugin was not successfully initialized." & vbNewLine & _
        "Its functions will not be available.", vbCritical
    End 'exit
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo EH
    WriteLog "Closing plugin..."

    Set oTansaPlugin = Nothing

    If Not (oNR Is Nothing) Then
        'just making sure that newsroom is unlocked
        'oNR.Unlock - not needed.. newsroom is being closed anyway
        Set oNR = Nothing
    End If

    'cleanup old logs
    DeleteExpiredLogs g_iLogRetentionDays

    WriteLog Format$(Now, TIMESTAMPFORMAT) & " *** END Newsroom Tansa Plugin ***"

    UnhookForm Me.hwnd
    Exit Sub
EH:
    ErrH Err.Number, Err.Description, Err.Source, "Form_Unload", True
End Sub

Private Function LoadConfig() As Boolean
On Error GoTo EH
    Dim sTemp As String
    Dim configDoc As MSXML2.DOMDocument30
    Dim sIniFile As String
    
    'load configuration file
    Set configDoc = New MSXML2.DOMDocument30
    With configDoc
        .async = False
        .Load App.Path & "\" & CONFIGFILE
    End With
    
    'set ini file path
    sIniFile = App.Path & "\" & INIFILE
    
    
    
    'load config info
    '----------------------------------------------
    If Not LoadConfigNode(configDoc, "//PluginConfiguration/LogPath", g_sLogPath) Then Exit Function
    If Trim$(g_sLogPath) = "" Then g_sLogPath = Environ("Temp") 'default to profile's temporary folder to make sure user has access
    If Not InitLog Then Exit Function 'initialize log
    
    '3.6.0.11 20090213 jpm: get Newsroom path from INI file, not from xml config
    ' since installer package can edit the INI file during installation, but not an xml config.
    'If Not LoadConfigNode(configDoc, "//PluginConfiguration/ParentAppBin", g_sParentAppBin) Then Exit Function
    g_sParentAppBin = GetIniStr(sIniFile, "Newsroom", "Path")
    If Trim(g_sParentAppBin) = "" Then
        MsgBox "ERROR: Cannot find path of parent application (Newsroom). Check the plugin INI file.", vbCritical
        Exit Function
    End If
    If Not FileExists(g_sParentAppBin) Then
        MsgBox "ERROR: Parent application " & g_sParentAppBin & " does not exist.", vbCritical
        Exit Function
    End If
    'get version of parent application (Newsroom) - use API to get it
    g_sParentAppVersion = GetProductVersion(g_sParentAppBin)
    
    If Not LoadConfigNode(configDoc, "//PluginConfiguration/ParentAppUserName", g_sParentAppUserName) Then Exit Function
    'cannot retrieve Hermes user through NRCOM, record Windows user instead
    If UCase$(Trim$(g_sParentAppUserName)) = "WINDOWSUSER" Or Trim$(g_sParentAppUserName) = "" Then _
        g_sParentAppUserName = Environ("UserName")
    
    If Not LoadConfigNode(configDoc, "//PluginConfiguration/ParentAppUiLanguageCode", g_sParentAppUiLanguageCode) Then Exit Function
    
    If Not LoadConfigNode(configDoc, "//PluginConfiguration/SoftHyphenCharCode", sTemp) Then Exit Function
    g_lSoftHyphenCharCode = Val(Trim$(sTemp))
    
    If Not LoadConfigNode(configDoc, "//PluginConfiguration/CommandSubCharCode", sTemp) Then Exit Function
    g_lCommandSubCharCode = Val(Trim$(sTemp))
    
    If Not LoadConfigNode(configDoc, "//PluginConfiguration/EscapeCommandChars", sTemp) Then Exit Function
    g_bEscapeCommandChars = IIf(Trim$(sTemp) = "1", True, False)
    
    If Not LoadConfigNode(configDoc, "//PluginConfiguration/CommandChars", g_sCommandChars) Then Exit Function
    
    If Not LoadConfigNode(configDoc, "//PluginConfiguration/NotesMarker", g_sNotesMarker) Then Exit Function
    
    If Not LoadConfigNode(configDoc, "//PluginConfiguration/SaveBeforeProofing", sTemp) Then Exit Function
    g_bSaveBeforeProofing = IIf(Trim$(sTemp) = "1", True, False)
    
    If Not LoadConfigNode(configDoc, "//PluginConfiguration/DebugMode", sTemp) Then Exit Function
    g_bDebug = IIf(Trim$(sTemp) = "1", True, False)

    If Not LoadConfigNode(configDoc, "//PluginConfiguration/LogRetentionDays", sTemp) Then Exit Function
    g_iLogRetentionDays = Val(Trim$(sTemp))
    
    
    'get list of tags to be used for proofing notes content
    Set g_xdomCheckTags = configDoc.selectNodes("//CheckTags/Tag")
    
    'get list of commands that will never be visible in Newsroom-commands that shouldn't be included in char pos counting
    Set g_xdomInvisibleCommands = configDoc.selectNodes("//InvisibleCommands/Command")
    
    'get list of commands that can represent space
    Set g_xdomSpaceCommands = configDoc.selectNodes("//SpaceCommands/Command")
    
    '4.0.0.1 20110520 put the special spaces in a collection for later use
    Dim node As IXMLDOMNode
    Dim attrCode As IXMLDOMNode
    Dim lSubCharCode As Long
    
    Set g_colSpecialSpaces = New Collection
     
    For Each node In g_xdomSpaceCommands
        Set attrCode = node.Attributes.getNamedItem("subCharCode")
        If Not (attrCode Is Nothing) Then
            If IsNumeric(Trim$(attrCode.nodeValue)) Then
                lSubCharCode = CLng(Val(Trim$(attrCode.nodeValue)))
                g_colSpecialSpaces.Add node.Text, CStr(lSubCharCode) 'use the sub char code as the key
            End If
        End If
    Next


    Set configDoc = Nothing

    LoadConfig = True
    Exit Function
EH:
    Set configDoc = Nothing
    Err.Raise Err.Number, "LoadConfig:" & Err.Source, Err.Description
End Function

Private Function LoadConfigNode(ByRef configDoc As MSXML2.DOMDocument30, _
    ByVal sNodeName As String, ByRef sNodeText As String) As Boolean
On Error GoTo EH
    Dim node As MSXML2.IXMLDOMNode
    
    Set node = configDoc.selectSingleNode(sNodeName)
    If Not node Is Nothing Then
        sNodeText = node.Text
        LoadConfigNode = True
    Else
        If Trim$(g_sLogFile) <> "" Then WriteLog "ERROR: Config node " & sNodeName & " cannot be found in " & CONFIGFILE
        MsgBox "ERROR: Configuration node " & sNodeName & " cannot be found in " & CONFIGFILE, vbCritical
        LoadConfigNode = False
    End If

    Set node = Nothing
    Exit Function
EH:
    Err.Raise Err.Number, "LoadConfigNode:" & Err.Source, Err.Description
End Function

Private Function InitNRCOM() As Boolean
On Error GoTo EH
    'get reference to the newsroom extension object
    Set oNR = Nothing
    Set oNR = GetObject(, "NewsRoom.Extension")
    InitNRCOM = True
    Exit Function
EH:
    If Err.Number = 429 Then
        WriteLog "Newsroom is not running."
        MsgBox "Newsroom is not running.", vbExclamation
    Else
        Err.Raise Err.Number, "InitNRCOM:" & Err.Source, Err.Description
    End If
End Function

Private Function ConnectToTansa() As Boolean
On Error GoTo EH
    Set oTansaPlugin = Nothing
    Set oTansaPlugin = New EdClientIntegration  'create class and connection to Tansa
    
    If Not (oTansaPlugin Is Nothing) Then
        ConnectToTansa = True
    Else
        WriteLog "Error encountered while connecting to Tansa."
        MsgBox "Error encountered while connecting to Tansa.", vbCritical
        ConnectToTansa = False
    End If
    Exit Function
EH:
    Err.Raise Err.Number, "ConnectToTansa:" & Err.Source, Err.Description
End Function

Private Sub ShowTansaMenu()
On Error GoTo EH
    'make sure newsroom is running
    If Not (oNR Is Nothing) Then
        'make sure an object is open
        If (oNR.IsTextOpened) Then
            Dim oTansaPopup As TS4C.PopupMenu
            Set oTansaPopup = New TS4C.PopupMenu
            
            'show Tansa menu at current mouse pos
            oTansaPopup.ShowMenuAtPointerPosition
            'just making sure that:
            oNR.Unlock  'newsroom is unlocked after Tansa function completed
        Else
            MsgBox "No Newsroom object is opened.", vbExclamation
        End If
    Else
        MsgBox "Newsroom is not running.", vbExclamation
    End If
    
    Exit Sub
EH:
    ErrH Err.Number, Err.Description, Err.Source, "ShowTansaMenu", True
End Sub

Public Sub GotMessage(lHandle As Long, lMsg As Long)
On Error GoTo EH
    If g_bDebug Then WriteLog ">>>GotMessage. received nroom handle=" & lHandle & ", msg=" & lMsg
    
    'plugin form "hide"
    Me.Height = 1
    Me.Width = 1

    'set plugin's forms to always stay on top of Newsroom
    Call SetWindowLong(Me.hwnd, SWW_HPARENT, lHandle)
    
    Select Case lMsg
        
        Case MSG_CLOSEPLUGIN 'msg code to close the plugin
            Unload Me
        
        Case MSG_SHOWTANSAMENU 'msg code to show menu
            bShowTansaMenu = True
        
        Case MSG_RUNTANSAPROOFING 'msg code to run proofing
            bRunTansaProofing = True
        
        Case MSG_RUNTANSAHYPHENATION 'msg code to run hyphenation
            bRunTansaHyphenation = True
        
        Case Else
            WriteLog ">>>GotMessage. Plugin received unknown message. value=" & lMsg
            MsgBox "Plugin received unknown message. value=" & lMsg, vbExclamation
    
    End Select
    Exit Sub
EH:
    ErrH Err.Number, Err.Description, Err.Source, "GotMessage", True
End Sub

Private Sub myTimer_Timer()
On Error GoTo EH
    'functions are called through a timer
    'GotMessage cannot call the functions directly because this causes an automation error
    'workaround recommeded by Microsoft is to use boolean flags (which are turned on in GotMessage)
    'and a timer component
    If bShowTansaMenu Then
        myTimer.Enabled = False
        ShowTansaMenu
        bShowTansaMenu = False
        myTimer.Enabled = True
        
    ElseIf bRunTansaProofing Then
        myTimer.Enabled = False
        RunTansaProofing
        bRunTansaProofing = False
        myTimer.Enabled = True
        
    ElseIf bRunTansaHyphenation Then
        myTimer.Enabled = False
        RunTansaHyphenation
        bRunTansaHyphenation = False
        myTimer.Enabled = True
    End If
    Exit Sub
EH:
    ErrH Err.Number, Err.Description, Err.Source, "myTimer_Timer", True
    myTimer.Enabled = True
End Sub

Private Function RunTansaProofing() As Boolean
On Error GoTo EH
    If g_bDebug Then WriteLog ">>>RunTansaProofing"
    
    Dim oServices As TS4C.Services
    Set oServices = New TS4C.Services
    Call oServices.RunTansaProofing(uiAll)
    
    RunTansaProofing = True
    Exit Function
EH:
    ErrH Err.Number, Err.Description, Err.Source, "RunTansaProofing", True
End Function

Private Function RunTansaHyphenation() As Boolean
On Error GoTo EH
    If g_bDebug Then WriteLog ">>>RunTansaHyphenation"
    
    Dim oServices As TS4C.Services
    Set oServices = New TS4C.Services
    Call oServices.RunTansaHyphenation(uiAll)

    RunTansaHyphenation = True
    Exit Function
EH:
    ErrH Err.Number, Err.Description, Err.Source, "RunTansaHyphenation", True
End Function

