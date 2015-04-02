VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EdClientIntegration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements TS4C.IExtension

'some notes:
'VB - 1 based
'Tansa - 0 based
'Nrcom - 0 based
'Regexp - 0 based

Private m_oConnection As TS4C.Connection

Private m_sTextToProof As String
Private m_sOrigText As String 'contains original text from Newsroom

Private m_colNewlinePos As Collection   'to store newline positions, 0-based
Private m_colEscPos As Collection       'to store escape chars positions
Private m_colNotesPos As Collection     'to store notes positions (opening and closing)
Private m_colProofTagPos As Collection  'to store position of proof tags (opening and closing)
Private m_colNotesMarkerPos As Collection   'to store position of proof note markers

'Newsroom menu statuses
Private m_bRevisionsShown As Boolean
Private m_bCommandsShown As Boolean
Private m_bNotesShown As Boolean

'constants
'notes opening and closing commands
Private Const NOTESOPENTAG As String = "<NO1>"
Private Const NOTESCLOSETAG As String = "<NO>"

'lowercase, to differentiate notes to be proofed from regular notes
Private Const PROOFNOTESOPENTAG As String = "<no1>"
Private Const PROOFNOTESCLOSETAG As String = "<no>"

'Newsroom menu id's
Private Const MENUID_SHOWREVISIONS As Integer = 10281
Private Const MENUID_COMMANDSSHOW As Integer = 10303
Private Const MENUID_VIEWNOTES As Integer = 10314

Private Sub Class_Initialize()
On Error GoTo EH
    Set m_oConnection = New TS4C.Connection
    m_oConnection.ShowSplash = True
    Call m_oConnection.OpenConnection(Me, "")
    Exit Sub
EH:
    Dim errNum As Long
    Dim errSrc As String
    errNum = Err.Number
    errSrc = Err.Source
    ErrH errNum, Err.Description, errSrc, "Class_Initialize", False  'local err handling
    Err.Raise errNum, ErrorMod.ErrHandler.MakeSource(errSrc, TypeName(Me), "Class_Initialize") 'Tansa err handling
End Sub

Private Sub Class_Terminate()
On Error GoTo EH
    If Not m_oConnection Is Nothing Then
        Call m_oConnection.CloseConnection
    End If
    Exit Sub
EH:
    Dim errNum As Long
    Dim errSrc As String
    errNum = Err.Number
    errSrc = Err.Source
    ErrH errNum, Err.Description, errSrc, "Class_Terminate", False  'local err handling
    Err.Raise errNum, ErrorMod.ErrHandler.MakeSource(errSrc, TypeName(Me), "Class_Terminate") 'Tansa err handling
End Sub

Private Sub IExtension_CheckShortcutKey(ByVal bShift As Boolean, ByVal bCtrl As Boolean, ByVal bAlt As Boolean, ByVal lKeyCode As Long, bUnavailable As Boolean, bInUse As Boolean, sOwner As String)
  'Since we return False in IExtension_IsShortcutKeysSupported,
  'we do not need any code here, since this method never will be called
End Sub

Private Function IExtension_CompareTexts(oTextObject As Object, ByVal sOriginalText As String) As Boolean
  'If the proof or hyphenation dialogs are shown modelessly, the user might alter the text in the
  'document while the Tansa dialog is open. When the Tansa dialog receives focus, it then checks each
  'of the items in its list to see if the corresponding text in the document has been altered.
  'This method is called for each of the items in the list, so the Extension can do the actual text
  'comparison. When the Tansa dialogs are shown modally, as in this example, there is no need to
  'implement any code in this method.
End Function

Private Function IExtension_CountHyphens() As Boolean
  'Since Tansa Text Editor do not support invisible hyphens,
  'the standard, visible hyphen character (-) will be used
  'just for the sake of this example. This character does
  'indeed occupy one character position, thus we return True.
    IExtension_CountHyphens = True
End Function

Private Sub IExtension_DetachTansa()
  'When this method is invoked, Tansa should be completely unloaded from the parent application.
  'The parent application should need to be restarted for the user to be able to use Tansa again.
On Error GoTo EH
    'Call m_oConnection.Disconnect 'Disconnect function removed in Tansa 4
    Set m_oConnection = Nothing
    Exit Sub
EH:
    Dim errNum As Long
    Dim errSrc As String
    errNum = Err.Number
    errSrc = Err.Source
    ErrH errNum, Err.Description, errSrc, "IExtension_DetachTansa", False  'local err handling
    Err.Raise errNum, ErrorMod.ErrHandler.MakeSource(errSrc, TypeName(Me), "IExtension_DetachTansa") 'Tansa err handling
End Sub

Private Sub IExtension_FinishSession(ByVal eTriggerAction As TS4C.TriggerAction, ByVal bUserCancelled As Boolean)
On Error GoTo EH
    If g_bDebug Then WriteLog Format$(Now, TIMESTAMPFORMAT) & " >>>IExtension_FinishSession"

    'clear collections
    Set m_colNewlinePos = Nothing
    Set m_colEscPos = Nothing
    Set m_colNotesPos = Nothing
    Set m_colProofTagPos = Nothing
    Set m_colNotesMarkerPos = Nothing

    'workaround for issue: issue when text within notes are replaced and then the Show/Hide Commands menu function is called,
    '  the space after the corrected text is removed
    '- found that toggling the View Notes command fixes the problem
    oNR.MenuFunc "IDM_VIEWNOTES"
    oNR.MenuFunc "IDM_VIEWNOTES"

    'set menu items to their original state
    If m_bCommandsShown Then
        oNR.MenuFunc "IDM_COMMANDSSHOW"
    Else
        oNR.MenuFunc "IDM_COMMANDSHIDE"
    End If
    If m_bNotesShown Then
        'do nothing
    Else
        oNR.MenuFunc "IDM_VIEWNOTES" 'call it to turn it off since it was turned on programmatically
    End If
    If m_bRevisionsShown Then
        oNR.MenuFunc "IDM_SHOWREVISIONS" 'call it to turn it on since it was turned off programmatically
    Else
        'do nothing
    End If

    'unlock newsroom
    oNR.Unlock
        
    WriteLog Format$(Now, TIMESTAMPFORMAT) & " End Tansa Session"
    Exit Sub
EH:
    Dim errNum As Long
    Dim errSrc As String
    errNum = Err.Number
    errSrc = Err.Source
    ErrH errNum, Err.Description, errSrc, "IExtension_FinishSession", False  'local err handling
    Err.Raise errNum, ErrorMod.ErrHandler.MakeSource(errSrc, TypeName(Me), "IExtension_FinishSession") 'Tansa err handling
End Sub

Private Sub IExtension_GetAutoHyphenationSupport(bDocOpen As Boolean, bLineFeed As Boolean, bInsertText As Boolean, bDocClose As Boolean)
  'None of these events can be caught through Tansa Text Editor's interface, so we leave all the parameters
  'as they are, with False as value
End Sub

Private Function IExtension_GetCompanyName() As String
    IExtension_GetCompanyName = App.CompanyName
End Function

Private Function IExtension_GetDocFullName(oDoc As Object) As String
On Error GoTo EH:
    'get document name from NRCOM
    'TODO: check whether need to parse returned text
    IExtension_GetDocFullName = oNR.GetViewCaption
    Exit Function
EH:
    Dim errNum As Long
    Dim errSrc As String
    errNum = Err.Number
    errSrc = Err.Source
    ErrH errNum, Err.Description, errSrc, "IExtension_GetDocFullName", False  'local err handling
    Err.Raise errNum, ErrorMod.ErrHandler.MakeSource(errSrc, TypeName(Me), "IExtension_GetDocFullName") 'Tansa err handling
End Function

Private Function IExtension_GetExtensionName() As String
    'return extension name
    IExtension_GetExtensionName = App.ProductName
End Function

Private Function IExtension_GetExtensionVersion() As String
    'return version
    IExtension_GetExtensionVersion = App.Major & "." & App.Minor & ".0." & App.Revision
End Function

Private Function IExtension_GetHyphenationDialogShowMode() As TS4C.ShowMode
    'should return modal. Newsroom does not support a dynamic Text Object Model (TOM)
    IExtension_GetHyphenationDialogShowMode = smModal
End Function

Private Function IExtension_GetHyphenCharUnicode() As Long
    IExtension_GetHyphenCharUnicode = g_lSoftHyphenCharCode    'code retrieved from config file
End Function

Private Function IExtension_GetParentAppName() As String
    IExtension_GetParentAppName = PARENTAPPNAME
End Function

Private Function IExtension_GetParentAppUiLanguageCode() As String
    IExtension_GetParentAppUiLanguageCode = g_sParentAppUiLanguageCode
End Function

Private Function IExtension_GetParentAppUserName() As String
    IExtension_GetParentAppUserName = g_sParentAppUserName
End Function

Private Function IExtension_GetParentAppVersion() As String
    IExtension_GetParentAppVersion = g_sParentAppVersion
End Function

Private Function IExtension_GetProofingDialogShowMode() As TS4C.ShowMode
    'should return modal. Newsroom does not support a dynamic Text Object Model (TOM)
    IExtension_GetProofingDialogShowMode = smModal
End Function

Private Function IExtension_GetReminderSupport() As TS4C.SupportedReminderType
    'Tansa Text Editor does neither support Save nor Print, so we can't have reminders for any of these events
    IExtension_GetReminderSupport = srtNone
End Function

Private Sub IExtension_GetSpecialSpaceSupport(oSpecialSpaces As TS4C.GenericCollection)
    'new in Tansa 4
On Error GoTo EH
    If g_bDebug Then WriteLog ">>>IExtension_GetSpecialSpaceSupport"
    
    'Real Unicodes
    '8194  En Space
    '8195  Em Space
    '8196  Three-Per-Em Space
    '8197  Quarter Em Space
    '8198  Six-Per-Em Space
    '8199  Figure Space
    '8200  Punctuation Space
    '8201  Thin Space
    '8202  Hair Space
    '160   Nonbreaking Space
    
    'Substitute Unicodes. These spaces are not part of the Unicode standard
    '32110 Nonbreaking Em Space
    '32111 Nonbreaking En Space
    '32112 Nonbreaking Three-Per-Em Space
    '32113 Nonbreaking Quarter Em Space
    '32114 Nonbreaking Six-Per-Em Space
    '32115 Nonbreaking Thin Space
    '32116 Nonbreaking Hair Space
    '32117 Nonbreaking Punctuation Space
    '32118 Nonbreaking Figure Space

    Dim oSpecialSpace As TS4C.SpecialSpace
  
    'For Each oSpecialSpace In oSpecialSpaces
    '    Select Case oSpecialSpace.Unicode
    '        Case 8196: oSpecialSpace.SubstUnicode = -1
    '        Case 8197: oSpecialSpace.SubstUnicode = -1
    '        Case 8198: oSpecialSpace.SubstUnicode = -1
    '        Case 8200: oSpecialSpace.SubstUnicode = -1
    '        Case 32110: oSpecialSpace.SubstUnicode = 8195   'Nonbreaking Em Space -> Em Space
    '        Case 32111: oSpecialSpace.SubstUnicode = 8194   'Nonbreaking En Space -> En Space
    '        Case 32112: oSpecialSpace.SubstUnicode = -1     'Nonbreaking Three-Per-Em Space -> Not supported
    '        Case 32113: oSpecialSpace.SubstUnicode = -1     'Nonbreaking Quarter Em Space -> Not supported
    '        Case 32114: oSpecialSpace.SubstUnicode = -1     'Nonbreaking Six-Per-Em Space -> Six-Per-Em Space
    '        Case 32115: oSpecialSpace.SubstUnicode = 8201   'Nonbreaking Thin Space -> Thin Space
    '        Case 32116: oSpecialSpace.SubstUnicode = 8202   'Nonbreaking Hair Space -> Hair Space
    '        Case 32117: oSpecialSpace.SubstUnicode = -1     'Nonbreaking Punctuation Space -> Not supported
    '        Case 32118: oSpecialSpace.SubstUnicode = 8199   'Nonbreaking Figure Space -> Figure Space
    '    End Select
    'Next
    
    '4.0.0.1 20110520 jpm: added support for Special Spaces
    For Each oSpecialSpace In oSpecialSpaces
        'disable special spaces not found in the config file
        If Not ExistsInCol(g_colSpecialSpaces, CStr(oSpecialSpace.Unicode)) Then
            oSpecialSpace.SubstUnicode = -1
        End If
        If g_bDebug Then WriteLog "Special space=" & oSpecialSpace.Unicode & ", Sub=" & oSpecialSpace.SubstUnicode
    Next
    
    Exit Sub
EH:
    Dim errNum As Long
    Dim errSrc As String
    errNum = Err.Number
    errSrc = Err.Source
    ErrH errNum, Err.Description, errSrc, "IExtension_GetSpecialSpaceSupport", False  'local err handling
    Err.Raise errNum, ErrorMod.ErrHandler.MakeSource(errSrc, TypeName(Me), "IExtension_GetSpecialSpaceSupport") 'Tansa err handling
End Sub

Private Function IExtension_GetSupportEMailAddress() As String
    IExtension_GetSupportEMailAddress = SUPPORTEMAILADDRESS
End Function

Private Function IExtension_GetTextObject(ByVal lTextNumber As Long, ByVal lStartPos As Long, ByVal lEndPos As Long, ByVal sText As String, ByVal eServiceType As TS4C.TansaService) As Object
'gets text to be corrected
On Error GoTo EH
    If g_bDebug Then WriteLog ">>>IExtension_GetTextObject: Text=" & sText & ", TextNum=" & lTextNumber & ", StartPos=" & lStartPos & ", EndPos=" & lEndPos
    'INFO: can also adjust start pos here instead of in Replacetext?
    
    Dim oTextObject As New TextObject
    
    oTextObject.StartPos = lStartPos
    oTextObject.TypoLength = lEndPos - lStartPos
    
    Set IExtension_GetTextObject = oTextObject
    Exit Function
EH:
    Dim errNum As Long
    Dim errSrc As String
    errNum = Err.Number
    errSrc = Err.Source
    ErrH errNum, Err.Description, errSrc, "IExtension_GetTextObject", False  'local err handling
    Err.Raise errNum, ErrorMod.ErrHandler.MakeSource(errSrc, TypeName(Me), "IExtension_GetTextObject") 'Tansa err handling
End Function

Private Function IExtension_GetTexts(oDoc As Object) As TS4C.GenericCollection
On Error GoTo EH
    If g_bDebug Then WriteLog ">>>IExtension_GetTexts"
    
    Dim oTextCollection As New TS4C.GenericCollection
    Dim i As Integer
    
    m_sTextToProof = ""
    m_sOrigText = ""
    
    
    'Note: use GetContentCom instead of GetContent so that typo commands/notes/tags
    'can be included in the text to be sent for proofing, in case they are shown on screen
    '20150215 use GetContentCom_utf16
    m_sTextToProof = oNR.GetContentCom_utf16
    If g_bDebug Then WriteLog "Text from Nroom: " & vbNewLine & m_sTextToProof
    
    
    'remove invisible commands
    RemoveInvisibleCommands m_sTextToProof
    If g_bDebug Then WriteLog "Text after removing invisible commands: " & vbNewLine & m_sTextToProof
    
    
    m_sOrigText = m_sTextToProof 'make a copy of original text, for later use
    
    
    'Newsroom notes handling
    HandleProofTags m_sTextToProof
    HandleNotes m_sTextToProof
    'remove the note commands from orig text too because they are not visible in Newsroom-not included in pos counts
    m_sOrigText = Replace(m_sOrigText, NOTESOPENTAG, "")
    m_sOrigText = Replace(m_sOrigText, NOTESCLOSETAG, "")
    For i = 1 To m_colNotesMarkerPos.Count
        m_sOrigText = Mid$(m_sOrigText, 1, m_colNotesMarkerPos.Item(i)) & _
            g_sNotesMarker & Mid$(m_sOrigText, m_colNotesMarkerPos.Item(i) + 1)  'note collection contains 0-based values
    Next i
    If g_bDebug Then
        WriteLog "Text after Notes handling: " & vbNewLine & m_sTextToProof
        WriteLog "Original text after Notes handling: " & vbNewLine & m_sOrigText
    End If
    
    'convert typo commands and tags to substitution char so they won't be proofed
    ConvertSpaceCommands m_sTextToProof 'special space commands
    ConvertCommands m_sTextToProof
    If g_bDebug Then WriteLog "Text after converting commands: " & vbNewLine & m_sTextToProof
    

    'get the newline positions and store in a collection
    'Note: for Newsroom, vbnewline is not considered in the counting of char positions
    'while Tansa considers it as one char
    StoreNewlinePos m_sTextToProof
    'also store any escape characters used becauses NRCOM returns them
    'but they are not visible on screen and not included in counting of char positions
    StoreEscCharPos m_sTextToProof


    oTextCollection.Add m_sTextToProof
    Set IExtension_GetTexts = oTextCollection

    Exit Function
EH:
    Dim errNum As Long
    Dim errSrc As String
    errNum = Err.Number
    errSrc = Err.Source
    ErrH errNum, Err.Description, errSrc, "IExtension_GetTexts", False  'local err handling
    Err.Raise errNum, ErrorMod.ErrHandler.MakeSource(errSrc, TypeName(Me), "IExtension_GetTexts") 'Tansa err handling
End Function

Private Function IExtension_GetTradeMarkInformation() As String
    IExtension_GetTradeMarkInformation = App.LegalTrademarks
End Function

Private Function IExtension_IsHyphenationSupported() As Boolean
    IExtension_IsHyphenationSupported = True
End Function

Private Function IExtension_IsModelessAllowed() As Boolean
    'return False
    IExtension_IsModelessAllowed = False
End Function

Private Function IExtension_IsShortcutKeysSupported() As Boolean
    'It's not possible to assign shortcut keys in the extension
    IExtension_IsShortcutKeysSupported = False
End Function

Private Function IExtension_IsTextProtected() As Boolean
    'TODO: determine if object if opened in read/write mode
    IExtension_IsTextProtected = False
End Function

Private Function IExtension_IsTextSelection() As Boolean
    'NRCOM doesn't have fuctionality to get selected text, it can only return all the text
    IExtension_IsTextSelection = False
End Function

Private Function IExtension_IsUnprotectTextPossible() As Boolean
    'new in Tansa 4
End Function

Private Function IExtension_LaunchUrl(ByVal sUrl As String) As Boolean
    'new in Tansa 4
End Function

Private Function IExtension_PrepareSession() As Boolean
On Error GoTo EH
    WriteLog Format$(Now, TIMESTAMPFORMAT) & " Start Tansa Session"
    If g_bDebug Then WriteLog Format$(Now, TIMESTAMPFORMAT) & " >>>IExtension_PrepareSession"
    
    If Not (oNR Is Nothing) Then
        If (oNR.IsTextOpened) Then
            'log document name
            WriteLog "Document: " & oNR.GetViewCaption
        
            'lock Newsroom
            oNR.Lock
            
            'save object
            If g_bSaveBeforeProofing Then oNR.MenuFunc "IDM_SAVETXT"
        
            'get Hermes menu statuses
            m_bCommandsShown = oNR.GetHermesMenuStatus(MENUID_COMMANDSSHOW)
            m_bNotesShown = oNR.GetHermesMenuStatus(MENUID_VIEWNOTES)
            m_bRevisionsShown = oNR.GetHermesMenuStatus(MENUID_SHOWREVISIONS)
            If g_bDebug Then WriteLog "Hermes Menu Status: CommandsShown=" & m_bCommandsShown & _
                ", NotesShown=" & m_bNotesShown & ", RevisionsShown=" & m_bRevisionsShown
        
            'show commands
            '-need to show commands to determine which "command-like" text
            ' are real commands and which are not
            oNR.MenuFunc "IDM_COMMANDSSHOW" 'always call function to show commands
            'show notes. note: menu func is a toggle
            If Not m_bNotesShown Then oNR.MenuFunc "IDM_VIEWNOTES"
            'do not show revisions. note: menu func is a toggle
            If m_bRevisionsShown Then oNR.MenuFunc "IDM_SHOWREVISIONS"
                        
            IExtension_PrepareSession = True
        Else
            MsgBox "No Newsroom object is opened.", vbExclamation
            IExtension_PrepareSession = False
        End If
    Else
        MsgBox "Newsroom is not running.", vbExclamation
        IExtension_PrepareSession = False
    End If
    
    Exit Function
EH:
    Dim errNum As Long
    Dim errSrc As String
    errNum = Err.Number
    errSrc = Err.Source
    ErrH errNum, Err.Description, errSrc, "IExtension_PrepareSession", False  'local err handling
    Err.Raise errNum, ErrorMod.ErrHandler.MakeSource(errSrc, TypeName(Me), "IExtension_PrepareSession") 'Tansa err handling
End Function

Private Sub IExtension_ProcessCustomMenuItemClick(ByVal eCustomMenuItem As TS4C.CustomMenuItem)
    'no custom menu items
End Sub

Private Sub IExtension_RebuildMenus()
  'If this Extension has built any Tansa menus fram scratch, utilizing the GetDoItYourselfKit method on
  'the PopupMenu object, then all these menus should be rebuilt to get fresh values for all properties.
  'If, however, this Extension only uses the menus provided by the ShowMenuAtPointerPosition and/or
  'ShowMenuAtScreenCoordinates methods, then don't put any code here.
End Sub

Private Sub IExtension_RemoveHyphenationPoints()
'some notes: VB - 1 based, Tansa - 0 based, Nrcom - 0 based
On Error GoTo EH
    WriteLog Format$(Now, TIMESTAMPFORMAT) & " Start Remove Hyphenation Points"
    If g_bDebug Then WriteLog ">>>IExtension_RemoveHyphenationPoints"
    
    Dim lHyphenPos As Long
            
        
    'log document name
    WriteLog "Document: " & oNR.GetViewCaption
        
    'lock Newsroom
    oNR.Lock
    'save object
    If g_bSaveBeforeProofing Then oNR.MenuFunc "IDM_SAVETXT"
    
    'get Hermes menu statuses
    m_bCommandsShown = oNR.GetHermesMenuStatus(MENUID_COMMANDSSHOW)
    m_bNotesShown = oNR.GetHermesMenuStatus(MENUID_VIEWNOTES)
    m_bRevisionsShown = oNR.GetHermesMenuStatus(MENUID_SHOWREVISIONS)
    If g_bDebug Then WriteLog "Hermes Menu Status: CommandsShown=" & m_bCommandsShown & _
        ", NotesShown=" & m_bNotesShown & ", RevisionsShown=" & m_bRevisionsShown

    'show commands
    '-need to show commands to determine which "command-like" text
    ' are real commands and which are not
    oNR.MenuFunc "IDM_COMMANDSSHOW" 'always call function to show commands
    'show notes. note: menu func is a toggle
    If Not m_bNotesShown Then oNR.MenuFunc "IDM_VIEWNOTES"
    'do not show revisions. note: menu func is a toggle
    If m_bRevisionsShown Then oNR.MenuFunc "IDM_SHOWREVISIONS"
            
            
    'Note: use GetContentCom instead of GetContent so that typo commands/notes/tags
    'can be included in the text to be sent for proofing, in case they are shown on screen
    '20150215 use GetContentCom_utf16
    m_sTextToProof = oNR.GetContentCom_utf16
    If g_bDebug Then WriteLog "Text from Nroom: " & vbNewLine & m_sTextToProof
    
    'remove invisible commands
    RemoveInvisibleCommands m_sTextToProof
    If g_bDebug Then WriteLog "Text after removing invisible commands: " & vbNewLine & m_sTextToProof
    
    'special handling for Newsroom notes
    HandleProofTags m_sTextToProof
    HandleNotes m_sTextToProof
    If g_bDebug Then WriteLog "Text after Notes handling: " & vbNewLine & m_sTextToProof
    
    
    lHyphenPos = 0
    Do
        lHyphenPos = InStr(lHyphenPos + 1, m_sTextToProof, ChrW(IExtension_GetHyphenCharUnicode))
                
        If lHyphenPos > 0 Then
            'store newline pos, 0-based. NRCOM doesn't count newlines
            StoreNewlinePos m_sTextToProof
            'store escape char pos, 0-based. NRCOM doesn't count esc chars
            StoreEscCharPos m_sTextToProof
                
            m_sTextToProof = Mid$(m_sTextToProof, 1, lHyphenPos - 1) & _
                "" & Mid$(m_sTextToProof, lHyphenPos + 1) 'update buffer
        
            'adjust newline pos
            'pass hyphenpos-1 because VB is 1-based, Tansa and NRCOM are 0-based
            lHyphenPos = AdjustStartPos(lHyphenPos - 1)
            
            '20150215 use Replace_utf16
            oNR.Replace_utf16 lHyphenPos, 1, ""   'Replace with empty string!
        End If
    Loop While lHyphenPos > 0
       
    
    'set menu items to their original state
    If m_bCommandsShown Then
        oNR.MenuFunc "IDM_COMMANDSSHOW"
    Else
        oNR.MenuFunc "IDM_COMMANDSHIDE"
    End If
    If m_bNotesShown Then
        'do nothing
    Else
        oNR.MenuFunc "IDM_VIEWNOTES" 'call it to turn it off since it was turned on programmatically
    End If
    If m_bRevisionsShown Then
        oNR.MenuFunc "IDM_SHOWREVISIONS" 'call it to turn it on since it was turned off programmatically
    Else
        'do nothing
    End If
    
    'unlock newsroom
    oNR.Unlock
    
    
    WriteLog Format$(Now, TIMESTAMPFORMAT) & " End Remove Hyphenation Points"
    Exit Sub
EH:
    Dim errNum As Long
    Dim errSrc As String
    errNum = Err.Number
    errSrc = Err.Source
    ErrH errNum, Err.Description, errSrc, "IExtension_RemoveHyphenationPoints", False  'local err handling
    Err.Raise errNum, ErrorMod.ErrHandler.MakeSource(errSrc, TypeName(Me), "IExtension_RemoveHyphenationPoints") 'Tansa err handling
End Sub

Private Sub IExtension_ReplaceText(oTextObject As Object, ByVal sNewText As String, ByVal oNewTextWordCollection As TS4C.GenericCollection)
'Select the correct piece of text (the error) and replace it with the value of sNewText (the correction)
'some notes: VB - 1 based, Tansa - 0 based, Nrcom - 0 based
On Error GoTo EH
    If g_bDebug Then WriteLog ">>>IExtension_ReplaceText: " & sNewText
  
    Dim lStartPos As Long
    Dim lTypoLen As Long
    Dim sOldText As String
    Dim bHandled As Boolean

    lStartPos = oTextObject.StartPos
    lTypoLen = oTextObject.TypoLength

    If g_bDebug Then WriteLog "Info from Tansa: StartPos=" & lStartPos & ", TypoLen=" & lTypoLen

    'escape any special characters
    If g_bEscapeCommandChars Then sNewText = PutEscChar(sNewText)

    sOldText = Mid$(m_sOrigText, lStartPos + 1, lTypoLen)
    
    bHandled = False
    If sNewText = " " Then
        'special handling for space correction
        sNewText = HandleSpaceReplacements(sOldText, sNewText, bHandled)
    End If
                        
    If Not bHandled Then
        'text corrections
                
        '4.0.0.1 20110520 jpm: added support for Special Spaces
        sNewText = HandleSpecialSpaceReplacements(sNewText)
        
        'handle Newsroom notes caught in between corrections
        sNewText = PushEmbeddedNotes(sOldText, sNewText, lStartPos, lTypoLen)
        
        'append any embedded tags in orig text to the correction
        sNewText = PushEmbeddedCommands(sOldText, sNewText)
    End If


    If g_bApplyNotesCommandsInCorrections Then  '4.0.0.5 20150402 jpm: adjust if setting is on
        'check if correction is part of a Newsroom note
        'if it is part of note, put in the Note commands with the correction
        sNewText = AdjustIfPartOfNote(sNewText, lStartPos)
    End If
    

    'adjust the start pos since hermes doesn't count some chars
    lStartPos = AdjustStartPos(lStartPos)
    
    
    'record modification
    WriteLog "Replace '" & sOldText & "' with '" & sNewText & _
            "', StartPos=" & lStartPos & ", TypoLen=" & lTypoLen

    'this line replaces the wrong text immediately in NRoom
    '20150215 use Replace_utf16
    oNR.Replace_utf16 lStartPos, lTypoLen, sNewText

    Exit Sub
EH:
    Dim errNum As Long
    Dim errSrc As String
    errNum = Err.Number
    errSrc = Err.Source
    ErrH errNum, Err.Description, errSrc, "IExtension_ReplaceText", False  'local err handling
    Err.Raise errNum, ErrorMod.ErrHandler.MakeSource(errSrc, TypeName(Me), "IExtension_ReplaceText") 'Tansa err handling
End Sub

Private Sub IExtension_ReprotectText()
    'do nothing: no need to reprotect text
End Sub

Private Sub IExtension_RestoreCursor()
    'do nothing: NRCOM is not able to record cursor position
End Sub

Private Sub IExtension_RestoreSelection(ByVal lEndPosModifier As Long)
    'do nothing: NRCOM cannot restore cursor position
End Sub

Private Sub IExtension_SelectText(oTextObject As Object)
    'highlight text in Newsroom while browsing through Tansa dialog
On Error GoTo EH
    Dim lStartPos As Long
    Dim lTypoLen As Long
    Dim sOldText As String

    lStartPos = oTextObject.StartPos
    lTypoLen = oTextObject.TypoLength
    
    'adjust the start pos since hermes doesn't count newlines
    lStartPos = AdjustStartPos(lStartPos)
    
    oNR.Highlight lStartPos, lTypoLen
    Exit Sub
EH:
    Dim errNum As Long
    Dim errSrc As String
    errNum = Err.Number
    errSrc = Err.Source
    ErrH errNum, Err.Description, errSrc, "IExtension_SelectText", False  'local err handling
    Err.Raise errNum, ErrorMod.ErrHandler.MakeSource(errSrc, TypeName(Me), "IExtension_SelectText") 'Tansa err handling
End Sub

Private Sub IExtension_SetAutoHyphenation(ByVal bDocOpen As Boolean, ByVal bLineFeed As Boolean, ByVal bInsertText As Boolean, ByVal bDocClose As Boolean, ByVal lProfileID As Long)
  'Since all parameters in IExtension_GetAutoHyphenationSupport was
  'returned with value False this method will never be invoked
End Sub

Private Sub IExtension_SetStatusText(ByVal sText As String, ByVal lProgressPercent As Long)
  'Since we return False in IExtension_IsModelessAllowed, this method will be invoked
  'do nothing: nothing to show
End Sub

Private Sub IExtension_StoreCursor()
    'do nothing
    'NRCOM cannot record current cursor position
End Sub

Private Sub IExtension_StoreSelection()
    'do nothing
    'NRCOM cannot record current selected text
End Sub

Private Function IExtension_UnprotectText() As Boolean
    'TODO: might need to determine if object was opened in write mode first
    IExtension_UnprotectText = True
End Function

'######################## Custom Procedures ##############################

Private Sub StoreNewlinePos(ByVal sText As String)
On Error GoTo EH
    If g_bDebug Then WriteLog ">>>StoreNewlinePos"
    
    Dim i As Long
    Dim lNewlinePosVB As Long
    Dim lNewlinePosTansa As Long

    Set m_colNewlinePos = New Collection

    i = 0
    lNewlinePosVB = 1
    lNewlinePosTansa = 0

    While (lNewlinePosVB > 0)
        lNewlinePosVB = InStr(sText, vbNewLine)
        
        If lNewlinePosVB > 0 Then 'found newline
            lNewlinePosTansa = lNewlinePosVB - 1 'VB is 1-based; Tansa is 0-based
            
            If m_colNewlinePos.Count > 0 Then
                lNewlinePosTansa = lNewlinePosTansa + m_colNewlinePos(m_colNewlinePos.Count) + 2
            End If
            If g_bDebug Then WriteLog "Store: NewlinePosVB=" & lNewlinePosVB & _
                    " NewlinePosTansa=" & lNewlinePosTansa

            m_colNewlinePos.Add lNewlinePosTansa 'store 0-based position
                      
            i = lNewlinePosVB + 2 'cr + lf
            sText = Mid$(sText, i)
        End If
    Wend
    Exit Sub
EH:
    Err.Raise Err.Number, "StoreNewlinePos:" & Err.Source, Err.Description
End Sub

Private Sub StoreEscCharPos(ByVal sText As String)
On Error GoTo EH
    If g_bDebug Then WriteLog ">>>StoreEscCharPos"
    
    Dim lSearchPos As Long
    Dim lEscCharPos As Long
    Dim sNextChar As String

    Set m_colEscPos = New Collection

    lSearchPos = 1

    Do
        lEscCharPos = InStr(lSearchPos, sText, "\")
        If lEscCharPos > 0 Then 'found escape char
            
            sNextChar = Mid$(sText, lEscCharPos + 1, 1) 'get char at next pos
            
            Select Case sNextChar
                Case "<", ">", "[", "]", "{", "}", "\"
                    'char at next pos is a special char that was 'escaped' by preceding \
                    m_colEscPos.Add (lEscCharPos - 1)   'store 0-based pos (VB is 1-based)
                    
                    If g_bDebug Then WriteLog "Store: EscCharPos=" & (lEscCharPos - 1)
                    
                    lSearchPos = lEscCharPos + 2
                Case Else
                    'no escaped character
                    lSearchPos = lEscCharPos + 1
            End Select
        End If
    Loop While lEscCharPos > 0
    
    Exit Sub
EH:
    Err.Raise Err.Number, "StoreEscCharPos:" & Err.Source, Err.Description
End Sub

Private Function AdjustStartPos(ByVal lStartPos As Long) As Long
On Error GoTo EH
    If g_bDebug Then WriteLog ">>>AdjustStartPos"

    Dim lPos As Long
    Dim lNewlineCtr As Long
    Dim lEscCtr As Long
    Dim lNotesMarkerCtr As Long
    Dim lAdjustedStartPos As Long
    Dim i As Long
    Dim s As String
    
    lNewlineCtr = 0
    lEscCtr = 0
    lNotesMarkerCtr = 0

    'newlines, note: collection contains 0-based pos
    For i = 1 To m_colNewlinePos.Count
        lPos = m_colNewlinePos.Item(i)
        If lStartPos >= lPos Then lNewlineCtr = lNewlineCtr + 1
    Next
    
    'escape chars, note: collection contains 0-based pos
    For i = 1 To m_colEscPos.Count
        lPos = m_colEscPos.Item(i)
        If lStartPos >= lPos Then lEscCtr = lEscCtr + 1
    Next

    'proof note marker, note: collection contains 0-based pos
    For i = 1 To m_colNotesMarkerPos.Count
        lPos = m_colNotesMarkerPos.Item(i)
        If lStartPos >= lPos Then lNotesMarkerCtr = lNotesMarkerCtr + 1
    Next

    'deduct number of newlines because Nrcom doesn't count newlines
    'deduct escape chars because Nrcom doesn't count them too
    lAdjustedStartPos = lStartPos - lNewlineCtr - lEscCtr - (lNotesMarkerCtr * Len(g_sNotesMarker))
    
    If g_bDebug Then WriteLog "Adjust: OrigStartPos=" & lStartPos & _
            " DeductNewlines=" & lNewlineCtr & _
            " DeductEscapes=" & lEscCtr & _
            " DeductNoteMarks=" & (lNotesMarkerCtr * Len(g_sNotesMarker)) & _
            " AdjustedStartPos=" & lAdjustedStartPos
    
    AdjustStartPos = lAdjustedStartPos
    Exit Function
EH:
    Err.Raise Err.Number, "AdjustStartPos:" & Err.Source, Err.Description
End Function

Private Sub RemoveInvisibleCommands(ByRef sText As String)
'Newsroom puts commands that will never be visible on screen
'but are included in the text retrieved through NRCOM
'these are:
'-<US> commands e.g. <USHEADLINE>, <USCAPTION>
'-carton delimiters {} e.g. {TEXT}, {HEADLINE}, {CAPTION}
'-<WC1> and <WC> commands from wirecenter
'they need to be removed so they will not be included in pos counting
On Error GoTo EH
    If g_bDebug Then WriteLog ">>>RemoveInvisibleCommands"
                
    Dim sSearchPat As String
    Dim node As IXMLDOMNode
    Dim attr As IXMLDOMNode
    
    For Each node In g_xdomInvisibleCommands
        sSearchPat = Trim$(node.Text)
        
        'this attribute indicates whether the command has parameters
        Set attr = node.Attributes.getNamedItem("withParam")
        If Trim$(attr.nodeValue) = "1" Then 'command has params, put wildcards in search pattern
            sSearchPat = RegExReplacePattern(sSearchPat, "([\>\]}])$", ".*?$1", True, True, True)
        End If
        
        'regex special chars
        sSearchPat = Replace(sSearchPat, "<", "\<")
        sSearchPat = Replace(sSearchPat, ">", "\>")
        sSearchPat = Replace(sSearchPat, "[", "\[")
        sSearchPat = Replace(sSearchPat, "]", "\[")
        
        'make sure command is not escaped
        sSearchPat = "([^\\]?)" & sSearchPat
        
        If g_bDebug Then WriteLog "Remove matches of: " & sSearchPat
        sText = RegExReplacePattern(sText, sSearchPat, "$1", True, True, True)
    Next
                
    Exit Sub
EH:
    Err.Raise Err.Number, "RemoveInvisibleCommands:" & Err.Source, Err.Description
End Sub

Private Sub ConvertSpaceCommands(ByRef sText As String)
'convert space commands
On Error GoTo EH
    If g_bDebug Then WriteLog ">>>ConvertSpaceCommands"
    
    Dim sSearchPat As String
    Dim sReplacement As String
    Dim lSubCharCode As Long
    Dim bHasSubChar As Boolean
    Dim node As IXMLDOMNode
    Dim attrCode As IXMLDOMNode
    Dim attrParam As IXMLDOMNode
    
    For Each node In g_xdomSpaceCommands
        sSearchPat = Trim$(node.Text)
        sReplacement = sSearchPat
        bHasSubChar = False 'init to false
        
        Set attrCode = node.Attributes.getNamedItem("subCharCode")
        If Not (attrCode Is Nothing) Then
            If IsNumeric(Trim$(attrCode.nodeValue)) Then
                lSubCharCode = CLng(Val(Trim$(attrCode.nodeValue)))
                bHasSubChar = True
            End If
        End If
        
        If bHasSubChar Then
            '4.0.0.1 20110520 jpm: added support for Special Spaces
            'space commands are transformed to specific unicode chars. the mapping list is in the config file
            'all the chars of the command converted to the substitution char, except for one which is converted to the special unicode char
            'e.g. "<EM"> becomes "?N??" where N is ChrW(8195)
            sReplacement = ChrW(g_lCommandSubCharCode) & ChrW(lSubCharCode) & String(Len(sReplacement) - 2, ChrW(g_lCommandSubCharCode))
        Else
            'replacement would be a space,
            'and all the chars of the command converted to the substitution char, except for one
            'typo commands are atleast 4 chars - put space at second char
            'e.g. "<EM>" becomes "? ??"
            sReplacement = ChrW(g_lCommandSubCharCode) & " " & String(Len(sReplacement) - 2, ChrW(g_lCommandSubCharCode))
        End If
        
        'this attribute indicates whether the command has parameters
        Set attrParam = node.Attributes.getNamedItem("withParam")
        If Trim$(attrParam.nodeValue) = "1" Then 'command has params, put wildcards in search pattern
            sSearchPat = RegExReplacePattern(sSearchPat, "\>$", ".*?\>", True, True, True)
        End If
        
        'regex special chars
        sSearchPat = Replace(sSearchPat, "<", "\<")
        sSearchPat = Replace(sSearchPat, ">", "\>")
        
        'make sure command is not escaped
        sSearchPat = "([^\\]?)" & sSearchPat
        
        If g_bDebug Then WriteLog "Convert matches of: " & sSearchPat & " to: " & sReplacement
        sText = RegExReplacePattern(sText, sSearchPat, "$1" & sReplacement, True, True, True)
    Next
    
    Exit Sub
EH:
    Err.Raise Err.Number, "ConvertSpaceCommands:" & Err.Source, Err.Description
End Sub

Private Sub ConvertCommands(ByRef sText As String)
'convert typographical commands and tags to substitution char
On Error GoTo EH
    If g_bDebug Then WriteLog ">>>ConvertCommands"

    Dim colMatches As MatchCollection
    Dim oMatch As Match
    Dim lFoundPos As Long
    Dim bConvert As Boolean
    
    Set colMatches = RegExGetMatches(sText, "\<.*?\>|\[.*?\]", True, True, True)
    
    If Not (colMatches Is Nothing) Then
        'note:
        'indexes/positions in colMatches is 0-based
        'VB string functions are 1-based
        For Each oMatch In colMatches
            lFoundPos = oMatch.FirstIndex + 1
            
            If lFoundPos = 1 Then
                bConvert = True 'command is at first pos-it's not escaped
            Else
                bConvert = IIf(Mid$(sText, lFoundPos - 1, 1) <> "\", True, False) 'check if command is escaped or not
            End If
            
            If bConvert Then 'not escaped - it's a real command - convert now
                sText = Mid$(sText, 1, lFoundPos - 1) & _
                    String(oMatch.length, ChrW(g_lCommandSubCharCode)) & _
                    Mid$(sText, lFoundPos + oMatch.length)
                If g_bDebug Then WriteLog "Convert: " & oMatch.Value & _
                    ", Pos=" & lFoundPos & ", Len=" & oMatch.length
            End If
        Next
    Else
        If g_bDebug Then WriteLog "No command found. Nothing converted."
    End If
    
    Exit Sub
EH:
    Err.Raise Err.Number, "ConvertCommands:" & Err.Source, Err.Description
End Sub

Private Function PushEmbeddedCommands(sOldText As String, sNewText As String) As String
'if there are any tags/commands within the text to be replaced, append them to the correction
On Error GoTo EH
    If g_bDebug Then WriteLog ">>>PushEmbeddedCommands: OrigText=" & sOldText
    
    Dim colMatches As MatchCollection
    Dim oMatch As Match
    Dim sCommands As String
    Dim lFoundPos As Long
    Dim bPush As Boolean
    
    sCommands = ""
    Set colMatches = RegExGetMatches(sOldText, "\<.*?\>|\[.*?\]", True, True, True)
    
    If Not (colMatches Is Nothing) Then
        'note:
        'indexes/positions in colMatches is 0-based
        'VB string functions are 1-based
        For Each oMatch In colMatches
            lFoundPos = oMatch.FirstIndex + 1
            
            If lFoundPos = 1 Then
                bPush = True 'command is at first pos-it's not escaped
            Else
                bPush = IIf(Mid$(sOldText, lFoundPos - 1, 1) <> "\", True, False) 'check if command is escaped or not
            End If
            
            If bPush Then sCommands = sCommands & oMatch.Value 'not escaped - it's a real command
        Next
    
        sNewText = sNewText & sCommands 'append embedded commands at the end of new text
        If g_bDebug Then WriteLog "Found embedded commands: " & sCommands & _
            ", Updated correction: " & sNewText
    Else
        If g_bDebug Then WriteLog "No embedded commands found. Nothing changed."
    End If
    
    PushEmbeddedCommands = sNewText
    Exit Function
EH:
    Err.Raise Err.Number, "PushEmbeddedCommands:" & Err.Source, Err.Description
End Function

Private Function PushEmbeddedNotes(sOldText As String, sNewText As String, _
    lStartPos As Long, lTypoLen As Long) As String
'if there are any notes caught in between the corrections, push them to the end of the correction
On Error GoTo EH
    If g_bDebug Then WriteLog ">>>PushEmbeddedNotes: OrigText=" & sOldText & _
        ", StartPos=" & lStartPos & ", TypoLen=" & lTypoLen
        
    Dim i As Integer
    Dim lEndPos As Long
    Dim sNotesPos() As String
    Dim lNoteStartPos As Long
    Dim lNoteEndPos As Long
    Dim sNote As String
    Dim sEmbeddedNotes As String
    
    sEmbeddedNotes = ""
    lEndPos = lStartPos + lTypoLen - 1

    'check if the note pos is within proof tags
    For i = 1 To m_colNotesPos.Count
        sNotesPos = Split(m_colNotesPos(i), "|")
        lNoteStartPos = Val(sNotesPos(0))
        lNoteEndPos = Val(sNotesPos(1))
                
        'info:
        'collection contains note positions in VB (1-based)
        'start position returned is from Tansa (0-based)
        If lNoteStartPos > (lStartPos + 1) And lNoteEndPos < (lEndPos + 1) Then
            'found a note that was caught within the correction
            sNote = Mid$(m_sOrigText, lNoteStartPos, (lNoteEndPos - lNoteStartPos + 1))
            If g_bDebug Then WriteLog "Found embedded note: " & sNote
            'record embedded note and wrap it with notes commands
            sEmbeddedNotes = sEmbeddedNotes & NOTESOPENTAG & sNote & NOTESCLOSETAG
        End If
    Next i
        
    If Trim$(sEmbeddedNotes) <> "" Then
        sNewText = sNewText & sEmbeddedNotes
        If g_bDebug Then WriteLog "Updated correction with embedded notes: " & sNewText
    Else
        If g_bDebug Then WriteLog "No embedded notes found. Nothing changed."
    End If
    
    PushEmbeddedNotes = sNewText
    Exit Function
EH:
    Err.Raise Err.Number, "PushEmbeddedNotes:" & Err.Source, Err.Description
End Function

Private Function HandleSpaceReplacements(sOldText As String, sNewText As String, _
    ByRef bSpaceHandled As Boolean) As String
On Error GoTo EH
    If g_bDebug Then WriteLog ">>>HandleSpaceReplacements"
    
    'note: space commands always have 4 chars. e.g. <EM>, <TH>, <EN>
    'no need to handle commands that are not 4 chars at this point
        
    bSpaceHandled = False
        
    If Len(sOldText) = 3 And RegExTest(sOldText, " <[a-zA-Z]{1}", True, True, True) Then
        'there's a space command following the space
        'e.g. hi <em>hello
        '---> hi ? ??hello
        'sOldText = " <e", set sNewText = "<e"
        'result: hi<em>hello
        sNewText = Mid$(sOldText, 2)
        bSpaceHandled = True
        If g_bDebug Then WriteLog "Space command after a space. New text changed to: " & sNewText
    ElseIf Len(sOldText) = 4 And RegExTest(sOldText, "[a-zA-Z]{2}> ", True, True, True) Then
        'there's a space command before the space
        'e.g. hi<em> hello
        '---> hi? ?? hello
        'sOldText = "em> ", set sNewText = "em>"
        'result: hi<em>hello
        sNewText = Mid$(sOldText, 1, 3)
        bSpaceHandled = True
        If g_bDebug Then WriteLog "Space command before a space. New text changed to: " & sNewText
    ElseIf Len(sOldText) = 5 And RegExTest(sOldText, "[a-zA-Z]{2}><[a-zA-Z]{1}", True, True, True) Then
        'there are 2 space commands
        'e.g. <th><em>
        '---> ? ??? ??
        'sOldText = "th><e", set sNewText = "e" - just retain 2nd space command
        'result: <em>
        sNewText = Mid$(sOldText, 5)
        bSpaceHandled = True
        If g_bDebug Then WriteLog "Two space commands found. New text changed to: " & sNewText
    End If

    HandleSpaceReplacements = sNewText
    Exit Function
EH:
    Err.Raise Err.Number, "HandleSpaceReplacements:" & Err.Source, Err.Description
End Function

Private Function HandleSpecialSpaceReplacements(sNewText As String) As String
On Error GoTo EH
    If g_bDebug Then WriteLog ">>>HandleSpecialSpaceReplacements"
    
    Dim sNewsroomCommand As String
    Dim lSubCharCode As Long
    Dim bHasSubChar As Boolean
    Dim node As IXMLDOMNode
    Dim attrCode As IXMLDOMNode
    
    For Each node In g_xdomSpaceCommands
        sNewsroomCommand = Trim$(node.Text)
        bHasSubChar = False 'init to false
        
        Set attrCode = node.Attributes.getNamedItem("subCharCode")
        If Not (attrCode Is Nothing) Then
            If IsNumeric(Trim$(attrCode.nodeValue)) Then
                lSubCharCode = CLng(Val(Trim$(attrCode.nodeValue)))
                bHasSubChar = True
            End If
        End If
        
        If bHasSubChar Then
            If InStr(sNewText, ChrW(lSubCharCode)) Then
                'sub char used in the replacement
                'need to convert these to the corresponding Newsroom commands
                sNewText = Replace(sNewText, ChrW(lSubCharCode), sNewsroomCommand)
                If g_bDebug Then WriteLog "Convert found matches of: ChrW(" & lSubCharCode & ") to: " & sNewsroomCommand
            End If
        End If
    Next

    HandleSpecialSpaceReplacements = sNewText
    Exit Function
EH:
    Err.Raise Err.Number, "HandleSpecialSpaceReplacements:" & Err.Source, Err.Description
End Function

Private Sub HandleProofTags(sText As String)
On Error GoTo EH
    If g_bDebug Then WriteLog ">>>HandleProofTags"

    Set m_colProofTagPos = New Collection   'to store proof tags pos

    Dim tag As IXMLDOMNode
    For Each tag In g_xdomCheckTags
        GetProofTagsPos sText, UCase(Trim$(tag.Text))
    Next
    Exit Sub
EH:
    Err.Raise Err.Number, "HandleProofTags:" & Err.Source, Err.Description
End Sub

Private Sub GetProofTagsPos(sText As String, sProofTag As String)
On Error GoTo EH
    If g_bDebug Then WriteLog ">>>GetProofTagsPos: ProofTag=" & sProofTag

    Dim lOpenPos As Long
    Dim lClosePos As Long
    Dim lCloseEndPos As Long
    Dim lSearchPos As Long
    Dim sOpenTag As String
    Dim sCloseTag As String
    Dim sTaggedText As String
    
    sOpenTag = "[" & sProofTag & "]"
    sCloseTag = "[/" & sProofTag & "]"
    lSearchPos = 1 'pos where search starts

    Do
        'search for opening tag
        lOpenPos = InStr(lSearchPos, sText, sOpenTag)

        If lOpenPos > 0 Then
            'search til find closing tag
            lClosePos = InStr(lOpenPos + Len(sOpenTag), sText, sCloseTag)

            If lClosePos > 0 Then 'found closing tag
                lCloseEndPos = lClosePos + Len(sCloseTag) - 1
            Else 'no closing tag - all chars following the opening tag are part of note
                lCloseEndPos = Len(sText)
            End If
    
            'store proof tag pos
            'e.g. [WEBHED]text[/WEBHED] - OpenPos=1, CloseEndPos=21
            m_colProofTagPos.Add (lOpenPos & "|" & lCloseEndPos)
    
            If g_bDebug Then
                sTaggedText = Mid$(sText, lOpenPos, (lCloseEndPos - lOpenPos + 1)) 'content including tag
                WriteLog "Found tagged text: " & sTaggedText & _
                    ", Store: OpenPos=" & lOpenPos & ", CloseEndPos=" & lCloseEndPos
            End If
            
            lSearchPos = lCloseEndPos 'update start pos of search
        End If
    Loop While lOpenPos > 0
    
    Exit Sub
EH:
    Err.Raise Err.Number, "GetProofTagsPos:" & Err.Source, Err.Description
End Sub

Private Sub HandleNotes(ByRef sText As String)
On Error GoTo EH
    If g_bDebug Then WriteLog ">>>HandleNotes"

    Dim lSearchPos As Long
    Dim lOpenPos As Long
    Dim lClosePos As Long
    Dim lCloseEndPos As Long
    Dim sNote As String
    Dim iNoteOpenTagCount As Integer
    Dim iNoteCloseTagCount As Integer
    Dim lNoteStartPos As Long
    Dim lNoteEndPos As Long
    Dim bProofNote As Boolean
    Dim sTempOpen As String
    Dim sTempClose As String

    Set m_colNotesPos = New Collection  'collection for notes pos
    Set m_colNotesMarkerPos = New Collection 'collection for proof notes marker
    
    lSearchPos = 1 'pos where search starts
    iNoteOpenTagCount = 0
    iNoteCloseTagCount = 0

    Do
        lOpenPos = InStr(lSearchPos, sText, NOTESOPENTAG)

        If lOpenPos > 0 Then
            iNoteOpenTagCount = iNoteOpenTagCount + 1 'increment counter
            lClosePos = InStr(lOpenPos + Len(NOTESOPENTAG), sText, NOTESCLOSETAG) 'search for closing

            If lClosePos > 0 Then
                iNoteCloseTagCount = iNoteCloseTagCount + 1 'increment counter
                lCloseEndPos = lClosePos + Len(NOTESCLOSETAG) - 1
                sNote = Mid$(sText, lOpenPos + Len(NOTESOPENTAG), (lClosePos - lOpenPos) - Len(NOTESOPENTAG))
            Else
                lCloseEndPos = Len(sText)
                sNote = Mid$(sText, lOpenPos + Len(NOTESOPENTAG))
            End If
                        
            If g_bDebug Then WriteLog "Found note: " & sNote & _
                ", TagOpenPos=" & lOpenPos & ", TagCloseEndPos=" & lCloseEndPos
            
            bProofNote = IsNoteWithinProofTags(lOpenPos, lCloseEndPos)
            If bProofNote Then
                'to differentiate from regular notes, use lowercase commands
                sTempOpen = PROOFNOTESOPENTAG
                sTempClose = PROOFNOTESCLOSETAG
            Else
                'it's a regular note - replace all chars in the note to the substitution char,
                'except escaped chars, the esc char and newline
                sNote = SubstituteNoteChars(sNote)
                If g_bDebug Then WriteLog "Converted note: " & sNote
                
                'regular notes commands
                sTempOpen = NOTESOPENTAG
                sTempClose = NOTESCLOSETAG
            End If
                                
            'at this point, the opening and closing note tags are still included in the text
            'so that pos can be compared with proof tags pos
            sText = Mid$(sText, 1, lOpenPos - 1) & _
                sTempOpen & sNote & IIf(lClosePos > 0, sTempClose, "") & _
                Mid$(sText, IIf(lClosePos > 0, lClosePos + Len(NOTESCLOSETAG), Len(sText) + 1))
            
            'store what newsroom thinks are the opening and closing pos of the note
            'newsroom does not count the opening and closing note tags
            'e.g. <NO1>note text<NO> - NoteStartPos=1, NoteEndPos=9
            lNoteStartPos = (lOpenPos + Len(NOTESOPENTAG)) - _
                (iNoteOpenTagCount * Len(NOTESOPENTAG)) - _
                IIf(iNoteCloseTagCount > 0, ((iNoteCloseTagCount - IIf(lClosePos > 0, 1, 0)) * Len(NOTESCLOSETAG)), 0)
            lNoteStartPos = lNoteStartPos + (m_colNotesMarkerPos.Count * Len(g_sNotesMarker))
            lNoteEndPos = lNoteStartPos + Len(sNote) - 1
            
            If bProofNote Then  'need to adjust if the note is to be proofed - for the note marker pos
                lNoteStartPos = lNoteStartPos + 1
                lNoteEndPos = lNoteEndPos + 1
            
                m_colNotesMarkerPos.Add (lNoteStartPos - 1 - 1) 'store 0-based pos
                If lClosePos > 0 Then m_colNotesMarkerPos.Add lNoteEndPos 'store 0-based pos
            End If
            
            'store in collection - used later in determining whether correction is part of note or not
            m_colNotesPos.Add (lNoteStartPos & "|" & lNoteEndPos)
            If g_bDebug Then WriteLog "Store: NoteStartPos=" & lNoteStartPos & ", NoteEndPos=" & lNoteEndPos
                
            lSearchPos = IIf(lClosePos > 0, lClosePos + Len(NOTESCLOSETAG), Len(sText) + 1)
        End If
    Loop While lOpenPos > 0
    
    'finally remove all instances of the opening and closing notes command
    'since they are not included by newsroom in pos counting
    sText = Replace(sText, NOTESOPENTAG, "", , , vbBinaryCompare)
    sText = Replace(sText, NOTESCLOSETAG, "", , , vbBinaryCompare)
    
    'replace notes commands of those notes to be proofed with the marker to separate them from regular text
    'to avoid: [webhed]<NO1>good<NO>[/webhed]becomes bad -> "goodbecomes bad"
    sText = Replace(sText, PROOFNOTESOPENTAG, g_sNotesMarker, , , vbBinaryCompare)
    sText = Replace(sText, PROOFNOTESCLOSETAG, g_sNotesMarker, , , vbBinaryCompare)
    
    Exit Sub
EH:
    Err.Raise Err.Number, "HandleNotes:" & Err.Source, Err.Description
End Sub

Private Function IsNoteWithinProofTags(lOpenPos As Long, lCloseEndPos As Long) As Boolean
On Error GoTo EH
    If g_bDebug Then WriteLog ">>>IsNoteWithinProofTags"
    
    Dim i As Integer
    Dim sProofTagPos() As String
    Dim bReturn As Boolean
    
    bReturn = False 'init to false

    'check if the note pos is within proof tags
    For i = 1 To m_colProofTagPos.Count
        sProofTagPos = Split(m_colProofTagPos(i), "|")
        
        If lOpenPos > Val(sProofTagPos(0)) And lCloseEndPos <= Val(sProofTagPos(1)) Then
            bReturn = True 'note is within proof tags
            Exit For
        End If
    Next i
    
    If g_bDebug Then WriteLog (IIf(bReturn, "Note is inside proof tags", "Note is not inside proof tags"))
    IsNoteWithinProofTags = bReturn
    Exit Function
EH:
    Err.Raise Err.Number, "IsNoteWithinProofTags:" & Err.Source, Err.Description
End Function

Private Function AdjustIfPartOfNote(sNewText As String, lStartPos As Long) As String
On Error GoTo EH
    If g_bDebug Then WriteLog ">>>AdjustIfPartOfNote: StartPos=" & lStartPos
    
    Dim i As Integer
    Dim sNotePos() As String
    
    'info:
    'collection contains note positions in VB (1-based)
    'start position returned is from Tansa (0-based)
    
    'check if the text is within a note
    For i = 1 To m_colNotesPos.Count
        sNotePos = Split(m_colNotesPos(i), "|")
        
        If (lStartPos + 1) >= Val(sNotePos(0)) And (lStartPos + 1) <= Val(sNotePos(1)) Then
            'correction is part of note - wrap with Newsroom Notes commands
            sNewText = NOTESOPENTAG & sNewText & NOTESCLOSETAG
            If g_bDebug Then WriteLog "Correction is part of note. Converted text=" & sNewText
            Exit For
        End If
    Next i

    AdjustIfPartOfNote = sNewText
    Exit Function
EH:
    Err.Raise Err.Number, "AdjustIfPartOfNote:" & Err.Source, Err.Description
End Function

Private Function PutEscChar(sNewText As String) As String
'put esc char before special chars in replacements
On Error GoTo EH
    If g_bDebug Then WriteLog ">>>PutEscChar: CommandChars=" & g_sCommandChars
    
    sNewText = RegExReplacePattern(sNewText, "([" & g_sCommandChars & "])", "\$1", True, True, True)
    If g_bDebug Then WriteLog "Escaped Text=" & sNewText
    
    PutEscChar = sNewText
    Exit Function
EH:
    Err.Raise Err.Number, "PutEscChar:" & Err.Source, Err.Description
End Function

Private Function SubstituteNoteChars(sNote) As String
'substitute note characters
On Error GoTo EH
    If g_bDebug Then WriteLog ">>>SubstituteNoteChars"

    Dim sCurC As String
    Dim sNextC As String
    Dim sNewNote As String
    Dim bSkipNext As Boolean
    Dim i As Long

    sNewNote = ""

    For i = 1 To Len(sNote)
        If Not bSkipNext Then
            sCurC = Mid$(sNote, i, 1)
            If sCurC = "\" Then 'do not substitute esc char and escaped char; their pos need to be recorded
                sNextC = Mid$(sNote, i + 1, 1)
                Select Case sNextC
                    Case "<", ">", "[", "]", "{", "}", "\"
                        sNewNote = sNewNote & sCurC & sNextC
                        bSkipNext = True
                    Case Else
                        sNewNote = sNewNote & ChrW(g_lCommandSubCharCode) 'substitute
                End Select
            Else
                Select Case sCurC
                    Case vbNewLine, vbCrLf, vbCr, vbLf 'do not substitute newline; its pos needs to be recorded
                        sNewNote = sNewNote & sCurC
                    Case Else
                        sNewNote = sNewNote & ChrW(g_lCommandSubCharCode) 'substitute
                End Select
            End If
        Else
            bSkipNext = False
        End If
    Next i

    SubstituteNoteChars = sNewNote
    Exit Function
EH:
    Err.Raise Err.Number, "SubstituteNoteChars:" & Err.Source, Err.Description
End Function

