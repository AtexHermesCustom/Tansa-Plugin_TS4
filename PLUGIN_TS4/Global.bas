Attribute VB_Name = "Global"
Option Explicit

'global NewsRoom extension object
Public oNR As Variant

'global Tansa integration object
Public oTansaPlugin As EdClientIntegration

'constants
Public Const PARENTAPPNAME As String = "Newsroom"
Public Const CONFIGFILE As String = "NewsroomTansaPlugin.xml"
Public Const INIFILE As String = "TS4NE.ini"
Public Const SUPPORTEMAILADDRESS As String = "support@tansa.no"

'global variables for configuration info
Public g_sParentAppBin As String
Public g_sParentAppVersion As String
Public g_sParentAppUserName As String
Public g_sParentAppUiLanguageCode As String
Public g_lSoftHyphenCharCode As Long 'code for Newsroom soft hyphe
Public g_lCommandSubCharCode As Long 'code for char to be used to substitute typo commands/tags. Tansa will not proof this char
Public g_bEscapeCommandChars As Boolean 'whether the plug-in would escape start/end command chars entered by the user
Public g_sCommandChars As String 'in regular expression format
Public g_sNotesMarker As String 'char to separate notes to be proofed from regular text content
Public g_bApplyNotesCommandsInCorrections As Boolean 'whether to wrap corrections in notice mode with Notes commands <NO1> and <NO>
Public g_bSaveBeforeProofing As Boolean 'whether to save object first before proofing or not
Public g_bDebug As Boolean
Public g_sLogPath As String
Public g_iLogRetentionDays As Integer

Public g_xdomCheckTags As IXMLDOMNodeList 'list of tags to be used for proofing notes content
Public g_xdomInvisibleCommands As IXMLDOMNodeList 'list of Newsroom invisible commands
Public g_xdomSpaceCommands As IXMLDOMNodeList 'list of Newsroom commands that can represent space
Public g_xdomNotesCommands As IXMLDOMNodeList 'list of Newsroom commands used for marking notes

Public g_sDefaultNotesOpenTag As String 'default notes open command/tag
Public g_sDefaultNotesCloseTag As String 'default notes close command/tag

Public g_colSpecialSpaces As Collection 'collection of special spaces, with the unicode substitution char as the key

Public Function ExistsInCol(oCol As Collection, sKey As String) As Boolean
On Error GoTo EH:
    Dim vVal As Variant
    vVal = oCol(sKey)
    ExistsInCol = True
    Exit Function
EH:
    ExistsInCol = False
End Function

