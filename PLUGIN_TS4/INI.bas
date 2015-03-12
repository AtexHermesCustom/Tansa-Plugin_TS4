Attribute VB_Name = "INI"
Option Explicit

Declare Function GetPrivateProfileString Lib "kernel32" _
        Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
        ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString _
        As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'***************************************************************************************
' Function: GetIniStr
' DESCRIPTION: This function reads settings from an INI file.
' Input:    sIniFile - full path of INI file
'           sSection - section location
'           sKey - key being sought
' Output:   returns the string value from INI file.  If section/key does not exist,
'           will return a blank string
'***************************************************************************************

Public Function GetIniStr(sIniFile As String, sSection As String, sKey As String)
On Error GoTo ErrH
    Dim sTemp As String * 256
    Dim iLen As Integer
    
    iLen = GetPrivateProfileString(sSection, sKey, "", sTemp, Len(sTemp), sIniFile)
    GetIniStr = Left(sTemp, iLen)
Exit Function
ErrH:
    Err.Raise Err.Number, "GetIniStr:" & Err.Source, Err.Description
End Function
