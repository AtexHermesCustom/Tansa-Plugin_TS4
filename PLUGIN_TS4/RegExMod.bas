Attribute VB_Name = "RegExMod"
Option Explicit

'Note:
'RegExp is 0-based

Public Function RegExTest(sStr As String, sSearchPat As String, _
    bIgnoreCase As Boolean, bGlobal As Boolean, bMultiLine As Boolean) As Boolean
On Error GoTo EH
    Dim oRegEx As RegExp
    
    Set oRegEx = New RegExp 'Create a regular expression object
    With oRegEx
        .IgnoreCase = bIgnoreCase
        .Global = bGlobal
        .MultiLine = bMultiLine
        .Pattern = sSearchPat
    End With
    
    RegExTest = oRegEx.Test(sStr)
    Exit Function
EH:
    Err.Raise Err.Number, "RegExTest:" & Err.Source, Err.Description
End Function

Public Function RegExGetMatches(sStr As String, sSearchPat As String, _
    bIgnoreCase As Boolean, bGlobal As Boolean, bMultiLine As Boolean) As MatchCollection
On Error GoTo EH
    Dim oRegEx As RegExp
    Dim colMatches As MatchCollection
    
    Set oRegEx = New RegExp 'Create a regular expression object
    With oRegEx
        .IgnoreCase = bIgnoreCase
        .Global = bGlobal
        .MultiLine = bMultiLine
        .Pattern = sSearchPat
    End With
    
    If (oRegEx.Test(sStr)) Then
        Set colMatches = oRegEx.Execute(sStr) 'execute search and save to collection
    Else
        Set colMatches = Nothing
    End If
    
    Set RegExGetMatches = colMatches
    Exit Function
EH:
    Err.Raise Err.Number, "RegExGetMatches:" & Err.Source, Err.Description
End Function

Public Function RegExReplacePattern(sStr As String, sSearchPat As String, sReplacePat As String, _
    bIgnoreCase As Boolean, bGlobal As Boolean, bMultiLine As Boolean) As String
On Error GoTo EH
    Dim oRegEx As RegExp
    
    Set oRegEx = New RegExp 'Create a regular expression object
    With oRegEx
        .IgnoreCase = bIgnoreCase
        .Global = bGlobal
        .MultiLine = bMultiLine
        .Pattern = sSearchPat
    End With
    
    If (oRegEx.Test(sStr)) Then
        sStr = oRegEx.Replace(sStr, sReplacePat)
    Else
        'do nothing: nothing changed in string
    End If
    
    RegExReplacePattern = sStr
    Exit Function
EH:
    Err.Raise Err.Number, "RegExReplacePattern:" & Err.Source, Err.Description
End Function
