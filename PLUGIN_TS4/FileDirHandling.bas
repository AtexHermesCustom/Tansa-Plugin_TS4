Attribute VB_Name = "FileDirHandling"
Option Explicit

Public Function FileExists(FileName As String) As Boolean
    On Error GoTo EH
    FileExists = (GetAttr(FileName) And vbDirectory) = 0
EH:
End Function

Function FolderExists(FolderName As String) As Boolean
    On Error GoTo EH
    FolderExists = GetAttr(FolderName) And vbDirectory
EH:
End Function
