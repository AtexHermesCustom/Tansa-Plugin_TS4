Attribute VB_Name = "WinMessaging"
Option Explicit

'***********************************************************************
' This module subclasses the client form. This means that every time
' windows sends a message to the form, it gets passed through one of
' our functions. We can then deal with the messages as appropriate.
'***********************************************************************

'-----------------------------------------------------------------------
' API CALLS
'These are used to subclass a control.
'-----------------------------------------------------------------------
'Changes the address for the windows procedure, and returns the original value
Public Declare Function SetWindowLong Lib "user32" _
    Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) _
    As Long

'Passes the message information on to the specified windows procedure
Private Declare Function CallWindowProc Lib "user32" _
    Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, _
    ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) _
    As Long

'Registers a message as windows, returning a unique ID that can be used
'to identify the message
Private Declare Function RegisterWindowMessage Lib "user32" _
    Alias "RegisterWindowMessageA" (ByVal lpString As String) _
    As Long
    
'-----------------------------------------------------------------------

'-----------------------------------------------------------------------
' VARIABLES
'-----------------------------------------------------------------------
'This const tells the SetWindowLong to change the address of the message
'handler
Const GWL_WNDPROC = -4
'This const is used by SetWindowLong to set a form to be always on top
Public Const SWW_HPARENT = -8

'This holds the memory location of the original windows handler
Private ProcPrev As Long

'This is the name of the message we're going to send
Private Const MSG_MYMESSAGE As String = "MSG_TSNE_MESSAGE"

'This holds the title (caption property) of the client form. This is used
'when it comes to send a message to we can identify the program
Public Const MSG_CLIENTTITLE As String = "Newsroom Tansa Plugin"

'message codes that plugin might receive from Newsroom
Public Const MSG_SHOWTANSAMENU As Integer = 10
Public Const MSG_RUNTANSAPROOFING As Integer = 20
Public Const MSG_RUNTANSAHYPHENATION As Integer = 30
Public Const MSG_CLOSEPLUGIN As Integer = 90
'-----------------------------------------------------------------------

'-----------------------------------------------------------------------
' WINDOWS HOOK/UNHOOK
Public Function HookForm(hwnd As Long)
On Error GoTo EH
    'This basically says that instead of using the normal windows handler
    'to deal with messages, use our WndProc function instead. It stores
    'the original memory location of the windows handler in ProcPrev
    ProcPrev = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WndProc)
    Exit Function
EH:
    Err.Raise Err.Number, "HookForm:" & Err.Source, Err.Description
End Function

Public Function UnhookForm(hwnd As Long)
On Error GoTo EH
    'Unhooks the form. Instead of going via our function, go back to
    'using your original windows handler
    SetWindowLong hwnd, GWL_WNDPROC, ProcPrev
    Exit Function
EH:
    Err.Raise Err.Number, "UnhookForm:" & Err.Source, Err.Description
End Function
'-----------------------------------------------------------------------

'-----------------------------------------------------------------------
' WINDOWS PROCEDURE
' Our own windows procedure. When HookForm is called, all messages are
' passed through this procedure.
'-----------------------------------------------------------------------
Private Function WndProc(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error GoTo EH
    'Depending on the message that is being sent, we decide how to deal with it...
    Select Case msg
        'Is it the message sent from the server form?
        Case WM_MYMESSAGE
            'If so, then call the function on the main form
            frmMain.GotMessage wParam, lParam
                     
        Case Else
            'All other messages we don't care about, so we want to send
            'it on to it's normal location (which is stored in ProcPrev)
            WndProc = CallWindowProc(ProcPrev, hwnd, msg, wParam, lParam)
    End Select
    Exit Function
EH:
    Err.Raise Err.Number, "WndProc:" & Err.Source, Err.Description
End Function

'This function returns the msg id of the windows message
Public Function WM_MYMESSAGE() As Long
On Error GoTo EH
    'Static variable that holds the unique id of our message
    Static msg As Long
    
    'If this is the first time we're running the function,
    'register the message. Results the unique ID of the registerd
    'message
    If msg = 0 Then
        msg = RegisterWindowMessage(MSG_MYMESSAGE)
    End If
    
    'Return the result
    WM_MYMESSAGE = msg
    Exit Function
EH:
    Err.Raise Err.Number, "WM_MYMESSAGE:" & Err.Source, Err.Description
End Function
