VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_lStartPos As Long
Private m_lTypoLength As Long

Public Property Get StartPos() As Long
    StartPos = m_lStartPos
End Property

Public Property Let StartPos(ByVal lNewValue As Long)
    m_lStartPos = lNewValue
End Property

Public Property Get TypoLength() As Long
    TypoLength = m_lTypoLength
End Property

Public Property Let TypoLength(ByVal lNewValue As Long)
    m_lTypoLength = lNewValue
End Property



