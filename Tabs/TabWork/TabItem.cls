VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TabItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_StrText As String
Private m_StrKey As String
Private m_bActive As Boolean
Private m_tRect As RECT

Friend Property Get Rectangle() As RECT
    LSet Rectangle = m_tRect
End Property

Friend Property Let Rectangle(tRect As RECT)
    LSet m_tRect = tRect
End Property

Public Property Get key() As String
    key = m_StrKey
End Property

Public Property Let key(ByVal StrValue As String)
    m_StrKey = StrValue
End Property

Public Property Get Text() As String
    Text = m_StrText
End Property

Public Property Let Text(ByVal StrValue As String)
    m_StrText = StrValue
End Property

Public Property Get Active() As Boolean
    Active = m_bActive
End Property

Public Property Let Active(ByVal bActive As Boolean)

    If m_bActive <> bActive Then
        m_bActive = bActive
    End If

End Property

Friend Function Hittest(x As Long, _
                        y As Long) As Boolean
    Hittest = PtInRect(m_tRect, x, y)
End Function
