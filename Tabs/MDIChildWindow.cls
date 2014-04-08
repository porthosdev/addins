VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MDIChildWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_SubClass As cSubclass
Private m_hWnd As Long
Private m_lMgrPtr As Long

Implements ISubclass

Property Get hwnd() As Long
    hwnd = m_hWnd
End Property

Sub Attach(ByVal hwnd As Long, _
           ByVal lPtr As Long)
    Detach
    m_lMgrPtr = lPtr
    Set m_SubClass = New cSubclass

    With m_SubClass
        .AddMsg WM_NCACTIVATE, MSG_AFTER
        .AddMsg WM_SETTEXT, MSG_AFTER
        #If Not IN_ADDIN Then
            .AddMsg WM_STYLECHANGED, MSG_AFTER
        #End If
        .Subclass hwnd, Me
    End With

    m_hWnd = hwnd
End Sub

Sub Detach()
    On Error Resume Next

    If Not m_SubClass Is Nothing Then
        m_SubClass.UnSubclass
        Set m_SubClass = Nothing
    End If

    m_lMgrPtr = 0
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Detach
End Sub

Private Sub ISubclass_Proc(ByVal bBefore As Boolean, _
                           bHandled As Boolean, _
                           lReturn As Long, _
                           hwnd As Long, _
                           uMsg As Long, _
                           wParam As Long, _
                           lParam As Long)
    Dim mgr As MDIMonitor
    
    Select Case uMsg

        Case WM_NCACTIVATE

            If wParam Then
                Set mgr = GetVbaManager

                If Not mgr Is Nothing Then
                    mgr.Notify m_hWnd, ACTION_ACTIVATE
                End If
            End If
    
        Case WM_SETTEXT, WM_STYLECHANGED
            Set mgr = GetVbaManager

            If Not mgr Is Nothing Then
                mgr.Notify m_hWnd, ACTION_SETTEXT
            End If

    End Select

End Sub

Private Function GetVbaManager() As MDIMonitor
    Dim oTemp As MDIMonitor

    ' Turn the pointer into an illegal, uncounted interface
    If m_lMgrPtr = 0 Then Exit Function
    CopyMemory oTemp, m_lMgrPtr, 4
    ' Do NOT hit the End button here! You will crash!
    ' Assign to legal reference
    Set GetVbaManager = oTemp
    ' Still do NOT hit the End button here! You will still crash!
    ' Destroy the illegal reference
    CopyMemory oTemp, 0&, 4
End Function

