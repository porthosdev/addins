VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_Monitor As MDIMonitor

Private Sub MDIForm_Load()
    Set m_Monitor = New MDIMonitor
    
    m_Monitor.StartMonitor False, Me.hwnd
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    m_Monitor.EndMonitor
End Sub

Private Sub mnuFileNew_Click()
    Dim a As New Form1
    a.Show
End Sub
