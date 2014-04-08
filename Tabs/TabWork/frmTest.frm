VERSION 5.00
Begin VB.Form frmTest 
   BackColor       =   &H80000001&
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next Tab"
      Height          =   525
      Left            =   1980
      TabIndex        =   0
      Top             =   1980
      Width           =   1245
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents tab1 As TabStrip
Attribute tab1.VB_VarHelpID = -1

Private Const TAB_COUNT As Long = 30
Private Sub cmdNext_Click()
    Static i As Long
    'If i < 90 Then i = 90
    i = i + 1&
    If i > TAB_COUNT Then i = 1
    
    tab1.ActivateItem "a" & i
End Sub

Private Sub Form_Load()
  Set tab1 = New TabStrip
  tab1.Create 0, 0, 300, 21
  Dim i As Long
  
  For i = 1 To TAB_COUNT
    tab1.AddItem "a" & i, "Tab " & i
  Next
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tab1.MouseDown Button, X \ Screen.TwipsPerPixelX + tab1.Left, tab1.Top + Y \ Screen.TwipsPerPixelY
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tab1.MouseMove Button, X \ Screen.TwipsPerPixelX + tab1.Left, tab1.Top + Y \ Screen.TwipsPerPixelY
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tab1.MouseUp Button, X \ Screen.TwipsPerPixelX + tab1.Left, tab1.Top + Y \ Screen.TwipsPerPixelY
End Sub

Private Sub Form_Paint()
  tab1.Redraw
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  tab1.Move 0, 20, ScaleWidth \ Screen.TwipsPerPixelX, 22
End Sub


Private Sub tab1_ItemClose(ByVal key As String)
    MsgBox "closed " & key
End Sub

Private Sub tab1_RequestPaint(hDC As Long)
  hDC = Me.hDC
End Sub
