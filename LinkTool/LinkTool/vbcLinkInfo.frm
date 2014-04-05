VERSION 5.00
Begin VB.Form frmLinkInfo 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "[MVBLC] Mathimagics VB Link Controller"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   5370
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   4185
      TabIndex        =   1
      Top             =   45
      Width           =   1095
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00F4FFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2355
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   4095
   End
End
Attribute VB_Name = "frmLinkInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'===========================================================
' mathimagics@yahoo.co.uk
'===========================================================
' MVBLC Link Control Tool:   Form "vbcLinkInfo"
'===========================================================
'
' This FORM is only used when displaying link error reports,
' or when showing the STATUS report.
'===========================================================

Sub ShowStatus(vbCommand As String)
   Dim j%, token$()
   j = InStr(vbCommand, "/STATUS:")
   EXEFILE = Mid(vbCommand, j + 8)
   j = InStrRev(EXEFILE, "\")
   EXENAME = Mid$(EXEFILE, j + 1)
   
   Show
   LoadDLL
   With frmLinkInfo.List1
      .AddItem ""
      .AddItem "  Export List: " & EXENAME
      token = Split(GetExports(), vbLf)
      For j = 0 To UBound(token): .AddItem token(j): Next
      .AddItem Format(Now, "  HH:MM:SS  DD MMM YY")
      .ListIndex = .ListCount - 1: DoEvents
      .ListIndex = -1
      End With
   UnMapAndLoad LoadImage
   End Sub

Sub ShowError(vbCommand As String)
   Dim F%, j%, temp$, fLine$
   j = InStr(vbCommand, "/STATUS:")
   EXEFILE = Mid(vbCommand, j + 8)
   j = InStrRev(EXEFILE, "\")
   EXENAME = Mid$(EXEFILE, j + 1)
   
   Me.Width = 13635
   
   Show
   List1.AddItem ""
   List1.AddItem "An unexpected link error has occurred"
   List1.AddItem EXENAME & " link failed"
   List1.AddItem ""
   F = FreeFile
   On Error GoTo BadSign
   Open "c:\vbLink.log" For Input As #F
   Do
      Line Input #F, fLine
      j = InStr(fLine, "error")
      If j Then
         fLine = Mid$(fLine, j)
         j = InStr(fLine, """")
         If j Then
            temp = Left$(fLine, j - 1)
            j = InStr(fLine, """ (")
            If j Then fLine = temp & Mid$(fLine, j + 2)
            End If
         List1.AddItem "> " & fLine
         End If
      Loop Until EOF(F)
   Close #F
   Exit Sub

BadSign:
   List1.AddItem "The log file is not available"
   End Sub
   
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Resize()
   On Error Resume Next
   Command1.Top = Me.Height - Command1.Height - 500
   List1.Move 0, 0, Me.ScaleWidth, Command1.Top - 200
End Sub
