VERSION 5.00
Begin VB.Form frmLazy 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   7890
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkStripSpace 
      Caption         =   "StripSpace"
      Height          =   255
      Left            =   6540
      TabIndex        =   9
      Top             =   5580
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "CR -> :"
      Height          =   435
      Left            =   6480
      TabIndex        =   8
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   2115
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   4980
      Width           =   6255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Swap   a = b - >               b = a"
      Height          =   555
      Left            =   6480
      TabIndex        =   6
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   2115
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   2820
      Width           =   6255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "w/ vbCrLf"
      Height          =   255
      Left            =   6480
      TabIndex        =   4
      Top             =   1260
      Width           =   1155
   End
   Begin VB.CommandButton Command2 
      Caption         =   "MultiLine"
      Height          =   495
      Left            =   6480
      TabIndex        =   3
      Top             =   660
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   2115
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   660
      Width           =   6255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Chr( ) "
      Height          =   315
      Left            =   6480
      TabIndex        =   1
      Top             =   180
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   6195
   End
End
Attribute VB_Name = "frmLazy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim s As String, i, l
    Dim ret As String
    
    s = Text1
    For i = 1 To Len(s)
        l = Mid(s, i, 1)
        ret = ret & "Chr(" & Asc(l) & ") & "
    Next
        
        Text1 = Mid(ret, 1, Len(ret) - 2)
End Sub

Private Sub Command2_Click()
    Dim tmp() As String, ret As String, i
    
    ret = Replace(Text2, """", """""")
    tmp() = Split(ret, vbCrLf)
    
    For i = 0 To UBound(tmp)
        tmp(i) = """" & tmp(i) & """ " & IIf(Check1.Value = 1, "& vbcrlf ", "") & "& _"
    Next
    
    ret = Join(tmp(), vbCrLf)
    
    Text2 = Mid(ret, 1, Len(ret) - 3)
    
End Sub

 

Private Sub Command3_Click()

Dim tmp, i, e

    tmp = Split(Text3, vbCrLf)
    For i = 0 To UBound(tmp)
        If InStr(tmp(i), "=") > 0 Then
            e = Split(tmp(i), "=", 2)
            tmp(i) = Trim(e(1)) & "=" & Trim(e(0))
        End If
    Next
    Text3 = Join(tmp, vbCrLf)
            
End Sub

Private Sub Command4_Click()
    Dim x, i, ret
    
    x = Split(Text4, vbCrLf)
    For i = 0 To UBound(x)
        If Len(Trim(x(i))) > 0 Then
            ret = ret & x(i) & ": "
        Else
            ret = ret & vbCrLf
        End If
    Next
    
    ret = Replace(ret, ": " & vbCrLf, vbCrLf)
    
    If chkStripSpace.Value = 1 Then
        ret = Replace(ret, "  ", "")
    End If
    
    Text4 = ret
End Sub

Private Sub Text1_DblClick()
    Clipboard.Clear
    Clipboard.SetText Text1
End Sub

Private Sub Text2_DblClick()
    Clipboard.Clear
    Clipboard.SetText Text2
End Sub

Private Sub Text3_DblClick()
    Clipboard.Clear
    Clipboard.SetText Text3
End Sub

