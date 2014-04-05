VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDebug 
   Caption         =   "Linker Command Line Debugger"
   ClientHeight    =   10800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16725
   LinkTopic       =   "Form1"
   ScaleHeight     =   10800
   ScaleWidth      =   16725
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkRunVisible 
      Caption         =   "Run Visible"
      Height          =   285
      Left            =   8865
      TabIndex        =   15
      Top             =   5670
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   10905
      Left            =   10980
      TabIndex        =   5
      Top             =   -90
      Width           =   5325
      Begin VB.CommandButton cmdSaveDef 
         Caption         =   "Save Changes"
         Height          =   330
         Left            =   3330
         TabIndex        =   13
         Top             =   5040
         Width           =   1725
      End
      Begin VB.TextBox txtDef 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4515
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   12
         Top             =   405
         Width           =   4875
      End
      Begin VB.TextBox txtVBC 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3705
         Left            =   270
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   9
         Top             =   6525
         Width           =   4875
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   420
         Left            =   1575
         TabIndex        =   6
         Top             =   10350
         Width           =   3570
         Begin VB.CommandButton cmdAbort 
            Caption         =   "Abort"
            Height          =   330
            Left            =   405
            TabIndex        =   8
            Top             =   45
            Width           =   1410
         End
         Begin VB.CommandButton cmdContinue 
            Caption         =   "Continue"
            Height          =   330
            Left            =   2070
            TabIndex        =   7
            Top             =   45
            Width           =   1365
         End
      End
      Begin VB.Label Label4 
         Caption         =   "Def File:"
         Height          =   240
         Left            =   225
         TabIndex        =   14
         Top             =   180
         Width           =   780
      End
      Begin VB.Label Label3 
         Caption         =   "VBC File: "
         Height          =   285
         Left            =   270
         TabIndex        =   11
         Top             =   6165
         Width           =   690
      End
      Begin VB.Label lblVBCFile 
         Caption         =   "Label4"
         Height          =   285
         Left            =   1035
         TabIndex        =   10
         Top             =   6165
         Width           =   3840
      End
   End
   Begin RichTextLib.RichTextBox txtIn 
      Height          =   5280
      Left            =   90
      TabIndex        =   2
      Top             =   315
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   9313
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmDebug.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtOut 
      Height          =   4650
      Left            =   45
      TabIndex        =   3
      Top             =   6030
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   8202
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmDebug.frx":007C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblNoChanges 
      Caption         =   "[ NO CHANGES FOUND ]"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   5220
      TabIndex        =   4
      Top             =   5760
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.Label Label2 
      Caption         =   "OutGoing Command Line (can be edited)"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   45
      TabIndex        =   1
      Top             =   5805
      Width           =   5235
   End
   Begin VB.Label Label1 
      Caption         =   "Incoming Command Line"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   135
      TabIndex        =   0
      Top             =   45
      Width           =   4785
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim abort As Boolean

Function DebugCommandLine() As String
        
    txtIn = AddNewLines(orgCmdLine)
    txtOut = AddNewLines(vbCommand)
    
    If FileExists(defFile) Then
        txtDef = ReadFile(defFile)
    Else
        txtDef = "[ No Def File found ]"
        cmdSaveDef.Enabled = False
    End If
    
    If orgCmdLine = vbCommand Then
        lblNoChanges.Visible = True
    Else
        difflines
    End If
    
    lblVBCFile.Caption = cmdFile
    
    If FileExists(cmdFile) Then
        txtVBC = ReadFile(cmdFile)
    Else
        txtVBC = "[ No VBC File Found ]"
    End If
    
    Me.Show 1
    
    If abort Then
        DebugCommandLine = Empty
    Else
        DebugCommandLine = Replace(txtOut.Text, vbCrLf, Empty)
    End If
    Unload Me
    
End Function

Sub difflines()
    
    Dim a() As String
    Dim b() As String
    Dim curb As Long, i As Long
    Dim changed As Boolean
    
    a = Split(txtIn.Text, vbCrLf)
    b = Split(txtOut.Text, vbCrLf)

    For i = 0 To UBound(b)
        
        changed = Not Exists(b(i), a)
        
        If changed Then
            txtOut.SelStart = curb
            txtOut.SelLength = Len(b(i))
            txtOut.SelColor = vbRed
        End If
        
        curb = curb + Len(b(i)) + 2
        
    Next
    
    'scroll back to top and deselect
    txtOut.SelStart = 1
    txtOut.SelLength = 0
    
End Sub

Function Exists(txt, a() As String) As Boolean
    Dim x
    For Each x In a
        If Trim(txt) = Trim(x) Then
            Exists = True
            Exit Function
        End If
    Next
End Function

Function AddNewLines(ByVal x As String) As String
    
    x = Replace(x, ".OBJ""", ".OBJ""" & vbCrLf, , , vbTextCompare)
    x = Replace(x, "/", vbCrLf & "/")
    AddNewLines = x
    
End Function

Private Sub chkRunVisible_Click()
   RunVisible = chkRunVisible.value
End Sub

Private Sub cmdAbort_Click()
    abort = True
    Me.Visible = False
End Sub

Private Sub cmdContinue_Click()
    Me.Visible = False 'breaks modal show
End Sub

 
Private Sub cmdSaveDef_Click()
    writeFile defFile, txtDef.Text
End Sub

Private Sub Command1_Click()
    On Error Resume Next
    Dim cmd As String
    cmd = Replace(txtOut.Text, vbCrLf, Empty)
    cmd = "cmd /k """ & VB6FOLDER & "vbLink.exe " & cmd
    Shell cmd
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Frame2.Left = Me.Width - Frame2.Width - 200
    txtIn.Width = Frame2.Left - 200
    txtOut.Width = txtIn.Width
End Sub
