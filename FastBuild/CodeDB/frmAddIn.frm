VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmAddIn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Code Database Addin -dzzie@yahoo.com"
   ClientHeight    =   6735
   ClientLeft      =   2175
   ClientTop       =   2220
   ClientWidth     =   13095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   13095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraAdd 
      Caption         =   " Add New Code "
      Height          =   4875
      Left            =   4140
      TabIndex        =   7
      Top             =   630
      Visible         =   0   'False
      Width           =   8700
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1620
         TabIndex        =   11
         Top             =   270
         Width           =   6570
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3570
         Left            =   1575
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   10
         Top             =   780
         Width           =   6660
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Extract"
         Height          =   255
         Index           =   0
         Left            =   6615
         TabIndex        =   9
         Top             =   4500
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
         Height          =   255
         Index           =   1
         Left            =   7515
         TabIndex        =   8
         Top             =   4500
         Width           =   855
      End
      Begin VB.Label lblClose 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   8325
         TabIndex        =   14
         Top             =   180
         Width           =   330
      End
      Begin VB.Label Label2 
         Caption         =   "Code Body"
         Height          =   375
         Left            =   225
         TabIndex        =   13
         Top             =   900
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Prototype"
         Height          =   285
         Left            =   270
         TabIndex        =   12
         Top             =   315
         Width           =   1005
      End
   End
   Begin VB.ListBox lstFilter 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   1080
      TabIndex        =   6
      Top             =   1305
      Visible         =   0   'False
      Width           =   2970
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      Height          =   255
      Index           =   5
      Left            =   11730
      TabIndex        =   5
      Top             =   6300
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copy"
      Height          =   255
      Index           =   4
      Left            =   10890
      TabIndex        =   4
      Top             =   6300
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Un-Com"
      Height          =   255
      Index           =   3
      Left            =   9930
      TabIndex        =   3
      Top             =   6300
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Comment"
      Height          =   255
      Index           =   2
      Left            =   9090
      TabIndex        =   2
      Top             =   6300
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   90
      TabIndex        =   1
      Top             =   6255
      Width           =   4365
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6060
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   4365
   End
   Begin RichTextLib.RichTextBox Text2 
      Height          =   6045
      Left            =   4500
      TabIndex        =   15
      Top             =   135
      Width           =   8430
      _ExtentX        =   14870
      _ExtentY        =   10663
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      RightMargin     =   50000
      TextRTF         =   $"frmAddIn.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuStrings 
         Caption         =   "Strings"
      End
      Begin VB.Menu mnuAddCode 
         Caption         =   "Add Code"
      End
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long
Const LB_FINDSTRING = &H18F

#Const IS_ADDIN = False

#If IS_ADDIN Then
    Public VBInstance As VBIDE.VBE
    Public Connect As Connect
#End If

Dim ws As Workspace
Dim db As Database
Dim rs As Recordset

 

Private Sub Form_Load()
    On Error GoTo oops
    
    With List1
        lstFilter.Move .Left, .top, .Width, .Height
        fraAdd.Move .Left, .top, Me.Width - 400, Text2.Height
    End With
    
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\db1.mdb")
    Set rs = db.OpenRecordset("CodeDB", dbOpenDynaset)
    rs.MoveFirst
    List1.Clear
    While Not rs.EOF
        List1.AddItem rs.Fields("NAME").Value & String(80, " ") & "@" & rs.Fields("ID").Value
        rs.MoveNext
    Wend
    
    Exit Sub
oops: MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    rs.Close: db.Close: ws.Close
    Set rs = Nothing: Set db = Nothing: Set ws = Nothing
End Sub

Private Sub Command1_Click(Index As Integer)
    On Error GoTo oops
    Select Case Index
        Case 0: Call extract
        Case 1: Call AddNewCode
        Case 2: Call comment
        Case 3: Call comment(False)
        Case 4: Clipboard.Clear: Clipboard.SetText Text2
        Case 5: Text2.Text = Empty
    End Select
    Exit Sub
oops: MsgBox Err.Description
End Sub

Private Sub AddNewCode()
    If Text3 = Empty Or Text4 = Empty Then MsgBox "Need code or name duh": Exit Sub
    
    q = Chr(34) 'quote
    dq = Chr(34) & Chr(34)
    v = """" & Replace(Text3, q, dq) & """,""" & Replace(Text4, q, dq) & """"
    mysql = "INSERT INTO CodeDB (NAME,CODE) VALUES(" & v & ");"
    db.Execute mysql
    
    Form_Unload 0
    Form_Load
End Sub

Private Sub extract()
    On Error Resume Next
    If Text4 = Empty Then MsgBox "Ughh need function to extract name from!": Exit Sub
    tmp = firstLine(Text4)
    fs = InStrRev(tmp, " ", InStr(tmp, "("))
    tmp = Mid(tmp, fs + 1, Len(tmp))
    If Len(tmp) > 254 Then tmp = Mid(tmp, 1, 254)
    Text3 = tmp
End Sub

Private Sub comment(Optional out As Boolean = True)
    'basic outline of sub from Palidan on pscode
    Dim StartLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long
    VBInstance.ActiveCodePane.GetSelection StartLine, StartColumn, EndLine, EndColumn
    If StartLine = EndLine And StartColumn = EndColumn Then Exit Sub
    For i = StartLine To EndLine
        If i = EndLine And EndColumn = 1 Then Exit For
        l = VBInstance.ActiveCodePane.CodeModule.Lines(i, 1)
        If out Then
            VBInstance.ActiveCodePane.CodeModule.ReplaceLine i, "'" + l
        Else
            VBInstance.ActiveCodePane.CodeModule.ReplaceLine i, Mid(l, 2)
        End If
    Next
    Connect.Hide
End Sub

Private Sub CopyCode()
    Dim rs2 As Recordset
    
    If lstFilter.Visible Then
        txt = lstFilter.List(lstFilter.ListIndex)
    Else
        txt = List1.List(List1.ListIndex)
    End If
    
    cid = Mid(txt, InStrRev(txt, "@") + 1, Len(txt))
    rs.Filter = "ID = " & cid
    Set rs2 = rs.OpenRecordset
    Text2.Text = Text2.Text & vbCrLf & vbCrLf & rs2.Fields("CODE")
    Set rs2 = Nothing
    
End Sub

Function firstLine(it)
    t = Split(it, vbCrLf)
    firstLine = t(0)
End Function

Private Sub lblClose_Click()
    fraAdd.Visible = False
End Sub



Private Sub mnuAddCode_Click()
    fraAdd.Visible = True
End Sub

Private Sub mnuStrings_Click()
    frmLazy.Show
End Sub

 

Private Sub Text2_Change()
    modSyntaxHighlighting.SyntaxHighlight Text2
    Text2.selStart = Len(Text2)
End Sub

Private Sub Text4_Change()
    Text4.selStart = 0
    Text4.selLength = 0
End Sub

Private Sub Text4_DblClick()
    c = Clipboard.GetText
    If c <> Empty Then Text4 = c: Command1_Click 0
End Sub

Private Sub Text1_Change()

    'List1.ListIndex = SendMessage(List1.hwnd, LB_FINDSTRING, -1, ByVal CStr(Text1.Text))
    
    If Len(Text1) = 0 Then
        lstFilter.Visible = False
    Else
        lstFilter.Visible = True
        Dim i As Long
        lstFilter.Clear
        For i = 0 To List1.ListCount - 1
            If InStr(1, List1.List(i), Text1, vbTextCompare) > 0 Then
                lstFilter.AddItem List1.List(i)
            End If
        Next
    End If
    
        
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call CopyCode
End Sub

Private Sub List1_DblClick()
    Call CopyCode
End Sub

Private Sub lstFilter_Click()
    Call CopyCode
End Sub

