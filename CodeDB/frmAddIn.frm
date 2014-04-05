VERSION 5.00
Begin VB.Form frmAddIn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Code Database Addin -dzzie@yahoo.com"
   ClientHeight    =   9030
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   11460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9030
   ScaleWidth      =   11460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Extras"
      Height          =   255
      Left            =   8460
      TabIndex        =   11
      Top             =   4365
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      Height          =   255
      Index           =   5
      Left            =   7320
      TabIndex        =   10
      Top             =   4365
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copy"
      Height          =   255
      Index           =   4
      Left            =   6480
      TabIndex        =   9
      Top             =   4365
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Un-Com"
      Height          =   255
      Index           =   3
      Left            =   5520
      TabIndex        =   8
      Top             =   4365
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Comment"
      Height          =   255
      Index           =   2
      Left            =   4680
      TabIndex        =   7
      Top             =   4365
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   255
      Index           =   1
      Left            =   10575
      TabIndex        =   6
      Top             =   4365
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Extract"
      Height          =   255
      Index           =   0
      Left            =   9675
      TabIndex        =   5
      Top             =   4365
      Width           =   855
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
      Left            =   4635
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   645
      Width           =   6660
   End
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
      Left            =   4680
      TabIndex        =   3
      Top             =   135
      Width           =   6570
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4140
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   4770
      Width           =   11250
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
      Top             =   4320
      Width           =   4455
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
      Height          =   4140
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4365
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long
Const LB_FINDSTRING = &H18F

Public VBInstance As VBIDE.VBE
Public Connect As Connect

Dim ws As Workspace
Dim db As Database
Dim rs As Recordset

Private Sub Command2_Click()
    frmLazy.Show
End Sub

Private Sub Form_Load()
    On Error GoTo oops
    
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
        Case 2: Call Comment
        Case 3: Call Comment(False)
        Case 4: Clipboard.Clear: Clipboard.SetText Text2
        Case 5: Text2 = Empty
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

Private Sub Comment(Optional out As Boolean = True)
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
    txt = List1.List(List1.ListIndex)
    cid = Mid(txt, InStrRev(txt, "@") + 1, Len(txt))
    rs.Filter = "ID = " & cid
    Set rs2 = rs.OpenRecordset
    Text2 = Text2 & vbCrLf & vbCrLf & rs2.Fields("CODE")
    Set rs2 = Nothing
End Sub

Function firstLine(it)
    t = Split(it, vbCrLf)
    firstLine = t(0)
End Function

Private Sub Text2_Change()
    Text2.SelStart = Len(Text2)
End Sub

Private Sub Text4_Change()
    Text4.SelStart = 0
    Text4.SelLength = 0
End Sub

Private Sub Text4_DblClick()
    c = Clipboard.GetText
    If c <> Empty Then Text4 = c: Command1_Click 0
End Sub

Private Sub Text1_Change()
    List1.ListIndex = SendMessage(List1.hwnd, LB_FINDSTRING, -1, ByVal CStr(Text1.Text))
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call CopyCode
End Sub

Private Sub List1_DblClick()
    Call CopyCode
End Sub
