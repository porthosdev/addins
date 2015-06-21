VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAddIn 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "API Add-In For Visual basic"
   ClientHeight    =   7920
   ClientLeft      =   2190
   ClientTop       =   1905
   ClientWidth     =   9930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   9930
   Begin VB.PictureBox picHold 
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   90
      ScaleHeight     =   555
      ScaleWidth      =   9420
      TabIndex        =   9
      Top             =   3870
      Width           =   9420
      Begin VB.PictureBox picHoldBtn 
         Height          =   555
         Left            =   5355
         ScaleHeight     =   495
         ScaleWidth      =   3960
         TabIndex        =   14
         Top             =   0
         Width           =   4020
         Begin VB.CommandButton cmdRemove 
            Caption         =   "&Remove"
            Height          =   375
            Index           =   0
            Left            =   90
            TabIndex        =   17
            Top             =   60
            Width           =   1215
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "Remove All"
            Height          =   375
            Index           =   1
            Left            =   1350
            TabIndex        =   16
            Top             =   60
            Width           =   1215
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   375
            Left            =   2700
            TabIndex        =   15
            Top             =   60
            Width           =   1215
         End
      End
      Begin VB.CheckBox chkPublicPrivate 
         Caption         =   "Constants Public"
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   13
         Top             =   45
         Value           =   1  'Checked
         Width           =   1635
      End
      Begin VB.CheckBox chkPublicPrivate 
         Caption         =   "Types Public"
         Height          =   195
         Index           =   1
         Left            =   1755
         TabIndex        =   12
         Top             =   45
         Value           =   1  'Checked
         Width           =   1320
      End
      Begin VB.CheckBox chkPublicPrivate 
         Caption         =   "Declares Public"
         Height          =   195
         Index           =   2
         Left            =   3195
         TabIndex        =   11
         Top             =   45
         Value           =   1  'Checked
         Width           =   1500
      End
      Begin VB.CheckBox chkPublicPrivate 
         Caption         =   "ALL Public/Private"
         Height          =   195
         Index           =   3
         Left            =   225
         TabIndex        =   10
         Top             =   315
         Value           =   1  'Checked
         Width           =   1860
      End
      Begin VB.Image imgMove2 
         Height          =   255
         Left            =   4800
         MousePointer    =   7  'Size N S
         Picture         =   "frmAddIn.frx":0000
         Stretch         =   -1  'True
         Top             =   255
         Width           =   255
      End
      Begin VB.Image imgMove 
         Height          =   255
         Left            =   4800
         MousePointer    =   7  'Size N S
         Picture         =   "frmAddIn.frx":0442
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   4590
         Y1              =   270
         Y2              =   270
      End
   End
   Begin MSComctlLib.ListView lstFindResult 
      Height          =   3390
      Left            =   90
      TabIndex        =   8
      Top             =   450
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   5980
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   4005
      Top             =   7200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.TextBox txtResult 
      Height          =   2715
      Left            =   90
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Top             =   4455
      Width           =   9420
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   375
      Left            =   8280
      TabIndex        =   5
      Top             =   45
      Width           =   1215
   End
   Begin VB.ComboBox cmbType 
      Height          =   315
      ItemData        =   "frmAddIn.frx":0884
      Left            =   990
      List            =   "frmAddIn.frx":0891
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   90
      Width           =   1455
   End
   Begin VB.TextBox txtFind 
      Height          =   285
      Left            =   2520
      TabIndex        =   2
      Text            =   "*"
      Top             =   90
      Width           =   5685
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6930
      TabIndex        =   1
      Top             =   7245
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Copy && Close"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8280
      TabIndex        =   0
      Top             =   7245
      Width           =   1215
   End
   Begin VB.Image imgOpen 
      Height          =   240
      Left            =   180
      Picture         =   "frmAddIn.frx":08B1
      Top             =   135
      Width           =   240
   End
   Begin VB.Label lblEntriesFound 
      AutoSize        =   -1  'True
      Caption         =   "0 Entries found,  0 Entries Selected"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   135
      TabIndex        =   7
      Top             =   7335
      Width           =   3615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Find:"
      Height          =   195
      Left            =   585
      TabIndex        =   3
      Top             =   135
      Width           =   345
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VBInstance As VBIDE.VBE
Public Connect As Connect

Option Explicit
Option Compare Text

Dim APIFileName As String

Dim Constants As New Collection
Dim Types As New Collection
Dim Declares As New Collection

Dim SConstants As New Collection
Dim STypes As New Collection
Dim SDeclares As New Collection

Dim LastFindType As Integer

Dim PY As Single

Private Sub CancelButton_Click()
    Unload Me
    
    Connect.Hide
End Sub

Private Sub chkPublicPrivate_Click(Index As Integer)
    Static ExitChk As Boolean
    Dim V As Integer
    
    If ExitChk Then Exit Sub
    ExitChk = True
    
    If Index <> 3 Then
        V = chkPublicPrivate(0).Value + chkPublicPrivate(1).Value + chkPublicPrivate(2).Value
        
        If V = 0 Then
            chkPublicPrivate(3).Value = 0
        ElseIf V = 3 Then
            chkPublicPrivate(3).Value = 1
        Else
            chkPublicPrivate(3).Value = 2
        End If
    Else
        chkPublicPrivate(0).Value = chkPublicPrivate(3).Value
        chkPublicPrivate(1).Value = chkPublicPrivate(3).Value
        chkPublicPrivate(2).Value = chkPublicPrivate(3).Value
    End If
    
    ExitChk = False
End Sub

Private Sub cmbType_Change()
    cmdFind.Enabled = cmbType.ListIndex <> -1 And txtFind.Text <> "" And APIFileName <> ""
    UpdateEntries
End Sub

Private Sub cmdAdd_Click()
    Dim K As Long, PConst As String, PTypes As String, PDeclare As String
    Dim AddStr As String, CTypes As New Collection
    
    Me.MousePointer = 11
    
    PConst = IIf(chkPublicPrivate(0).Value = 1, "Public ", "Private ")
    PTypes = IIf(chkPublicPrivate(1).Value = 1, "Public ", "Private ")
    PDeclare = IIf(chkPublicPrivate(2).Value = 1, "Public ", "Private ")
    
    On Error Resume Next
    For K = 1 To lstFindResult.ListItems.Count
        If lstFindResult.ListItems(K).Selected Then
            Select Case LastFindType
            Case 0
                AddStr = PConst & Trim(Constants(lstFindResult.ListItems(K).Text))
                SConstants.Add AddStr, AddStr
            Case 1
                AddStr = Trim(Types(lstFindResult.ListItems(K).Text))
                CTypes.Add AddStr, AddStr
            Case 2
                AddStr = PDeclare & Trim(Declares(lstFindResult.ListItems(K).Text))
                SDeclares.Add AddStr, AddStr
                
                GetTypes CTypes, AddStr
            End Select
        End If
    Next K
    
    If LastFindType > 0 Then AddTypes CTypes, PTypes
    On Error GoTo 0
    
    UpdateTextBox
    UpdateEntries
    
    
    Me.MousePointer = 0
End Sub

Private Sub cmdFind_Click()
    Dim FNum As Integer, StrLine As String
    Dim LI As ListItem, FuncName As String, K As Integer
    
    If Dir(APIFileName) = "" Then APIFileName = ""
    If APIFileName = "" Then Exit Sub
    
    If Dir(APIFileName) = "" Then
        MsgBox "Error, can not open file: " & APIFileName, vbExclamation, "API File Missing"
        APIFileName = ""
        Exit Sub
    End If
    
    FNum = FreeFile
    Open APIFileName For Input Access Read Lock Write As FNum
        Do Until Constants.Count = 0
            Constants.Remove 1
        Loop
        
        Do Until Types.Count = 0
            Types.Remove 1
        Loop
        
        Do Until Declares.Count = 0
            Declares.Remove 1
        Loop
        
        lstFindResult.ListItems.Clear
        lstFindResult.Visible = False
        lstFindResult.Sorted = False
        
        LastFindType = cmbType.ListIndex
        Select Case cmbType.ListIndex
        Case 0
            Do Until lstFindResult.ColumnHeaders.Count = 0
                lstFindResult.ColumnHeaders.Remove 1
            Loop
            
            lstFindResult.ColumnHeaders.Add , , "Constants", lstFindResult.Width - 400
            Do
                Line Input #FNum, StrLine
                
                If (StrLine Like "*Const *=*") And (Not Trim(StrLine) Like "'*") Then
                    If StrLine Like "*Const " & txtFind.Text Then
                        If Trim(StrLine) Like "Private*" Then _
                            StrLine = Trim(Mid(StrLine, InStr(1, StrLine, "Private", vbTextCompare) + 8, Len(StrLine)))
                        
                        If Trim(StrLine) Like "Public*" Then _
                            StrLine = Trim(Mid(StrLine, InStr(1, StrLine, "Public", vbTextCompare) + 7, Len(StrLine)))
                        
                        On Error Resume Next
                        FuncName = Mid(StrLine, 6, Len(StrLine))
                        Constants.Add StrLine, FuncName
                        On Error GoTo 0
                        lstFindResult.ListItems.Add , , FuncName
                    End If
                End If
            Loop Until EOF(FNum)
        Case 1
            Do Until lstFindResult.ColumnHeaders.Count = 0
                lstFindResult.ColumnHeaders.Remove 1
            Loop
            
            lstFindResult.ColumnHeaders.Add , , "Types", lstFindResult.Width - 400
            Do
                Line Input #FNum, StrLine
                StrLine = Trim(StrLine)
                
                If ((StrLine Like "Type *") Or (StrLine Like "Private Type *") Or (StrLine Like "Public Type *")) And _
                        (Not Trim(StrLine) Like "'*") _
                        And (Not Trim(StrLine) Like "*Declare Function*") _
                        And (Not Trim(StrLine) Like "*Declare Sub*") _
                        And (Not Trim(StrLine) Like "*Const*") Then
                        
                    If StrLine Like "*Type " & txtFind.Text Then
                        If Trim(StrLine) Like "Private*" Then _
                            StrLine = Trim(Mid(StrLine, InStr(1, StrLine, "Private", vbTextCompare) + 8, Len(StrLine)))
                        
                        If Trim(StrLine) Like "Public*" Then _
                            StrLine = Trim(Mid(StrLine, InStr(1, StrLine, "Public", vbTextCompare) + 7, Len(StrLine)))
                        
                        FuncName = Trim(Mid(StrLine, 5, Len(StrLine)))
                        lstFindResult.ListItems.Add , , FuncName
                        
                        Do
                            Line Input #FNum, StrLine
                        Loop Until EOF(FNum) Or (StrLine Like "End Type*")
                        
                        On Error Resume Next
                        Types.Add FuncName, FuncName
                        On Error GoTo 0
                    End If
                End If
            Loop Until EOF(FNum)
        Case 2
            Do Until lstFindResult.ColumnHeaders.Count = 0
                lstFindResult.ColumnHeaders.Remove 1
            Loop
            
            lstFindResult.ColumnHeaders.Add , , "Function/Sub Name", 2100
            lstFindResult.ColumnHeaders.Add , , "Declares", lstFindResult.Width - 2500
            
            Do
                Line Input #FNum, StrLine
                
                If (StrLine Like "*Declare *") And (Not Trim(StrLine) Like "'*") Then
                    If Trim(StrLine) Like "Private*" Then _
                        StrLine = Trim(Mid(StrLine, InStr(1, StrLine, "Private", vbTextCompare) + 8, Len(StrLine)))
                    
                    If Trim(StrLine) Like "Public*" Then _
                        StrLine = Trim(Mid(StrLine, InStr(1, StrLine, "Public", vbTextCompare) + 7, Len(StrLine)))
                    
                    StrLine = Trim(Mid(StrLine, 8, Len(StrLine)))
                    
                    If Mid(StrLine, InStr(1, StrLine, " ") + 1, Len(StrLine)) Like txtFind.Text Then
                        K = IIf(LCase(Left(StrLine, 4)) = "sub ", 4, 9)
                        FuncName = Trim(Mid(StrLine, K, InStr(K + 1, StrLine, " ") - K))
                        
                        On Error Resume Next
                        Declares.Add "Declare " & StrLine, FuncName
                        On Error GoTo 0
                        
                        Set LI = lstFindResult.ListItems.Add(, , FuncName)
                        LI.SubItems(1) = StrLine
                    End If
                End If
            Loop Until EOF(FNum)
        End Select
        
        UpdateEntries
        lstFindResult.Visible = True
    Close FNum
End Sub

Private Sub cmdRemove_Click(Index As Integer)
    Dim TText As String, SText As String, Q As Variant
    Dim SS As Long, SL As Long, SR As Long
    
    If Trim(txtResult.Text) = "" Then Exit Sub
    
    TText = txtResult.Text
    
    If Index = 0 Then
        If txtResult.SelLength > 0 Then
            SS = txtResult.SelStart
            SL = txtResult.SelLength
            
            SR = InStr(SS + SL - 2, TText, vbNewLine, vbBinaryCompare)
            SS = InStrRev(TText, vbNewLine, SS + 2, vbBinaryCompare)
        Else
            SS = txtResult.SelStart
            
            SR = InStr(SS - 2, TText, vbNewLine, vbBinaryCompare)
            SS = InStrRev(TText, vbNewLine, SS + 2, vbBinaryCompare)
        End If
        
        txtResult.SelStart = SS
        txtResult.SelLength = SR - SS
        SText = txtResult.SelText
        
        For Each Q In SConstants
            If InStr(1, SText, Q) > 0 Then SConstants.Remove Q
        Next Q
        
        For Each Q In STypes
            If InStr(1, SText, Q) > 0 Then STypes.Remove Q
        Next Q
        
        For Each Q In SDeclares
            If InStr(1, SText, Q) > 0 Then SDeclares.Remove Q
        Next Q
        
        UpdateTextBox
    Else
        Do Until SConstants.Count = 0
            SConstants.Remove 1
        Loop
        
        Do Until STypes.Count = 0
            STypes.Remove 1
        Loop
        
        Do Until SDeclares.Count = 0
            SDeclares.Remove 1
        Loop
        
        txtResult.Text = ""
    End If
    
    UpdateEntries
End Sub

Private Sub Form_Load()
    Dim K
    APIFileName = Trim(GetSetting("API Add-In For VB", "APIFileName", "FileName", ""))
    
    chkPublicPrivate(0).Value = Val(GetSetting("API Add-In For VB", "APIFileName", "ConstantsPublic", "1"))
    chkPublicPrivate(1).Value = Val(GetSetting("API Add-In For VB", "APIFileName", "TypesPublic", "1"))
    chkPublicPrivate(2).Value = Val(GetSetting("API Add-In For VB", "APIFileName", "DeclaresPublic", "1"))
    
    Me.Top = Val(GetSetting("API Add-In For VB", "Pos", "Top", Me.Top))
    Me.Left = Val(GetSetting("API Add-In For VB", "Pos", "Left", Me.Left))
    Me.Width = Val(GetSetting("API Add-In For VB", "Pos", "Width", Me.Width))
    Me.Height = Val(GetSetting("API Add-In For VB", "Pos", "Height", Me.Height))
    picHold.Top = Val(GetSetting("API Add-In For VB", "Pos", "Divider", picHold.Top))
    
    On Error GoTo Err_Exit
    If APIFileName <> "" Then
        If Dir(APIFileName) = "" Then APIFileName = ""
    End If
    
    If APIFileName = "" Then
        CDialog.Filter = "WIN32API File (*.txt)|*.txt"
        CDialog.FileName = "WIN32API.TXT"
        CDialog.ShowOpen
        
        APIFileName = CDialog.FileName
    End If
    
    If APIFileName <> "" Then
        Me.Caption = "API Add-In For Visual Basic - " & Mid(APIFileName, _
            InStrRev(APIFileName, "\", , vbBinaryCompare) + 1, Len(APIFileName))
    Else
        Me.Caption = "API Add-In For Visual Basic"
    End If
    
    cmbType.ListIndex = 2
    Connect.Show
    
    Exit Sub
Err_Exit:
    Err.Clear
    
    For Each K In Me.Controls
        If TypeOf K Is CommandButton Then
            If K.Caption <> "&Cancel" Then K.Enabled = False
        End If
    Next K
    
    Connect.Show
    Exit Sub
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug.Print "MouseMove"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    txtResult.Text = ""
    txtFind.Text = "*"
    lstFindResult.ListItems.Clear
    
    Do Until Constants.Count = 0
        Constants.Remove 1
    Loop
    
    Do Until Types.Count = 0
        Types.Remove 1
    Loop
    
    Do Until Declares.Count = 0
        Declares.Remove 1
    Loop
    
    
    Do Until SConstants.Count = 0
        SConstants.Remove 1
    Loop
    
    Do Until STypes.Count = 0
        STypes.Remove 1
    Loop
    
    Do Until SDeclares.Count = 0
        SDeclares.Remove 1
    Loop
    
    SaveSetting "API Add-In For VB", "Pos", "Top", Me.Top
    SaveSetting "API Add-In For VB", "Pos", "Left", Me.Left
    SaveSetting "API Add-In For VB", "Pos", "Width", Me.Width
    SaveSetting "API Add-In For VB", "Pos", "Height", Me.Height
    SaveSetting "API Add-In For VB", "Pos", "Divider", picHold.Top
    
    SaveSetting "API Add-In For VB", "APIFileName", "FileName", APIFileName
    
    SaveSetting "API Add-In For VB", "APIFileName", "ConstantsPublic", chkPublicPrivate(0).Value
    SaveSetting "API Add-In For VB", "APIFileName", "TypesPublic", chkPublicPrivate(1).Value
    SaveSetting "API Add-In For VB", "APIFileName", "DeclaresPublic", chkPublicPrivate(2).Value
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    cmdFind.Left = ScaleWidth - cmdFind.Width - 50
    txtFind.Width = cmdFind.Left - txtFind.Left - 50
    lstFindResult.Width = ScaleWidth - 100
    txtResult.Width = ScaleWidth - 100
    picHold.Width = ScaleWidth - 100
    picHoldBtn.Left = picHold.ScaleWidth - picHoldBtn.Width
    
    OKButton.Move ScaleWidth - OKButton.Width - 50, ScaleHeight - OKButton.Height - 50
    CancelButton.Move OKButton.Left - CancelButton.Width - 50, ScaleHeight - CancelButton.Height - 50
    
    lblEntriesFound.Top = ScaleHeight - lblEntriesFound.Height - 100
    
    If picHold.Top > ScaleHeight - 2000 Then
        If ScaleHeight - 2000 > 1500 Then
            picHold.Top = ScaleHeight - 2000
        Else
            picHold.Top = 1500
        End If
    End If
    
    lstFindResult.Height = picHold.Top - lstFindResult.Top
    txtResult.Top = picHold.Top + picHold.Height
    txtResult.Height = OKButton.Top - txtResult.Top - 50
    
    If picHold.ScaleWidth / 2 > 4800 Then
        imgMove.Left = picHold.ScaleWidth / 2
        imgMove2.Left = imgMove.Left
    End If
End Sub

Private Sub imgOpen_Click()
    Dim K
    
    On Error GoTo Err_Exit
    CDialog.Filter = "WIN32API File (*.txt)|*.txt"
    CDialog.FileName = APIFileName
    CDialog.ShowOpen
    
    APIFileName = CDialog.FileName
    
    If APIFileName <> "" Then
        Me.Caption = "API Add-In For Visual Basic - " & Mid(APIFileName, _
            InStrRev(APIFileName, "\", , vbBinaryCompare) + 1, Len(APIFileName))
        
        For Each K In Me.Controls
            If TypeOf K Is CommandButton Then
                If K.Caption <> "Copy && Close" Then K.Enabled = True
            End If
        Next K
    Else
        Me.Caption = "API Add-In For Visual Basic"
        
        For Each K In Me.Controls
            If TypeOf K Is CommandButton Then
                If K.Caption <> "&Cancel" Then K.Enabled = False
            End If
        Next K
    End If
    
    lstFindResult.ListItems.Clear
    cmbType.ListIndex = 0
    UpdateEntries
    
    Connect.Show
    Exit Sub
Err_Exit:
    Err.Clear
    Exit Sub
End Sub

Private Sub imgMove_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PY = Y
End Sub

Private Sub imgMove_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 0 And PY <> 0 Then
        If (picHold.Top - (PY - Y)) < 1500 Then
            picHold.Top = 1500
        ElseIf (picHold.Top - (PY - Y)) > ScaleHeight - 2000 Then
            picHold.Top = ScaleHeight - 2000
        Else
            picHold.Top = picHold.Top - (PY - Y)
        End If
        
        lstFindResult.Height = picHold.Top - lstFindResult.Top
        txtResult.Top = picHold.Top + picHold.Height
        txtResult.Height = OKButton.Top - txtResult.Top - 50
    End If
End Sub

Private Sub imgMove_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PY = 0
End Sub

Private Sub imgMove2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgMove_MouseDown Button, Shift, X, imgMove2.Top + Y
End Sub

Private Sub imgMove2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgMove_MouseMove Button, Shift, X, imgMove2.Top + Y
End Sub

Private Sub imgMove2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgMove_MouseUp 0, 0, 0, 0
End Sub

Private Sub lstFindResult_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static SortOrder As Boolean, PrevColumnName As String
    
    If ColumnHeader.Text <> PrevColumnName Then
        SortOrder = False
    Else
        SortOrder = Not SortOrder
    End If
    
    lstFindResult.SortKey = ColumnHeader.Index - 1
    lstFindResult.SortOrder = IIf(SortOrder, lvwDescending, lvwAscending)
    lstFindResult.Sorted = True
    
    PrevColumnName = ColumnHeader.Text
End Sub

Private Sub lstFindResult_DblClick()
    cmdAdd_Click
End Sub

Private Sub OKButton_Click()
    Me.VBInstance.ActiveCodePane.CodeModule.AddFromString txtResult.Text
    
    Unload Me
    Connect.Hide
End Sub

Private Sub txtFind_Change()
    cmdFind.Enabled = cmbType.ListIndex <> -1 And txtFind.Text <> "" And APIFileName <> ""
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmdFind.Enabled Then cmdFind_Click
        KeyAscii = 0
    End If
End Sub

Public Sub UpdateTextBox()
    Dim Q, OutText As String
    
    If SConstants.Count > 0 Then
        For Each Q In SConstants
            OutText = OutText & vbNewLine & Q
        Next Q
        
        OutText = OutText & vbNewLine
    End If
    
    If STypes.Count > 0 Then
        For Each Q In STypes
            OutText = OutText & vbNewLine & Q & vbNewLine
        Next Q
    End If
    
    If SDeclares.Count > 0 Then
        For Each Q In SDeclares
            OutText = OutText & vbNewLine & Q
        Next Q
        
        OutText = OutText & vbNewLine
    End If
    
    txtResult.Text = OutText
End Sub

Private Sub txtResult_Change()
    OKButton.Enabled = txtResult.Text <> ""
End Sub

Public Sub UpdateEntries()
    lblEntriesFound.Caption = lstFindResult.ListItems.Count & " Entries found, " & _
        (SConstants.Count + STypes.Count + SDeclares.Count) & " Entries Selected"
End Sub

Public Sub GetTypes(TypesCol As Collection, ByVal DeclareStr As String)
    Dim K As Integer, StrType As String
    
    DeclareStr = Mid(DeclareStr, InStr(1, DeclareStr, "(") + 1)
    DeclareStr = Left(DeclareStr, InStr(1, DeclareStr, ")"))
    
    Do While InStr(1, DeclareStr, "As ", vbTextCompare) > 0
        DeclareStr = Mid(DeclareStr, InStr(1, DeclareStr, "As ", vbTextCompare) + 3)
        
        K = InStr(1, DeclareStr, ",") - 1
        If K = -1 Then K = InStr(1, DeclareStr, ")") - 1
        If K = -1 Then K = InStr(1, DeclareStr, " ") - 1
        
        If K > 0 Then
            StrType = UCase(Trim(Left(DeclareStr, K)))
            
            Select Case StrType
            Case "STRING", "INTEGER", "LONG", "BYTE", "ANY"
            
            Case Else
                On Error Resume Next
                TypesCol.Add StrType, StrType
                On Error GoTo 0
            End Select
        End If
    Loop
End Sub

Public Sub AddTypes(TypesCol As Collection, PP As String)
    Dim StrFile As String, FNum As Integer, K As Long, Q As Long, I As Integer
    Dim StrType As String
    
    If TypesCol.Count = 0 Then Exit Sub
    
    FNum = FreeFile
    Open APIFileName For Binary Access Read Lock Write As FNum
        StrFile = String(LOF(FNum), 0)
        Get #FNum, , StrFile
    Close FNum
    
    Do
        I = I + 1
        K = InStr(1, StrFile, vbNewLine & "Type " & TypesCol(I), vbTextCompare) + 2
        If K = 2 Then K = InStr(1, StrFile, vbNewLine & "Private Type " & TypesCol(I), vbTextCompare) + 2
        If K = 2 Then K = InStr(1, StrFile, vbNewLine & "Public Type " & TypesCol(I), vbTextCompare) + 2
        
        If K > 2 Then
            Q = InStr(K + 5, StrFile, "End Type", vbTextCompare) + 8
            
            If Q > 8 Then
                StrType = PP & Trim(Mid(StrFile, K, Q - K))
                On Error Resume Next
                If STypes.Count = 0 Then
                    STypes.Add StrType, StrType
                Else
                    STypes.Add StrType, StrType, 1
                End If
                On Error GoTo 0
                
                ' find types inside other types...
                
                Dim TLines() As String, TLine As Variant, VR As String
                TLines = Split(StrType, vbNewLine, , vbBinaryCompare)
                
                For Each TLine In TLines
                    K = InStr(1, TLine, " As ", vbTextCompare)
                    
                    If K > 0 Then
                        Q = InStr(K + 5, TLine, " ")
                        If Q = 0 Then Q = Len(TLine)
                        
                        VR = UCase(Trim(Mid(TLine, K + 4, Q - K)))
                        Select Case VR
                        Case "STRING", "INTEGER", "LONG", "BYTE"
                            
                        Case Else
                            On Error Resume Next
                            TypesCol.Add VR, VR
                            On Error GoTo 0
                        End Select
                    End If
                Next TLine
            End If
        End If
    Loop Until I = TypesCol.Count
End Sub
