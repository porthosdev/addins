VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFindAll 
   Caption         =   "  Source Search"
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11190
   Icon            =   "frmFindAll.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3765
   ScaleWidth      =   11190
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkMatchCase 
      Caption         =   "Match Case"
      Height          =   330
      Left            =   9855
      TabIndex        =   6
      Top             =   90
      Width           =   1275
   End
   Begin VB.CheckBox chkWholeWord 
      Caption         =   "Whole Word"
      Height          =   285
      Left            =   8505
      TabIndex        =   5
      Top             =   90
      Width           =   1230
   End
   Begin MSComctlLib.ListView lvMod 
      Height          =   3210
      Left            =   45
      TabIndex        =   4
      Top             =   450
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   5662
      View            =   3
      LabelEdit       =   1
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Hits"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7155
      TabIndex        =   2
      Top             =   45
      Width           =   1230
   End
   Begin VB.TextBox txtFind 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1080
      TabIndex        =   1
      Top             =   45
      Width           =   6000
   End
   Begin MSComctlLib.ListView lv 
      Height          =   3210
      Left            =   3375
      TabIndex        =   3
      Top             =   450
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   5662
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Compline"
         Text            =   "Line"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "procedure"
         Text            =   "Procedure"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "codeline"
         Text            =   "Code"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   825
   End
End
Attribute VB_Name = "frmFindAll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'These routines are simplified versions of code from CodeFixer addin
'by Roger Gilchrist <rojagilkrist@hotmail.com> Copyright 2003


Private bCancel As Boolean
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Const HWND_TOPMOST = -1

Sub SetWindowTopMost(f As Form)
   SetWindowPos f.hwnd, HWND_TOPMOST, f.Left / 15, _
        f.Top / 15, f.Width / 15, _
        f.Height / 15, Empty
End Sub

Public Function GetActiveCodeModule() As CodeModule
  On Error Resume Next
  Set GetActiveCodeModule = GetActiveCodePane.CodeModule
  On Error GoTo 0
End Function

Public Function GetActiveCodePane() As CodePane
  On Error Resume Next
  Set GetActiveCodePane = g_VBInstance.ActiveCodePane
  On Error GoTo 0
End Function

Public Function GetActiveModuleName() As String
  On Error Resume Next
  GetActiveModuleName = GetActiveCodePane.CodeModule.Name
  On Error GoTo 0
End Function

Public Function GetActiveProject() As VBProject
  On Error Resume Next
  Set GetActiveProject = VBInstance.ActiveVBProject
  On Error GoTo 0
End Function

Private Function GetCurrentProcedureName() As String
  On Error Resume Next
  Dim StartLine As Long
  Dim lJunk     As Long
  With GetActiveCodePane
        .GetSelection StartLine, lJunk, lJunk, lJunk
        lJunk = 0
        GetLineData .CodeModule, StartLine, GetCurrentProcedureName, lJunk, lJunk, lJunk, lJunk, lJunk
  End With
End Function

Public Sub GetLineData(cmpMod As CodeModule, _
                       ByVal cdeline As Long, _
                       ProcName As String, _
                       ProcLineNo As Long, _
                       PRocStartLine As Long, _
                       ProcEndLine As Long, _
                       ProcHeadLine As Long, _
                       Proc1stInsertLine As Long)

  Dim I         As Long
  Dim K         As Long
  Dim CleanElem As Variant

  On Error Resume Next
    
  For I = 1 To 4
        
        K = Choose(I, vbext_pk_Proc, vbext_pk_Get, vbext_pk_Let, vbext_pk_Set)
        CleanElem = Null
        
        On Error Resume Next
        CleanElem = cmpMod.ProcOfLine(cdeline, K) 'IF you crash here first check that Error Trapping is not ON
        On Error GoTo 0
        
        If Not IsNull(CleanElem) Then
            ProcName = CleanElem
            
            If Len(ProcName) = 0 Then
                ProcName = "(Declarations)"
                ProcLineNo = cdeline
                PRocStartLine = 1
                ProcEndLine = cmpMod.CountOfDeclarationLines
             Else
                ProcLineNo = cmpMod.ProcBodyLine(ProcName, K)
                PRocStartLine = cmpMod.PRocStartLine(ProcName, K)
                ProcHeadLine = PRocStartLine
                Proc1stInsertLine = ProcHeadLine
                Proc1stInsertLine = Proc1stInsertLine + 1
                ProcEndLine = PRocStartLine + cmpMod.ProcCountLines(ProcName, K)
            End If
            
            Exit For
        End If
        
  Next I

End Sub


Public Function GetProcName(cMod As CodeModule, ByVal Sline As Long) As String

  On Error Resume Next
  
  GetProcName = cMod.ProcOfLine(Sline, vbext_pk_Proc)
  
  If LenB(GetProcName) = 0 Then
        GetProcName = cMod.ProcOfLine(Sline, vbext_pk_Let)
  End If
  
  If LenB(GetProcName) = 0 Then
        GetProcName = cMod.ProcOfLine(Sline, vbext_pk_Get)
  End If
  
  If LenB(GetProcName) = 0 Then
        GetProcName = cMod.ProcOfLine(Sline, vbext_pk_Set)
  End If
  
  If LenB(GetProcName) = 0 Then
        'dummy for detecting that item is in Declaration section
        GetProcName = "(Declarations)"
  End If
  
  On Error GoTo 0

End Function

Public Function GetProcLineNumber(cmpMod As CodeModule, CodeLineNo As Long) As String

  Dim LProcName As String
  Dim I         As Long
  Dim CleanElem As Variant

  On Error Resume Next
    
  LProcName = GetProcName(cmpMod, CodeLineNo)
  If LProcName = "(Declarations)" Then
     GetProcLineNumber = CodeLineNo
  Else
  
     'The + 1 is because ProcBodyLine returns a 0 based count but most people like 1 based counts
     'Oddly CodeLineNo which is generated by VB's Find is 1 based
     For I = 1 To 4
        
        CleanElem = Null
        
        On Error Resume Next
        
        'IF you crash here first check that Error Trapping is not ON
        CleanElem = CodeLineNo - cmpMod.ProcBodyLine(LProcName, Choose(I, vbext_pk_Proc, vbext_pk_Get, vbext_pk_Let, vbext_pk_Set)) + 1
        
        On Error GoTo 0
        
        If Not IsNull(CleanElem) Then
            GetProcLineNumber = CleanElem
            Exit For
        End If
        
     Next I
     
  End If

End Function

Private Function ModuleExt(Modtype As Long) As String
  
  Select Case Modtype
        Case vbext_ct_StdModule:  ModuleExt = ".bas"  '1 standard module.
        Case vbext_ct_ClassModule: ModuleExt = ".cls" '2 class module.
        Case vbext_ct_MSForm: ModuleExt = ".frm"      '3 form.
        Case vbext_ct_ResFile: ModuleExt = ".res"     '4 standard resource file.
        Case vbext_ct_VBForm: ModuleExt = ".frm"      '5 Visual Basic form.
        Case vbext_ct_VBMDIForm: ModuleExt = ".frm"   '6 The component is an MDI form.
        Case vbext_ct_PropPage: ModuleExt = ".pag"    '7 property page.
        Case vbext_ct_UserControl: ModuleExt = ".ctl" '8 user control.
        Case vbext_ct_DocObject: ModuleExt = ".dob"    '9 RelatedDocument.
        Case vbext_ct_ActiveXDesigner: ModuleExt = ".dsr" '11 ActiveX designer.
  End Select
  
End Function

Public Sub DoSearch(lv As ListView, strfind As String, Optional wholeWord As Boolean, Optional matchCase As Boolean)

  
    Dim PrevCurCodePane As Long
    Dim SelstartLine    As Long
    Dim SeEndLine       As Long
    Dim SelStartCol     As Long
    Dim SelEndCol       As Long
    Dim ProcLineNo      As Long
    Dim OldPosition     As Long
    Dim StartLine       As Long
    Dim startCol        As Long
    Dim EndLine         As Long
    Dim endCol          As Long
    Dim strStrComTest   As String
    Dim code            As String
    Dim ProcName        As String
    Dim li              As ListItem
    Dim CompMod         As CodeModule
    Dim comp            As VBComponent
    Dim proj            As VBProject
    Dim modules As Long
    
    Dim parent As CModule
    Dim result As CResult
    Dim hits As Long
    
     On Error Resume Next
     bCancel = False
     lv.ListItems.Clear
     lvMod.ListItems.Clear
      
      For Each proj In g_VBInstance.VBProjects
                  
          For Each comp In proj.VBComponents
                
                Set parent = New CModule
                parent.module = comp.Name & ModuleExt(comp.Type)
                parent.proj = proj.Name
                
                modules = modules + 1
                Me.Caption = "Searching Component " & comp.Name
                
                If LenB(comp.Name) > 0 Then
                
                    Set CompMod = comp.CodeModule
    
                    StartLine = 1 'initialize search range
                    startCol = 1
                    EndLine = -1
                    endCol = -1
                    
                    Do While CompMod.Find(strfind, StartLine, startCol, EndLine, endCol, wholeWord, matchCase, False)
                        
                        DoEvents
                        code = CompMod.Lines(StartLine, 1)
                        If bCancel Then Exit Do
                        
                        If LenB(code) > 0 Then
                            Set result = New CResult
                            result.proc = GetProcName(CompMod, StartLine)
                            result.lineNo = StartLine
                            result.text = Trim$(code)
                            result.ComponentName = comp.Name
                            parent.hits.Add result
                            hits = hits + 1
                            code = Empty
                        End If
    
                        StartLine = StartLine + 1
                        If StartLine > CompMod.CountOfLines Then Exit Do
                        
                        startCol = 1
                        EndLine = -1
                        endCol = -1
                        
                    Loop
              
              End If

              If bCancel Then Exit For
              
              If parent.hits.Count > 0 Then
                    Set li = lvMod.ListItems.Add(, , IIf(parent.hits.Count < 10, " ", "") & parent.hits.Count)
                    li.SubItems(1) = parent.module 'if proj.count > 1 then prepend proj name..
                    Set li.Tag = parent
              End If
              
        Next comp
            
        If bCancel Then Exit For
    Next proj
      
    Me.Caption = "Searched " & modules & " Modules found " & hits & " results"
    cmdSearch.Caption = "Search"
    
    lvMod_ItemClick lvMod.ListItems(1)
    
End Sub
 
Private Sub cmdSearch_Click()
  On Error Resume Next
  
    Dim wholeWord As Boolean
    Dim matchCase As Boolean
    
    If cmdSearch.Caption = "Search" Then
        
        If Len(txtFind) = 0 Or txtFind = " " Then
            MsgBox "Enter a string to search for!"
            Exit Sub
        End If
        
        If chkMatchCase.value = 1 Then matchCase = True
        If chkWholeWord.value = 1 Then wholeWord = True
        
        bCancel = False
        cmdSearch.Caption = "Cancel"
        DoSearch lv, txtFind, wholeWord, matchCase
        
    Else
        bCancel = True
        cmdSearch.Caption = "Search"
    End If
    
End Sub

Private Sub Form_Load()
    FormPos Me, True
    SetWindowTopMost Me
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    lv.Width = Me.Width - lv.Left - 200
    lv.ColumnHeaders(3).Width = lv.Width - lv.ColumnHeaders(3).Left - 200
    lvMod.ColumnHeaders(2).Width = lvMod.Width - lvMod.ColumnHeaders(2).Left - 200
    lv.Height = Me.Height - lv.Top - 200
    lvMod.Height = lv.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FormPos Me, True, True
End Sub

Private Sub Label1_Click()
    txtFind.text = Empty
End Sub

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim r As CResult
    Dim proj As VBProject
    Dim cp As CodePane
    Dim comp As VBComponent
    
    On Error GoTo hell
    
1    Set r = Item.Tag

2    For Each proj In g_VBInstance.VBProjects
3        For Each comp In proj.VBComponents
4            If comp.Name = r.ComponentName Then
5                comp.CodeModule.CodePane.Show
6                comp.CodeModule.CodePane.TopLine = r.lineNo
'7                comp.CodeModule.CodePane.SetSelection r.lineNo, InStr(1, r.text, txtFind, vbTextCompare), r.lineNo, Len(txtFind)
                Exit Sub
            End If
        Next
    Next
    
    Exit Sub
hell:
    MsgBox "Error in lv_ItemClick: " & Erl & " - " & Err.Description
    
End Sub

Private Sub lvMod_ItemClick(ByVal Item As MSComctlLib.ListItem)
  On Error Resume Next
  
    Dim p As CModule
    Dim r As CResult
    Dim li As ListItem
    
    Set p = Item.Tag
    lv.ListItems.Clear
    
    For Each r In p.hits
        Set li = lv.ListItems.Add(, , r.lineNo)
        li.SubItems(1) = r.proc
        li.SubItems(2) = r.text
        Set li.Tag = r
    Next
    
End Sub

Private Sub txtFind_KeyPress(KeyCode As Integer)
  On Error Resume Next
  
    If KeyCode = 13 Then
        cmdSearch_Click
        KeyCode = 0
    End If
End Sub

Sub FormPos(fform As Form, Optional andSize As Boolean = False, Optional save_mode As Boolean = False)
    
    On Error Resume Next
    
    Dim f, sz
    f = Split(",Left,Top,Height,Width", ",")
    
    If fform.WindowState = vbMinimized Then Exit Sub
    If andSize = False Then sz = 2 Else sz = 4
    
    For I = 1 To sz
        If save_mode Then
            ff = CallByName(fform, f(I), VbGet)
            SaveSetting App.EXEName, fform.Name & ".FormPos", f(I), ff
        Else
            def = CallByName(fform, f(I), VbGet)
            ff = GetSetting(App.EXEName, fform.Name & ".FormPos", f(I), def)
            CallByName fform, f(I), VbLet, ff
        End If
    Next
    
End Sub

