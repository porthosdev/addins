VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFindAll 
   Caption         =   "Source Search"
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11190
   LinkTopic       =   "Form1"
   ScaleHeight     =   3765
   ScaleWidth      =   11190
   StartUpPosition =   2  'CenterScreen
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
         Size            =   12
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
      Left            =   9585
      TabIndex        =   2
      Top             =   90
      Width           =   1455
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
      Width           =   8385
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
         Size            =   12
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
Private bCancel As Boolean

'These routines are very simplified versions of code from CodeFixer addin
'by Roger Gilchrist <rojagilkrist@hotmail.com> Copyright 2003




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


Public Sub DoSearch(lv As ListView, strfind As String)

  
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
    Dim Comp            As VBComponent
    Dim proj            As VBProject
    Dim modules As Long
    
    Dim parent As CModule
    Dim result As CResult
    
     ' On Error Resume Next
      bCancel = False
      lv.ListItems.Clear
      lvMod.ListItems.Clear
      
      For Each proj In g_VBInstance.VBProjects
                  
          For Each Comp In proj.VBComponents
                
                Set parent = New CModule
                parent.module = Comp.Name
                parent.proj = proj.Name
                
                modules = modules + 1
                Me.Caption = "Searching Component " & Comp.Name
                
                If LenB(Comp.Name) > 0 Then
                
                    Set CompMod = Comp.CodeModule
    
                    StartLine = 1 'initialize search range
                    startCol = 1
                    EndLine = -1
                    endCol = -1
                    
                    Do While CompMod.Find(strfind, StartLine, startCol, EndLine, endCol, False, False, False)
                        
                        DoEvents
                        code = CompMod.Lines(StartLine, 1)
                        If bCancel Then Exit Do
                        
                        If LenB(code) > 0 Then
                            Set result = New CResult
                            result.proc = GetProcName(CompMod, StartLine)
                            result.lineNo = StartLine
                            result.text = Trim$(code)
                            parent.hits.Add result
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
                    Set li = lvMod.ListItems.Add(, , parent.hits.Count)
                    li.SubItems(1) = parent.module 'if proj.count > 1 then prepend proj name..
                    Set li.Tag = parent
              End If
              
        Next Comp
            
        If bCancel Then Exit For
    Next proj
      
    Me.Caption = "Searched " & modules & " Modules found " & lv.ListItems.Count & " results"
    cmdSearch.Caption = "Search"
 
End Sub
 
Private Sub cmdSearch_Click()
    
    If cmdSearch.Caption = "Search" Then
        
        If Len(txtFind) = 0 Or txtFind = " " Then
            MsgBox "Enter a string to search for!"
            Exit Sub
        End If

        bCancel = False
        cmdSearch.Caption = "Cancel"
        DoSearch lv, txtFind
        
    Else
        bCancel = True
        cmdSearch.Caption = "Search"
    End If
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    lv.Width = Me.Width - lv.Left - 100
    lv.ColumnHeaders(3).Width = lv.Width - lv.ColumnHeaders(3).Left - 100
    lvMod.ColumnHeaders(2).Width = lvMod.Width - lvMod.ColumnHeaders(2).Left - 100
End Sub

Private Sub lvMod_ItemClick(ByVal Item As MSComctlLib.ListItem)

    Dim p As CModule
    Dim r As CResult
    Dim li As ListItem
    
    Set p = Item.Tag
    lv.ListItems.Clear
    
    For Each r In p.hits
        Set li = lv.ListItems.Add(, , r.lineNo)
        li.SubItems(1) = r.proc
        li.SubItems(2) = r.text
    Next
    
End Sub
















