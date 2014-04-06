VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmAddIn 
   Caption         =   "Memory Window              http://sandsprite.com"
   ClientHeight    =   4560
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9810
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   9810
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Top Most"
      Height          =   285
      Left            =   6075
      TabIndex        =   6
      Top             =   45
      Width           =   1005
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   4
      Top             =   45
      Width           =   3525
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Dump"
      Height          =   330
      Left            =   8955
      TabIndex        =   3
      Top             =   0
      Width           =   825
   End
   Begin VB.ComboBox cboType 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4365
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   45
      Width           =   1635
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "Prev"
      Height          =   330
      Left            =   7290
      TabIndex        =   1
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   330
      Left            =   8055
      TabIndex        =   0
      Top             =   0
      Width           =   645
   End
   Begin RichTextLib.RichTextBox Text2 
      Height          =   4065
      Left            =   0
      TabIndex        =   7
      Top             =   405
      Width           =   9690
      _ExtentX        =   17092
      _ExtentY        =   7170
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmAddIn.frx":0000
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
   Begin VB.Label Label1 
      Caption         =   "Address"
      Height          =   240
      Left            =   0
      TabIndex        =   5
      Top             =   45
      Width           =   645
   End
   Begin VB.Menu mnUTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuSaveMem 
         Caption         =   "Save Memory Block"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuNewWindow 
         Caption         =   "New Window"
      End
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Name:   VB6 Memory Window Addin
' Author: David Zimmer
' Site:   http://sandsprite.com
'
' notes:
'   might have a couple bugs, but its a decent tool and does what it needs to
'   olly.dll is open source if missing download here:
'          http://sandsprite.com/CodeStuff/olly_dll.html
'
' features:
'   view data as: hexdump, longs, long address, ascii, unicode, disasm
'   next/previous memory block
'   always on top
'   hit escape and it takes you back through the displayed address history
'   in long, long address, and disasm mode, if you ctrl + mouse over a valid address it will hyperlink it (like lazarus)
'
'
' todo:
'       savemem command (binary)
'       allow user to select external process..

#If IS_ADDIN Then
    Public VBInstance As VBIDE.VBE
    Public Connect As Connect
#End If



Public lastVA As Long
Public nextVa As Long
Public lastText As String

Dim history As New Collection
Dim curView As CView


Const szHexDump = &H200
Const szDwordDump = 25
Const szString = &H1000

Private DontReact As Boolean
Private working As Boolean
Private selA As New Collection
Private HighLightRunning As Boolean
Private ctrlDown As Boolean
Private lastWord As String
Private hilight As CSelection



Private Sub cboType_Click()
    If DontReact Then Exit Sub
    If lastVA = 0 Then Exit Sub
    If lastText <> Text1 Then 'address changed but they didnt hit return or dump button..
        Command1_Click
    Else
        DisplayData lastVA
    End If
End Sub

Private Sub Check1_Click()
    TopMost Me, (Check1.value = 1)
End Sub

Private Sub cmdNext_Click()
    
    If curView Is Nothing Then Exit Sub
    
    Select Case cboType.ListIndex
        Case 0: 'hexdump
                DisplayData curView.addr + szHexDump
 
        Case 1: 'dwords
                DisplayData curView.addr + (szDwordDump * 4 * 4)

        Case 2: 'long address
                DisplayData curView.addr + (szDwordDump * 4)

        Case 3: 'ascii
                DisplayData curView.addr + szString
                
        Case 4: 'unicode
                DisplayData curView.addr + szString
                
        Case 5: 'disasm
                DisplayData curView.nextVa
                
    End Select
End Sub

Private Sub cmdPrev_Click()

    If curView Is Nothing Then Exit Sub
    
    Select Case cboType.ListIndex
        Case 0: 'hexdump
                DisplayData curView.addr - szHexDump
 
        Case 1: 'long
                DisplayData curView.addr - (szDwordDump * 4 * 4)
                
        Case 2: 'long address
                DisplayData curView.addr - (szDwordDump * 4)

        Case 3: 'ascii
                DisplayData curView.addr - szString
                
        Case 4: 'unicode
                DisplayData curView.addr - szString
                
        Case 5: 'disasm
                'didnt implement the disasm back code its messy..
                
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 17 Then
        ctrlDown = True
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 17 Then
        ctrlDown = False
        If Not hilight Is Nothing Then
            hilight.Undo Text2
            Set hilight = Nothing
        End If
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim cv As CView
    
    On Error Resume Next
    
    If KeyAscii = 27 Then 'they hit escape, go back in history..
        
        KeyAscii = 0
        Set cv = GetHistory()
        
        If cv Is Nothing Then Exit Sub
        
        If cv.addr = lastVA Then
            Set cv = GetHistory()
            If cv Is Nothing Then Exit Sub
        End If
        
        If cv.addr <> 0 Then
            
            DontReact = True
            cboType.ListIndex = cv.viewMode
            DontReact = False
            
            DisplayData cv.addr, False
            Text1 = "0x" & Hex(cv.addr)
            
            Set curView = cv
            
        End If
   
    End If
End Sub

Private Function GetHistory() As CView
    Dim tmp
    Dim c As Long
    
    On Error Resume Next
    If history.Count = 0 Then Exit Function
     
    Set GetHistory = history.Item(history.Count)
    history.Remove history.Count
    'List1.RemoveItem List1.ListCount - 1
    
End Function



Private Sub mnuNewWindow_Click()
    On Error Resume Next
    Dim f As New frmAddIn
    f.Show
End Sub

Private Sub mnuSaveMem_Click()
    MsgBox "todo"
End Sub

Private Sub Text1_KeyPress(KeyCode As Integer)
    
    If KeyCode = 13 Then 'if they hit return dump data
        Command1_Click
        KeyCode = 0
    End If
    
End Sub

Private Sub Command1_Click()
    Dim va As Long
    Dim tmp As String
    Dim exp As String
    
    On Error Resume Next
    
    value = 0
    Text2.Text = Empty
    lastText = Text1
    tmp = Replace(Text1, "0x", "&h", , , vbTextCompare)
    
    'If Left(Text1, 1) = "?" Then     'as awesome as this would have been..doesnt work from an addin :_(
    '    exp = Mid(tmp, 2)
    '    ExecuteLine "Form1.Value=" & exp
    '    va = CLng(Value)
    'Else
        va = CLng(tmp)
        If va = 0 Then
            Err.Clear
            va = CLng("&h" & tmp) 'maybe it was a hex address with no prefix..
        End If
    'End If
    
    If va = 0 Or Err.Number <> 0 Then
        If Len(exp) > 0 Then
            Text2.Text = "Could not evaluate expression: " & exp & vbCrLf & "return was: " & value
        Else
            Text2.Text = "Could not convert to long address.."
        End If
        Exit Sub
    End If
        
    DisplayData va
  
End Sub

Function DisplayData(va As Long, Optional recordHistory As Boolean = True)

    Dim b() As Byte
    Dim l() As Long
    Dim tmp() As String
    Dim i As Long
    Dim t2 As String
    
    'this list1 stuff was to visually debug the history logic..
    
    If recordHistory Then
    
        If Not curView Is Nothing And va = lastVA Then
            curView.viewMode = cboType.ListIndex
            'List1.List(List1.ListCount - 1) = List1.List(List1.ListCount - 1) & "," & cboType.ListIndex
        End If
        
        Set curView = New CView
        curView.addr = va
        curView.viewMode = cboType.ListIndex
        history.Add curView
        'List1.AddItem Hex(va) & "," & cboType.ListIndex
        
    End If
     
    ResetRTF Text2
    lastVA = va
        
     Select Case cboType.ListIndex
        Case 0: 'hexdump
                nextVa = va + szHexDump
                If Not ReadMemBuf(va, szHexDump, b()) Then GoTo failed
                t2 = hexdump(va, b())
                HighlightOffsets Text2, t2
                
                
        Case 1: 'long
                nextVa = va + (szDwordDump * 4)  '4 per line
                
                If Not ReadMemLongs(va, szDwordDump * 4, l()) Then GoTo failed
                
                For i = 0 To (szDwordDump * 4) - 1
                    If i Mod 4 = 0 Then
                        If Len(t2) > 0 Then push tmp, t2
                        t2 = tHex(va + (i * 4)) 'add address
                    End If
                    t2 = t2 & "  " & tHex(l(i))
                Next
                
                t2 = Join(tmp, vbCrLf)
                HighlightOffsets Text2, t2
                
        Case 2: 'long address
                nextVa = va + (szDwordDump * 4)
                If Not ReadMemLongs(va, szDwordDump, l()) Then GoTo failed
                
                For i = 0 To szDwordDump - 1
                    push tmp, tHex(va + (i * 4)) & "  " & tHex(l(i)) & "  " & GetMemory(l(i), True)
                Next
                t2 = Join(tmp, vbCrLf)
                HighlightOffsets Text2, t2

        Case 3: 'ascii
                nextVa = va + szString
                If Not ReadMemBuf(va, szString, b()) Then GoTo failed
                killNonPrintable b, True
                Text2 = StrConv(b, vbUnicode, LANG_US)
                
        Case 4: 'unicode
                nextVa = va + szString
                If Not ReadMemBuf(va, szString, b()) Then GoTo failed
                killNonPrintable b, False
                t2 = StrConv(b, vbUnicode, LANG_US)
                Text2 = Replace(t2, Chr(0), Empty)
                
        Case 5: 'disasm
                t2 = DisasmBlock(va)
                HighlightOffsets Text2, Replace(t2, vbTab, "    ")
                
        
    End Select
    
    
    Exit Function
    
failed:
     Text2 = "Could not read memory at address..."
     Exit Function

End Function

Private Sub Form_Load()

    Check1.value = GetSetting(App.EXEName, "settings", "onTop", 0)
    FormPos Me
        
    Dim h As Long
    h = LoadLibrary("olly.dll")
    
    DontReact = True
    cboType.AddItem "Hexdump"
    cboType.AddItem "Long"
    cboType.AddItem "Long Address"
    cboType.AddItem "ASCII"
    cboType.AddItem "Unicode"
    
    If h <> 0 Then
        cboType.AddItem "Disasm"
    Else
        Text2 = "Disasm mode not available could not find olly.dll"
    End If
    
    cboType.ListIndex = 0
    DontReact = False
        
    hProcess = OpenProcess(PROCESS_VM_READ, False, GetCurrentProcessId())
         
    If IsIde() Then
        Text1 = "0x" & Hex(h)
        Command1_Click
    End If
     
     
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CloseHandle hProcess
    FormPos Me, True, True
    SaveSetting App.EXEName, "settings", "onTop", Check1.value
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Text2.Width = Me.Width - Text2.Left - 200
    Text2.Height = Me.Height - Text2.Top - 800
End Sub

Private Sub Text2_Click()
    If Not hilight Is Nothing Then
        Text1 = "0x" & hilight.selWord
        hilight.Undo Text2
        Set hilight = Nothing
        Command1_Click
    End If
End Sub

Private Sub Text2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
   Dim i As Long
   Dim curWord As String
   Dim curTick As Long
   Dim lngPos As Long
   Dim ss As Long
   Dim topLine As Long
   Dim startPos As Long
   Dim address As Long
   Dim value As Long
   
   If working Then Exit Sub
   If Not ctrlDown Then Exit Sub
   If HighLightRunning Then Exit Sub

   i = cboType.ListIndex
   
   'only for long, long addr, and disasm modes
   If i <> 1 And i <> 2 And i <> 5 Then Exit Sub

   curWord = WordUnderCursor(Text2, x, y, startPos)
   
   If Len(curWord) = 0 Then
        Deselect
        Exit Sub
   End If
   
   If curWord = lastWord Then Exit Sub
   lastWord = curWord
   
   address = isHexNum(curWord)
   If address = 0 Then
        Deselect
        Exit Sub
   End If
   
   If Not ReadLng(address, value) Then
        Deselect
        Exit Sub
   End If
   
   If Not hilight Is Nothing Then
        If hilight.SelStart = startPos Then Exit Sub
        hilight.Undo Text2
        Set hilight = Nothing
   End If
   
   working = True
   LockWindowUpdate Text2.hWnd

   'save current selection offsets
   'topLine = TopLineIndex(Text2)  'currently a bug in this..but we usually have small display size so probably no scrolling anyway fuck it..
   ss = Text2.SelStart
   
   Text2.SelStart = startPos
   Text2.SelLength = Len(curWord)
   'Me.Caption = curWord

   Set hilight = New CSelection
   hilight.LoadSel Text2
    
   Text2.SelColor = vbBlue
   Text2.SelUnderline = True
   Screen.MousePointer = vbArrow
   
   Text2.SelStart = ss
   'ScrollToLine Text2, CInt(topLine)
   LockWindowUpdate 0
   working = False

End Sub


Function Deselect()
    If Not hilight Is Nothing Then
        hilight.Undo Text2
        Set hilight = Nothing
   End If
End Function

