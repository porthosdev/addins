VERSION 5.00
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
      TabIndex        =   7
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
      TabIndex        =   5
      Top             =   45
      Width           =   3525
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Dump"
      Height          =   330
      Left            =   8955
      TabIndex        =   4
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
      TabIndex        =   3
      Top             =   45
      Width           =   1635
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4110
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   405
      Width           =   9690
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
   Begin VB.Label Label1 
      Caption         =   "Address"
      Height          =   240
      Left            =   0
      TabIndex        =   6
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
'
'
' todo:
'       ctrl click on an address to goto it would be nice
'       highlight address of data in blue would be nice..
'       savemem command (binary)

#If IS_ADDIN Then
    Public VBInstance As VBIDE.VBE
    Public Connect As Connect
#End If

'this one only works in the IDE, doesnt work from an addin? (running or breakpoint)
'Private Declare Function EbExecuteLine Lib "vba6.dll" (ByVal pStringToExec As Long, ByVal Unknownn1 As Long, ByVal Unknownn2 As Long, ByVal fCheckOnly As Long) As Long

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, lpBuffer As Byte, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function ReadProcessLongs Lib "kernel32" Alias "ReadProcessMemory" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, lpBuffer As Long, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

Private Declare Function Disasm Lib "olly.dll" (ByRef src As Byte, ByVal srcsize As Long, ByVal ip As Long, Disasm As t_Disasm, Optional disasmMode As Long = 4) As Long

Private Type t_Disasm
  ip As Long
  dump As String * 256
  result As String * 256
  unused(1 To 308) As Byte
End Type

Public value
Public hProcess As Long
Public lastVA As Long
Public nextVa As Long
Public lastText As String

Dim history As New Collection
Dim curView As CView

Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Const PROCESS_VM_READ = (&H10)
Const LANG_US = &H409

Const szHexDump = &H200
Const szDwordDump = 25
Const szString = &H1000
Private DontReact As Boolean

'Public Function ExecuteLine(sCode As String) As Boolean
'   ExecuteLine = EbExecuteLine(StrPtr(sCode), 0, 0, 0) = 0
'End Function

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
            
        End If
   
    End If
End Sub

Private Function GetHistory() As CView
    Dim tmp
    Dim c As Long
    
    On Error Resume Next
    If history.count = 0 Then Exit Function
     
    Set GetHistory = history.Item(history.count)
    history.Remove history.count
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

Private Sub Text1_KeyPress(Keycode As Integer)
    
    If Keycode = 13 Then 'if they hit return dump data
        Command1_Click
        Keycode = 0
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
    
    'this list1 stuff is just to visually debug the history logic..
    
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
     
    lastVA = va
        
     Select Case cboType.ListIndex
        Case 0: 'hexdump
                nextVa = va + szHexDump
                ReadMemBuf va, szHexDump, b()
                Text2 = hexdump(va, b())
                
        Case 1: 'long
                nextVa = va + (szDwordDump * 4)  '4 per line
                ReadMemLongs va, szDwordDump * 4, l()
                For i = 0 To (szDwordDump * 4) - 1
                    If i Mod 4 = 0 Then
                        If Len(t2) > 0 Then push tmp, t2
                        t2 = tHex(va + (i * 4)) 'add address
                    End If
                    t2 = t2 & "  " & tHex(l(i))
                Next
                Text2 = Join(tmp, vbCrLf)
                
        Case 2: 'long address
                nextVa = va + (szDwordDump * 4)
                ReadMemLongs va, szDwordDump, l()
                For i = 0 To szDwordDump - 1
                    push tmp, tHex(va + (i * 4)) & "  " & tHex(l(i)) & "  " & GetMemory(l(i), True)
                Next
                Text2 = Join(tmp, vbCrLf)

        Case 3: 'ascii
                nextVa = va + szString
                ReadMemBuf va, szString, b()
                killNonPrintable b, True
                Text2 = StrConv(b, vbUnicode, LANG_US)
                
        Case 4: 'unicode
                nextVa = va + szString
                ReadMemBuf va, szString, b()
                killNonPrintable b, False
                t2 = StrConv(b, vbUnicode, LANG_US)
                Text2 = Replace(t2, Chr(0), Empty)
                
        Case 5: 'disasm
                
                Text2 = DisasmBlock(va)
                
        
    End Select
    
End Function

Function ReadMemBuf(start As Long, count As Long, out() As Byte) As Boolean
    Dim ret As Long
    ReDim out(count - 1)
    ret = ReadProcessMemory(hProcess, start, out(0), count, count)
    ReadMemBuf = IIf(ret <> 0, True, False)
End Function

Function ReadMemLongs(start As Long, count As Long, out() As Long) As Boolean
    Dim ret As Long
    ReDim out(count - 1)
    ret = ReadProcessLongs(hProcess, start, out(0), count * 4, count)
    ReadMemLongs = IIf(ret <> 0, True, False)
End Function

Function DisasmBlock(ByVal va As Long, Optional instCount As Long = 20) As String
    Dim tmp() As String
    Dim tmpVa As Long
    Dim instAfterVa As Long
    Dim bytesBack As Long
    Dim n As Long
    Dim x As String

    On Error Resume Next

    'MsgBox "VA: " & Hex(va)

    tmpVa = va

    Dim n1 As String, d As String, n2 As String, n3 As Long

    Do While 1
        x = DisasmVA(tmpVa, n)
        If InStr(x, "??") > 0 Then Exit Do

        push tmp, Hex(tmpVa) & vbTab & x
        instAfterVa = instAfterVa + 1

        If n = 0 Or instAfterVa = instCount Then 'bad disasm or max reached..
            nextVa = tmpVa
            curView.nextVa = tmpVa
            Exit Do
        Else
            tmpVa = tmpVa + n
        End If
 
    Loop

    DisasmBlock = Join(tmp, vbCrLf)

End Function

Function DisasmVA(ByVal va As Long, Optional leng_out As Long, Optional dump_out) As String
    Dim da As t_Disasm
    Dim b()  As Byte
    Dim x As Long
    On Error Resume Next

    If Not ReadMemBuf(va, 20, b) Then
        DisasmVA = "?????"
    Else
        leng_out = Disasm(b(0), UBound(b) + 1, va, da)
        dump_out = da.dump
        x = InStr(dump_out, Chr(0))
        If x > 0 Then dump_out = Mid(dump_out, 1, x - 1)
        DisasmVA = Mid(da.result, 1, InStr(da.result, Chr(0)) - 1)
    End If
End Function

'reads a long from memory and returns it as hex with optional ascii/unicode text dereference..not BSTR though..
Function GetMemory(ByVal va As Long, Optional ByVal asciiDump As Boolean = False) As String
    
    If va = 0 Then Exit Function
    
    Dim r As Long
    Dim b() As Byte
    Dim tmp As String
    Dim i As Long
    Dim isUnicode As Boolean
    Dim oneChance As Boolean
    Dim scanAt As Long
    Dim firstScan As Boolean
    
    If Not ReadLng(va, r) Then Exit Function
    
    GetMemory = " -> " & tHex(r)
    
    firstScan = True
    scanAt = va 'first try direct pointer to string
    
tryAgain:
    
    If Not firstScan Then 'we already tried first mechanism and failed
        If scanAt = r Then 'we failed 2nd too
            Exit Function
        Else
            scanAt = r
        End If
    End If
    
    firstScan = False
    
    If asciiDump Then
        If ReadMemBuf(va, 50, b) Then
            For i = 0 To UBound(b)
                If b(i) > 20 And b(i) < 120 Then
                    If oneChance Then
                        isUnicode = True
                        oneChance = False
                    End If
                    tmp = tmp & Chr(b(i))
                Else
                    If b(i) = 0 And oneChance = False Then 'needs another ascii to reset so 00 00 will terminate
                        oneChance = True
                    Else
                        Exit For
                    End If
                End If
            Next
            If Len(tmp) > 3 Then
                If isUnicode Then tmp = Replace(tmp, Chr(0), Empty)
                tmp = " -> " & IIf(isUnicode, "Uni: ", "Asc: ") & tmp
                If scanAt = r Then tmp = " -> " & tHex(r) & tmp  '**eax=str
                GetMemory = tmp
            Else
                GoTo tryAgain
            End If
        Else
            i = 1 'marker to move to next trial
            GoTo tryAgain
        End If
    End If
            
            
End Function

Sub killNonPrintable(b() As Byte, Optional nullToo As Boolean = False)
    Dim dot As Byte, x As Byte
    
    dot = Asc(".")
    
      For i = 0 To UBound(b)
            x = b(i)
            If x = 0 Then
                If nullToo Then b(i) = dot
            ElseIf x > 32 And x < 127 Then
                'its printable do nothing
            Else
                b(i) = dot
            End If
       Next
            
End Sub

Function hexdump(ByVal base As Long, it() As Byte) As String
    Dim my, i, c, s, a As Byte, b
    Dim lines() As String
    
    my = ""
    For i = 0 To UBound(it)
        a = it(i)
        c = Hex(a)
        c = IIf(Len(c) = 1, "0" & c, c)
        'b = b & IIf(a > 65 And a < 120, Chr(a), ".")
        b = b & IIf((a > 32 And a < 127), Chr(a), ".")
        my = my & c & " "
        If (i + 1) Mod 16 = 0 Then
            push lines(), Hex(base) & " " & my & " [" & b & "]"
            base = base + 16
            my = Empty
            b = Empty
        End If
    Next
    
    If Len(b) > 0 Then
        If Len(my) < 48 Then
            my = my & String(48 - Len(my), " ")
        End If
        If Len(b) < 16 Then
             b = b & String(16 - Len(b), " ")
        End If
        push lines(), my & " [" & b & "]"
    End If
        
    If UBound(it) < 16 Then
        hexdump = Hex(base) & " " & my & " [" & b & "]" & vbCrLf
    Else
        hexdump = Join(lines, vbCrLf)
    End If
    
    
End Function

Function ReadLng(ByVal va As Long, retLng As Long) As Boolean
    Dim b(4) As Byte
    Dim tmp As Long
    If ReadProcessMemory(hProcess, va, b(0), 4, 0) > 0 Then
        CopyMemory tmp, b(0), 4
        retLng = tmp
        ReadLng = True
    End If
End Function

Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

Private Function tHex(x As Long) As String
    Dim t As String
    
    t = Hex(x)
    While Len(t) < 8
        t = "0" & t
    Wend
    tHex = t
    
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


Sub FormPos(fform As Object, Optional andSize As Boolean = True, Optional save_mode As Boolean = False)
    
    On Error Resume Next
    
    Dim f, sz
    f = Split(",Left,Top,Height,Width", ",")
    
    If fform.WindowState = vbMinimized Then Exit Sub
    If andSize = False Then sz = 2 Else sz = 4
    
    For i = 1 To sz
        If save_mode Then
            ff = CallByName(fform, f(i), VbGet)
            SaveSetting "MyAddin", fform.Name & ".FormPos", f(i), ff
        Else
            def = CallByName(fform, f(i), VbGet)
            ff = GetSetting("MyAddin", fform.Name & ".FormPos", f(i), def)
            CallByName fform, f(i), VbLet, ff
        End If
    Next
    
End Sub

Function TopMost(frm As Object, Optional ontop As Boolean = True)
    On Error Resume Next
    s = IIf(ontop, HWND_TOPMOST, HWND_NOTOPMOST)
    SetWindowPos frm.hWnd, s, frm.Left / 15, frm.Top / 15, frm.Width / 15, frm.Height / 15, 0
End Function

Function pop(ary)
    On Error GoTo isEmpty
    x = UBound(ary)
    pop = ary(x)
    ReDim Preserve ary(x - 1)
    Exit Function
isEmpty: Erase ary
End Function

Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    x = UBound(ary)
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function


