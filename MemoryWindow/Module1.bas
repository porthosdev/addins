Attribute VB_Name = "Module1"
'this one only works in the IDE, doesnt work from an addin? (running or breakpoint)
'Private Declare Function EbExecuteLine Lib "vba6.dll" (ByVal pStringToExec As Long, ByVal Unknownn1 As Long, ByVal Unknownn2 As Long, ByVal fCheckOnly As Long) As Long

Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, lpBuffer As Byte, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function ReadProcessLongs Lib "kernel32" Alias "ReadProcessMemory" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, lpBuffer As Long, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function Disasm Lib "olly.dll" (ByRef src As Byte, ByVal srcsize As Long, ByVal ip As Long, Disasm As t_Disasm, Optional disasmMode As Long = 4) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Type t_Disasm
  ip As Long
  dump As String * 256
  result As String * 256
  unused(1 To 308) As Byte
End Type

Const EM_LINESCROLL = &HB6
Const EM_GETFIRSTVISIBLELINE = &HCE
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2

Const EM_CHARFROMPOS& = &HD7
 
Private Type POINTAPI
    x As Long
    y As Long
End Type

Global hProcess As Long
Global Const PROCESS_VM_READ = (&H10)
Global Const LANG_US = &H409

'Public Function ExecuteLine(sCode As String) As Boolean
'   ExecuteLine = EbExecuteLine(StrPtr(sCode), 0, 0, 0) = 0
'End Function

Function isHexNum(s As String) As Long
    On Error Resume Next
    Dim l As Long
    l = CLng("&h" & s)
    If Err.Number = 0 Then isHexNum = l
End Function

Function TopLineIndex(x As Object) As Long
    TopLineIndex = SendMessage(x.hwnd, EM_GETFIRSTVISIBLELINE, 0, ByVal 0&)
End Function

Sub ScrollToLine(t As Object, x As Integer)
     x = x - TopLineIndex(t)
     ScrollIncremental t, , x
End Sub

Sub ScrollIncremental(t As Object, Optional horz As Integer = 0, Optional vert As Integer = 0)
    'lParam&  The low-order 2 bytes specify the number of vertical
    '          lines to scroll. The high-order 2 bytes specify the
    '          number of horizontal columns to scroll. A positive
    '          value for lParam& causes text to scroll upward or to the
    '          left. A negative value causes text to scroll downward or
    '          to the right.
    ' r&       Indicates the number of lines actually scrolled.
    
    Dim r As Long
    r = CLng(&H10000 * horz) + vert
    r = SendMessage(t.hwnd, EM_LINESCROLL, 0, ByVal r)

End Sub

Function WordUnderCursor(MyRTB As RichTextBox, x As Single, y As Single, startPos As Long) As String
    Dim MyPoint As POINTAPI
    Dim MyPos As Integer
    Dim MyStartPos As Integer
    Dim MyEndPos As Integer
    Dim MyCharacter As Integer
    Dim TextLength As Integer
    
    On Error Resume Next
    MyPoint.x = x \ Screen.TwipsPerPixelX
    MyPoint.y = y \ Screen.TwipsPerPixelY
    MyPos = SendMessage(MyRTB.hwnd, EM_CHARFROMPOS, 0&, MyPoint)
    
    If MyPos <= 0 Then Exit Function
    
    MyText = MyRTB.Text
    
    For MyStartPos = MyPos To 1 Step -1
        MyCharacter = Asc(Mid$(MyRTB.Text, MyStartPos, 1))
        If Not isAlpha(MyCharacter) Then Exit For
           
    Next
    
    MyStartPos = MyStartPos + 1
    TextLength = Len(MyText)
    
    For MyEndPos = MyPos To TextLength
        MyCharacter = Asc(Mid$(MyText, MyEndPos, 1))
        If Not isAlpha(MyCharacter) Then Exit For
    Next
    
    MyEndPos = MyEndPos - 1
    If MyStartPos <= MyEndPos Then
        startPos = MyStartPos - 1
        WordUnderCursor = Mid$(MyText, MyStartPos, MyEndPos - MyStartPos + 1)
    End If
        
End Function
 

Function isAlpha(c As Integer) As Boolean

    If c >= Asc("a") And c <= Asc("z") Then
        isAlpha = True
        Exit Function
    End If
    
    If c >= Asc("A") And c <= Asc("Z") Then
        isAlpha = True
        Exit Function
    End If
    
    If c >= Asc("0") And c <= Asc("9") Then
        isAlpha = True
        Exit Function
    End If
    
End Function

'remove all old formatting..
Sub ResetRTF(tb As RichTextBox)
    tb.Text = " "
    tb.selStart = 0
    tb.selLength = 1
    tb.SelColor = vbBlack
    tb.SelBold = False
End Sub

Sub HighlightOffsets(tb As RichTextBox, txt As String)
    
    On Error Resume Next
    
    Dim tmp() As String
    Dim x, i As Long
    
    ResetRTF tb
    tb.Text = txt
    tmp() = Split(txt, vbCrLf)
    
    HighLightRunning = True
    LockWindowUpdate tb.hwnd
    
    Dim curPos As Long 'global offset in buffer
    Dim a As Long
    
    For i = 0 To UBound(tmp) 'each line
        x = tmp(i)
        a = InStr(x, " ")
        If a > 0 Then
            tb.selStart = curPos
            tb.selLength = a
            tb.SelColor = vbBlue
        End If
        
        curPos = curPos + Len(x) + 2 'for crlf
    Next
            
    tb.selStart = 0
    tb.selLength = 0
    LockWindowUpdate 0
    HighLightRunning = False
    
End Sub


''based on selstart..
'Function CurrentWord(rtb As RichTextBox) As String
'
'    Dim startAt As Long
'    Dim endAt As Long
'    Dim c As Integer
'
'    On Error Resume Next
'
'    startAt = rtb.selStart
'
'    Do While startAt > 1
'        c = Asc(Mid(rtb.Text, startAt, 1))
'        If isAlpha(c) Then
'            startAt = startAt - 1
'        Else
'            startAt = startAt + 1
'            Exit Do
'        End If
'    Loop
'
'    endAt = rtb.selStart
'
'    Do While endAt < Len(rtb.Text)
'        c = Asc(Mid(rtb.Text, endAt, 1))
'        If isAlpha(c) Then
'            endAt = endAt + 1
'        Else
'            Exit Do
'        End If
'    Loop
'
'    CurrentWord = Mid(rtb.Text, startAt, endAt - startAt)
'
'End Function


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
    SetWindowPos frm.hwnd, s, frm.Left / 15, frm.Top / 15, frm.Width / 15, frm.Height / 15, 0
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

Function IsIde() As Boolean
    On Error GoTo out
    Debug.Print 1 / 0
out: IsIde = Err
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

Function tHex(x As Long) As String
    tHex = Right("00000000" & Hex(x), 8)
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
            ElseIf x = 13 Or x = 10 Or x = 9 Then
                'we will leave \t\r\n
            Else
                b(i) = dot
            End If
       Next
            
End Sub

'note: this implementation was designed for % 16 data inputs..
Function hexdump(ByVal base As Long, it() As Byte) As String
    Dim my, i, c, s, a As Byte, b
    Dim lines() As String
    
    my = ""
    For i = 0 To UBound(it)
        a = it(i)
        c = Hex(a)
        c = IIf(Len(c) = 1, "0" & c, c)
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

