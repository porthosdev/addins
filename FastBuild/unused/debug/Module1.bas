Attribute VB_Name = "Module1"
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Declare Sub DebugBreak Lib "kernel32" ()

Enum hookType
    ht_jmp = 0
    ht_pushret = 1
    ht_jmp5safe = 2
    ht_jmpderef = 3
    ht_micro = 4
End Enum

Enum hookErrors
    he_None = 0
    he_cantDisasm
    he_cantHook
    he_maxHooks
    he_UnknownHookType
    he_Other
End Enum

'BOOL __stdcall HookFunction(ULONG_PTR OriginalFunction, ULONG_PTR NewFunction, char *name, hookType ht)
Public Declare Function HookFunction Lib "hooklib.dll" (ByVal lpOrgFunc As Long, ByVal lpNewFunc As Long, ByVal name As String, ByVal ht As hookType) As Long

'char* __stdcall GetHookError(void)
Public Declare Function GetHookError Lib "hooklib.dll" () As Long

'void __stdcall SetDebugHandler(ULONG_PTR lpfn); --> callback prototype: void  (__stdcall *debugMsgHandler)(char* msg);
Public Declare Sub SetDebugHandler Lib "hooklib.dll" (ByVal lpCallBack As Long, Optional ByVal logLevel As Long = 0)

'VOID __stdcall DisableHook(ULONG_PTR Function)
Public Declare Function DisableHook Lib "hooklib.dll" (ByVal lpOrgFunc As Long) As Long

'VOID __stdcall EnableHook(ULONG_PTR Function)
Public Declare Function EnableHook Lib "hooklib.dll" (ByVal lpOrgFunc As Long) As Long

'VOID __stdcall RemoveHook(ULONG_PTR Function)
Public Declare Function RemoveHook Lib "hooklib.dll" (ByVal lpOrgFunc As Long) As Long

'int __stdcall CallOriginal(int orgFunc, int arg1)
Public Declare Function CallOriginal Lib "hooklib.dll" (ByVal lpOrgFunc As Long, ByVal arg1 As Long) As Long

Public Declare Sub RemoveAllHooks Lib "hooklib.dll" ()
Public Declare Sub UnInitilizeHookLib Lib "hooklib.dll" Alias "UnInitilize" ()

'these are the modified declares that we will use in our hook proc
'so vb doesnt do any string translations..

Private Type lng_OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As Long           'String
        lpstrCustomFilter As Long     'String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As Long             'String
        nMaxFile As Long
        lpstrFileTitle As Long        'String
        nMaxFileTitle As Long
        lpstrInitialDir As Long       'String
        lpstrTitle As Long            'String
        flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As Long           'String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As Long        'String
End Type


Const LANG_US = &H409
Global hHookLib As Long
Global lpGetSaveFileName As Long

'void  (__stdcall *debugMsgHandler)(char* msg);
Public Sub DebugMsgHandler(ByVal lpMsg As Long)

    Dim msg As String
    Dim tmp() As String
    
    msg = CStringToVBString(lpMsg)
    
    If InStr(msg, vbTab) > 0 Then msg = Replace(msg, vbTab, "    ")
    
    If InStr(msg, vbLf) Then
        msg = Replace(msg, vbCr, Empty)
        tmp() = Split(msg, vbLf)
        For Each x In tmp
            If Len(Trim(x)) > 0 Then Form1.List1.AddItem "DebugMsg: " & x
        Next
    Else
        Form1.List1.AddItem "DebugMsg: " & msg
    End If
    
End Sub


'were mostly operating as a passthrough, so we have converted arguments to as long
'so VB6 doesnt hose with anything by doing any unexpected automatic conversions
'
'http://msdn.microsoft.com/en-us/library/windows/desktop/ms646839%28v=vs.85%29.aspx
'  lpstrFile, nMaxFile, lpstrInitialDir (can be null see link for defaults)
'  Flags: OFN_OVERWRITEPROMPT 0x00000002
'

Public Function My_GetSaveFileName(ByRef ofn As lng_OPENFILENAME) As Long
    Dim ext As String
    Dim outputName As String
    
    Form1.List1.AddItem "Inside My_GetSaveFileName lpOFN = 0x" & VarPtr(lpOFN)
    
    'make our structure reference the actual one passed in..
    'CopyMemory ByVal VarPtr(ofn), ByVal lpOFN, 4
    
    If ofn.lpstrFile <> 0 Then
        outputName = CStringToVBString(ofn.lpstrFile)
        If Len(outputName) > 3 Then ext = LCase(Right(outputName, 3))
        If ext = "exe" Or ext = "dll" Or ext = "ocx" Then 'were interested
            MsgBox "Todo: lookup default name, and output dir from vbp file and just return it.."
            'todo: set fields in ofn to match
            'Exit Function
        End If
    End If
    
    'if we make it down to here..they we should proceed as normal
    Form1.List1.AddItem "Calling real GetSaveFileName"
    My_GetSaveFileName = CallOriginal(lpGetSaveFileName, lpOFN)
    
End Function


Function CStringToVBString(lpCstr As Long) As String

    Dim x As Long
    Dim sBuffer As String
    Dim lpBuffer As Long
    Dim b() As Byte
    
    If lpCstr <> 0 Then
        x = lstrlen(lpCstr)
        If x > 0 Then
            ReDim b(x)
            CopyMemory ByVal VarPtr(b(0)), ByVal lpCstr, x
            CStringToVBString = StrConv(b, vbUnicode, LANG_US)
        End If
    End If
    
    CStringToVBString = Replace(CStringToVBString, Chr(0), Empty) 'just in case..
    
End Function
