Attribute VB_Name = "Module1"
' so after a solid day of hard labor hooking the api and working out all the bugs..i find
' vb6 has a built in event that fires to prompt for a file name and can override it.
' i never imagined they would have a hook for that built in..sooo, anyway I will include this
' module, even though it isnt used, but it might come in handy someday and I am not going to
' lose the work..
'
' live and learn!

'todo:
'      detect if project home directory has been moved and alert?

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
Private Declare Function HookFunction Lib "hooklib.dll" (ByVal lpOrgFunc As Long, ByVal lpNewFunc As Long, ByVal name As String, ByVal ht As hookType) As Long

'char* __stdcall GetHookError(void)
Public Declare Function GetHookError Lib "hooklib.dll" () As Long

'void __stdcall SetDebugHandler(ULONG_PTR lpfn); --> callback prototype: void  (__stdcall *debugMsgHandler)(char* msg);
Private Declare Sub SetDebugHandler Lib "hooklib.dll" (ByVal lpCallBack As Long, Optional ByVal logLevel As Long = 0)

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

Public VBInstance As VBIDE.VBE
Public Connect As Connect

Public hookLog As New Collection

Function SetHook() As Boolean

    Dim h As Long
    Dim ret As Long
    Dim lpMsg As Long
    
    hHookLib = LoadLibrary("hooklib.dll")
    If hHookLib = 0 Then hHookLib = LoadLibrary(App.path & "\hooklib.dll")
    If hHookLib = 0 Then hHookLib = LoadLibrary(App.path & "\..\hooklib.dll")
    
    If hHookLib = 0 Then
        hookLog.Add "Could not find hooklib.dll compile or download from github."
        Exit Function
    End If
    
    hookLog.Add "Hooklib base address: 0x" & Hex(hHookLib)
    
    'this is optional but were debugging the library so..
    'SetDebugHandler AddressOf DebugMsgHandler, 1
    
    h = LoadLibrary("comdlg32.dll")
    lpGetSaveFileName = GetProcAddress(h, "GetSaveFileNameA")
    
    If lpGetSaveFileName = 0 Then
        hookLog.Add "Could not GetProcAddress(GetSaveFileNameA)"
        Exit Function
    End If
        
    ret = HookFunction(lpGetSaveFileName, AddressOf My_GetSaveFileName, "GetSaveFileNameA", ht_jmp)
    
    'ret = HookFunction(0, 0, "GetSaveFileNameA", ht_jmp) 'to test GetHookError()
    If ret = 0 Then
        lpMsg = GetHookError()
        hookLog.Add "Hook Function failed msg: " & CStringToVBString(lpMsg)
        Exit Function
    End If
    
    hookLog.Add "GetSaveFileNameA Successfully Hooked.."
    
End Function
    
    
'void  (__stdcall *debugMsgHandler)(char* msg);
Public Sub DebugMsgHandler(ByVal lpMsg As Long)

    Dim msg As String
    Dim tmp() As String
    
    msg = CStringToVBString(lpMsg)
    
    If InStr(msg, vbTab) > 0 Then msg = Replace(msg, vbTab, "    ")
    
    If InStr(msg, vbLf) Then
        msg = Replace(msg, vbCr, Empty)
        tmp() = Split(msg, vbLf)
        For Each X In tmp
            If Len(Trim(X)) > 0 Then hookLog.Add "DebugMsg: " & X
        Next
    Else
        hookLog.Add "DebugMsg: " & msg
    End If
    
End Sub


'were mostly operating as a passthrough, so we have converted arguments to as long
'so VB6 doesnt hose with anything by doing any unexpected automatic conversions
'
'http://msdn.microsoft.com/en-us/library/windows/desktop/ms646839%28v=vs.85%29.aspx
'  lpstrFile, nMaxFile, lpstrInitialDir (can be null see link for defaults)
'  Flags: OFN_OVERWRITEPROMPT 0x00000002
'

'you can not debug this in the IDE, because it will hook its own instance of the IDE and not the target..
Public Function My_GetSaveFileName(ByRef ofn As lng_OPENFILENAME) As Long
    Dim ext As String
    Dim outputName As String
    Dim fastBuildPath As String
    Dim b() As Byte
    Dim lastSlash As Long
    
    On Error Resume Next
    
    hookLog.Add "Inside My_GetSaveFileName lpOfn = 0x" & VarPtr(ofn)
    
    If ofn.lpstrFile <> 0 Then
1        outputName = CStringToVBString(ofn.lpstrFile)
2        hookLog.Add "Checking outputname: " & outputName
3        If Len(outputName) > 3 Then ext = LCase(Right(outputName, 3))
    End If

4    If isBuildPathSet() Then
    
        If ext = "exe" Or ext = "dll" Or ext = "ocx" Then 'were interested
        
5            fastBuildPath = VBInstance.ActiveVBProject.ReadProperty("fastBuild", "fullPath")
            If Len(fastBuildPath) > 0 Then
                hookLog.Add "Automatically setting build path: " & fastBuildPath
                fastBuildPath = fastBuildPath & Chr(0) & Chr(0)
6                b() = StrConv(fastBuildPath, vbFromUnicode, LANG_US)
7                CopyMemory ByVal ofn.lpstrFile, ByVal VarPtr(b(0)), UBound(b)
                My_GetSaveFileName = 1
8                lastSlash = InStrRev(fastBuildPath, "\")
                If lastSlash > 0 Then ofn.nFileOffset = lastSlash
                Exit Function
            End If
            
        End If
        
    End If
    
    'if we make it down to here..they we should proceed as normal
    hookLog.Add "Calling real GetSaveFileName 0x" & Hex(lpGetSaveFileName)
    
    'call the real windows api and save dialog..
9    My_GetSaveFileName = CallOriginal(lpGetSaveFileName, VarPtr(ofn))
    
    If Len(ext) = 0 Or (ext = "exe" Or ext = "dll" Or ext = "ocx") Then
        'user didnt hit cancel and pointer is ok..
        If My_GetSaveFileName = 1 And ofn.lpstrFile <> 0 Then
10            outputName = CStringToVBString(ofn.lpstrFile)
11            hookLog.Add "Checking outputname: " & outputName
12            If Len(outputName) > 3 Then ext = LCase(Right(outputName, 3))
        
            If ext = "exe" Or ext = "dll" Or ext = "ocx" Then 'were interested
                'this must be the first time they set a full compile path..we will save it..
13                VBInstance.ActiveVBProject.WriteProperty "fastBuild", "fullPath", outputName
            End If
        End If
    End If
             
    If Err.Number <> 0 Then
        hookLog.Add "Had error in hook line: " & Erl & "  desc: " & Err.Description
    End If
    
End Function

'checks if full path has been set yet, can set it if it has..
'this way we dont have to specify path every damn time, and
'the lastDefDir wont get mixed up sometimes like it does..
Function isBuildPathSet() As Boolean

    On Error Resume Next
    Dim fastBuildPath As String
    Dim defPath As String
    
    If VBInstance.ActiveVBProject Is Nothing Then Exit Function
    
    'going to leave the auto set disabled, so user can specify the file path manually the first time..
    'defPath = VBInstance.ActiveVBProject.BuildFileName
    
    fastBuildPath = VBInstance.ActiveVBProject.ReadProperty("fastBuild", "fullPath")
    
    If Len(fastBuildPath) = 0 Then
        'If Len(defPath) > 0 Then
        '    If InStr(defPath, "\") > 0 Then 'just project1.exe for unsaved we wont record that yet..
        '        VBInstance.ActiveVBProject.WriteProperty "fastBuild", "fullPath", defPath
        '        isBuildPathSet = True
        '    End If
        'End If
        Exit Function
    End If
    
    isBuildPathSet = True
    
End Function

Function GetPostBuildCommand() As String
    On Error Resume Next
    If VBInstance.ActiveVBProject Is Nothing Then Exit Function
    GetPostBuildCommand = VBInstance.ActiveVBProject.ReadProperty("fastBuild", "PostBuild")
End Function

Function CStringToVBString(lpCstr As Long) As String

    Dim X As Long
    Dim sBuffer As String
    Dim lpBuffer As Long
    Dim b() As Byte
    
    If lpCstr <> 0 Then
        X = lstrlen(lpCstr)
        If X > 0 Then
            ReDim b(X)
            CopyMemory ByVal VarPtr(b(0)), ByVal lpCstr, X
            CStringToVBString = StrConv(b, vbUnicode, LANG_US)
        End If
    End If
    
    CStringToVBString = Replace(CStringToVBString, Chr(0), Empty) 'just in case..
    
End Function

Function IsIde() As Boolean
    On Error GoTo out
    Debug.Print 1 / 0
out: IsIde = Err
End Function

Function FileExists(path) As Boolean
  On Error Resume Next
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True _
  Else FileExists = False
End Function

Function FolderExists(path) As Boolean
  On Error Resume Next
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbDirectory) <> "" Then FolderExists = True _
  Else FolderExists = False
End Function

