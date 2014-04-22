Attribute VB_Name = "Module2"
Public LastCommandOutput As String
Public VBInstance As VBIDE.VBE
Public Connect As Connect
Public ClearImmediateOnStart As Long
Public ShowPostBuildOutput As Long

Function ExpandVars(ByVal cmd As String, exeFullPath As String) As String
    Dim appDir As String
    Dim fName As String
    
    cmd = Trim(cmd)
    appDir = GetParentFolder(exeFullPath)
    fName = FileNameFromPath(exeFullPath)
    
    ExpandVars = Replace(cmd, "%1", exeFullPath)
    ExpandVars = Replace(ExpandVars, "%app", appDir, , , vbTextCompare)
    ExpandVars = Replace(ExpandVars, "%fname", fName, , , vbTextCompare)
    
End Function

Function isBuildPathSet() As Boolean

    On Error Resume Next
    Dim fastBuildPath As String
    
    If VBInstance.ActiveVBProject Is Nothing Then Exit Function
    fastBuildPath = VBInstance.ActiveVBProject.ReadProperty("fastBuild", "fullPath")
    If Len(fastBuildPath) = 0 Then Exit Function
    isBuildPathSet = True
    
End Function

'set the current directory to be parent folder as vbp folder path...
Sub SetHomeDir()
    On Error Resume Next
    Dim homeDir As String
    homeDir = VBInstance.ActiveVBProject.FileName 'path to vbp file
    homeDir = GetParentFolder(homeDir)
    If Len(homeDir) > 0 Then ChDir homeDir
End Sub

Function GetPostBuildCommand() As String
    On Error Resume Next
    If VBInstance.ActiveVBProject Is Nothing Then Exit Function
    GetPostBuildCommand = VBInstance.ActiveVBProject.ReadProperty("fastBuild", "PostBuild")
End Function

Function IsIde() As Boolean
    On Error GoTo out
    Debug.Print 1 / 0
out: IsIde = Err
End Function

Function FileExists(path) As Boolean
  On Error Resume Next
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
  If Err.Number <> 0 Then FileExists = False
End Function

Function FolderExists(path) As Boolean
  On Error Resume Next
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbDirectory) <> "" Then FolderExists = True
  If Err.Number <> 0 Then FolderExists = False
End Function

Function GetParentFolder(path) As String
    On Error Resume Next
    Dim tmp() As String
    Dim ub As String
    If Len(path) = 0 Then Exit Function
    If InStr(path, "\") < 1 Then Exit Function
    If Right(path, 1) = "\" Then path = Mid(path, 1, Len(path) - 1)
    tmp = Split(path, "\")
    ub = tmp(UBound(tmp))
    GetParentFolder = Replace(Join(tmp, "\"), "\" & ub, "")
    If Err.Number <> 0 Then GetParentFolder = Empty
End Function

Function FileNameFromPath(fullpath) As String
    If InStr(fullpath, "\") > 0 Then
        tmp = Split(fullpath, "\")
        FileNameFromPath = CStr(tmp(UBound(tmp)))
    Else
        FileNameFromPath = fullpath
    End If
End Function

Function GetFileReport(fpath As String) As String
    On Error Resume Next
    
    Dim MyStamp As Date
    Dim ret() As String
    
    If Not FileExists(fpath) Then
        GetFileReport = "Build Failed: " & fpath
        Exit Function
    End If
    
    MyStamp = FileDateTime(fpath)
    
    push ret, "Output File: " & fpath & "  (" & FileSize(fpath) & ")"
    push ret, "Last Modified: " & Format(MyStamp, "dddd, mmmm dd, yyyy - h:mm:ss AM/PM")
    
    GetFileReport = Join(ret, vbCrLf)

End Function

Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init: ReDim ary(0): ary(0) = value
End Sub

Public Function FileSize(fpath As String) As String
    Dim fsize As Long
    Dim szName As String
    On Error GoTo hell
    
    fsize = FileLen(fpath)
    
    szName = " bytes"
    If fsize > 1024 Then
        fsize = fsize / 1024
        szName = " Kb"
    End If
    
    If fsize > 1024 Then
        fsize = fsize / 1024
        szName = " Mb"
    End If
    
    FileSize = fsize & szName
    
    Exit Function
hell:
    
End Function

