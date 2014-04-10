Attribute VB_Name = "Module2"
Public LastCommandOutput As String
Public VBInstance As VBIDE.VBE
Public Connect As Connect
Public ClearImmediateOnStart As Long

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

