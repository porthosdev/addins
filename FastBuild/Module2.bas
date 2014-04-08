Attribute VB_Name = "Module2"
Global LastCommandOutput As String
Public VBInstance As VBIDE.VBE
Public Connect As Connect

Function ExpandVars(ByVal cmd As String, exeFullPath As String) As String
    Dim appDir As String
    Dim fName As String
    
    cmd = Trim(cmd)
    appDir = GetParentFolder(exeFullPath)
    fName = FileNameFromPath(exeFullPath)
    
    ExpandVars = Replace(cmd, "%1", exeFullPath)
    ExpandVars = Replace(ExpandVars, "%apppath", appDir, , , vbTextCompare)
    ExpandVars = Replace(ExpandVars, "%outname", fName, , , vbTextCompare)
    'ExpandVars = Replace(ExpandVars, "%vb", VB6FOLDER, , , vbTextCompare)
    
End Function

Function isBuildPathSet() As Boolean

    On Error Resume Next
    Dim fastBuildPath As String
    
    If VBInstance.ActiveVBProject Is Nothing Then Exit Function
    fastBuildPath = VBInstance.ActiveVBProject.ReadProperty("fastBuild", "fullPath")
    If Len(fastBuildPath) = 0 Then Exit Function
    isBuildPathSet = True
    
End Function

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
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True _
  Else FileExists = False
End Function

Function FolderExists(path) As Boolean
  On Error Resume Next
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbDirectory) <> "" Then FolderExists = True _
  Else FolderExists = False
End Function

Function GetParentFolder(path) As String
    On Error Resume Next
    Dim tmp() As String
    Dim ub As Long
    tmp = Split(path, "\")
    ub = tmp(UBound(tmp))
    GetParentFolder = Replace(Join(tmp, "\"), "\" & ub, "")
End Function

Function FileNameFromPath(fullpath) As String
    If InStr(fullpath, "\") > 0 Then
        tmp = Split(fullpath, "\")
        FileNameFromPath = CStr(tmp(UBound(tmp)))
    Else
        FileNameFromPath = fullpath
    End If
End Function

