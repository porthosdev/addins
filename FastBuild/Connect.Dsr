VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   11325
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   17535
   _ExtentX        =   30930
   _ExtentY        =   19976
   _Version        =   393216
   Description     =   "Streamline Build Process"
   DisplayName     =   "Fast Build"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'this used to use the API Hooking code in Module1.bas, until i learned
'vb actually already exposed the necessary events as part of the addin model..
'oops there goes a solid days labor..

Private FormDisplayed           As Boolean
Private VBInstance              As VBIDE.VBE
Dim mfrmAddIn                   As New frmAddIn

Dim mcbFastBuildUI                As Office.CommandBarControl
Private WithEvents mnuFastBuildUI As CommandBarEvents
Attribute mnuFastBuildUI.VB_VarHelpID = -1

'Dim mcbFastBuild                As Office.CommandBarControl
'Private WithEvents mnuFastBuild As CommandBarEvents

Private WithEvents FileEvents As VBIDE.FileControlEvents
Attribute FileEvents.VB_VarHelpID = -1

Dim mcbExecute                As Office.CommandBarControl
Private WithEvents mnuExecute As CommandBarEvents
Attribute mnuExecute.VB_VarHelpID = -1

Dim mcbAddref                As Office.CommandBarControl
Private WithEvents mnuAddref As CommandBarEvents
Attribute mnuAddref.VB_VarHelpID = -1


Sub Hide()
    
    On Error Resume Next
    
    FormDisplayed = False
    mfrmAddIn.Hide
   
End Sub

Sub Show()
  
    On Error Resume Next
    Dim needsRefresh As Boolean
    
    If mfrmAddIn Is Nothing Then
        Set mfrmAddIn = New frmAddIn
    Else
        needsRefresh = True
    End If
    
    If Not VBInstance.ActiveVBProject Is Nothing Then
        Debug.Print "OnConnect Project: " & VBInstance.ActiveVBProject.FileName
    Else
        Debug.Print "VBInstance.ActiveVBProject is nothing "
    End If
    
    FormDisplayed = True
    mfrmAddIn.Show
    
    'If needsRefresh Then mfrmAddIn.cmdRefresh_Click
    
End Sub

'------------------------------------------------------
'this method adds the Add-In to VB
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error GoTo error_handler
    
    'save the vb instance
    Set VBInstance = Application
    
    'this is a good place to set a breakpoint and
    'test various addin objects, properties and methods
    Debug.Print "FullName: " & VBInstance.FullName
     
    If ConnectMode = ext_cm_External Then
        'Used by the wizard toolbar to start this wizard
        Me.Show
    Else
        Set mcbFastBuildUI = AddButton("Fast Build", 101) 'AddToAddInCommandBar("Fast Build")
        Set mnuFastBuildUI = VBInstance.Events.CommandBarEvents(mcbFastBuildUI)
        
        'Set mcbFastBuild = AddButton("Compile", 103)
        'Set mnuFastBuild = VBInstance.Events.CommandBarEvents(mcbFastBuild)
        
        Set mcbExecute = AddButton("Execute", 102)
        Set mnuExecute = VBInstance.Events.CommandBarEvents(mcbExecute)
 
        Set mcbAddref = AddrefMenu("Quick AddRef")
        Set mnuAddref = VBInstance.Events.CommandBarEvents(mcbAddref)
        
        Set FileEvents = Application.Events.FileControlEvents(Nothing)
        
        Set Module2.VBInstance = VBInstance
        Set Module2.Connect = Me
        
        'If Module1.IsIde() Then
        '    Debug.Print "YOU CAN NOT DEBUG THE HOOK CODE IN THE IDE" 'it hooks the debugger instance not the remote one..
        'Else
        '    'SetHook 'detours style hook of GetSaveFileName
        'End If
        
    End If

    Exit Sub
    
error_handler:
    
    MsgBox Err.Description
    
End Sub

'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    
    'mcbFastBuild.Delete
    mcbFastBuildUI.Delete
    mcbExecute.Delete
    mcbAddref.Delete
    
    Unload frmAddRefs
    Unload mfrmAddIn
    Set mfrmAddIn = Nothing

    'release all references so object can shut down and remove itself..
    'otherwise you wont be able to unload and compile, you will have to restart ide
    Set Module2.VBInstance = Nothing
    Set Module2.Connect = Nothing
    Set VBInstance = Nothing
    
    'RemoveAllHooks
    'UnInitilizeHookLib
    'FreeLibrary hHookLib
    
End Sub

Private Sub FileEvents_AfterWriteFile(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, ByVal FileName As String, ByVal result As Integer)
        
    If FileType <> vbext_ft_Exe Then Exit Sub
       
    If Not isBuildPathSet() Then
        VBInstance.ActiveVBProject.WriteProperty "fastBuild", "fullPath", FileName
    End If
    
    Dim postbuild As String
    postbuild = GetPostBuildCommand()
    
    If Len(postbuild) > 0 Then
        SetHomeDir
        postbuild = ExpandVars(postbuild, FileName)
        LastCommandOutput = GetCommandOutput("cmd /c " & postbuild, True, True)
    End If
    
    
End Sub

Private Sub FileEvents_DoGetNewFileName(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, NewName As String, ByVal OldName As String, CancelDefault As Boolean)
    Dim fastBuildPath As String
    
    If FileType <> vbext_ft_Exe Then
        'MsgBox "Filetype: " & FileType
        Exit Sub
    End If
    
    If Not isBuildPathSet() Then
        'MsgBox "Build path not set"
        Exit Sub
    End If
     
    fastBuildPath = VBInstance.ActiveVBProject.ReadProperty("fastBuild", "fullPath")
    If Len(fastBuildPath) = 0 Then
        'MsgBox "fast build path empty?"
        Exit Sub
    End If
    
    'MsgBox "overriding path!"
    NewName = fastBuildPath
    OldName = fastBuildPath
    CancelDefault = True
 
End Sub

Private Sub mnuAddref_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    frmAddRefs.Show
End Sub

Private Sub mnuFastBuildUI_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Me.Show
End Sub

'I am removing this method..it has bugs in how MakeCompiledFile is implemented..
'if the path you specify in BuildFileName is not valid, then it will fail without error
'I could work around it, but its better to manually add a Build tool bar button from the command bar editor.
'
'Private Sub mnuFastBuild_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
'    On Error Resume Next
'
'    If isBuildPathSet() Then
'        VBInstance.ActiveVBProject.BuildFileName = VBInstance.ActiveVBProject.ReadProperty("fastBuild", "fullPath")
'    End If
'
'    'apparently calling this method manually like this just uses the default and skips DoGetNewFileName hooks..
'    VBInstance.ActiveVBProject.MakeCompiledFile
'
'End Sub

Private Sub mnuExecute_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    Dim fastBuildPath As String
     
    If Not isBuildPathSet() Then
        MsgBox "Can not launch the executable, path not yet set", vbInformation
        Exit Sub
    End If
    
    fastBuildPath = VBInstance.ActiveVBProject.ReadProperty("fastBuild", "fullPath")
    
    If Not FileExists(fastBuildPath) Then
        MsgBox "File not found: " & fastBuildPath, vbInformation
        Exit Sub
    End If
    
    Shell fastBuildPath, vbNormalFocus
    
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbExclamation
    End If
    
End Sub


Private Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl
    Dim cbMenuCommandBar As Office.CommandBarControl
    Dim cbMenu As Object

    On Error GoTo hell

    Set cbMenu = VBInstance.CommandBars("Add-Ins")
    If cbMenu Is Nothing Then Exit Function

    Set cbMenuCommandBar = cbMenu.Controls.Add(1)
    cbMenuCommandBar.caption = sCaption
    Set AddToAddInCommandBar = cbMenuCommandBar

    Exit Function
hell:
End Function

Private Function AddButton(caption As String, resImg As Long) As Office.CommandBarControl
    Dim cbMenu As Object
    Dim orgData As String
    
    orgData = Clipboard.GetText
    
    VBInstance.CommandBars(2).Visible = True
    Set cbMenu = VBInstance.CommandBars(2).Controls.Add(1, , , VBInstance.CommandBars(2).Controls.Count)
    cbMenu.caption = caption
    Clipboard.SetData LoadResPicture(resImg, 0)
    cbMenu.PasteFace
    Set AddButton = cbMenu
    
    Clipboard.Clear
    Clipboard.SetText orgData
End Function

Private Function AddrefMenu(caption As String) As Office.CommandBarControl

    Dim cbProjMenu As Office.CommandBarControl
    Dim cbSubMenu As Office.CommandBarControl
    Dim i As Long
    
    On Error GoTo hell

    Set cbProjMenu = VBInstance.CommandBars(1).Controls("Project")   'menu bar is always first command bar
    
    If cbProjMenu Is Nothing Then Exit Function

    For Each cbSubMenu In cbProjMenu.Controls
        i = i + 1
        If cbSubMenu.caption = "Refere&nces..." Then Exit For
    Next
    If i = cbProjMenu.Controls.Count Then Exit Function

    Set AddrefMenu = cbProjMenu.Controls.Add(, , , i + 1) 'add the menu before the References ... menu
    AddrefMenu.caption = caption

hell:

End Function

