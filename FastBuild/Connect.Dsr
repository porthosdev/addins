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

Public FormDisplayed          As Boolean
Public VBInstance             As VBIDE.VBE
Dim mcbMenuCommandBar         As Office.CommandBarControl
Dim mcbMenuCommandBar2         As Office.CommandBarControl
Dim mfrmAddIn                 As New frmAddIn
Public WithEvents MenuHandler As CommandBarEvents          'command bar event handler
Attribute MenuHandler.VB_VarHelpID = -1
Public WithEvents MenuHandler2 As CommandBarEvents          'command bar event handler
Attribute MenuHandler2.VB_VarHelpID = -1
Public WithEvents FileEvents As VBIDE.FileControlEvents
Attribute FileEvents.VB_VarHelpID = -1

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
    
    '?VBInstance.ActiveVBProject.BuildFileName
    'C:\Documents and Settings\david\Desktop\test\Project1.exe
    
    '?VBInstance.ActiveVBProject.FileName
    'C:\Documents and Settings\david\Desktop\test\Project1.vbp
    
    '?VBInstance.ActiveVBProject.ReadProperty("tacobell","blah")
    'if key doesnt exist it will throw error..saved to vbp file..
    
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
        Set mcbMenuCommandBar = AddToAddInCommandBar("Fast Build")
        Set mcbMenuCommandBar2 = Addit()
        'sink the event
        Set Me.MenuHandler = VBInstance.Events.CommandBarEvents(mcbMenuCommandBar)
        Set Me.MenuHandler2 = VBInstance.Events.CommandBarEvents(mcbMenuCommandBar2)
        Set Me.FileEvents = Application.Events.FileControlEvents(Nothing)
        
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
    
    'delete the command bar entry
    mcbMenuCommandBar.Delete
    mcbMenuCommandBar2.Delete
    Unload mfrmAddIn
    Set mfrmAddIn = Nothing

    'RemoveAllHooks
    'UnInitilizeHookLib
    'FreeLibrary hHookLib
    
End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
End Sub


Private Sub FileEvents_AfterWriteFile(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, ByVal FileName As String, ByVal Result As Integer)
        
    If FileType <> vbext_ft_Exe Then Exit Sub
       
    If Not isBuildPathSet() Then
        VBInstance.ActiveVBProject.WriteProperty "fastBuild", "fullPath", FileName
    End If
    
    Dim postbuild As String
    postbuild = GetPostBuildCommand()
    
    If Len(postbuild) > 0 Then
        postbuild = ExpandVars(postbuild, FileName)
        LastCommandOutput = RunCommand(postbuild)
    End If
    
    
End Sub

Private Sub FileEvents_DoGetNewFileName(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, NewName As String, ByVal OldName As String, CancelDefault As Boolean)
    Dim fastBuildPath As String
    
    If FileType <> vbext_ft_Exe Then Exit Sub  'not interested...
    
    If Not isBuildPathSet() Then Exit Sub
     
    fastBuildPath = VBInstance.ActiveVBProject.ReadProperty("fastBuild", "fullPath")
    If Len(fastBuildPath) > 0 Then
        NewName = fastBuildPath
        CancelDefault = True
    End If
 
End Sub

Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Me.Show
End Sub

Private Sub MenuHandler2_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Me.Show
End Sub


Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl
    Dim cbMenuCommandBar As Office.CommandBarControl  'command bar object
    Dim cbMenu As Object

    On Error GoTo AddToAddInCommandBarErr

    'see if we can find the Add-Ins menu
    Set cbMenu = VBInstance.CommandBars("Add-Ins")
    If cbMenu Is Nothing Then
        'not available so we fail
        Exit Function
    End If

    'add it to the command bar
    Set cbMenuCommandBar = cbMenu.Controls.Add(1)
    'set the caption
    cbMenuCommandBar.Caption = sCaption

    Set AddToAddInCommandBar = cbMenuCommandBar

    Exit Function

AddToAddInCommandBarErr:

End Function

Private Function Addit() As Office.CommandBarControl
    Dim cbMenu As Object
    Dim orgData As String
    
    orgData = Clipboard.GetText
    
    VBInstance.CommandBars(2).Visible = True
    Set cbMenu = VBInstance.CommandBars(2).Controls.Add(1, , , VBInstance.CommandBars(2).Controls.Count)
    cbMenu.Caption = "Fast Build"
    Clipboard.SetData LoadResPicture(101, 0)
    cbMenu.PasteFace
    Set Addit = cbMenu
    
    Clipboard.Clear
    Clipboard.SetText orgData
End Function

