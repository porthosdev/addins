VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   12600
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   14865
   _ExtentX        =   26220
   _ExtentY        =   22225
   _Version        =   393216
   Description     =   "Displays Raw Memory"
   DisplayName     =   "Memory Window"
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

Public FormDisplayed          As Boolean
Public VBInstance             As VBIDE.VBE
Dim mcbMenuCommandBar         As Office.CommandBarControl
Dim mcbMenuCommandBar2         As Office.CommandBarControl
Dim mfrmAddIn                 As New frmAddIn
Public WithEvents MenuHandler As CommandBarEvents          'command bar event handler
Attribute MenuHandler.VB_VarHelpID = -1
Public WithEvents MenuHandler2 As CommandBarEvents          'command bar event handler
Attribute MenuHandler2.VB_VarHelpID = -1

Sub Hide()
    
    On Error Resume Next
    
    FormDisplayed = False
    mfrmAddIn.Hide
   
End Sub

Sub Show()
  
    On Error Resume Next
    
    If mfrmAddIn Is Nothing Then
        Set mfrmAddIn = New frmAddIn
    End If
    
    Set mfrmAddIn.VBInstance = VBInstance
    Set mfrmAddIn.Connect = Me
    FormDisplayed = True
    mfrmAddIn.Show
   
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
    Debug.Print VBInstance.FullName

    If ConnectMode = ext_cm_External Then
        'Used by the wizard toolbar to start this wizard
        Me.Show
    Else
        Set mcbMenuCommandBar = AddToAddInCommandBar("Memory Window")
        Set mcbMenuCommandBar2 = Addit()
        'sink the event
        Set Me.MenuHandler = VBInstance.Events.CommandBarEvents(mcbMenuCommandBar)
        Set Me.MenuHandler2 = VBInstance.Events.CommandBarEvents(mcbMenuCommandBar2)
    End If

    Exit Sub
    
error_handler:
    
    MsgBox "MemoryWindow.OnConnect: " & Err.Description
    
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

End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
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
    
    On Error Resume Next
    
    orgData = Clipboard.GetText
    Clipboard.Clear
    
    Dim ci As Long
    ci = VBInstance.CommandBars.Count
    
    If ci = 0 Then
        MsgBox "No Command Bars?", vbInformation, "MemoryWindow Addin"
        VBInstance.CommandBars.Add
    End If
    
    If ci > 2 Then ci = 2 Else ci = 1
    
    VBInstance.CommandBars(ci).Visible = True
    Set cbMenu = VBInstance.CommandBars(ci).Controls.Add(1, , , VBInstance.CommandBars(2).Controls.Count)
    cbMenu.Caption = "Memory"
    Clipboard.SetData LoadResPicture(101, 0)
    cbMenu.PasteFace
    Set Addit = cbMenu
    
    Clipboard.Clear
    If Len(orgData) > 0 Then Clipboard.SetText orgData
End Function

