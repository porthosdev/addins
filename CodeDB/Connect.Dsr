VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   11130
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   18450
   _ExtentX        =   32544
   _ExtentY        =   19632
   _Version        =   393216
   Description     =   "Misc Code Tools"
   DisplayName     =   "CodeDB"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 98 (ver 6.0)"
   LoadName        =   "Command Line / Startup"
   LoadBehavior    =   5
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public FormDisplayed          As Boolean
Public VBInstance             As VBIDE.VBE
Dim mcbMenuCommandBar         As Office.CommandBarControl
Dim mfrmAddIn                 As New frmAddIn
Public WithEvents MenuHandler As CommandBarEvents          'command bar event handler
Attribute MenuHandler.VB_VarHelpID = -1


Sub Hide()
    On Error Resume Next
    FormDisplayed = False
    mfrmAddIn.Hide
End Sub

Sub Show()
    On Error Resume Next
    If mfrmAddIn Is Nothing Then Set mfrmAddIn = New frmAddIn
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
    
    Set VBInstance = Application
    If ConnectMode = ext_cm_External Then 'wizard toolbar to start this wizard
        Me.Show
    Else
        Set mcbMenuCommandBar = Addit()
        Set Me.MenuHandler = VBInstance.Events.CommandBarEvents(mcbMenuCommandBar)
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
    mcbMenuCommandBar.Delete
    Unload mfrmAddIn
    Set mfrmAddIn = Nothing
End Sub

'this event fires when the menu is clicked in the IDE
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Me.Show
End Sub

Private Function Addit() As Office.CommandBarControl
    Dim cbMenu As Object
    Dim orgData As String
    
    orgData = Clipboard.GetText
    
    VBInstance.CommandBars(2).Visible = True
    Set cbMenu = VBInstance.CommandBars(2).Controls.Add(1, , , VBInstance.CommandBars(2).Controls.Count)
    cbMenu.Caption = "Code db"
    Clipboard.SetData LoadResPicture(102, 0)
    cbMenu.PasteFace
    Set Addit = cbMenu
    
    Clipboard.Clear
    Clipboard.SetText orgData
End Function
