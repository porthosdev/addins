VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   10605
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   23370
   _ExtentX        =   41222
   _ExtentY        =   18706
   _Version        =   393216
   Description     =   $"Connect.dsx":0000
   DisplayName     =   "CodeView"
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

Public FormDisplayed As Boolean
Dim mcbMenuCommandBar As Office.CommandBarControl
Public WithEvents MenuHandler As CommandBarEvents
Attribute MenuHandler.VB_VarHelpID = -1

Dim mcbMenuCommandBar2 As Office.CommandBarControl
Public WithEvents MenuHandler2 As CommandBarEvents
Attribute MenuHandler2.VB_VarHelpID = -1

Public WithEvents ComponentHandler As VBComponentsEvents
Attribute ComponentHandler.VB_VarHelpID = -1
Public WithEvents ProjectHandler As VBProjectsEvents
Attribute ProjectHandler.VB_VarHelpID = -1

Dim mToolCodeView As ToolCodeView
Dim wToolCodeView As VBIDE.Window
Const GuidCodeView As String = "05745B8A-E341-11E3-9712-70581D5D46B0"

Sub Hide()
    wToolCodeView.Visible = False
    FormDisplayed = False
End Sub

Sub Show()
    wToolCodeView.Visible = True
    FormDisplayed = True
End Sub

Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error GoTo error_handler
    
1    Set g_VBInstance = Application
2    Set mcbMenuCommandBar = AddToAddInCommandBar("CodeView - Source Navigation")
3    Set MenuHandler = g_VBInstance.Events.CommandBarEvents(mcbMenuCommandBar)

9    Set mcbMenuCommandBar2 = AddToAddInCommandBar("Find-All")
10   Set MenuHandler2 = g_VBInstance.Events.CommandBarEvents(mcbMenuCommandBar2)

4    Set ComponentHandler = g_VBInstance.Events.VBComponentsEvents(Nothing)
5    Set ProjectHandler = g_VBInstance.Events.VBProjectsEvents()

6    Set wToolCodeView = g_VBInstance.Windows.CreateToolWindow(AddInInst, "CodeView.ToolCodeView", "CodeView", GuidCodeView, mToolCodeView)
8    Me.Show
     
    Exit Sub
error_handler:
    MsgBox "AddinInstance_OnConnection: Line: " & Erl() & " Desc: " & Err.Description
End Sub

Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    
    mcbMenuCommandBar.Delete
    mcbMenuCommandBar2.Delete
    FormDisplayed = False
    Unload mToolCodeView
    Unload frmFindAll
    Set mToolCodeView = Nothing
    Set g_VBInstance = Nothing
    
End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
    '
End Sub

Private Sub ComponentHandler_ItemSelected(ByVal VBComponent As VBIDE.VBComponent)
    
    If Not VBComponent.CodeModule Is Nothing Then
        Set mToolCodeView.ActiveCodeModule = VBComponent.CodeModule
        mToolCodeView.Reload
    End If
    
End Sub

Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Me.Show
    mToolCodeView.Reload
End Sub

Private Sub MenuHandler2_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    frmFindAll.Show
End Sub



Function AddToAddInCommandBar(sCaption As String, Optional menuName As String = "Add-Ins") As Office.CommandBarControl
    Dim cbMenuCommandBar As Office.CommandBarControl
    Dim cbMenu As Object
  
    On Error GoTo AddToAddInCommandBarErr
    Set cbMenu = g_VBInstance.CommandBars(menuName)
    
    If cbMenu Is Nothing Then Exit Function
    
    Set cbMenuCommandBar = cbMenu.Controls.Add(1)
    cbMenuCommandBar.Caption = sCaption
    Set AddToAddInCommandBar = cbMenuCommandBar
    
    Exit Function
    
AddToAddInCommandBarErr:
End Function

Private Sub ProjectHandler_ItemRemoved(ByVal VBProject As VBIDE.VBProject)
    Set mToolCodeView.ActiveCodeModule = Nothing
    Set mToolCodeView.CurMod = Nothing
End Sub
