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
'Dim mcbMenuCommandBar As Office.CommandBarControl
'Public WithEvents MenuHandler As CommandBarEvents

Dim mcbMenuCommandBar2 As Office.CommandBarControl
Public WithEvents MenuHandler2 As CommandBarEvents
Attribute MenuHandler2.VB_VarHelpID = -1

Public WithEvents ComponentHandler As VBComponentsEvents
Attribute ComponentHandler.VB_VarHelpID = -1

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
    If ConnectMode = ext_cm_External Then
        'Used by the wizard toolbar to start this wizard
8        Me.Show

    Else
2        'Set mcbMenuCommandBar = AddToAddInCommandBar("CodeView")
         'If Not mcbMenuCommandBar Is Nothing Then
3        '    Set MenuHandler = g_VBInstance.Events.CommandBarEvents(mcbMenuCommandBar)
         'End If
    
9        Set mcbMenuCommandBar2 = AddToAddInCommandBar("Find-All")
         If Not mcbMenuCommandBar2 Is Nothing Then
10            Set MenuHandler2 = g_VBInstance.Events.CommandBarEvents(mcbMenuCommandBar2)
         End If
    
4        Set ComponentHandler = g_VBInstance.Events.VBComponentsEvents(Nothing)
6        Set wToolCodeView = g_VBInstance.Windows.CreateToolWindow(AddInInst, "CodeView.ToolCodeView", "CodeView", GuidCodeView, mToolCodeView)
11       Me.Show

    End If
     
    Exit Sub
error_handler:
    MsgBox "AddinInstance_OnConnection: Line: " & Erl() & " Desc: " & Err.Description
End Sub

Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    
    On Error GoTo hell
    Dim f As Form
    
'    If Not mcbMenuCommandBar Is Nothing Then
'        mcbMenuCommandBar.Delete
'        Set mcbMenuCommandBar = Nothing
'        If Not MenuHandler  Is Nothing Then Set MenuHandler = Nothing
'    End If
    
    If Not mcbMenuCommandBar2 Is Nothing Then
        mcbMenuCommandBar2.Delete
        Set mcbMenuCommandBar2 = Nothing
        If Not MenuHandler2 Is Nothing Then Set MenuHandler2 = Nothing
    End If

    If Not ComponentHandler Is Nothing Then Set ComponentHandler = Nothing
    If Not mToolCodeView Is Nothing Then Set mToolCodeView = Nothing
    If Not g_VBInstance Is Nothing Then Set g_VBInstance = Nothing
    If Not wToolCodeView Is Nothing Then Set wToolCodeView = Nothing
    
    For Each f In Forms
        Unload f
    Next
    
    Exit Sub
    
hell:
    MsgBox "CodeView.AddinInstance_OnDisconnection " & Err.Description

End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
    '
End Sub

Private Sub ComponentHandler_ItemRemoved(ByVal VBComponent As VBIDE.VBComponent)
    Set mToolCodeView.ActiveCodeModule = Nothing
    mToolCodeView.InitCodeView
End Sub

Private Sub ComponentHandler_ItemSelected(ByVal VBComponent As VBIDE.VBComponent)
    
    On Error GoTo hell
    
    If VBComponent.Type = vbext_ct_RelatedDocument Then Exit Sub
    
    'for related documents, VB6 IDE crashs here Method CodeModule of object VBComponent failed..
    'apparently the on error resume next can not catch this...(but ok while in IDE debugging):
    'HRESULT: 0x80010105 (RPC_E_SERVERFAULT) question Automation error the server threw an exception..
    If Not VBComponent.CodeModule Is Nothing Then
        If Err.Number <> 0 Then Exit Sub
        Set mToolCodeView.ActiveCodeModule = VBComponent.CodeModule
        mToolCodeView.Reload
    End If
    
    Exit Sub
hell:
    MsgBox "Codeview.ComponentHandler_ItemSelected: " & Err.Description
    
End Sub

'Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
'    On Error Resume Next
'    Me.Show
'    mToolCodeView.Reload
'End Sub

Private Sub MenuHandler2_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    frmFindAll.Show
End Sub

Function AddToAddInCommandBar(sCaption As String, Optional menuName As String = "Add-Ins") As Office.CommandBarControl
    Dim cbMenuCommandBar  As Office.CommandBarControl
    Dim cbMenu As Object
    Dim ctl As CommandBarControls
    
    On Error GoTo AddToAddInCommandBarErr
    Set cbMenu = g_VBInstance.CommandBars(menuName)
    
    If cbMenu Is Nothing Then Exit Function
    Set ctl = cbMenu.Controls
    Set cbMenuCommandBar = ctl.Add(, , , 1)
    cbMenuCommandBar.Caption = sCaption
    Set AddToAddInCommandBar = cbMenuCommandBar
    
    Exit Function
    
AddToAddInCommandBarErr:
End Function

