VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Loader 
   ClientHeight    =   6990
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   11520
   _ExtentX        =   20320
   _ExtentY        =   12330
   _Version        =   393216
   Description     =   "VB IDE Extender"
   DisplayName     =   "CodeHelp VB Extender"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Loader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
#If IS_DEBUG Then
    Private Declare Function GetDesktopWindow _
                    Lib "user32.dll" () As Long
#End If

'Private Const ADDIN_NAME      As String = "CodeHelp IDE Extender"
Private m_VBE                           As VBIDE.VBE
Private m_VbaWinMgr                     As MDIMonitor
Private WithEvents IDEFileEvents        As FileControlEvents
Attribute IDEFileEvents.VB_VarHelpID = -1

'------------------------------------------------------
'this method adds the Add-In to VB
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, _
                                       ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, _
                                       ByVal AddInInst As Object, _
                                       custom() As Variant)
    On Error GoTo error_handler
    'save the vb instance
    Set m_VBE = Application
    Set IDEFileEvents = m_VBE.Events.FileControlEvents(Nothing)
    StartExtender (ConnectMode = ext_cm_AfterStartup)
    
    Exit Sub
  
error_handler:
    MsgBox Err.Description
End Sub

'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, _
                                          custom() As Variant)
    On Error Resume Next
    'shut down the Add-In
    'Debug.Print "Disconnect"
    StopExtender
End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
    'set this to display the form on connect
    Debug.Print "IDT Startup"
End Sub

Friend Property Get VBInstance() As VBIDE.VBE
    Set VBInstance = m_VBE
End Property

Private Sub StartExtender(ByVal afterStartUp As Boolean)
    Dim hWndMain As Long
    
    Set myLoader = Me
    
    'This for debugging purpose
    'So we can run this add in in separate IDE instance
    'Set IS_DEBUG = 1 in conditional compilation argument in Make Tab of the Project properties dialog
    #If IS_DEBUG Then
        Dim lastWnd As Long
        Dim sCaption As String
        lastWnd = GetWindow(GetDesktopWindow, GW_CHILD)

        Do While lastWnd <> 0
            sCaption = GetWinText(lastWnd)

            If Left$(sCaption, 8) = "CodeHelp" Then
                hWndMain = lastWnd
                Exit Do
            End If

            lastWnd = GetWindow(lastWnd, GW_HWNDNEXT)
        Loop

    #Else
        hWndMain = m_VBE.MainWindow.hwnd
    #End If
    Set m_VbaWinMgr = New MDIMonitor
    m_VbaWinMgr.StartMonitor afterStartUp, hWndMain, True
End Sub

Private Sub StopExtender()
    m_VbaWinMgr.EndMonitor
    Set m_VbaWinMgr = Nothing
    Set myLoader = Nothing
End Sub

Private Sub IDEFileEvents_BeforeLoadFile(ByVal VBProject As VBIDE.VBProject, _
                                         FileNames() As String)
    Static lastProject As VBIDE.VBProject

    If Not VBProject Is lastProject Then
        m_VbaWinMgr.ResetLockCount
        Set lastProject = VBProject
    End If

End Sub
