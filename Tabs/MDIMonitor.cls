VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MDIMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const WIN_NAME As String = "VbaWindow"
Private Const ID_CLOSE As Long = 1
Private Const ID_CLOSE_ALL As Long = 2
Private Const ID_CLOSE_BUT_ACTIVE As Long = 3
Private Const ID_SHOW_CLOSE As Long = 4
Private Const MENU_SHOW_CLOSE As String = "Show Close Button On Active Tab"

Private Type NCCALCSIZE_PARAMS
    rgrc(0 To 2) As RECT
    lppos As Long 'WINDOWPOS
End Type

Implements ISubclass

Private m_MainWindowHook As cSubclass
Private m_Hook As cSubclass
Private m_MDIChildWindows As Collection
Private WithEvents m_tabStrip As TabStrip
Attribute m_tabStrip.VB_VarHelpID = -1

Private m_hMDIClient As Long
Private m_hMenu As Long
Private m_hWndMain As Long
'MDIMonitor can also be use in Normal MDI application
'So this next var is used to specify that we only want to track VbaWindow in IDE
Private m_bUseInAddin As Boolean
Private m_eWinVer As enWinVersion
Private m_bShowCloseButton As Boolean

Sub StartMonitor(ByVal afterStartUp As Boolean, _
                 ByVal hWndMain As Long, _
                 Optional ByVal VBIDEAddin As Boolean = False)
    'Dim itemAddin As AddIn
    
    EndMonitor
    
    m_bUseInAddin = VBIDEAddin
    m_hWndMain = hWndMain
    m_hMDIClient = FindWindowEx(hWndMain, 0, "MDIClient", vbNullString)
    
    If m_hMDIClient = 0 Then Exit Sub
  
    Set m_MDIChildWindows = New Collection
  
    Set m_Hook = New cSubclass

    With m_Hook
        .AddMsg WM_PARENTNOTIFY, MSG_AFTER
        .AddMsg WM_SIZE, MSG_AFTER
        .AddMsg WM_MDIACTIVATE, MSG_AFTER
        .AddMsg WM_NCCALCSIZE, MSG_AFTER
        .AddMsg WM_NCPAINT, MSG_AFTER
        .AddMsg WM_NCLBUTTONDOWN, MSG_AFTER
        .AddMsg WM_NCRBUTTONDOWN, MSG_AFTER
        .AddMsg WM_NCMOUSEMOVE, MSG_AFTER
        .AddMsg WM_NCLBUTTONUP, MSG_AFTER
        .AddMsg WM_NCHITTEST, MSG_AFTER
        .AddMsg WM_NCMOUSELEAVE, MSG_AFTER
        .Subclass m_hMDIClient, Me
    End With
  
    Set m_tabStrip = New TabStrip
  
    m_tabStrip.Create 0, 0, 100, 10
    
    #If IN_ADDIN Then ' Conditional compilation, we set this to 1 in our AddIn
        Dim item As VBIDE.AddIn

        For Each item In myLoader.VBInstance.Addins

            If InStr(1, item.Description, "MZ-Tools") Then
                mzToolsDetected = True
            End If

        Next

    #End If
    
    If mzToolsDetected = False Then
        Set m_MainWindowHook = New cSubclass
        m_MainWindowHook.AddMsg WM_SETTINGCHANGE, MSG_AFTER
        m_MainWindowHook.Subclass hWndMain, Me
    End If
    
    If afterStartUp Then
        RegisterExistingWindows
    End If
    
    SetWindowPos m_hMDIClient, 0, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_FRAMECHANGED
    
    'update client window if child is not in maximized state
    RedrawWindow m_hMDIClient, ByVal 0, 0, RDW_INVALIDATE Or RDW_UPDATENOW Or RDW_ERASE
    
    LoadRegSetting
    CreatePopup
End Sub

Sub EndMonitor()
    If m_hMDIClient = 0 Then Exit Sub
    If m_hWndMain = 0 Then Exit Sub
    
    TrackMouse = False
    If Not m_Hook Is Nothing Then
        If m_eWinVer < enWin2000 Then
            'Place extra lock
            Set lockMDIHook = m_Hook
        End If
        m_Hook.UnSubclass
        Set m_Hook = Nothing
    End If

    If Not m_MainWindowHook Is Nothing Then
        If m_eWinVer < enWin2000 Then
            'Place extra lock
            Set lockMDIHook = m_MainWindowHook
        End If
        m_MainWindowHook.UnSubclass
        Set m_MainWindowHook = Nothing
    End If

    Set m_MDIChildWindows = Nothing
    Set m_tabStrip = Nothing
  
    If m_hMenu <> 0 Then
        DestroyMenu m_hMenu
        m_hMenu = 0
    End If

    SetWindowPos m_hMDIClient, 0, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_FRAMECHANGED
    m_hWndMain = 0
    m_hMDIClient = 0
    
    SaveRegSetting
End Sub

Public Property Get TabHeight() As Long
    TabHeight = 22&
End Property

Public Sub ResetLockCount()

    If Not m_tabStrip Is Nothing Then
        m_tabStrip.ResetLockCount
    End If

End Sub

Private Sub Class_Initialize()
    m_eWinVer = GetOSVersion
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    EndMonitor
End Sub

Private Sub ISubclass_Proc(ByVal bBefore As Boolean, _
                           bHandled As Boolean, _
                           lReturn As Long, _
                           hwnd As Long, _
                           uMsg As Long, _
                           wParam As Long, _
                           lParam As Long)
    Dim key As String
    Dim hActive As Long
    Dim Button As Integer
    Dim tPt As POINTL
    
    Select Case uMsg

        Case WM_SETTINGCHANGE 'Main Window msg, this is not called if MZTools loaded
            m_tabStrip.InitColors
      
        Case WM_PARENTNOTIFY

            Select Case LoWord(wParam)

                Case WM_CREATE
                    AddNewWindow lParam
          
                Case WM_DESTROY
                    key = "#" & lParam
                    If m_eWinVer < enWin2000 Then
                        'Place extra lock
                        Set lockChild = m_MDIChildWindows.item(key)
                    End If
                    'now it's safe to remove the subclass
                    m_MDIChildWindows.Remove key
                    m_tabStrip.LockUpdate True
                    m_tabStrip.DeleteItem key
          
                    If m_tabStrip.TabCount > 0 Then
                        'This windows is in the process of unloading, but it still exists
                        'so any call to find activeWindow still returns this window
                        hActive = GetWindow(lParam, GW_HWNDNEXT)

                        If hActive <> 0 Then
                            m_tabStrip.ActivateItem "#" & CStr(hActive)
                        Else
                            hActive = SendMessage(m_hMDIClient, WM_MDIGETACTIVE, 0, ByVal 0)

                            If hActive <> 0 Then m_tabStrip.ActivateItem "#" & CStr(hActive)
                        End If
                    End If

                    m_tabStrip.LockUpdate False

                    'refresh color in case user have changed their system color preference
                    If mzToolsDetected Then m_tabStrip.InitColors
                    m_tabStrip.Redraw
            End Select
    
        Case WM_SIZE

            If wParam <> SIZE_MINIMIZED Then MoveTabStrip
        
        Case WM_NCHITTEST
            tPt.x = LoWord(lParam): tPt.y = HiWord(lParam)
            MapWindowPoints 0, m_hMDIClient, tPt, 1
            
            If tPt.y < 0 Then
                lReturn = HTBORDER
                If m_eWinVer >= enWin2000 Then
                    'still have some problem with TrackMouseEvent in win98
                    TrackMouse = True
                End If
            Else
                m_tabStrip.MouseLeave
            End If
            
        Case WM_NCCALCSIZE

            If wParam <> 0 Then OnNCCalcSize lParam
        
        Case WM_NCPAINT
            'refresh color in case user have changed their system color preference
            If mzToolsDetected Then m_tabStrip.InitColors
            m_tabStrip.Redraw

        Case WM_MDIACTIVATE
            key = "#" & wParam

            If m_tabStrip.IsExists(key) Then m_tabStrip.ActivateItem key
            
        Case WM_NCLBUTTONDOWN, WM_NCRBUTTONDOWN, WM_NCMOUSEMOVE, WM_NCLBUTTONUP
            
            tPt.x = LoWord(lParam)
            tPt.y = HiWord(lParam)
            
            MapWindowPoints 0, hwnd, tPt, 1
            
            tPt.y = tPt.y + TabHeight + 2&
            
            Select Case uMsg

                Case WM_NCLBUTTONDOWN
                    m_tabStrip.MouseDown vbLeftButton, tPt.x, tPt.y

                Case WM_NCRBUTTONDOWN
                    m_tabStrip.MouseDown vbRightButton, tPt.x, tPt.y
                
                Case WM_NCMOUSEMOVE
                    If GetAsyncKeyState(VK_LBUTTON) Then
                        Button = vbLeftButton
                    End If

                    m_tabStrip.MouseMove Button, tPt.x, tPt.y

                Case WM_NCLBUTTONUP
                    m_tabStrip.MouseUp Button, tPt.x, tPt.y
            End Select
        
        Case WM_NCMOUSELEAVE
            'MsgBox "Leave"
            TrackMouse = False
            m_tabStrip.MouseLeave
            
    End Select

End Sub

Public Sub Notify(ByVal itemhWnd As Long, _
                  ByVal action As Long)

    Select Case action

        Case ACTION_SETTEXT
            'Refresh text
            '            Dim item As MDIChildWindow
            '
            '            For Each item In m_MDIChildWindows
            '                m_tabStrip.ItemText "#" & item.hwnd, GetCaption(GetWinText(item.hwnd))
            '            Next
            m_tabStrip.ItemText "#" & itemhWnd, GetCaption(GetWinText(itemhWnd))

            'refresh color in case user have changed their system color preference
            If mzToolsDetected Then m_tabStrip.InitColors

        Case ACTION_ACTIVATE
            m_tabStrip.ActivateItem "#" & itemhWnd
    End Select

End Sub

Private Property Let TrackMouse(ByVal Value As Boolean)
    Static tracked As Boolean
    Dim tME As tagTRACKMOUSEEVENT
    If Value <> tracked Then
        tracked = Value
        tME.cbSize = Len(tME)
        tME.hwndTrack = m_hMDIClient
        tME.dwHoverTime = 10
        tME.dwFlags = TME_NONCLIENT Or TME_LEAVE
        If Value = False Then
            tME.dwFlags = tME.dwFlags Or TME_CANCEL
        End If
        Call TrackMouseEvent(tME)
        
    End If
End Property

Private Sub MoveTabStrip()
    Dim tR As RECT
    Dim lWidth As Long
  
    GetWindowRect m_hMDIClient, tR
    lWidth = tR.Right - tR.Left
    
    m_tabStrip.Move 2&, 1&, lWidth - 2&, TabHeight

    'refresh color in case user have changed their system color preference
    If mzToolsDetected Then m_tabStrip.InitColors
    m_tabStrip.Redraw
End Sub

Private Sub m_tabStrip_CompletePaint(ByVal hDC As Long)
    ReleaseDC m_hMDIClient, hDC
End Sub

Private Sub m_tabStrip_ItemClick(ByVal key As String, _
                                 ByVal prevKey As String, _
                                 ByVal Button As MouseButtonConstants)
  
    Dim itemhWnd As Long
    Dim menuID As Long
    Dim item As TabItem
    Dim hWnds() As Long
    
    If Button = vbLeftButton Then
    
        itemhWnd = CLng(Mid$(key, 2))
        PostMessage m_hMDIClient, WM_MDIACTIVATE, ByVal itemhWnd, ByVal 0
    
    Else
        menuID = ShowPopup(key)

        If menuID = 0 Then Exit Sub
        
        Select Case menuID

            Case ID_CLOSE
                itemhWnd = CLng(Mid$(key, 2))
                PostMessage itemhWnd, WM_CLOSE, ByVal 0, ByVal 0
        
            Case ID_CLOSE_ALL
                m_tabStrip.LockUpdate True
                itemhWnd = GetWindow(m_hMDIClient, GW_CHILD)

                Do While itemhWnd <> 0
                    SendMessage itemhWnd, WM_CLOSE, ByVal 0, ByVal 0
                    itemhWnd = GetWindow(m_hMDIClient, GW_CHILD)
                Loop

                m_tabStrip.LockUpdate False
                m_tabStrip.Redraw

            Case ID_CLOSE_BUT_ACTIVE
                'we want to use SendMessage to Close these childs, so we need this array
                ReDim hWnds(m_tabStrip.TabCount - 2&) As Long
                menuID = 0
                For Each item In m_tabStrip
                    If item.Active = False Then
                        hWnds(menuID) = m_MDIChildWindows.item(item.key).hwnd
                        menuID = menuID + 1&
                    End If
                Next
                m_tabStrip.LockUpdate True
                For menuID = 0 To UBound(hWnds)
                    SendMessage hWnds(menuID), WM_CLOSE, 0, ByVal 0
                Next
                m_tabStrip.LockUpdate False
                #If IN_ADDIN Then
                    'make sure the last item always active
                    m_tabStrip.item(1).Active = True
                #End If
                m_tabStrip.Redraw
            
            Case ID_SHOW_CLOSE
                m_bShowCloseButton = Not m_bShowCloseButton
                If m_bShowCloseButton Then
                    ModifyMenu m_hMenu, ID_SHOW_CLOSE, MF_STRING Or MF_CHECKED, ID_SHOW_CLOSE, ByVal MENU_SHOW_CLOSE
                Else
                    ModifyMenu m_hMenu, ID_SHOW_CLOSE, MF_STRING, ID_SHOW_CLOSE, ByVal MENU_SHOW_CLOSE
                End If
                m_tabStrip.ShowCloseButton = m_bShowCloseButton
        End Select

    End If

End Sub

Private Sub m_tabStrip_ItemClose(ByVal key As String)
    Dim itemhWnd As Long
    itemhWnd = CLng(Mid$(key, 2))
  
    PostMessage itemhWnd, WM_CLOSE, ByVal 0, ByVal 0
End Sub

Private Sub m_tabStrip_RequestPaint(hDC As Long)
    hDC = GetWindowDC(m_hMDIClient)
    SetBkMode hDC, TRANSPARENT
End Sub

Private Function GetCaption(ByVal Text As String) As String
    #If IN_ADDIN Then

        If Len(Text) > 0 Then
            
            If InStr(1, Text, "-") Then
                'Thanks to Zirro Tolerance
                GetCaption = Trim$(Replace(Split(Text, "-")(1), "(Code)", ""))
            Else
                GetCaption = Trim$(Replace(Text, "(Code)", ""))
            End If
        End If

    #Else
        GetCaption = Text
    #End If
End Function

Private Function GetWinText(hwnd As Long, _
                            Optional className As Boolean = False) As String
    'some static vars to speed up things, this func will be called many times
    Static sBuffer As String * 128& 'is it safe to use 128 bytes? should be enough..
    Static textLength As Long
  
    If className Then
        textLength = GetClassName(hwnd, sBuffer, 129&)
    Else
        textLength = GetWindowText(hwnd, sBuffer, 129&)
    End If
  
    If textLength > 0 Then
        GetWinText = Left$(sBuffer, textLength)
    End If

End Function

Private Sub AddNewWindow(ByVal hwnd As Long)
    Dim key As String
    Dim vbaItem As MDIChildWindow
  
    key = "#" & CStr(hwnd)
  
    If m_tabStrip.IsExists(key) Then Exit Sub
  
    Set vbaItem = New MDIChildWindow
    m_tabStrip.AddItem key, GetCaption(GetWinText(hwnd, False))
    vbaItem.Attach hwnd, ObjPtr(Me)
    m_MDIChildWindows.Add vbaItem, key
    'ShowWindow hwnd, SW_MAXIMIZE
End Sub

Private Sub RegisterExistingWindows()
    Dim lastWin As Long
  
    lastWin = GetWindow(m_hMDIClient, GW_CHILD)
    'disable update to speed up loading large project
    m_tabStrip.LockUpdate True

    Do While lastWin <> 0

        If m_bUseInAddin Then
            If GetWinText(lastWin, True) = WIN_NAME Then
                AddNewWindow lastWin
            End If

        Else
            'Add all childs we can find
            AddNewWindow lastWin
        End If
        
        lastWin = GetWindow(lastWin, GW_HWNDNEXT)
    Loop

    m_tabStrip.LockUpdate False
    m_tabStrip.Redraw
End Sub

Private Sub CreatePopup()
    m_hMenu = CreatePopupMenu()
    AppendMenu m_hMenu, MF_STRING, ID_CLOSE, "Close"
    AppendMenu m_hMenu, MF_STRING, ID_CLOSE_ALL, "Close All"
    AppendMenu m_hMenu, MF_STRING, ID_CLOSE_BUT_ACTIVE, "Close All But Active"
    AppendMenu m_hMenu, MF_SEPARATOR, 0, ""
    
    If m_bShowCloseButton Then
        AppendMenu m_hMenu, MF_STRING Or MF_CHECKED, ID_SHOW_CLOSE, MENU_SHOW_CLOSE
    Else
        AppendMenu m_hMenu, MF_STRING, ID_SHOW_CLOSE, MENU_SHOW_CLOSE
    End If
End Sub

Private Function ShowPopup(ByVal key As String) As Long
    Dim tPt As POINTL
  
    If m_tabStrip.TabCount > 1 Then
        EnableMenuItem m_hMenu, ID_CLOSE_ALL, MF_ENABLED
        EnableMenuItem m_hMenu, ID_CLOSE_BUT_ACTIVE, MF_ENABLED
    Else
        EnableMenuItem m_hMenu, ID_CLOSE_ALL, MF_GRAYED
        EnableMenuItem m_hMenu, ID_CLOSE_BUT_ACTIVE, MF_GRAYED
    End If
    
    If Len(key) > 0 Then
        EnableMenuItem m_hMenu, ID_CLOSE, MF_ENABLED
    Else
        EnableMenuItem m_hMenu, ID_CLOSE, MF_GRAYED
    End If
    
    GetCursorPos tPt
    ShowPopup = TrackPopupMenu(m_hMenu, TPM_RETURNCMD Or TPM_NONOTIFY, tPt.x, tPt.y, ByVal 0, m_hWndMain, ByVal 0)
End Function

Private Sub OnNCCalcSize(ByRef lParam As Long)
    Dim tNC As NCCALCSIZE_PARAMS
    
    CopyMemory tNC, ByVal lParam, Len(tNC)
    tNC.rgrc(0).Top = tNC.rgrc(0).Top + TabHeight
    LSet tNC.rgrc(1) = tNC.rgrc(0)
    CopyMemory ByVal lParam, tNC, Len(tNC)
    MoveTabStrip
End Sub

Private Sub SaveRegSetting()
    Dim reg As cRegistry
    Set reg = New cRegistry
    
    reg.ClassKey = HKEY_CURRENT_USER
    #If IN_ADDIN Then
        reg.SectionKey = "Software\Microsoft\Visual Basic\6.0\Addins\CodeHelp.Loader"
    #Else
        reg.SectionKey = "Software\" & App.EXEName & "\TabbedMDI"
    #End If
    reg.ValueKey = "HideClose"
    reg.ValueType = REG_DWORD
    reg.Value = Abs(CInt(Not m_bShowCloseButton))
    
    Set reg = Nothing
End Sub

Private Sub LoadRegSetting()
    Dim reg As cRegistry
    Set reg = New cRegistry
    
    reg.ClassKey = HKEY_CURRENT_USER
    #If IN_ADDIN Then
        reg.SectionKey = "Software\Microsoft\Visual Basic\6.0\Addins\CodeHelp.Loader"
    #Else
        reg.SectionKey = "Software\" & App.EXEName & "\TabbedMDI"
    #End If
    reg.ValueKey = "HideClose"
    reg.ValueType = REG_DWORD
    m_bShowCloseButton = Not CBool(reg.Value)
    
    m_tabStrip.ShowCloseButton = m_bShowCloseButton
    Set reg = Nothing
End Sub
