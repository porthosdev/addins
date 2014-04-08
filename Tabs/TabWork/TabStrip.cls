VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TabStrip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function OleTranslateColor2 _
                Lib "oleaut32.dll" _
                Alias "OleTranslateColor" (ByVal Clr As Long, _
                                           ByVal hPal As Long, _
                                           ByVal lpcolorref As Long) As Long

Private m_lClrSysHilite As Long
Private m_lClrSys3DFace As Long
Private m_lClrSys3DHilite As Long
Private m_lClrSys3DFaceLite As Long
Private m_lClrSysShadow As Long

Private m_lClrText As Long
Public Event ItemClick(ByVal key As String, ByVal prevKey As String, ByVal Button As MouseButtonConstants)
Public Event ItemClose(ByVal key As String)
Public Event RequestPaint(ByRef hDC As Long)
Public Event CompletePaint(ByVal hDC As Long)

Private m_Items As Collection
Private m_lWidth As Long
Private m_lHeight As Long

Private m_lastActiveKey As String
Private m_lLeft As Long
Private m_lTop As Long

Private m_oFont As IFont
Private m_oActiveFont As IFont
Private m_oMarlettFont As IFont

Private m_lVirtualLeft As Long
Private m_lVirtualWidth As Long
Private m_bScrollRight As Boolean
Private m_bScrollLeft As Boolean

Private m_closeIcon As StdPicture
Private m_closeIconHover As StdPicture
Private m_closeIconPush As StdPicture

Private m_iLockCount As Long

Private m_tRNav As RECT
Private m_tRVisibleRect As RECT
Private m_tRClose As RECT
Private m_bDownOnClose As Boolean
Private m_bShowCloseButton As Boolean

Public Property Let ShowCloseButton(ByVal value As Boolean)
    If value <> m_bShowCloseButton Then
        m_bShowCloseButton = value
        Redraw
    End If
End Property

Public Property Get Left() As Long
    Left = m_lLeft
End Property

Public Property Let Left(ByVal lLeft As Long)

    If m_lLeft <> lLeft Then
        m_lLeft = lLeft
        Redraw
    End If

End Property

Public Property Get Top() As Long
    Top = m_lTop
End Property

Public Property Let Top(ByVal lTop As Long)

    If Top <> m_lTop Then
        m_lTop = lTop
        Redraw
    End If

End Property

Public Property Get TabCount() As Long

    If m_Items Is Nothing Then Exit Property
    TabCount = m_Items.Count
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = m_Items.[_NewEnum]
End Property

Public Property Get item(vntIndexKey As Variant) As TabItem
    Set item = m_Items(vntIndexKey)
End Property

Public Sub Create(ByVal lLeft As Long, _
                  ByVal Top As Long, _
                  ByVal Width As Long, _
                  ByVal height As Long)
    
    Set m_Items = New Collection
  
    Set m_oFont = New StdFont

    With m_oFont
        .Name = "Tahoma"
        .Bold = False
        .Size = 8
    End With
  
    Set m_oActiveFont = New StdFont

    With m_oActiveFont
        .Name = "Tahoma"
        .Bold = True
        .Size = 8
    End With
  
    Set m_oMarlettFont = New StdFont

    With m_oMarlettFont
        .Name = "Marlett"
        .Bold = False
        .Size = 10
    End With
  
    Set m_closeIcon = LoadResPicture(101, vbResIcon)
    Set m_closeIconHover = LoadResPicture(102, vbResIcon)
    Set m_closeIconPush = LoadResPicture(103, vbResIcon)
    
    Move lLeft, Top, Width, height
End Sub

Public Sub Move(ByVal lLeft As Long, _
                ByVal Top As Long, _
                ByVal Width As Long, _
                ByVal height As Long)
    m_lLeft = lLeft
    m_lTop = Top
    m_lWidth = Width
    m_lHeight = height
    m_lVirtualLeft = m_lWidth - m_lVirtualWidth
    
    If m_lVirtualWidth > m_lWidth Then m_lVirtualLeft = m_lVirtualLeft - 28&
    If m_lVirtualLeft > 0 Then m_lVirtualLeft = 3&
    Redraw
End Sub

Public Sub ItemText(ByVal key As Variant, _
                    ByVal Text As String)
    On Error Resume Next
    Dim oTabItem As TabItem
  
    Set oTabItem = m_Items(key)
  
    If Not oTabItem Is Nothing Then
        oTabItem.Text = Text
        Redraw
    End If

End Sub

Public Sub Redraw()
    Dim hDC As Long
    
    Dim tPt As POINTL

    'Only update if no lock have been placed
    If m_iLockCount <> 0 Then Exit Sub
    
    RaiseEvent RequestPaint(hDC)
    
    If hDC <> 0 Then
        RecalculateRects hDC
        SetWindowOrgEx hDC, -m_lVirtualLeft - m_lLeft, 0, tPt
        DrawBackground hDC
        DrawTabItems hDC ', tR
        SetWindowOrgEx hDC, tPt.x, tPt.y, ByVal 0
        
        RaiseEvent CompletePaint(hDC)
    End If

End Sub

Public Sub AddItem(ByVal key As String, _
                   ByVal Text As String)
    Dim newItem As TabItem
  
    Set newItem = New TabItem
    newItem.key = key
    newItem.Text = Text
  
    m_Items.Add newItem, key
    ActivateItem key
  
End Sub

Public Sub LockUpdate(ByVal value As Boolean)

    'this can be called by different branch of code, so we need to place a reference counting
    'But please call it in pairs, always reenable update after disable update call
    If value Then
        m_iLockCount = m_iLockCount + 1&
    Else
        m_iLockCount = m_iLockCount - 1&
    End If

    If m_iLockCount < 0 Then m_iLockCount = 0
End Sub

Public Sub ResetLockCount()
    m_iLockCount = 0
End Sub

Public Sub DeleteItem(key As String)
    On Error Resume Next
    m_Items.Remove key
    Redraw
End Sub

Public Sub ActivateItem(key As String)
    Dim oTabItem As TabItem
    
    If m_lastActiveKey <> key Then
        
        If Len(m_lastActiveKey) > 0 Then
            If IsExists(m_lastActiveKey) Then
                Set oTabItem = m_Items(m_lastActiveKey)
                oTabItem.Active = False
            End If
        End If
        'extra check for Win98
        If IsExists(key) Then
            Set oTabItem = m_Items(key)
            oTabItem.Active = True
    
            m_lastActiveKey = key
    
            EnsureVisible key
        End If
        Redraw
        
    End If

End Sub

Public Function IsExists(ByVal key As String) As Boolean
    On Error Resume Next
    Dim myItem As TabItem
  
    Set myItem = m_Items(key)
    IsExists = (Err.Number = 0)
End Function

Friend Sub MouseDown(Button As Integer, _
                     ByVal x As Long, _
                     ByVal y As Long)
    Dim oTabItem As TabItem
    Dim prevKey As String
    Dim tR As RECT
    
    x = x - m_lVirtualLeft
    y = y - m_lTop
    m_bDownOnClose = False
    
    If m_tRNav.Right > 0 Then
        If PtInRect(m_tRNav, x, y) Then
            If x < (m_tRNav.Right - 15) Then

                'scroll left
                If m_bScrollLeft Then
                    Set oTabItem = m_Items(1)
                    tR = oTabItem.Rectangle
                    m_lVirtualLeft = m_lVirtualLeft + (tR.Right - tR.Left)

                    If m_lVirtualLeft > 0 Then m_lVirtualLeft = 3&
                    Redraw
                End If

            Else

                If m_bScrollRight Then
                    Set oTabItem = m_Items(m_Items.Count)
                    tR = oTabItem.Rectangle
                    m_lVirtualLeft = m_lVirtualLeft - (tR.Right - tR.Left)
                    Redraw
                End If
            End If

            Exit Sub
        End If
    End If
  
    For Each oTabItem In m_Items

        If oTabItem.Hittest(x, y) Then
            prevKey = m_lastActiveKey

            If Button = vbLeftButton Then
                If m_lastActiveKey <> oTabItem.key Then
                    ActivateItem oTabItem.key
                    RaiseEvent ItemClick(oTabItem.key, prevKey, Button)
                ElseIf m_bShowCloseButton Then

                    'hit on already active item
                    'check for close button
                    If PtInRect(m_tRClose, x, y) Then
                        m_bDownOnClose = True
                        
                        CloseHovered(Button) = False
                        CloseHovered(Button) = True

                    End If
                End If

            Else
                RaiseEvent ItemClick(oTabItem.key, "", vbRightButton)
            End If

            Exit Sub
        End If

    Next
    
    If Button = vbRightButton Then
        RaiseEvent ItemClick("", "", vbRightButton)
    End If

End Sub

Friend Sub MouseMove(Button As Integer, _
                     ByVal x As Long, _
                     ByVal y As Long)

    If m_bShowCloseButton Then
        x = x - m_lVirtualLeft
        y = y - m_lTop
        
        CloseHovered(Button) = PtInRect(m_tRClose, x, y)
    End If
End Sub

Friend Sub MouseUp(Button As Integer, _
                   ByVal x As Long, _
                   ByVal y As Long)
    If m_bShowCloseButton Then
        x = x - m_lVirtualLeft
        y = y - m_lTop
            
        If m_bDownOnClose Then
            m_bDownOnClose = False
            CloseHovered(0) = False
            If PtInRect(m_tRClose, x, y) Then
                RaiseEvent ItemClose(m_lastActiveKey)
            End If
        Else
            CloseHovered(0) = False
        End If
    End If
End Sub

Friend Sub MouseLeave()
    If m_bShowCloseButton Then
        CloseHovered(0) = False
    End If
End Sub

Private Property Let CloseHovered(Button As Integer, ByVal value As Boolean)
    Static isHovered As Boolean
    
    Dim hDC As Long
    Dim tPt As POINTL
    
    If value <> isHovered Then
        RaiseEvent RequestPaint(hDC)

        If hDC <> 0 Then
            SetWindowOrgEx hDC, -m_lVirtualLeft - m_lLeft, 0, tPt

            If value Then
                If Button = vbLeftButton And m_bDownOnClose Then
                    DrawIconEx hDC, m_tRClose.Left, m_tRClose.Top, m_closeIconPush.Handle, 16, 16, 0, 0, DI_NORMAL
                Else
                    DrawIconEx hDC, m_tRClose.Left, m_tRClose.Top, m_closeIconHover.Handle, 16, 16, 0, 0, DI_NORMAL
                End If

            Else
                
                DrawIconEx hDC, m_tRClose.Left, m_tRClose.Top, m_closeIcon.Handle, 16, 16, 0, 0, DI_NORMAL
            End If

            SetWindowOrgEx hDC, tPt.x, tPt.y, ByVal 0
            RaiseEvent CompletePaint(hDC)
        End If

        isHovered = value
    End If

End Property

Public Sub InitColors()
    
    OleTranslateColor vbHighlight, 0, m_lClrSysHilite
    OleTranslateColor vbButtonShadow, 0, m_lClrSysShadow
    OleTranslateColor vbButtonFace, 0, m_lClrSys3DFace
    OleTranslateColor vb3DHighlight, 0, m_lClrSys3DHilite

    If m_lClrSys3DFace = 13160660 Then
        m_lClrSys3DFaceLite = RGB(247, 243, 233)
    Else
        m_lClrSys3DFaceLite = AlphaBlend(vbButtonFace, vbWindowBackground, 128&)
    End If

    OleTranslateColor vbButtonText, 0, m_lClrText
End Sub

Private Sub Class_Initialize()
    InitColors
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Set m_Items = Nothing
    Set m_closeIcon = Nothing
    Set m_closeIconHover = Nothing
    Set m_closeIconPush = Nothing
    Set m_oMarlettFont = Nothing
    Set m_oFont = Nothing
    Set m_oActiveFont = Nothing
End Sub

Private Sub DrawBackground(ByVal hDC As Long)

    Dim hBr As Long, tRLine As RECT
    Dim tR As RECT
    
    tR = m_tRVisibleRect
    
    hBr = CreateSolidBrush(m_lClrSys3DFaceLite)
    FillRect hDC, tR, hBr
    DeleteObject hBr
    hBr = CreateSolidBrush(m_lClrSys3DFace)
    tRLine = tR
    tRLine.Top = tRLine.Bottom - 2&
    FillRect hDC, tRLine, hBr
    DeleteObject hBr
    hBr = SelectObject(hDC, CreatePen(PS_SOLID, 1, m_lClrSys3DHilite))
    MoveToEx hDC, tRLine.Left, tRLine.Top - 1&, ByVal 0
    LineTo hDC, tRLine.Right, tRLine.Top - 1&
    DeleteObject SelectObject(hDC, hBr)
End Sub

Private Sub DrawTabItems(hDC As Long)
    Dim oTabItem As TabItem
    Dim itemRect As RECT
    Dim hOldFont As Long
    Dim hBrActive As Long, hBrInActive As Long
    Dim hOldPen As Long
    Dim tRTest As RECT

    Dim iCount As Long
      
    hOldFont = SelectObject(hDC, m_oFont.hFont)
    hBrActive = CreateSolidBrush(m_lClrSys3DFace)
    hBrInActive = CreateSolidBrush(m_lClrSys3DFaceLite)
    
    hOldPen = SelectObject(hDC, CreatePen(PS_SOLID, 1, m_lClrSysShadow))
    
    m_bScrollLeft = False
    m_bScrollRight = False
    
    If m_Items.Count > 0 Then
   
        For Each oTabItem In m_Items
            iCount = iCount + 1&
            itemRect = oTabItem.Rectangle
            
            If IntersectRect(tRTest, itemRect, m_tRVisibleRect) Then   'Not Off screen
                
                FillRect hDC, itemRect, hBrInActive
                itemRect.Top = itemRect.Top + 3&
                
                If oTabItem.Active = False Then
                    SelectObject hDC, m_oFont.hFont
                    SetTextColor hDC, m_lClrSysShadow
    
                    If iCount > 1 Then
                        MoveToEx hDC, itemRect.Left - 1&, itemRect.Top + 3&, ByVal 0
                        LineTo hDC, itemRect.Left - 1&, itemRect.Bottom - 2&
              
                        If iCount = m_Items.Count Then
                            MoveToEx hDC, itemRect.Right - 1&, itemRect.Top + 3&, ByVal 0
                            LineTo hDC, itemRect.Right - 1&, itemRect.Bottom - 2&
                        End If
                    End If
                
                Else
                    SelectObject hDC, m_oActiveFont.hFont
                    SetTextColor hDC, m_lClrText
                    
                    itemRect.Bottom = itemRect.Bottom + 1&
                    FillRect hDC, itemRect, hBrActive
                    itemRect.Bottom = itemRect.Bottom - 1&
                    
                    MoveToEx hDC, itemRect.Right - 1, itemRect.Top, ByVal 0
                    LineTo hDC, itemRect.Right - 1, itemRect.Bottom
                    
                    If m_bShowCloseButton Then
                        DrawIconEx hDC, m_tRClose.Left, m_tRClose.Top, m_closeIcon.Handle, 16, 16, 0, 0, DI_NORMAL
        
                        'restore width for drawing caption
                        itemRect.Right = itemRect.Right - 20&
                    End If
                End If
          
                DrawText hDC, oTabItem.Text, Len(oTabItem.Text), itemRect, DT_SINGLELINE Or DT_CENTER Or DT_VCENTER
                '        Else
                '            Debug.Print oTabItem.Text
            End If
            
            'off or partially off
            If (itemRect.Left > m_tRVisibleRect.Right - 30&) Or (itemRect.Right > m_tRVisibleRect.Right - 30&) Then
                m_bScrollRight = True
            End If
            
            If (itemRect.Left < m_tRVisibleRect.Left) Or (itemRect.Right <= m_tRVisibleRect.Left) Then
                m_bScrollLeft = True
            End If
    
        Next
    
        If m_lVirtualWidth > (m_tRVisibleRect.Right - m_tRVisibleRect.Left) Then
            SetRect tRTest, m_tRVisibleRect.Right - 25&, m_tRVisibleRect.Top + 2&, m_tRVisibleRect.Right, m_tRVisibleRect.Bottom - 2&
    
            FillRect hDC, tRTest, hBrInActive
    
            'Scroll Left
            If m_bScrollLeft Then
                SetTextColor hDC, m_lClrSysShadow
            Else
                SetTextColor hDC, m_lClrSys3DFace
            End If
    
            SelectObject hDC, m_oMarlettFont.hFont
            DrawText hDC, "3", 1, tRTest, DT_LEFT Or DT_SINGLELINE Or DT_VCENTER
    
            'Scroll Right
            If m_bScrollRight Then
                SetTextColor hDC, m_lClrSysShadow
            Else
                SetTextColor hDC, m_lClrSys3DFace
            End If
    
            DrawText hDC, "4", 1, tRTest, DT_RIGHT Or DT_SINGLELINE Or DT_VCENTER
            m_tRNav = tRTest
    
        Else
            m_tRNav.Right = 0
        End If
    
    End If
    SelectObject hDC, hOldFont
    DeleteObject SelectObject(hDC, hOldPen)
    DeleteObject hBrActive
    DeleteObject hBrInActive
    
End Sub

Private Function GetTextWidth(ByVal hDC As Long, _
                              ByVal Text As String) As Long
    Dim tS As POINTL
    Call GetTextExtentPoint32(hDC, Text, Len(Text), tS)
    GetTextWidth = tS.x
End Function

Private Function AlphaBlend(ByVal clrFirst As Long, _
                            ByVal clrSecond As Long, _
                            ByVal lAlpha As Long) As Long
    Dim clrFore         As RGBQUAD
    Dim clrBack         As RGBQUAD
    Dim lDif As Long
    On Error Resume Next
    OleTranslateColor2 clrFirst, 0, VarPtr(clrFore)
    OleTranslateColor2 clrSecond, 0, VarPtr(clrBack)

    If lAlpha < 0& Then lAlpha = 0&
    If lAlpha > 255& Then lAlpha = 255&

    With clrFore
        lDif = CLng(.rgbRed) - CLng(clrBack.rgbRed)
        .rgbRed = (lAlpha * lDif) \ 255 + clrBack.rgbRed
        lDif = CLng(.rgbGreen) - CLng(clrBack.rgbGreen)
        .rgbGreen = (lAlpha * lDif) \ 255 + clrBack.rgbGreen
        lDif = CLng(.rgbBlue) - CLng(clrBack.rgbBlue)
        .rgbBlue = (lAlpha * lDif) \ 255 + clrBack.rgbBlue
    End With

    CopyMemory ByVal VarPtr(AlphaBlend), ByVal VarPtr(clrFore), 4
End Function

Private Sub RecalculateRects(ByVal hDC As Long)
    Dim item As TabItem
    Dim tR As RECT
    Dim lastWidth As Long
    
    SetRect m_tRVisibleRect, m_lLeft, m_lTop, m_lLeft + m_lWidth, m_lTop + m_lHeight + 1&
    
    LSet tR = m_tRVisibleRect
    tR.Bottom = tR.Bottom - 3&
    
    m_lVirtualWidth = 0
    SetRectEmpty m_tRClose

    For Each item In m_Items
        tR.Left = lastWidth

        If item.Active Then
            
            SelectObject hDC, m_oActiveFont.hFont
            tR.Right = tR.Left + GetTextWidth(hDC, item.Text) + 10
            'Add space for close icon
            If m_bShowCloseButton Then
                tR.Right = tR.Right + 20&
                m_tRClose = tR
                m_tRClose.Left = m_tRClose.Right - 20&
                m_tRClose.Right = m_tRClose.Left + 16&
                m_tRClose.Top = m_tRClose.Top + 4&
                m_tRClose.Bottom = m_tRClose.Top + 16&
            End If
        Else
            SelectObject hDC, m_oFont.hFont
            tR.Right = tR.Left + GetTextWidth(hDC, item.Text) + 10&
        End If

        lastWidth = tR.Right
        item.Rectangle = tR
        m_lVirtualWidth = m_lVirtualWidth + (tR.Right - tR.Left)
    Next
    
    With m_tRVisibleRect
        .Left = -m_lVirtualLeft
        .Right = .Left + m_lWidth
    End With
    
End Sub

Sub EnsureVisible(ByVal key As String)

    Dim tRItem As RECT
    Dim item As TabItem
    
    Set item = m_Items(key)
    LSet tRItem = item.Rectangle
    
    If (tRItem.Left < m_tRVisibleRect.Left) Then
        
        m_lVirtualLeft = -tRItem.Left + 3&

        If m_lVirtualLeft > 0 Then m_lVirtualLeft = 3
    
    ElseIf (tRItem.Right >= m_tRVisibleRect.Right) Then
        m_lVirtualLeft = m_lVirtualLeft - (tRItem.Right - m_tRVisibleRect.Right) - 28&

        If m_lVirtualLeft > 0 Then m_lVirtualLeft = 3
    End If

End Sub
