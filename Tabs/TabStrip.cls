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
Public Event RequestPaint(ByRef hdc As Long)
Public Event CompletePaint(ByVal hdc As Long)
'Public Event RequestBounds(ByRef lLeft As Long, ByRef Top As Long)

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
'Private m_bLocked As Boolean

Private m_closeIcon As StdPicture
Private m_tRNav As RECT
Private m_iLockCount As Long

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

    If m_lVirtualLeft > 0 Then m_lVirtualLeft = 0
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
    Dim hdc As Long
    Dim tR As RECT
    
    'Only update if no lock have been placed
    If m_iLockCount <> 0 Then Exit Sub
    
    RaiseEvent RequestPaint(hdc)
  
    If hdc <> 0 Then
        SetRect tR, m_lLeft, m_lTop, m_lLeft + m_lWidth, m_lTop + m_lHeight + 1&
        DrawBackground hdc, tR
        DrawTabItems hdc, tR
        RaiseEvent CompletePaint(hdc)
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

Public Sub DeleteItem(key As String)
    On Error Resume Next
    m_Items.Remove key
    Redraw
End Sub

Public Sub ActivateItem(key As String)
    Dim oTabItem As TabItem

    If m_lastActiveKey <> key Then

        'Set oTabItem = m_Items(Key)
        For Each oTabItem In m_Items
            oTabItem.Active = (key = oTabItem.key)
        Next

        m_lastActiveKey = key
        
        EnsureVisible key
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
  
    If m_tRNav.Right > 0 Then
        If PtInRect(m_tRNav, x, y) Then
            If x < (m_tRNav.Right - 15) Then

                'scroll left
                If m_bScrollLeft Then
                    Set oTabItem = m_Items(1)
                    tR = oTabItem.Rectangle
                    m_lVirtualLeft = m_lVirtualLeft + (tR.Right - tR.Left)

                    If m_lVirtualLeft > 0 Then m_lVirtualLeft = 0
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
                Else
                    'hit on already active item
                    'check for close button
                    tR = oTabItem.Rectangle

                    If (x > (tR.Right - 20&)) And (x < tR.Right - 2&) Then
                        RaiseEvent ItemClose(oTabItem.key)
                    End If
                End If

            Else
                RaiseEvent ItemClick(oTabItem.key, "", vbRightButton)
            End If

            Exit For
        End If

    Next

End Sub

Private Sub DrawBackground(ByVal hdc As Long, _
                           tR As RECT)
    Dim hBr As Long, tRLine As RECT
    hBr = CreateSolidBrush(m_lClrSys3DFaceLite)
    FillRect hdc, tR, hBr
    DeleteObject hBr
    hBr = CreateSolidBrush(m_lClrSys3DFace)
    tRLine = tR
    tRLine.Top = tRLine.Bottom - 2&
    FillRect hdc, tRLine, hBr
    DeleteObject hBr
    hBr = SelectObject(hdc, CreatePen(PS_SOLID, 1, m_lClrSys3DHilite))
    MoveToEx hdc, tRLine.Left, tRLine.Top - 1&, ByVal 0
    LineTo hdc, tRLine.Right, tRLine.Top - 1&
    DeleteObject SelectObject(hdc, hBr)
End Sub

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
    Set m_closeIcon = LoadResPicture(101, vbResIcon)
End Sub

Private Sub Class_Terminate()
    Set m_Items = Nothing
    Set m_closeIcon = Nothing
    Set m_oMarlettFont = Nothing
    Set m_oFont = Nothing
    Set m_oActiveFont = Nothing
End Sub

Private Sub DrawTabItems(hdc As Long, _
                         tR As RECT)
    Dim oTabItem As TabItem
    Dim itemRect As RECT, lastWidth As Long
    Dim hOldFont As Long
    Dim hBrActive As Long, hBrInActive As Long
    Dim hOldPen As Long
    Dim tRTest As RECT

    Dim iCount As Long
  
    'set the same bounds as tabstrip
    itemRect = tR
    itemRect.Bottom = itemRect.Bottom - 3&
    hOldFont = SelectObject(hdc, m_oFont.hFont)
    hBrActive = CreateSolidBrush(m_lClrSys3DFace)
    hBrInActive = CreateSolidBrush(m_lClrSys3DFaceLite)
    lastWidth = m_lLeft + m_lVirtualLeft + 3&
  
    hOldPen = SelectObject(hdc, CreatePen(PS_SOLID, 1, m_lClrSysShadow))
    
    m_bScrollLeft = False
    m_bScrollRight = False
    m_lVirtualWidth = 0
    
    For Each oTabItem In m_Items
        iCount = iCount + 1&
        itemRect.Left = lastWidth
        
        If oTabItem.Active Then
            SelectObject hdc, m_oActiveFont.hFont
        Else
            SelectObject hdc, m_oFont.hFont
        End If
        
        itemRect.Right = lastWidth + GetTextWidth(hdc, oTabItem.Text) + 10
        lastWidth = itemRect.Right
    
        If IntersectRect(tRTest, itemRect, tR) Then 'Not Off screen
            If oTabItem.Active = False Then
                SetTextColor hdc, m_lClrSysShadow
                FillRect hdc, itemRect, hBrInActive

                If iCount > 1 Then
                    MoveToEx hdc, itemRect.Left - 1&, itemRect.Top + 3&, ByVal 0
                    LineTo hdc, itemRect.Left - 1&, itemRect.Bottom - 2&
          
                    If iCount = m_Items.Count Then
                        MoveToEx hdc, itemRect.Right - 1&, itemRect.Top + 3&, ByVal 0
                        LineTo hdc, itemRect.Right - 1&, itemRect.Bottom - 2&
                    End If
                End If

                oTabItem.Rectangle = itemRect
            End If

            itemRect.Top = itemRect.Top + 3&
      
            If oTabItem.Active Then
                SetTextColor hdc, m_lClrText
                'Add space for close icon
                itemRect.Right = itemRect.Right + 20&
                lastWidth = itemRect.Right
            
                oTabItem.Rectangle = itemRect
            
                itemRect.Bottom = itemRect.Bottom + 1&
                FillRect hdc, itemRect, hBrActive
                itemRect.Bottom = itemRect.Bottom - 1&
                MoveToEx hdc, itemRect.Right - 1, itemRect.Top, ByVal 0
                LineTo hdc, itemRect.Right - 1, itemRect.Bottom
                DrawIconEx hdc, itemRect.Right - 20&, itemRect.Top + 2&, m_closeIcon.Handle, 16, 16, 0, 0, DI_NORMAL
            
                'restore width for drawing caption
                itemRect.Right = itemRect.Right - 20&
            End If
      
            DrawText hdc, oTabItem.Text, Len(oTabItem.Text), itemRect, DT_SINGLELINE Or DT_CENTER Or DT_VCENTER
            itemRect.Top = itemRect.Top - 3&
        End If
        
        'off or partially off
        If (itemRect.Left > tR.Right - 30&) Or (itemRect.Right > tR.Right - 30&) Then
            m_bScrollRight = True
        End If
        
        If (itemRect.Left < tR.Left) Or (itemRect.Right <= tR.Left) Then
            m_bScrollLeft = True
        End If

        m_lVirtualWidth = m_lVirtualWidth + (itemRect.Right - itemRect.Left)
    Next

    m_lVirtualWidth = m_lVirtualWidth + 20&

    If m_lVirtualWidth > (tR.Right - tR.Left) Then
        SetRect tRTest, tR.Right - 25&, tR.Top + 2&, tR.Right, tR.Bottom - 2&
    
        FillRect hdc, tRTest, hBrInActive
        
        'Scroll Left
        If m_bScrollLeft Then
            SetTextColor hdc, m_lClrSysShadow
        Else
            SetTextColor hdc, m_lClrSys3DFace
        End If

        SelectObject hdc, m_oMarlettFont.hFont
        DrawText hdc, "3", 1, tRTest, DT_LEFT Or DT_SINGLELINE Or DT_VCENTER
        
        'Scroll Right
        If m_bScrollRight Then
            SetTextColor hdc, m_lClrSysShadow
        Else
            SetTextColor hdc, m_lClrSys3DFace
        End If

        DrawText hdc, "4", 1, tRTest, DT_RIGHT Or DT_SINGLELINE Or DT_VCENTER
        m_tRNav = tRTest

    Else
        m_tRNav.Right = 0
    End If
  
    SelectObject hdc, hOldFont
    DeleteObject SelectObject(hdc, hOldPen)
    DeleteObject hBrActive
    DeleteObject hBrInActive
End Sub

Private Function GetTextWidth(ByVal hdc As Long, _
                              ByVal Text As String) As Long
    Dim tS As POINTL
    Call GetTextExtentPoint32(hdc, Text, Len(Text), tS)
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

Private Sub EnsureVisible(ByVal key As String)
'Not Working yet
'    Dim tRItem As RECT, tR As RECT, tRTest As RECT
'    Dim item As TabItem
'
'    Set item = m_Items(key)
'    LSet tRItem = item.Rectangle
'    'SetRect tR, m_lVirtualLeft + m_lLeft, m_lTop, m_lVirtualLeft + m_lLeft + m_lWidth, m_lTop + m_lHeight + 1&
'
'    'If Not IntersectRect(tRTest, tRItem, tR) Then
'    'If (tRItem.Left < tR.Left) Or (tRItem.Right > tR.Right) Then
'    Debug.Print item.Text, tRItem.Left, m_lVirtualLeft, m_lLeft
'    If tRItem.Left < 0 Then
'        Debug.Print m_lVirtualLeft
'        m_lVirtualLeft = m_lVirtualLeft - tRItem.Left
'        Debug.Print m_lVirtualLeft
'    End If

End Sub
