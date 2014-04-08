Attribute VB_Name = "mPublic"
Option Explicit

Public Const ACTION_SETTEXT As Long = 1
Public Const ACTION_ACTIVATE As Long = 2

Public Const LF_FACESIZE As Long = 32
Public Const WM_PARENTNOTIFY As Long = &H210
Public Const WM_SIZE As Long = &H5
Public Const WM_NCACTIVATE As Long = &H86
Public Const WM_NCCALCSIZE As Long = &H83
Public Const WM_CREATE As Long = &H1
Public Const WM_DESTROY As Long = &H2
Public Const WM_NCPAINT As Long = &H85

Public Const WM_NCMOUSEMOVE As Long = &HA0
Public Const WM_NCLBUTTONDOWN As Long = &HA1
Public Const WM_NCLBUTTONUP As Long = &HA2
Public Const WM_NCRBUTTONDOWN As Long = &HA4
Public Const WM_NCHITTEST As Long = &H84

Public Const WM_SETTEXT As Long = &HC
Public Const WM_SETTINGCHANGE As Long = &H1A
Public Const WM_MDIACTIVATE As Long = &H222
Public Const WM_MDIGETACTIVE As Long = &H229
Public Const WM_CLOSE As Long = &H10
Public Const WM_NCMOUSELEAVE As Long = 674 '&H2A2
Public Const WM_STYLECHANGED As Long = &H7D


Public Const GW_CHILD As Long = 5
Public Const GW_HWNDNEXT As Long = 2

Public Const SIZE_MINIMIZED As Long = 1
'Public Const SIZE_MAXIMIZED As Long = 2

Public Const SWP_FRAMECHANGED As Long = &H20
Public Const SWP_NOACTIVATE As Long = &H10
Public Const SWP_NOMOVE As Long = &H2
Public Const SWP_NOSIZE As Long = &H1
Public Const SWP_NOZORDER As Long = &H4

Public Const TRANSPARENT As Long = 1
Public Const DT_LEFT As Long = &H0
Public Const DT_CENTER As Long = &H1
Public Const DT_VCENTER As Long = &H4
Public Const DT_RIGHT As Long = &H2
Public Const DT_SINGLELINE As Long = &H20

'Public Const SW_MAXIMIZE As Long = 3
Public Const MF_STRING As Long = &H0&
Public Const MF_GRAYED As Long = &H1&
Public Const MF_ENABLED As Long = &H0&
Public Const MF_SEPARATOR As Long = &H800&
Public Const MF_CHECKED As Long = &H8&

Public Const TPM_NONOTIFY As Long = &H80&
Public Const TPM_RETURNCMD As Long = &H100&
Public Const GWL_WNDPROC As Long = -4
Public Const PS_SOLID As Long = 0
Public Const DI_NORMAL = 3
Public Const VK_LBUTTON As Long = &H1
Public Const HTBORDER As Long = 18
Public Const RDW_INVALIDATE As Long = &H1
Public Const RDW_UPDATENOW As Long = &H100
Public Const RDW_ERASE As Long = &H4
Public Const TME_CANCEL As Long = &H80000000
Public Const TME_LEAVE As Long = &H2
Public Const TME_NONCLIENT As Long = &H10

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type POINTL
    x As Long
    y As Long
End Type

Public Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Public Type tagTRACKMOUSEEVENT
    cbSize As Long
    dwFlags As Long
    hwndTrack As Long
    dwHoverTime As Long
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128 ' Maintenance string for PSS usage
End Type

Public Declare Function GetProcAddress _
               Lib "kernel32.dll" (ByVal hModule As Long, _
                                   ByVal lpProcName As String) As Long
Public Declare Function MapWindowPoints _
               Lib "user32.dll" (ByVal hwndFrom As Long, _
                                 ByVal hwndTo As Long, _
                                 ByRef lppt As Any, _
                                 ByVal cPoints As Long) As Long
Public Declare Function FindWindowEx _
               Lib "user32.dll" _
               Alias "FindWindowExA" (ByVal hWnd1 As Long, _
                                      ByVal hWnd2 As Long, _
                                      ByVal lpsz1 As String, _
                                      ByVal lpsz2 As String) As Long
Public Declare Function GetWindow _
               Lib "user32.dll" (ByVal hwnd As Long, _
                                 ByVal wCmd As Long) As Long
Public Declare Function SelectObject _
               Lib "gdi32.dll" (ByVal hDC As Long, _
                                ByVal hObject As Long) As Long
Public Declare Function DeleteObject _
               Lib "gdi32.dll" (ByVal hObject As Long) As Long
Public Declare Function FillRect _
               Lib "user32.dll" (ByVal hDC As Long, _
                                 ByRef lpRect As RECT, _
                                 ByVal hBrush As Long) As Long
Public Declare Function DrawText _
               Lib "user32.dll" _
               Alias "DrawTextA" (ByVal hDC As Long, _
                                  ByVal lpStr As String, _
                                  ByVal nCount As Long, _
                                  ByRef lpRect As RECT, _
                                  ByVal wFormat As Long) As Long
Public Declare Function DrawIconEx _
               Lib "user32.dll" (ByVal hDC As Long, _
                                 ByVal xLeft As Long, _
                                 ByVal yTop As Long, _
                                 ByVal hIcon As Long, _
                                 ByVal cxWidth As Long, _
                                 ByVal cyWidth As Long, _
                                 ByVal istepIfAniCur As Long, _
                                 ByVal hbrFlickerFreeDraw As Long, _
                                 ByVal diFlags As Long) As Long
Public Declare Function SetRect _
               Lib "user32.dll" (ByRef lpRect As RECT, _
                                 ByVal X1 As Long, _
                                 ByVal Y1 As Long, _
                                 ByVal X2 As Long, _
                                 ByVal Y2 As Long) As Long
Public Declare Function IntersectRect _
               Lib "user32.dll" (ByRef lpDestRect As RECT, _
                                 ByRef lpSrc1Rect As RECT, _
                                 ByRef lpSrc2Rect As RECT) As Long
Public Declare Function DestroyMenu _
               Lib "user32.dll" (ByVal hMenu As Long) As Long
Public Declare Function CreatePopupMenu _
               Lib "user32.dll" () As Long
Public Declare Function TrackPopupMenu _
               Lib "user32.dll" (ByVal hMenu As Long, _
                                 ByVal wFlags As Long, _
                                 ByVal x As Long, _
                                 ByVal y As Long, _
                                 ByVal nReserved As Long, _
                                 ByVal hwnd As Long, _
                                 ByRef lprc As Any) As Long
Public Declare Function AppendMenu _
               Lib "user32.dll" _
               Alias "AppendMenuA" (ByVal hMenu As Long, _
                                    ByVal wFlags As Long, _
                                    ByVal wIDNewItem As Long, _
                                    ByVal lpNewItem As Any) As Long
Public Declare Function EnableMenuItem _
               Lib "user32.dll" (ByVal hMenu As Long, _
                                 ByVal wIDEnableItem As Long, _
                                 ByVal wEnable As Long) As Long
Public Declare Function SendMessage _
               Lib "user32.dll" _
               Alias "SendMessageA" (ByVal hwnd As Long, _
                                     ByVal wMsg As Long, _
                                     ByVal wParam As Long, _
                                     ByRef lParam As Any) As Long
Public Declare Function PostMessage _
               Lib "user32.dll" _
               Alias "PostMessageA" (ByVal hwnd As Long, _
                                     ByVal wMsg As Long, _
                                     ByVal wParam As Long, _
                                     ByVal lParam As Long) As Long
Public Declare Function SetWindowPos _
               Lib "user32.dll" (ByVal hwnd As Long, _
                                 ByVal hWndInsertAfter As Long, _
                                 ByVal x As Long, _
                                 ByVal y As Long, _
                                 ByVal cx As Long, _
                                 ByVal cy As Long, _
                                 ByVal wFlags As Long) As Long
Public Declare Function GetWindowDC _
               Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC _
               Lib "user32.dll" (ByVal hwnd As Long, _
                                 ByVal hDC As Long) As Long
Public Declare Function GetWindowRect _
               Lib "user32.dll" (ByVal hwnd As Long, _
                                 ByRef lpRect As RECT) As Long
Public Declare Function SetBkMode _
               Lib "gdi32.dll" (ByVal hDC As Long, _
                                ByVal nBkMode As Long) As Long
Public Declare Function SetTextColor _
               Lib "gdi32.dll" (ByVal hDC As Long, _
                                ByVal crColor As Long) As Long
Public Declare Function GetClassName _
               Lib "user32.dll" _
               Alias "GetClassNameA" (ByVal hwnd As Long, _
                                      ByVal lpClassName As String, _
                                      ByVal nMaxCount As Long) As Long
Public Declare Function GetWindowText _
               Lib "user32.dll" _
               Alias "GetWindowTextA" (ByVal hwnd As Long, _
                                       ByVal lpString As String, _
                                       ByVal cch As Long) As Long
'Public Declare Function ShowWindow _
'               Lib "user32.dll" (ByVal hwnd As Long, _
'                                 ByVal nCmdShow As Long) As Long
Public Declare Function GetCursorPos _
               Lib "user32.dll" (ByRef lpPoint As POINTL) As Long
Public Declare Function GlobalAlloc _
               Lib "kernel32.dll" (ByVal wFlags As Long, _
                                   ByVal dwBytes As Long) As Long
Public Declare Sub CopyMemory _
               Lib "kernel32.dll" _
               Alias "RtlMoveMemory" (ByRef Destination As Any, _
                                      ByRef Source As Any, _
                                      ByVal Length As Long)
Public Declare Function GlobalFree _
               Lib "kernel32.dll" (ByVal hMem As Long) As Long
Public Declare Function CallWindowProc _
               Lib "user32.dll" _
               Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                        ByVal hwnd As Long, _
                                        ByVal msg As Long, _
                                        ByVal wParam As Long, _
                                        ByVal lParam As Long) As Long
Public Declare Function SetWindowLong _
               Lib "user32.dll" _
               Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                       ByVal nIndex As Long, _
                                       ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong _
               Lib "user32.dll" _
               Alias "GetWindowLongA" (ByVal hwnd As Long, _
                                       ByVal nIndex As Long) As Long
Public Declare Function GetModuleHandle _
               Lib "kernel32.dll" _
               Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function IsWindow _
               Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function PtInRect _
               Lib "user32.dll" (ByRef lpRect As RECT, _
                                 ByVal x As Long, _
                                 ByVal y As Long) As Long
Public Declare Function CreateSolidBrush _
               Lib "gdi32.dll" (ByVal crColor As Long) As Long
Public Declare Function CreatePen _
               Lib "gdi32.dll" (ByVal nPenStyle As Long, _
                                ByVal nWidth As Long, _
                                ByVal crColor As Long) As Long
Public Declare Function MoveToEx _
               Lib "gdi32.dll" (ByVal hDC As Long, _
                                ByVal x As Long, _
                                ByVal y As Long, _
                                ByRef lpPoint As Any) As Long
Public Declare Function LineTo _
               Lib "gdi32.dll" (ByVal hDC As Long, _
                                ByVal x As Long, _
                                ByVal y As Long) As Long
Public Declare Function OleTranslateColor _
               Lib "oleaut32.dll" (ByVal lOleColor As Long, _
                                   ByVal lHPalette As Long, _
                                   ByRef lColorRef As Long) As Long
Public Declare Function GetTextExtentPoint32 _
               Lib "gdi32.dll" _
               Alias "GetTextExtentPoint32A" (ByVal hDC As Long, _
                                              ByVal lpsz As String, _
                                              ByVal cbString As Long, _
                                              ByRef lpSize As POINTL) As Long

Public Declare Function SetWindowOrgEx _
               Lib "gdi32.dll" (ByVal hDC As Long, _
                                ByVal nX As Long, _
                                ByVal nY As Long, _
                                ByRef lpPoint As Any) As Long
               
Public Declare Function GetAsyncKeyState _
               Lib "user32.dll" (ByVal vKey As Long) As Integer
Public Declare Function SetRectEmpty _
               Lib "user32.dll" (ByRef lpRect As RECT) As Long
Public Declare Function RedrawWindow _
               Lib "user32.dll" (ByVal hwnd As Long, _
                                 ByRef lprcUpdate As Any, _
                                 ByVal hrgnUpdate As Long, _
                                 ByVal fuRedraw As Long) As Long
Public Declare Function TrackMouseEvent _
               Lib "user32.dll" (ByRef lpEventTrack As tagTRACKMOUSEEVENT) As Long

Public Declare Function ModifyMenu _
               Lib "user32.dll" _
               Alias "ModifyMenuA" (ByVal hMenu As Long, _
                                    ByVal nPosition As Long, _
                                    ByVal wFlags As Long, _
                                    ByVal wIDNewItem As Long, _
                                    lpString As Any) As Long

Private Declare Function GetVersionEx _
                Lib "kernel32.dll" _
                Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long

#If in_addin Then
    Public myLoader As Loader
#End If

Public Enum enWinVersion
  enWin95 = 1
  enWinNT = 2
  enWin98 = 3
  enWin2000 = 4
  enWinXP = 5
End Enum

'This is for Win98, we need to place a reference lock before we can safely
'remove the subclass in Win98
Public lockChild As MDIChildWindow
Public lockMDIHook As cSubclass

'This one public var is dedicated to MZTools AddIn :)
Public mzToolsDetected As Boolean

Function LoWord(lDWord As Long) As Integer
    If lDWord And &H8000& Then
        LoWord = lDWord Or &HFFFF0000
    Else
        LoWord = lDWord And &HFFFF&
    End If
End Function

Function HiWord(lDWord As Long) As Integer
    HiWord = (lDWord And &HFFFF0000) \ &H10000
End Function

Function GetWinText(hwnd As Long, Optional className As Boolean = False) As String
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

Function GetOSVersion() As enWinVersion
  'Get Windows version
  Dim tOS As OSVERSIONINFO
  
  tOS.dwOSVersionInfoSize = Len(tOS)
  GetVersionEx tOS
  
  If tOS.dwMajorVersion > 4& Then
    If tOS.dwMinorVersion > 0& Then
      GetOSVersion = enWinXP
    ElseIf tOS.dwMinorVersion = 0& Then
      GetOSVersion = enWin2000
    End If
  Else
    If tOS.dwPlatformId = 1& Then
      If tOS.dwMinorVersion > 0& Then
        GetOSVersion = enWin98
      Else
        GetOSVersion = enWin95
      End If
    ElseIf tOS.dwPlatformId = 2& Then
      GetOSVersion = enWinNT 'Should be check for NT 3.5 but we're not going that far
    End If
  End If
End Function
