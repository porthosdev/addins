VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEnableHook 
      Caption         =   "Enable Hook"
      Height          =   510
      Left            =   7470
      TabIndex        =   4
      Top             =   3375
      Width           =   1545
   End
   Begin VB.CommandButton cmdDisableHook 
      Caption         =   "Disable Hook"
      Height          =   510
      Left            =   5940
      TabIndex        =   3
      Top             =   3375
      Width           =   1410
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove Hook"
      Height          =   510
      Left            =   9135
      TabIndex        =   2
      Top             =   3375
      Width           =   1230
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GetSaveFileName"
      Height          =   465
      Left            =   1395
      TabIndex        =   1
      Top             =   3420
      Width           =   1410
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   10365
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'standard declares for calling it normally
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (ofn As OPENFILENAME) As Long

Private Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type



Private Sub cmdDisableHook_Click()
    List1.AddItem "Disable hook: " & DisableHook(lpGetSaveFileName)
End Sub

Private Sub cmdEnableHook_Click()
    List1.AddItem "Enable hook: " & EnableHook(lpGetSaveFileName)
End Sub

Private Sub cmdRemove_Click()
    List1.AddItem "Remove hook: " & RemoveHook(lpGetSaveFileName)
End Sub



Private Sub Command1_Click()
    Dim o As OPENFILENAME
    Dim ret As Long
    
    o.lStructSize = LenB(o)
    o.lpstrFile = "Project1.exe" & Chr(0) & Space(255)
    o.nMaxFile = 255
    o.lpstrDefExt = ".exe"
    
    'DebugBreak
    ret = GetSaveFileName(o)
    
    MsgBox ret
    
End Sub

 

Private Sub Form_Load()
    
    Dim h As Long
    Dim ret As Long
    Dim lpMsg As Long
    
    hHookLib = LoadLibrary("hooklib.dll")
    If hHookLib = 0 Then hHookLib = LoadLibrary(App.Path & "\hooklib.dll")
    If hHookLib = 0 Then hHookLib = LoadLibrary(App.Path & "\..\hooklib.dll")
    
    If hHookLib = 0 Then
        List1.AddItem "Could not find hooklib.dll compile or download from github."
        Exit Sub
    End If
    
    List1.AddItem "Hooklib base address: 0x" & Hex(h)
    
    'this is optional but were debugging the library so..
    SetDebugHandler AddressOf DebugMsgHandler, 1
    
    h = LoadLibrary("comdlg32.dll")
    lpGetSaveFileName = GetProcAddress(h, "GetSaveFileNameA")
    
    If lpGetSaveFileName = 0 Then
        List1.AddItem "Could not GetProcAddress(GetSaveFileNameA)"
        Exit Sub
    End If
        
    ret = HookFunction(lpGetSaveFileName, AddressOf My_GetSaveFileName, "GetSaveFileNameA", ht_jmp)
    
    'ret = HookFunction(0, 0, "GetSaveFileNameA", ht_jmp) 'to test GetHookError()
    If ret = 0 Then
        lpMsg = GetHookError()
        List1.AddItem "Hook Function failed msg: " & CStringToVBString(lpMsg)
        Exit Sub
    End If
    
    List1.AddItem "GetSaveFileNameA Successfully Hooked.."
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    RemoveAllHooks
    UnInitilizeHookLib
    FreeLibrary hHookLib
End Sub
