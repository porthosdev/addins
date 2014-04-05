VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   7410
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   7080
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'for accessing linked in C code
Private Declare Function to64 Lib "project1.exe" (ByVal hi As Long, ByVal lo As Long) As Currency
Private Declare Function add64 Lib "project1.exe" (ByVal a As Currency, ByVal b As Currency) As Currency
Private Declare Function sub64 Lib "project1.exe" (ByVal a As Currency, ByVal b As Currency) As Currency
Private Declare Function hex64 Lib "project1.exe" (ByVal a As Currency) As String
Private Declare Function init Lib "project1.exe" (ByVal lpfnGetProc As Long, ByVal lpfnLoadLib As Long) As Long

'for debugging in the IDE
Private Declare Function dll_to64 Lib "c_obj.dll" Alias "to64" (ByVal hi As Long, ByVal lo As Long) As Currency
Private Declare Function dll_add64 Lib "c_obj.dll" Alias "add64" (ByVal a As Currency, ByVal b As Currency) As Currency
Private Declare Function dll_sub64 Lib "c_obj.dll" Alias "sub64" (ByVal a As Currency, ByVal b As Currency) As Currency
Private Declare Function dll_hex64 Lib "c_obj.dll" Alias "hex64" (ByVal a As Currency) As String
Private Declare Function dll_init Lib "c_obj.dll" Alias "init" (ByVal lpfnGetProc As Long, ByVal lpfnLoadLib As Long) As Long

'for initilization of C functions
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

Dim IDE As Boolean

Private Sub Form_Load()
    
    Dim h As Long
    Dim lpfnGetProc As Long
    Dim lpfnLoadLib As Long
    
    IDE = isIde()
    
    h = LoadLibrary("kernel32.dll")
    lpfnGetProc = GetProcAddress(h, "GetProcAddress")
    lpfnLoadLib = GetProcAddress(h, "LoadLibraryA")
    
    If Not doinit(lpfnGetProc, lpfnLoadLib) Then
        MsgBox "Failed to initilize C function pointers!", vbInformation
        Exit Sub
    End If
    
    Dim a As Currency
    Dim b As Currency
    Dim c As Currency
    
    List1.AddItem "All functions will run from: " & IIf(IDE, "DLL", "INTERNAL")
    
    List1.AddItem "Creating two 64 bit numbers"
    a = doto64(&H11223344, 0)
    b = doto64(0, &H55667788)
    
    List1.AddItem "Adding them together"
    c = doAdd64(a, b)
    
    List1.AddItem "Converting to hex string"
    List1.AddItem "a + b = " & dohex64(c)
    
End Sub


Function doAdd64(a As Currency, b As Currency) As Currency
    If IDE Then
        doAdd64 = dll_add64(a, b)
    Else
        doAdd64 = add64(a, b)
    End If
End Function

Function dohex64(a As Currency) As String
    If IDE Then
        dohex64 = dll_hex64(a)
    Else
        dohex64 = hex64(a)
    End If
End Function

Function doto64(hi As Long, lo As Long) As Currency
    If IDE Then
        doto64 = dll_to64(hi, lo)
    Else
        doto64 = to64(hi, lo)
    End If
End Function

Function doinit(ByVal lpfnGetProc As Long, ByVal lpfnLoadLib As Long) As Boolean
    Dim v As Long
    
    If lpfnGetProc = 0 Then
        MsgBox "GetProcAddress can not be 0!"
        Exit Function
    End If
    
    If lpfnLoadLib = 0 Then
        MsgBox "LoadLibrary can not be 0!"
        Exit Function
    End If
    
    If IDE Then
        v = dll_init(lpfnGetProc, lpfnLoadLib)
    Else
        v = init(lpfnGetProc, lpfnLoadLib)
    End If
    
    If v = 0 Then doinit = True
    
End Function


Public Function isIde() As Boolean
    On Error GoTo hell
    Debug.Print 1 / 0
    isIde = False
    Exit Function
hell:
    isIde = True
End Function


