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

'when running in the IDE, module 1 functions will work as coded using the external dll
'however when compiled, module1.obj will be swapped out with the obj file for the dll
'and its functions coded in C will be compiled directly into the vb6 exe.

'for initilization of C functions
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

Private Sub Form_Load()
    
    Dim h As Long
    Dim lpfnGetProc As Long
    Dim lpfnLoadLib As Long

    h = LoadLibrary("kernel32.dll")
    lpfnGetProc = GetProcAddress(h, "GetProcAddress")
    lpfnLoadLib = GetProcAddress(h, "LoadLibraryA")
    
    If init(lpfnGetProc, lpfnLoadLib) <> 0 Then
        MsgBox "Failed to initilize C function pointers!", vbInformation
        Exit Sub
    End If
    
    Dim a As Currency
    Dim b As Currency
    Dim c As Currency
        
    List1.AddItem "Creating two 64 bit numbers"
    a = to64(&H11223344, 0)
    b = to64(0, &H55667788)
    
    List1.AddItem "Adding them together"
    c = add64(a, b)
    
    'MsgBox "VarPtr(c) = " & Hex(VarPtr(c))
    
    List1.AddItem "Converting to hex string"
    List1.AddItem "a + b = " & hex64(c)
    
    h = GetModuleHandle("c_module.dll")
    List1.AddItem "Functions were run from " & IIf(h = 0, "INTERNAL", "DLL")
    
End Sub

