VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   ScaleHeight     =   2445
   ScaleWidth      =   9600
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
      Left            =   225
      TabIndex        =   0
      Top             =   135
      Width           =   9150
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'C prototype is:
' void __stdcall MD5(char* sIn, char* bufOut)

'for accessing linked in C code
Private Declare Sub MD5 Lib "project1.exe" (ByVal lpCStr As Long, ByVal lpBuf As Long)

'for debugging in the IDE
Private Declare Sub dll_MD5 Lib "md5.dll" Alias "MD5" (ByVal lpCStr As Long, ByVal lpBuf As Long)

'this is a real world example of when we might use this technique with more complex C code that
'we wouldnt want to port to VB6. This example also tests using subfunctions in C, and it has
'several runtime clibrary functions which were compiled into the obj file as well
'namely strlen, memcpy, and memset.
'
'This should be a sufficiently complex example to prove the stability of the technique..
'
'also note we didnt have to manually export the MD5 function in our vbc file
'because the request to export it was defined in the obj file and the linker
'automatically handled it for it. We could have still defined it it wouldnt hurt..

Private Sub Form_Load()
    
    Dim hash As String
    
    List1.AddItem "Known: MD5(test!) = c4d354440cb41ee38e162bc1f431e99b"
    
    hash = doMD5("test!")
    List1.AddItem "Our:   MD5(test!) = " & hash
    
    List1.AddItem IIf(isIde(), "DLL", "INTERNAL") & " implementation was used.."
    
End Sub


Function doMD5(str As String) As String
    Dim bIn() As Byte
    Dim bOut() As Byte
    Dim ret As String
    
    'convert input to C string (note null terminator was added)
    bIn() = StrConv(str & Chr(0), vbFromUnicode, &H409)
    
    'allocate space for return hash
    ReDim bOut(16)
    
    If isIde() Then
        dll_MD5 VarPtr(bIn(0)), VarPtr(bOut(0))
    Else
        MD5 VarPtr(bIn(0)), VarPtr(bOut(0))
    End If
    
    'convert raw binary hash to hex string...
    For i = 0 To UBound(bOut)
        ret = ret & Right("0" & Hex(bOut(i)), 2)
    Next
    
    doMD5 = LCase(ret)
    
End Function

Public Function isIde() As Boolean
    On Error GoTo hell
    Debug.Print 1 / 0
    isIde = False
    Exit Function
hell:
    isIde = True
End Function


