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
Private Declare Sub inc Lib "project1.exe" (ByVal a As Long)
Private Declare Function retrieve Lib "project1.exe" () As Long

'for debugging in the IDE
Private Declare Sub dll_inc Lib "c_obj.dll" Alias "inc" (ByVal a As Long)
Private Declare Function dll_retrieve Lib "c_obj.dll" Alias "retrieve" () As Long

Private Sub Form_Load()
    
    For i = 1 To 10
        do_inc 1
    Next
    
    List1.AddItem "Running from: " & IIf(IDE, "DLL", "INTERNAL")
    
    List1.AddItem "End value: " & do_retrieve()
    
    
End Sub


Sub do_inc(a As Long)
    If isIde Then
        Call dll_inc(a)
    Else
        Call inc(a)
    End If
End Sub

Function do_retrieve() As Long
    If isIde Then
        do_retrieve = dll_retrieve()
    Else
        do_retrieve = retrieve()
    End If
End Function

Public Function isIde() As Boolean
    On Error GoTo hell
    Debug.Print 1 / 0
    isIde = False
    Exit Function
hell:
    isIde = True
End Function


