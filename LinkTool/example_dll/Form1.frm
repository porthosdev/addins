VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "VB6 Standard DLL Test Project"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   3555
      Top             =   1665
   End
   Begin MSComctlLib.ListView lv 
      Height          =   1815
      Left            =   270
      TabIndex        =   4
      Top             =   2745
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   3201
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ListView Test"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   270
      TabIndex        =   3
      Top             =   810
      Width           =   2940
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Call C Export"
      Height          =   375
      Left            =   3420
      TabIndex        =   2
      Top             =   855
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   465
      Left            =   315
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   180
      Width           =   2940
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test Timer"
      Height          =   420
      Left            =   3420
      TabIndex        =   0
      Top             =   225
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'example of how to trigger code back in main (single threaded) exe..
Private Declare Sub MyCExport Lib "use_dll.exe" ()

Private Sub Command1_Click()
    Timer1.Enabled = Not Timer1.Enabled
End Sub

Private Sub Command2_Click()
    MyCExport
End Sub

Private Sub Form_Load()

    'note: center screen property doesnt work..
    Me.Left = (Screen.Width / 2) - (Me.Width / 2)
    Me.Top = (Screen.Height / 2) - (Me.Height / 2)
    
    For i = 0 To 10
        List1.AddItem "item " & i
    Next
    
    For i = 0 To 10
        lv.ListItems.Add , , "item " & i
    Next
    
End Sub

Private Sub Timer1_Timer()
    Text1 = Now
End Sub
