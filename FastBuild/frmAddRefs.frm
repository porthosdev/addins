VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAddRefs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fast Build - Add References"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   5475
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDetails 
      Height          =   1050
      Left            =   135
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   4320
      Width           =   5145
   End
   Begin MSComctlLib.ListView lvFiltered 
      Height          =   1950
      Left            =   1395
      TabIndex        =   6
      Top             =   990
      Visible         =   0   'False
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   3440
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   765
      TabIndex        =   3
      Top             =   3195
      Width           =   3795
      Begin VB.TextBox txtSearch 
         Height          =   285
         Left            =   855
         TabIndex        =   4
         Top             =   90
         Width           =   2625
      End
      Begin VB.Label Label2 
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   3600
         TabIndex        =   8
         Top             =   135
         Width           =   195
      End
      Begin VB.Label Label1 
         Caption         =   "Search"
         Height          =   285
         Left            =   135
         TabIndex        =   5
         Top             =   135
         Width           =   645
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4155
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   7329
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Components"
      TabPicture(0)   =   "frmAddRefs.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lv"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "References"
      TabPicture(1)   =   "frmAddRefs.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lv2"
      Tab(1).ControlCount=   1
      Begin MSComctlLib.ListView lv 
         Height          =   3030
         Left            =   180
         TabIndex        =   1
         Top             =   135
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   5345
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lv2 
         Height          =   2850
         Left            =   -74865
         TabIndex        =   2
         Top             =   135
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   5027
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
      End
   End
End
Attribute VB_Name = "frmAddRefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Author: David Zimmer
' Site:   http://sandsprite.com
'
Dim reg As New CReg
Dim tlbs As Collection
Dim selEntry As CEntry

Private Sub Form_Load()
   
    With lv
        lv2.Move .Left, .Top, .Width, .Height
        lvFiltered.Move .Left + SSTab1.Left, .Top + SSTab1.Top, .Width, .Height
    End With
    
    lv.ColumnHeaders(1).Width = lv.Width
    lv2.ColumnHeaders(1).Width = lv2.Width
    lvFiltered.ColumnHeaders(1).Width = lv.Width
    
    Set tlbs = New Collection
    
    'Me.Visible = True
    'Me.Refresh
    'DoEvents
    
    reg.hive = HKEY_CLASSES_ROOT
    BuildComponentList
    
End Sub


Function BuildReferenceList()
    
    Dim clsids() As String
    Dim clsid
    Dim li As ListItem
    Dim e As CEntry
    Dim tmp As CEntry
    Dim vers() As String
    Dim revs() As String
    Dim c As New Collection
    Dim lia As ListItem
    
    If reg.hive = HKEY_CLASSES_ROOT Then
        clsids = reg.EnumKeys("\TypeLib")
    Else
        'Stop
        'clsids = reg.EnumKeys("\SOFTWARE\Classes\CLSID")
    End If
    
    For Each clsid In clsids

        Set e = New CEntry
        e.clsid = clsid
        
        'If clsID = "{00025E01-0000-0000-C000-000000000046}" Then Stop
                
         vers() = reg.EnumKeys("\TypeLib\" & clsid)
         If AryIsEmpty(vers) Then GoTo nextone
                
         revs() = reg.EnumKeys("\TypeLib\" & clsid & "\" & vers(UBound(vers)))
         If AryIsEmpty(revs) Then GoTo nextone
         
         With e
            e.name = reg.ReadValue("\TypeLib\" & clsid & "\" & vers(UBound(vers)), "")
            e.path = reg.ReadValue("\TypeLib\" & clsid & "\" & vers(UBound(vers)) & "\" & revs(0) & "\win32", "")
            .version = vers(UBound(vers)) & "." & revs(0)
            
            e.path = ValidatePath(e.path)
            
            If FileExists(e.path) And Len(.name) > 0 Then
                If Not KeyExistsInCollection(.name, c) Then
                    Set li = lv2.ListItems.Add(, , .name)
                    Set li.Tag = e
                    If RefAlreadyExists(.clsid) Then
                        .AlreadyReferenced = True
                        li.Checked = True
                    End If
                    c.Add e, .name
                End If
            End If
            
        End With
        
nextone:

   Next
End Function

Function ValidatePath(fpath As String) As String
    
    Dim a As Long
    Dim b As Long
    'example input: C:\WINDOWS\system32\catsrvut.dll\2
    
    a = InStrRev(fpath, ".")
    b = InStrRev(fpath, "\")
    If b > a Then
        ValidatePath = Mid(fpath, 1, b - 1)
    Else
        ValidatePath = fpath
    End If
    
End Function



Function GetExtension(path) As String
    If Len(path) = 0 Then Exit Function
    tmp = Split(path, "\")
    ub = tmp(UBound(tmp))
    If InStr(1, ub, ".") > 0 Then
       GetExtension = LCase(Mid(ub, InStrRev(ub, "."), Len(ub)))
    Else
       GetExtension = ""
    End If
End Function


Function BuildComponentList()
    
    Dim clsids() As String
    Dim clsid
    Dim li As ListItem
    Dim e As CEntry
    Dim tmp As CEntry
    
    'Const full = "\SOFTWARE\Classes\CLSID\{0DE63042-EB7E-4449-BF13-7FF73866F20E}\Implemented Categories\{40FC6ED5-2438-11CF-A3DB-080036F12502}"
    Const catid_control = "\Implemented Categories\{40FC6ED4-2438-11cf-A3DB-080036F12502}"
    Const catid_programmable = "\Implemented Categories\{40FC6ED5-2438-11CF-A3DB-080036F12502}"
    Const server = "\InprocServer32"
    
    If reg.hive = HKEY_CLASSES_ROOT Then
        clsids = reg.EnumKeys("\CLSID")
    Else
        clsids = reg.EnumKeys("\SOFTWARE\Classes\CLSID")
    End If
    
    For Each clsid In clsids

        Set e = New CEntry
        e.clsid = clsid
        
        'If clsID = "{66CBC149-A49F-48F9-B17A-6A3EA9B42A87}" Then Stop
        
        If reg.hive = HKEY_CLASSES_ROOT Then
            clsid = "\CLSID\" & clsid
        Else
            clsid = "\SOFTWARE\Classes\CLSID\" & clsid
        End If
        
        With e
        
            .isControl = reg.keyExists(clsid & "\Control")
            If .isControl = False Then
                If reg.keyExists(clsid & catid_control) Then .isControl = True
            End If
            
            '.isProgrammable = reg.keyExists(clsID & "\Programmable")
            'If .isProgrammable = False Then
            '    If reg.keyExists(clsID & catid_programmable) Then .isProgrammable = True
            'End If
            
            .typeLib = reg.ReadValue(clsid & "\typeLib", "")
            
            If Len(.typeLib) > 0 Then
            
                'If e.isControl And KeyExistsInCollection(.typeLib, tlbs) Then
                '    Set tmp = tlbs(.typeLib)
                '    If Not tmp.isControl Then tlbs.Remove tmp.typeLib   'we will update it below..
                'End If
                        
                'If (.isControl Or .isProgrammable) And Not KeyExistsInCollection(.typeLib, tlbs) Then
                If .isControl And Not KeyExistsInCollection(.typeLib, tlbs) Then
                    .name = GetName(.typeLib)
                    If Len(.name) > 0 Then
                        .path = reg.ReadValue(clsid & "\InprocServer32", "")
                        .progID = reg.ReadValue(clsid & "\ProgID", "")
                        .version = reg.ReadValue(clsid & "\version", "")
                        tlbs.Add e, .typeLib
                        
                        If e.isControl Then
                            Set li = lv.ListItems.Add(, , .name)
                            Set li.Tag = e
                            .AlreadyReferenced = RefAlreadyExists(.clsid)
                            li.Checked = .AlreadyReferenced
                        'ElseIf e.isProgrammable Then
                        '    Set li = lv2.ListItems.Add(, , .name)
                        '    Set li.Tag = e
                        End If
                    End If
                    
                End If
            End If
            
        End With

   Next
End Function

Function ExistsInLV(s, lv As ListView) As Boolean
    Dim li As ListItem
    For Each li In lv.ListItems
        If li.Text = s Then ExistsInLV = True
    Next
End Function

Private Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

Function KeyExistsInCollection(key As String, c As Collection) As Boolean
    On Error GoTo hell
    Dim x
    Set x = c(key)
    KeyExistsInCollection = True
    Exit Function
hell:
End Function

Function GetName(typeLibID As String)
    Dim keys() As String
    Dim k, v As String, base As String

    If reg.hive = HKEY_CLASSES_ROOT Then
        base = "\TypeLib\" & typeLibID
    Else
        base = "\SOFTWARE\Classes\TypeLib\" & typeLibID
    End If
    
    keys() = reg.EnumKeys(base)
    If AryIsEmpty(keys) Then Exit Function
    For Each k In keys
       v = reg.ReadValue(base & "\" & k, "")
       If Len(v) > 0 Then
           GetName = v
           Exit Function
       End If
    Next
End Function



Function FileExists(path) As Boolean
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True _
  Else FileExists = False
End Function

Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function
 

Private Sub Label2_Click()
    MsgBox "Case insensitive search. Type 'checked' to see active references", vbInformation
End Sub

Private Sub lvFiltered_ItemCheck(ByVal item As MSComctlLib.ListItem)
    
    Dim li As ListItem
    Dim llv As ListView
    
    'we need to sync parent list..
    If SSTab1.Tab = 0 Then Set llv = lv Else Set llv = lv2
    
    For Each li In llv.ListItems
        If li.Text = item.Text Then
            li.Checked = True 'apparently this doesnt fire the _ItemCheck event like it thought..
            Exit For
        End If
    Next
    
    HandleItemCheck item
    
End Sub

Private Sub lv2_ItemCheck(ByVal item As MSComctlLib.ListItem)
    HandleItemCheck item
End Sub

Private Sub lv_ItemCheck(ByVal item As MSComctlLib.ListItem)
    HandleItemCheck item
End Sub

Sub HandleItemCheck(ByVal item As MSComctlLib.ListItem)

    On Error GoTo hell
    
    Set selEntry = item.Tag
    txtDetails = selEntry.ToString()
    
    If selEntry Is Nothing Then Exit Sub
       
    Dim r As Reference
    Dim guid As String
    
    guid = selEntry.clsid
    
    If item.Checked Then
        Set r = VBInstance.ActiveVBProject.References.AddFromGuid(guid, 0, 0)
        selEntry.AlreadyReferenced = True
        If r Is Nothing Then
            MsgBox "Could not add reference to " & guid
            Exit Sub
        End If
    Else
        Set r = GetReference(guid)
        If r Is Nothing Then
            MsgBox "Could not find reference to " & guid
            Exit Sub
        End If
        'this can fail for default references..a boolean return would have been nice..
        'we should recheck getreference and recheck box if it failed..but to lazy for small bug..
        VBInstance.ActiveVBProject.References.Remove r
        selEntry.AlreadyReferenced = False
    End If
    
    
    Exit Sub
hell:
    MsgBox "Error: " & Err.Description
    
End Sub

Private Sub lv_ItemClick(ByVal item As MSComctlLib.ListItem)
    Set selEntry = item.Tag
    txtDetails = selEntry.ToString()
End Sub

Private Sub lv2_ItemClick(ByVal item As MSComctlLib.ListItem)
    Set selEntry = item.Tag
    txtDetails = selEntry.ToString()
End Sub

Private Sub lvFiltered_ItemClick(ByVal item As MSComctlLib.ListItem)
    Set selEntry = item.Tag
    txtDetails = selEntry.ToString()
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)

    'load on demand to reduce startup time
    If SSTab1.Tab = 1 And lv2.ListItems.Count = 0 Then BuildReferenceList
    
    If Len(txtSearch) > 0 Then txtSearch_Change

End Sub

Private Sub txtSearch_Change()

    If Len(txtSearch) = 0 Then
        lvFiltered.Visible = False
        Exit Sub
    End If
    
    Dim li As ListItem
    Dim li2 As ListItem
    Dim llv As ListView
    
    If SSTab1.Tab = 0 Then Set llv = lv Else Set llv = lv2
    
    lvFiltered.Visible = True
    lvFiltered.ListItems.Clear
    
    For Each li In llv.ListItems
        If txtSearch = "checked" And li.Checked Then
            Set li2 = lvFiltered.ListItems.Add(, , li.Text)
            Set li2.Tag = li.Tag
            li2.Checked = li.Checked
        ElseIf InStr(1, li.Text, txtSearch, vbTextCompare) > 0 Then
            Set li2 = lvFiltered.ListItems.Add(, , li.Text)
            Set li2.Tag = li.Tag
            li2.Checked = li.Checked
        End If
    Next
        
        
End Sub

Private Function RefAlreadyExists(clsid As String) As Boolean
    Dim r As Reference
    For Each r In VBInstance.ActiveVBProject.References
        If InStr(1, r.guid, clsid, vbTextCompare) > 0 Then
            RefAlreadyExists = True
            Exit Function
        End If
    Next
End Function

Private Function GetReference(clsid As String) As Reference
    Dim r As Reference
    For Each r In VBInstance.ActiveVBProject.References
        If InStr(1, r.guid, clsid, vbTextCompare) > 0 Then
            Set GetReference = r
            Exit Function
        End If
    Next
End Function

