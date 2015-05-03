Attribute VB_Name = "Utils"
Option Explicit

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, _
    ByVal lParam As Long) As Long
    
Private Const WM_SETREDRAW As Long = &HB

Public g_VBInstance As VBIDE.VBE

Const BLU = "<FONT COLOR='#000088'>"
Const GRN = "<FONT COLOR='#008800'>"
Const CF = "</FONT>"

Dim RW() As String
Dim Special() As String



Sub Freeze(hwnd As Long)
    SendMessageLong hwnd, WM_SETREDRAW, False, &O0
End Sub

Sub Unfreeze(hwnd As Long)
    SendMessageLong hwnd, WM_SETREDRAW, True, &O0
End Sub

Sub SaveMySetting(key, value)
    On Error Resume Next
    SaveSetting "CodeView", "General", key, value
End Sub

Function GetMySetting(key, def)
    On Error Resume Next
    GetMySetting = GetSetting("CodeView", "General", key, def)
End Function

Sub ClearChildNodes(Tree As TreeView, NodeName As String, Optional NodeObject As Node)
    
    On Error Resume Next 'GoTo hell

    Dim n As Node
    
    If Not NodeObject Is Nothing Then
        Set n = NodeObject
    Else
        Set n = Tree.Nodes(NodeName)
    End If
    
    If n Is Nothing Then Exit Sub
    
    While n.Children > 0
        If n.Child.Children > 0 Then
            ClearChildNodes Tree, n.Child
        End If
        Tree.Nodes.Remove n.Child.Index
    Wend
    
        Exit Sub
hell:
    'getting invalid node key error sometimes..
    'MsgBox "Err in ClearChildNodes: " & Err.Description

End Sub

'simple and bulky but ok..ripped from somewhere eons ago..
Function htmlize(txt As String) As String

  'On Error Resume Next
  
    If AryIsEmpty(RW) Then
        'case & space after word is important !
        RW = Split("Const ,Else ,ElseIf ,If ,Alias ,And , As,Base ,Binary ,Boolean," & _
                    "Byte ,ByVal ,Call ,Case ,CBool ,CByte ,CCur ,CDate ,CDbl ,CDec ," & _
                    "CInt ,CLng ,Close ,Compare ,Const ,CSng ,CStr ,Currency ,CVar ," & _
                    "CVErr ,Decimal ,Declare ,DefBool ,DefByte ,DefCur ,DefDate ," & _
                    "DefDbl ,DefDec ,DefInt ,DefLng ,DefObj ,DefSng ,DefStr ,DefVar ," & _
                    "Dim ,Do ,Double ,Each ,End ,Enum ,Eqv ,Erase ,Error ," & _
                    "Exit ,Explicit ,False ,For ,Function ,Get ,Global ,GoSub ,GoTo ," & _
                    "Imp ,In ,Input ,Input ,Integer ,Is ,LBound ,Let ,Lib ,Like ,Line ,Lock ," & _
                    "Long ,Loop ,LSet ,Name ,New ,Next ,Not ,Object ,Open ,Option ,On ,Or ," & _
                    "Output ,Preserve ,Print ,Private ,Property ,Public ,Put ,Random ," & _
                    "Read ,ReDim ,Resume ,Return ,RSet ,Seek ,Select ,Set ,Single ,Spc ," & _
                    "Static ,String,Stop ,Sub ,Tab ,Then ,True ,UBound ,Variant ,While ," & _
                    "Wend ,With ,Empty " _
              , ",")
              
        'these handle some other common casekeywords that dont fit the word<space> profile
        'necessary because this search is done on a macro scale and not by trying to anlyze
        'each word or character it comes across
        Special = Split("CLng(,CInt(,CBool(,CByte(,CStr(,True),False),Empty),(True,(False,(Empty", ",")
    End If
    

    Dim comment, code, lastDq, lastSq, tmp, i, it
    
    tmp = Split(txt, vbCrLf)
    For i = 0 To UBound(tmp)
        'Debug.Print i
        comment = Empty
        code = parseHTMLChars(tmp(i) & " ")
        If lineIsComment(CStr(code)) Then
            comment = code
            code = Empty
        Else
            lastDq = InStrRev(code, """")
            lastSq = InStrRev(code, "'")
            If lastSq > lastDq Then
                If lastDq <= 0 Then lastDq = lastSq
                If lastDq > 0 Then
                    comment = Mid(code, lastDq)
                    code = Mid(code, 1, lastDq - 1)
                End If
            End If
        End If
        tmp(i) = ParseCode(code) & ParseComment(comment)
    Next
     
    Dim header As String
    
    header = "<div style='background: #ffffff; overflow:auto;width:auto;border:solid gray;border-width:.1em .1em .1em .8em;padding:.2em .6em;'>" & _
             "<pre style='margin: 0; line-height: 125%'>" & vbCrLf
    
    it = Join(tmp, vbCrLf)
    htmlize = header & RemoveRedundantTags(it) & vbCrLf & "</div></pre>"

End Function

Private Function lineIsComment(txt As String) As Boolean
    Dim tmp As String
    tmp = Replace(txt, " ", Empty)
    tmp = Replace(tmp, vbTab, Empty)
    If Len(tmp) > 0 Then
        If Mid(tmp, 1, 1) = "'" Then lineIsComment = True
    End If
End Function

Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
  Dim i As Long
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function

Private Function parseHTMLChars(it)
    Dim t As String
    t = Replace(it, "&", "&amp;")
    t = Replace(t, "<", "&lt;")
    t = Replace(t, ">", "&gt;")
    parseHTMLChars = t
End Function

Private Function ParseCode(it)
    Dim i As Long
    If it = Empty Then Exit Function
    For i = 0 To UBound(RW)
        it = Replace(it, RW(i), BLU & RW(i) & CF)
    Next
    For i = 0 To UBound(Special)
        it = Replace(it, Special(i), BLU & Special(i) & CF)
    Next
    ParseCode = it
End Function

Private Function ParseComment(it)
    If it = Empty Then Exit Function
    ParseComment = GRN & it & CF
End Function

Private Function RemoveRedundantTags(it)
    'it = Replace(it, CF & BLU, Empty)
    'it = Replace(it, CF & GRN, Empty)
    'it = Replace(it, CF & vbCrLf & BLU, vbCrLf)
    'it = Replace(it, CF & vbCrLf & GRN, vbCrLf)
    RemoveRedundantTags = it
End Function

