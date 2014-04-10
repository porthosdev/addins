VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public path As String
Public progID As String
Public typeLib As String
Public version As String
Public name As String
Public isControl As Boolean
Public isProgrammable As Boolean
Public clsid As String
'Public ProxyStubClsid32 As String 'used in \Interfaces\
'Public InProcServer32 As String   'is an inprocess dll
'Public LocalServer32 As String    'is an activex exe
Public AlreadyReferenced As Boolean

Function ToString() As String
    Dim tmp() As String
    
    push tmp, "Name:    " & name
    
    If Len(progID) > 0 Then push tmp, "ProgID:  " & progID
    
    push tmp, "Path:    " & path
    push tmp, "Version: " & version
    
    If Len(typeLib) > 0 Then push tmp, "TypeLib: " & typeLib
    
    push tmp, "CLSID:   " & clsid
    
    ToString = Join(tmp, vbCrLf)
    
End Function



Private Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

