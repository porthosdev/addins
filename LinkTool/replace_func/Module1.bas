Attribute VB_Name = "Module1"

'in the vb function we do nothing just to show it worked..
Function add(ByVal a As Long, ByVal b As Long) As Long
    add = 0
End Function

'so we know other functions we unaffected..
Function GetResult(x As Long) As String
    GetResult = "Result was " & x
End Function

