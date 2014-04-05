Attribute VB_Name = "Module1"
Option Explicit
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function CoInitialize Lib "ole32.dll" (Optional ByVal pvReserved As Long = 0) As Long

'note: you must trap all errors at the vb level, if they make it back
'to the C caller it wont know what to make of the error codes

Function DllMain(ByVal hInstance As Long, ByVal lReason As Long, ByVal lReserved As Long) As Long
    Const firstCall = 1
    Call vbaS(hInstance, lReason, lReserved)
    If lReason = firstCall Then Call Init_VB
    DllMain = 1
End Function

' IID_IClassFactory {00000001-0000-0000-C000-000000000046}
' DllGetClassObject has be utilized from a tlb file as vb declare syntax wont work yet..
' CoInitialize is called so we can use other ActiveX components in our code..
Sub Init_VB()
    Dim pDummy As Long
    Dim pIID As IID
    pIID.Data1 = 1
    pIID.Data4(0) = &HC0
    pIID.Data4(7) = &H46
    Call DllGetClassObject(pDummy, pIID, pDummy)
    Call CoInitialize
End Sub

Public Function retVal(ByVal N As Long) As Long
    retVal = N
End Function

Public Function retVal2(ByRef N As Long) As Long
    retVal2 = N
    N = N + 1
End Function
    
'modal is ok always..does not require message pump..
Sub ModalForm()
    Form1.Show 1
End Sub

Function NonModalForm() As Long
    
    On Error Resume Next
    
    'non modal needs a messagepump running in same thread..(done in C in this example)
    Load Form1
    ShowWindow Form1.hwnd, 1
    NonModalForm = Form1.hwnd
    
    If Err.Number <> 0 Then MsgBox "Err in nonmodalForm line: " & Erl & " " & Err.Description
End Function

