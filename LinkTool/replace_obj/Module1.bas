Attribute VB_Name = "Module1"

Private Declare Function dll_init Lib "c_module.dll" Alias "init" (ByVal lpfnGetProc As Long, ByVal lpfnLoadLib As Long) As Long
Private Declare Function dll_to64 Lib "c_module.dll" Alias "to64" (ByVal hi As Long, ByVal lo As Long) As Currency
Private Declare Function dll_add64 Lib "c_module.dll" Alias "add64" (ByVal a As Currency, ByVal b As Currency) As Currency
Private Declare Function dll_sub64 Lib "c_module.dll" Alias "sub64" (ByVal a As Currency, ByVal b As Currency) As Currency
Private Declare Function dll_hex64 Lib "c_module.dll" Alias "hex64" (ByVal a As Currency) As String


Function init(ByVal lpfnGetProc As Long, ByVal lpfnLoadLib As Long) As Long
    init = dll_init(lpfnGetProc, lpfnLoadLib)
End Function

Function to64(ByVal hi As Long, ByVal lo As Long) As Currency
    to64 = dll_to64(hi, lo)
End Function

Function add64(ByVal a As Currency, ByVal b As Currency) As Currency
    add64 = dll_add64(a, b)
End Function

Function sub64(ByVal a As Currency, ByVal b As Currency) As Currency
    sub64 = dll_sub64(a, b)
End Function

'not sure why but when going through an API declare, our BSTR return is expanded again on return?
'but when run from the internal obj file is functions as expected...work around for now..
Function hex64(ByVal a As Currency) As String
    hex64 = dll_hex64(a)
    hex64 = StrConv(hex64, vbFromUnicode)
End Function


 

