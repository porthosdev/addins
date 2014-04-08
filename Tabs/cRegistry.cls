VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' =========================================================
' Class:    cRegistry
' Author:   Steve McMahon
' Date  :   21 Feb 1997
'

' ---------------------------------------------------------------------------
' vbAccelerator - free, advanced source code for VB programmers.
'     http://vbaccelerator.com
' =========================================================

'Lf
'Stripped to include only what we need

'Registry Specific Access Rights
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_ALL_ACCESS = &H3F

'Open/Create Options
Private Const REG_OPTION_NON_VOLATILE = 0&

'Define severity codes
Private Const ERROR_SUCCESS = 0&
Private Const ERROR_INVALID_DATA = 13&
Private Const ERROR_MORE_DATA = 234 '  dderror


'Structures Needed For Registry Prototypes
Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Boolean
End Type

Private Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

'Registry Function Prototypes
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" _
  (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
  ByVal samDesired As Long, phkResult As Long) As Long

Private Declare Function RegSetValueExStr Lib "advapi32" Alias "RegSetValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
   ByVal dwType As Long, ByVal szData As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32" Alias "RegSetValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
   ByVal dwType As Long, szData As Long, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExByte Lib "advapi32" Alias "RegSetValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
   ByVal dwType As Long, szData As Byte, ByVal cbData As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Declare Function RegQueryValueExStr Lib "advapi32" Alias "RegQueryValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
   ByRef lpType As Long, ByVal szData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32" Alias "RegQueryValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
   ByRef lpType As Long, szData As Long, ByRef lpcbData As Long) As Long
Private Declare Function RegQueryValueExByte Lib "advapi32" Alias "RegQueryValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
   ByRef lpType As Long, szData As Byte, ByRef lpcbData As Long) As Long
   
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" _
  (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, _
   ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, _
   lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, _
   lpdwDisposition As Long) As Long

'Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" _
  (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, _
   lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, _
   lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long

Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" ( _
    ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, _
    ByVal cbName As Long) As Long

Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" _
  (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
   lpcbValueName As Long, ByVal lpReserved As Long, ByVal lpType As Long, _
   ByVal lpData As Long, ByVal lpcbData As Long) As Long
   
'Private Declare Function RegEnumValueLong Lib "advapi32.dll" Alias "RegEnumValueA" _
  (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
   lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, _
   lpData As Long, lpcbData As Long) As Long
'Private Declare Function RegEnumValueStr Lib "advapi32.dll" Alias "RegEnumValueA" _
  (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
   lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, _
   ByVal lpData As String, lpcbData As Long) As Long
'Private Declare Function RegEnumValueByte Lib "advapi32.dll" Alias "RegEnumValueA" _
  (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
   lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, _
   lpData As Byte, lpcbData As Long) As Long

Private Declare Function RegConnectRegistry Lib "advapi32.dll" Alias "RegConnectRegistryA" (ByVal lpMachineName As String, ByVal hKey As Long, phKey As Long) As Long

Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" _
   (ByVal hKey As Long, ByVal lpClass As String, _
   lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, _
   lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, _
   lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, _
   lpftLastWriteTime As Any) As Long

Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
  (ByVal hKey As Long, ByVal lpSubKey As String) As Long

Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
  (ByVal hKey As Long, ByVal lpValueName As String) As Long

' Other declares:
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long


Public Enum ERegistryClassConstants
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
End Enum

Public Enum ERegistryValueTypes
'Predefined Value Types
    REG_SZ = (1)                           'Unicode nul terminated string
    REG_EXPAND_SZ = (2)                    'Unicode nul terminated string w/enviornment var
    REG_BINARY = (3)                       'Free form binary
    REG_DWORD = (4)                        '32-bit number
    REG_DWORD_LITTLE_ENDIAN = (4)          '32-bit number (same as REG_DWORD)
    REG_DWORD_BIG_ENDIAN = (5)             '32-bit number
    REG_MULTI_SZ = (7)                     'Multiple Unicode strings
End Enum

Private m_hClassKey             As Long
Private m_sSectionKey           As String
Private m_sValueKey             As String
Private m_vValue                As Variant
'Private m_sSetValue            As String
Private m_vDefault              As Variant
Private m_eValueType            As ERegistryValueTypes


Public Property Get KeyExists() As Boolean
    Dim hKey As Long
    
    If RegOpenKeyEx(m_hClassKey, m_sSectionKey, 0, 1, hKey) = ERROR_SUCCESS Then
        KeyExists = True
        RegCloseKey hKey
    Else
        KeyExists = False
    End If
    
End Property

Public Function CreateKey() As Boolean
    Dim tSA             As SECURITY_ATTRIBUTES
    Dim hKey            As Long
    Dim lCreate         As Long
    Dim e               As Long

    'Open or Create the key
    e = RegCreateKeyEx(m_hClassKey, m_sSectionKey, 0, "", REG_OPTION_NON_VOLATILE, _
                 KEY_ALL_ACCESS, tSA, hKey, lCreate)
    If e Then
        Err.Raise 26001, App.EXEName & ".cRegistry", "Failed to create registry Key: '" & m_sSectionKey
    Else
        CreateKey = (e = ERROR_SUCCESS)
        'Close the key
        RegCloseKey hKey
    End If
End Function

Public Function DeleteKey() As Boolean
    Dim e As Long
    
    e = RegDeleteKey(m_hClassKey, m_sSectionKey)
    If e Then
        Err.Raise 26001, App.EXEName & ".cRegistry", "Failed to delete registry Key: '" & m_hClassKey & "',Section: '" & m_sSectionKey
    Else
        DeleteKey = (e = ERROR_SUCCESS)
    End If
    
End Function

Public Function DeleteValue() As Boolean
    Dim e           As Long
    Dim hKey        As Long

    e = RegOpenKeyEx(m_hClassKey, m_sSectionKey, 0, KEY_ALL_ACCESS, hKey)
    If e Then
        Err.Raise 26001, App.EXEName & ".cRegistry", "Failed to open key '" & m_hClassKey & "',Section: '" & m_sSectionKey & "' for delete access"
    Else
        e = RegDeleteValue(hKey, m_sValueKey)
        If e Then
            Err.Raise 26001, App.EXEName & ".cRegistry", "Failed to delete registry Key: '" & m_hClassKey & "',Section: '" & m_sSectionKey & "',Key: '" & m_sValueKey
        Else
            DeleteValue = (e = ERROR_SUCCESS)
        End If
    End If
End Function

Public Property Get Value() As Variant
    Dim vValue              As Variant
    Dim cData               As Long
    Dim sData               As String
    Dim ordType             As Long
    Dim e                   As Long
    Dim hKey                As Long
        
    e = RegOpenKeyEx(m_hClassKey, m_sSectionKey, 0, KEY_QUERY_VALUE, hKey)
    
    e = RegQueryValueExLong(hKey, m_sValueKey, 0&, ordType, 0&, cData)
    If e And e <> ERROR_MORE_DATA Then
        Value = m_vDefault
        Exit Property
    End If
    
    m_eValueType = ordType
    Select Case ordType
    Case REG_DWORD, REG_DWORD_LITTLE_ENDIAN
        Dim iData As Long
        e = RegQueryValueExLong(hKey, m_sValueKey, 0&, _
                               ordType, iData, cData)
        vValue = CLng(iData)
        
    Case REG_DWORD_BIG_ENDIAN  ' Unlikely, but you never know
        Dim dwData As Long
        e = RegQueryValueExLong(hKey, m_sValueKey, 0&, _
                               ordType, dwData, cData)
        vValue = SwapEndian(dwData)
        
    Case REG_SZ, REG_MULTI_SZ ' Same thing to Visual Basic
        sData = String$(cData - 1, 0)
        e = RegQueryValueExStr(hKey, m_sValueKey, 0&, _
                               ordType, sData, cData)
        vValue = sData
        
    Case REG_EXPAND_SZ
        sData = String$(cData - 1, 0)
        e = RegQueryValueExStr(hKey, m_sValueKey, 0&, _
                               ordType, sData, cData)
        vValue = ExpandEnvStr(sData)
        
    ' Catch REG_BINARY and anything else
    Case Else
        Dim abData() As Byte
        ReDim abData(cData)
        e = RegQueryValueExByte(hKey, m_sValueKey, 0&, _
                                ordType, abData(0), cData)
        vValue = abData
        
    End Select
    Value = vValue
    'Close the key
    RegCloseKey hKey
End Property

Public Property Let Value(ByVal vValue As Variant)
    Dim ordType         As Long
    Dim c               As Long
    Dim hKey            As Long
    Dim e               As Long
    Dim lCreate         As Long
    Dim tSA             As SECURITY_ATTRIBUTES
    

    'Open or Create the key
    e = RegCreateKeyEx(m_hClassKey, m_sSectionKey, 0, "", REG_OPTION_NON_VOLATILE, _
                 KEY_ALL_ACCESS, tSA, hKey, lCreate)
    
    If e Then
        Err.Raise 26001, App.EXEName & ".cRegistry", "Failed to set registry value Key: '" & m_hClassKey & "',Section: '" & m_sSectionKey & "',Key: '" & m_sValueKey & "' to value: '" & m_vValue & "'"
    Else

        Select Case m_eValueType
        Case REG_BINARY
            If (VarType(vValue) = vbArray + vbByte) Then
                Dim ab() As Byte
                ab = vValue
                ordType = REG_BINARY
                c = UBound(ab) - LBound(ab) + 1 ' Bugfix by Morten Egelund Rasmussen (Dec. 11/2003)
                e = RegSetValueExByte(hKey, m_sValueKey, 0&, ordType, ab(0), c)
            Else
                Err.Raise 26001
            End If
        Case REG_DWORD, REG_DWORD_BIG_ENDIAN, REG_DWORD_LITTLE_ENDIAN
            If (VarType(vValue) = vbInteger) Or (VarType(vValue) = vbLong) Then
                Dim i As Long
                i = vValue
                ordType = REG_DWORD
                e = RegSetValueExLong(hKey, m_sValueKey, 0&, ordType, i, 4)
            End If
        Case REG_SZ, REG_EXPAND_SZ
            Dim s As String, iPos As Long
            s = vValue
            ordType = REG_SZ
            ' Assume anything with two non-adjacent percents is expanded string
            iPos = InStr(s, "%")
            If iPos Then
                If InStr(iPos + 2, s, "%") Then ordType = REG_EXPAND_SZ
            End If
            c = Len(s) + 1
            s = s & vbNullChar ' Thanks to Shane Marsden
            e = RegSetValueExStr(hKey, m_sValueKey, 0&, ordType, s, c)
            
        ' User should convert to a compatible type before calling
        Case Else
            e = ERROR_INVALID_DATA
            
        End Select
        
        If Not e Then
            m_vValue = vValue
        Else
            Err.Raise vbObjectError + 1048 + 26001, App.EXEName & ".cRegistry", "Failed to set registry value Key: '" & m_hClassKey & "',Section: '" & m_sSectionKey & "',Key: '" & m_sValueKey & "' to value: '" & m_vValue & "'"
        End If
            
    End If
    'Close the key
    If hKey <> 0 Then RegCloseKey hKey
    
End Property

Private Sub pSetClassValue(ByVal sValue As String)
    Dim sSection As String
    
    ClassKey = HKEY_CLASSES_ROOT
    Value = sValue
    sSection = SectionKey
    ClassKey = HKEY_LOCAL_MACHINE
    SectionKey = "SOFTWARE\Classes\" & sSection
    Value = sValue
    SectionKey = sSection
End Sub

Public Property Get ValueType() As ERegistryValueTypes
    ValueType = m_eValueType
End Property

Public Property Let ValueType(ByVal eValueType As ERegistryValueTypes)
    m_eValueType = eValueType
End Property

Public Property Get ClassKey() As ERegistryClassConstants
    ClassKey = m_hClassKey
End Property

Public Property Let ClassKey(ByVal eKey As ERegistryClassConstants)
    m_hClassKey = eKey
End Property

Public Property Get SectionKey() As String
    SectionKey = m_sSectionKey
End Property

Public Property Let SectionKey(ByVal sSectionKey As String)
    m_sSectionKey = sSectionKey
End Property

Public Property Get ValueKey() As String
    ValueKey = m_sValueKey
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' A Wrapper Property for the VALUE property .
' Assign all passed values to the class properties and call
' ME.VALUE to return value
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get ValueEx(ClassKey As ERegistryClassConstants, _
                             SectionKey As String, _
                             ValueKey As String, _
                             ValueType As ERegistryValueTypes, _
                             Default As Variant) As Variant
    With Me
       .ClassKey = ClassKey
       .SectionKey = SectionKey
       .ValueKey = ValueKey
       .ValueType = ValueType
       .Default = Default
        ValueEx = .Value
    End With
End Property

Public Property Let ValueEx(ClassKey As ERegistryClassConstants, _
                            SectionKey As String, _
                            ValueKey As String, _
                            ValueType As ERegistryValueTypes, _
                            Default As Variant, _
                            newValue As Variant)
    With Me
       .ClassKey = ClassKey
       .SectionKey = SectionKey
       .ValueKey = ValueKey
       .ValueType = ValueType
       .Default = Default
       
       If .ValueType = REG_SZ Then
          If CStr(newValue) = "" Then
             newValue = Default
          End If
       End If
       .Value = newValue
    End With
End Property

Public Property Let ValueKey(ByVal sValueKey As String)
    m_sValueKey = sValueKey
End Property

Public Property Get Default() As Variant
    Default = m_vDefault
End Property

Public Property Let Default(ByVal vDefault As Variant)
    m_vDefault = vDefault
End Property

Private Function SwapEndian(ByVal dw As Long) As Long
    CopyMemory ByVal VarPtr(SwapEndian) + 3, dw, 1
    CopyMemory ByVal VarPtr(SwapEndian) + 2, ByVal VarPtr(dw) + 1, 1
    CopyMemory ByVal VarPtr(SwapEndian) + 1, ByVal VarPtr(dw) + 2, 1
    CopyMemory SwapEndian, ByVal VarPtr(dw) + 3, 1
End Function

Private Function ExpandEnvStr(sData As String) As String
    Dim c As Long
    Dim s As String
    
    ' Get the length
    s = "" ' Needed to get around Windows 95 limitation
    c = ExpandEnvironmentStrings(sData, s, c)
    
    ' Expand the string
    s = String$(c - 1, 0)
    c = ExpandEnvironmentStrings(sData, s, c)
    
    ExpandEnvStr = s
End Function

