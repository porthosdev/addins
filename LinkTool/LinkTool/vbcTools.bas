Attribute VB_Name = "vbcTools"
Option Explicit
Option Compare Text

'=========================================================
' mathimagics@yahoo.co.uk
'=========================================================
'
' MVBLC Link Control Tool:   Module "vbcTools"
'=======================================================
'
'  This module has some DLL utility functions:
'
'  a) GETEXPORTS: reads the newly linked DLL's Export
'       Table (VBC "STATUS" option).
'
'  b) FIXDLL: looks at the DLL file to see if it imports
'       functions from the "fixed" DLL name hardcoded in
'       the Type Library "vbLibrary.TLB".  We "fix"
'       the DLL by changing that fixed name to be the
'       real name of the DLL itself.
'
'=======================================================

Type LIST_ENTRY         ' 8 bytes
   FLink As Long
   Blink As Long
   End Type

Type LOADED_IMAGE       ' 48 bytes (46 bytes packed)
   ModuleName As Long
   hFile As Long
   MappedAddress As Long         ' Base address of mapped file
   pFileHeader As Long           ' Pointer to IMAGE_PE_FILE_HEADER
   pLastRvaSection As Long       ' Pointer to first COFF section header (section table)??
   NumberOfSections As Long
   pSections As Long             ' Pointer to first COFF section header (section table)??
   Characteristics As Long       ' Image characteristics value
   fSystemImage As Byte
   fDOSImage As Byte
   FLink As Long
   Blink As Long
   SizeOfImage As Long
   End Type
   
Declare Function MapAndLoad Lib "Imagehlp.dll" ( _
   ByVal ImageName As String, _
   ByVal DLLPath As String, _
   LoadedImage As LOADED_IMAGE, _
   DotDLL As Long, _
   ReadOnly As Long) As Long

Declare Function UnMapAndLoad Lib "Imagehlp.dll" ( _
   LoadedImage As LOADED_IMAGE) As Long

Declare Function ImageRvaToVa Lib "Imagehlp.dll" ( _
   ByVal NTHeaders As Long, _
   ByVal Base As Long, _
   ByVal RVA As Long, _
   ByVal LastRvaSection As Long) As Long

Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
   lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Declare Sub Sleep Lib "kernel32" (ByVal nMilliseconds As Long)

Declare Function lstrlenA Lib "kernel32" (ByVal lpsz As Long) As Long

'=============================================================

Type IMAGE_DATA_DIRECTORY  ' 8 bytes
   RVA  As Long
   size  As Long
   End Type

Type IMAGE_OPTIONAL_HEADER       ' 232 bytes
   Magic  As Integer
   MajorLinkerVersion  As Byte
   MinorLinkerVersion  As Byte
   SizeOfCode  As Long
   SizeOfInitializedData  As Long
   SizeOfUninitializedData  As Long
   AddressOfEntryPoint  As Long
   BaseOfCode  As Long
   BaseOfData  As Long
   ImageBase  As Long
   SectionAlignment  As Long
   FileAlignment  As Long
   MajorOperatingSystemVersion  As Integer
   MinorOperatingSystemVersion  As Integer
   MajorImageVersion  As Integer
   MinorImageVersion  As Integer
   MajorSubsystemVersion  As Integer
   MinorSubsystemVersion  As Integer
   Win32VersionValue  As Long
   SizeOfImage  As Long
   SizeOfHeaders  As Long
   CheckSum  As Long
   Subsystem  As Integer
   DllCharacteristics  As Integer
   SizeOfStackReserve  As Long
   SizeOfStackCommit  As Long
   SizeOfHeapReserve  As Long
   SizeOfHeapCommit  As Long
   LoaderFlags  As Long
   NumberOfRvaAndSizes  As Long
   DataDirectory(0 To 15) As IMAGE_DATA_DIRECTORY
   End Type

Type IMAGE_COFF_HEADER     ' 20 bytes
   Machine  As Integer
   NumberOfSections  As Integer
   TimeDateStamp  As Long
   PointerToSymbolTable  As Long
   NumberOfSymbols  As Long
   SizeOfOptionalHeader  As Integer
   Characteristics  As Integer
   End Type

Type IMAGE_PE_FILE_HEADER   ' 256 bytes
   Signature  As Long                           ' 4 bytes -- PE signature
   FileHeader As IMAGE_COFF_HEADER            ' 20 bytes -- This is the COFF header
   OptionalHeader As IMAGE_OPTIONAL_HEADER    ' 232 bytes
   End Type

Type IMAGE_EXPORT_DIRECTORY_TABLE  ' 40 bytes
   Characteristics   As Long
   TimeDateStamp     As Long
   MajorVersion      As Integer
   MinorVersion      As Integer
   Name              As Long
   Base              As Long
   NumberOfFunctions As Long
   NumberOfNames     As Long
   pAddressOfFunctions        As Long
   ExportNamePointerTableRVA  As Long
   pAddressOfNameOrdinals     As Long
   End Type
   

Public LoadImage  As LOADED_IMAGE
Dim peheader   As IMAGE_PE_FILE_HEADER
Dim exportdir  As IMAGE_EXPORT_DIRECTORY_TABLE

Dim vaEntryPoint   As Long
Dim rvaEntryPoint  As Long
Dim dllBaseAddress As Long
Dim procTable      As Long
Dim procAddress    As Long
Dim ExportNamePointerTableVA As Long
Dim ImportNamePointerTableVA As Long
Dim rvaImportDirTable As Long
Dim rvaExportDirTable As Long
Dim vaImportDirTable As Long
Dim vaExportDirTable As Long

'
' the fixed DLL name used by the generic "self-referencing"
'   type library (vbLibraryHelper.tlb)
'
Const FixDLLname = "VBLIBRARYHELPER_MATHIMAGICS.DLL"

Function GetExports() As String
   
   Dim i      As Long
   Dim nNames As Long
   Dim sName  As String
   Dim pNext  As Long
   Dim lNext  As Long
   Dim epFlag As Boolean
   Dim xpFlag As Boolean, nxp As Integer
   Dim epName As String
   Dim xList  As String
   Dim iTag   As String
   
   xList = GetImports  ' get list of self-Imported names, if any
   
   If vaExportDirTable = 0 Then
      GetExports = GetExports & vbLf & "  No information (the Export Table could not be accessed)"
      Exit Function
   End If
      
   CopyMemory ByVal VarPtr(exportdir), ByVal vaExportDirTable, LenB(exportdir)
   procTable = ImageRvaToVa(LoadImage.pFileHeader, LoadImage.MappedAddress, exportdir.pAddressOfFunctions, 0)
   nNames = exportdir.NumberOfNames
   
   If nNames = 0 Then Exit Function
     
   ExportNamePointerTableVA = ImageRvaToVa(LoadImage.pFileHeader, LoadImage.MappedAddress, exportdir.ExportNamePointerTableRVA, 0&)

   pNext = ExportNamePointerTableVA
   CopyMemory lNext, ByVal pNext, 4
   
   For i = 0 To nNames - 1
      lNext = ImageRvaToVa(LoadImage.pFileHeader, LoadImage.MappedAddress, lNext, 0&)
      sName = CSTRtoVBSTR(lNext)
      CopyMemory procAddress, ByVal procTable, 4
      epFlag = (procAddress = rvaEntryPoint)  ' is this the entrypoint?
      xpFlag = InStr(xList, vbLf & sName)
      iTag = "    "
      If epFlag Then iTag = "  * ": epName = sName
      If xpFlag Then iTag = "  ~ ": nxp = nxp + 1
      iTag = iTag & Hex(procAddress + dllBaseAddress) & ":  " & sName
      GetExports = GetExports & vbLf & iTag
      pNext = pNext + 4
      procTable = procTable + 4
      CopyMemory lNext, ByVal pNext, 4
   Next
   
   If Len(epName) Or nxp Then   ' add auto-import info"
      GetExports = GetExports & vbLf & "  -------------------"
      If nxp Then GetExports = GetExports & vbLf & "  ~  auto-import"
      If Len(epName) Then GetExports = GetExports & vbLf & "  *  entrypoint"
   Else
      GetExports = GetExports & vbLf & "    " & Hex(rvaEntryPoint + dllBaseAddress) & ":  <entrypoint>"
   End If
   
   GetExports = GetExports & vbLf & "  -------------------"

End Function

Function GetImports() As String

   Dim i      As Long
   Dim nNames As Long
   Dim sName  As String
   Dim pNext  As Long
   Dim lNext  As Long
   Dim epFlag As Boolean
   Dim pNames As String
   
   If vaImportDirTable = 0 Then Exit Function
   
   Dim pImportTable   As Long
   Dim pLookupTable   As Long
   Dim pLookupEntry   As Long
   Dim LookupTableRVA As Long
   Dim DLLNameRVA     As Long
   Dim DLLname        As String
   
   pImportTable = vaImportDirTable  ' set by LoadDLL
   
   Do
      CopyMemory LookupTableRVA, ByVal pImportTable, 4
      CopyMemory DLLNameRVA, ByVal pImportTable + 12, 4
      If LookupTableRVA = 0 And DLLNameRVA = 0 Then Exit Do
      pLookupTable = ImageRvaToVa(LoadImage.pFileHeader, LoadImage.MappedAddress, LookupTableRVA, 0&)
      DLLNameRVA = ImageRvaToVa(LoadImage.pFileHeader, LoadImage.MappedAddress, DLLNameRVA, 0&)
      DLLname = CSTRtoVBSTR(DLLNameRVA)
      If DLLname = EXENAME Then
         GoSub GetProcList
         Exit Function
      End If
      pImportTable = pImportTable + 20
   Loop
      
   Exit Function

GetProcList:
   '
   ' Get all imported functions from one import DLL
   '
   Do While pLookupTable
      CopyMemory pLookupEntry, ByVal pLookupTable, 4
      If pLookupEntry = 0 Then Exit Do
      ' Check most significant bit
      ' If not 0 then avoid this entry, it's an ordinal reference, not a name
      If pLookupEntry Then
         pNext = ImageRvaToVa(LoadImage.pFileHeader, LoadImage.MappedAddress, pLookupEntry, 0&)
         sName = CSTRtoVBSTR(pNext + 2)
         pNames = pNames & vbLf & sName
         nNames = nNames + 1
         pLookupTable = pLookupTable + 4
      End If
   Loop
   
   GetImports = pNames
   Return

End Function
   
Function CSTRtoVBSTR(ByVal lpsz As Long) As String
   Dim i As Long, cChars As Long
   cChars = lstrlenA(lpsz)
   CSTRtoVBSTR = String$(cChars, 0)
   CopyMemory ByVal StrPtr(CSTRtoVBSTR), ByVal lpsz, cChars
   CSTRtoVBSTR = StrConv(CSTRtoVBSTR, vbUnicode)
   i = InStr(CSTRtoVBSTR, Chr$(0))
   If i > 0 Then CSTRtoVBSTR = Left$(CSTRtoVBSTR, i - 1)
   End Function

Sub LoadDLL()
   
   If MapAndLoad(EXEFILE, "", LoadImage, True, True) = 0 Then Exit Sub
   
   CopyMemory ByVal VarPtr(peheader), ByVal LoadImage.pFileHeader, 256
   
   rvaEntryPoint = peheader.OptionalHeader.AddressOfEntryPoint
   dllBaseAddress = peheader.OptionalHeader.ImageBase
   
   If rvaEntryPoint <> 0 Then
      vaEntryPoint = ImageRvaToVa(LoadImage.pFileHeader, LoadImage.MappedAddress, rvaEntryPoint, 0&)
   End If
   
   rvaExportDirTable = peheader.OptionalHeader.DataDirectory(0).RVA
   
   If rvaExportDirTable Then
      vaExportDirTable = ImageRvaToVa(LoadImage.pFileHeader, LoadImage.MappedAddress, rvaExportDirTable, 0&)
   End If
   
   rvaImportDirTable = peheader.OptionalHeader.DataDirectory(1).RVA
   
   If rvaImportDirTable Then
      vaImportDirTable = ImageRvaToVa(LoadImage.pFileHeader, LoadImage.MappedAddress, rvaImportDirTable, 0&)
   End If

End Sub

Sub FixDLL()
   '
   ' If required, make this DLL self-referencing. If there is an
   '   export table entry that references the DLL name we have
   '   hardcoded in vbLibrary.TLB, then replace it with the name
   '   of this DLL.
   '
   Dim buf() As Byte, oldkey() As Byte, keylen As Integer
   Dim newkey() As Byte
   Dim i&, j&, k&, b&, f%, fsize&
   Dim pImportTable   As Long
   Dim LookupTableRVA As Long
   Dim DLLNameRVA     As Long
   Dim DLLname        As String
   
   LoadDLL                    ' load the DLL image (this will get
   UnMapAndLoad LoadImage     ' the offset of its Import Table)
   
   keylen = Len(FixDLLname)
   oldkey = StrConv(FixDLLname & Chr$(0), vbFromUnicode)
   newkey = StrConv(UCase(EXENAME & ".dll") & Chr$(0), vbFromUnicode)
   f = FreeFile
   
   Open EXEFILE For Binary As #f
   fsize = LOF(f)
   ReDim buf(fsize - 1)
   Get #f, , buf

   pImportTable = rvaImportDirTable  ' set by LoadDLL
   
   Do
      CopyMemory LookupTableRVA, buf(pImportTable), 4
      CopyMemory DLLNameRVA, buf(pImportTable + 12), 4
      If LookupTableRVA = 0 And DLLNameRVA = 0 Then Exit Do
      DLLname = CSTRtoVBSTR(VarPtr(buf(DLLNameRVA)))
      If DLLname = FixDLLname Then
         Seek f, DLLNameRVA + 1
         Put f, , newkey  ' Fix the DLL Import Table entry
         Exit Do
      End If
      pImportTable = pImportTable + 20
   Loop
   
   Close #f

End Sub

