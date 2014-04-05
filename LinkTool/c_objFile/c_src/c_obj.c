#include <stdio.h>
#include <windows.h>

/*
	apparently all winapi has to go through dynamically loaded function pointers
	because I cant get the VB link to include any other lib files without it wacking
	the import table of the VB exe corrupting the executable..really though this technique
	is most useful for math like functions anyway which shouldnt need much if any api at all.
	i tried lib files from VC6 and VS2008. The Obj files from either work as long as they in
	debug mode.
*/

typedef FARPROC  (__stdcall *GetProc)(HMODULE a0,LPCSTR a1);
typedef HMODULE  (__stdcall *LoadLib)(LPCSTR a0);
typedef int (__stdcall *SysAllocSBL)(void* str, int sz);
typedef int (__cdecl *Sprnf)(char *, const char *, ...);
typedef int (__cdecl *Strlen)(char *);

GetProc getproc;
LoadLib loadlib;
Sprnf sprnf;
Strlen strln;
SysAllocSBL sysAlloc;

#define EXPORT comment(linker, "/EXPORT:"__FUNCTION__"="__FUNCDNAME__)

/*

 these will help you with strings..
	http://progtutorials.tripod.com/COM.htm
	http://forums.codeguru.com/showthread.php?257924-Access-the-BSTR-from-VB

 Compiler settings
	character set: multibyte
	runtime library: multithreaded (compiled in not dll)
	buffer security check = false
	disable randomization
	not compatiable with dep
	basic runt time checks = default
	/noentry  

*/ 

int __stdcall init(int lpfnGetProc, int lpfnLoadLib){
#pragma EXPORT 	 
	 int h = 0;
	 int failed = 0;

	 //_asm int 3
	 getproc = (GetProc)lpfnGetProc;
     loadlib = (LoadLib)lpfnLoadLib;

	 h = loadlib("msvcrt.dll");
	 sprnf = (Sprnf)getproc(h,"sprintf");
	 strln = (Strlen)getproc(h,"strlen");
	 
	 h = loadlib("oleaut32");
	 sysAlloc = (SysAllocSBL)getproc(h,"SysAllocStringByteLen");
	 
	 if( (int)sprnf == 0) failed++;
	 if( (int)strln == 0) failed++;
	 if( (int)sysAlloc == 0) failed++;

	 return failed;
}

unsigned __int64 __stdcall to64(unsigned int hi, unsigned int lo){
#pragma EXPORT
	unsigned __int64 ret=0;
	ret = hi;
	ret = ret << 32;
	ret += lo;  
	return ret;
}

unsigned __int64 __stdcall add64(unsigned __int64 base, unsigned __int64 val){
#pragma EXPORT
	return base + val;
}

unsigned __int64 __stdcall sub64(unsigned __int64 v0, unsigned __int64 v1){
#pragma EXPORT
	return v0 - v1;
}

BSTR __stdcall hex64(unsigned __int64 v0){
#pragma EXPORT
	char buf[100];
	sprnf(buf, "0x%016I64x", v0);
	return sysAlloc(buf, strln(buf));
}

/* 

another way to do it..

'Private Declare Function hex64 Lib "project1.exe" (ByVal a As Currency, ByVal buf As String, ByVal sz As Long) As Long
'Private Declare Function dll_hex64 Lib "c_obj.dll" Alias "hex64" (ByVal a As Currency, ByVal buf As String, ByVal sz As Long) As L

'Function doString64(a As Currency) As String
'    Dim ret As String
'    ret = Space(40)
'
'    If isIde() Then
'        sz = dll_string64(a, ret, 40)
'    Else
'        sz = string64(a, ret, 40)
'    End If
'
'    If sz > 0 Then ret = VBA.Left(ret, sz - 1)
'    doString64 = ret
'
'End Function

int __stdcall hex64(unsigned __int64 v0, char* buf, int bufsz ){
#pragma EXPORT
	sprnf(buf, "0x%016I64x", v0);
	return strln(buf); //strlen can actually get compiled in but doesnt show in .obj file? 
}

*/