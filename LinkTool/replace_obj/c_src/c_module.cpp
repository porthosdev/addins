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

//this is the name of the module were replacing..
#define MODNAME Module1  

//this structure is required to match the vb6 name mangling of AAGXXZ
class MODNAME
{
private:
   void __stdcall MODNAME::init();  
   void __stdcall MODNAME::to64();
   void __stdcall MODNAME::sub64();
   void __stdcall MODNAME::add64();
   void __stdcall MODNAME::hex64();
};



typedef FARPROC  (__stdcall *GetProc)(HMODULE a0,LPCSTR a1);
typedef HMODULE  (__stdcall *LoadLib)(LPCSTR a0);
typedef BSTR (__stdcall *SysAllocLen)(void* str, int sz);
typedef int (__cdecl *Sprnf)(char *, const char *, ...);
typedef int (__cdecl *Strlen)(char *);
typedef int (__stdcall *Mb2wc)(UINT CodePage, DWORD dwFlags, LPCSTR lpMultiByteStr, int cbMultiByte, LPWSTR lpWideCharStr, int cchWideChar);

GetProc getproc;
LoadLib loadlib;
Sprnf sprnf;
Strlen strln;
SysAllocLen sysAlloc;
Mb2wc mb2wc;

//apparently __FUNCTION__ will ignore the leading underscore on our function
//names which is ok. if the exports show up with a leading underscore with your
//version just add it in to the vb declares, or change prefix..
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

int __stdcall _init(int lpfnGetProc, int lpfnLoadLib){
#pragma EXPORT 	 
	 HMODULE h = 0;
	 int failed = 0;

	 //_asm int 3
	 getproc = (GetProc)lpfnGetProc;
     loadlib = (LoadLib)lpfnLoadLib;

	 h = loadlib("kernel32.dll");
	 mb2wc = (Mb2wc)getproc(h,"MultiByteToWideChar");

	 h = loadlib("msvcrt.dll");
	 sprnf = (Sprnf)getproc(h,"sprintf");
	 strln = (Strlen)getproc(h,"strlen");
	 
	 h = loadlib("oleaut32");
	 sysAlloc = (SysAllocLen)getproc(h,"SysAllocStringLen");
	 
	 if( (int)mb2wc == 0) failed++;
	 if( (int)sprnf == 0) failed++;
	 if( (int)strln == 0) failed++;
	 if( (int)sysAlloc == 0) failed++;

	 return failed;
}

unsigned __int64 __stdcall _to64(unsigned int hi, unsigned int lo){
#pragma EXPORT
	unsigned __int64 ret=0;
	ret = hi;
	ret = ret << 32;
	ret += lo;  
	return ret;
}

unsigned __int64 __stdcall _add64(unsigned __int64 base, unsigned __int64 val){
#pragma EXPORT
	return base + val;
}

unsigned __int64 __stdcall _sub64(unsigned __int64 v0, unsigned __int64 v1){
#pragma EXPORT
	return v0 - v1;
}

BSTR __stdcall _hex64(unsigned __int64 v0){
#pragma EXPORT
	char buf[100];
	sprnf(buf, "0x%016I64x", v0);
	int wslen = mb2wc(CP_ACP, 0, buf, strlen(buf), 0, 0);
	BSTR bstr = sysAlloc(0, wslen);
    mb2wc(CP_ACP, 0, buf, strlen(buf), bstr, wslen);
	return bstr;
}


//since our private class functions actually do have arguments, the implementation of the private
//functions has to redirect to the correct implementation. The real implementations are also exported
//from the dll so that they can be used when in dll form.
#define JMPFUNC(Name)   __declspec(naked) void __stdcall MODNAME::Name () { _asm jmp _##Name }

JMPFUNC(init)
JMPFUNC(to64)
JMPFUNC(sub64)
JMPFUNC(add64)
JMPFUNC(hex64)