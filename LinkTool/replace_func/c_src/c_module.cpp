#include <stdio.h>
#include <windows.h>

#define JMPFUNC(Name)   __declspec(naked) void __stdcall Module1::Name () { _asm jmp _##Name }

class Module1
{
private:
   void __stdcall Module1::add();  
};

//interesting..if you mark a function for export here, the linker will make sure to export it
//even if you only use the obj file itself.
//#define EXPORT comment(linker, "/EXPORT:"__FUNCTION__"="__FUNCDNAME__)

int __stdcall _add(int a, int b){
//#pragma EXPORT 	
	 return a + b;
}

JMPFUNC(add)
