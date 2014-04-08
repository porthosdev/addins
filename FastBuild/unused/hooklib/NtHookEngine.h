
//This is a modified version of the open source x86/x64 
//NTCore Hooking Engine written by:
//Daniel Pistelli <ntcore@gmail.com>
//http://www.ntcore.com/files/nthookengine.htm
//
//It uses the x86/x64 GPL disassembler engine
//diStorm was written by Gil Dabah. 
//Copyright (C) 2003-2012 Gil Dabah. diStorm at gmail dot com.
//
//Mods by David Zimmer <dzzie@yahoo.com>

enum hookType{ ht_jmp = 0, ht_pushret=1, ht_jmp5safe=2, ht_jmpderef=3, ht_micro };
enum hookErrors{ he_None=0, he_cantDisasm, he_cantHook, he_maxHooks, he_UnknownHookType, he_Other  };
extern hookErrors lastErrorCode;
extern int logLevel;

extern void  (__stdcall *debugMsgHandler)(char* msg);
extern char* __stdcall GetHookError(void);
extern char* __stdcall GetDisasm(ULONG_PTR pAddress, int* retLen = NULL);
extern void __stdcall DisableHook(ULONG_PTR Function);
extern void __stdcall EnableHook(ULONG_PTR Function);
extern ULONG_PTR __stdcall GetOriginalFunction(ULONG_PTR Hook);
extern BOOL __stdcall HookFunction(ULONG_PTR OriginalFunction, ULONG_PTR NewFunction, char* name, enum hookType ht);
extern void __stdcall SetDebugHandler(ULONG_PTR lpfn);
extern int __stdcall RemoveHook(ULONG_PTR Function);
extern void __stdcall RemoveAllHooks(void);
extern void __stdcall UnInitilize(void);