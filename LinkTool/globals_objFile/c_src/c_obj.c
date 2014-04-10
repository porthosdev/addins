#include <stdio.h>
#include <windows.h>

#define EXPORT comment(linker, "/EXPORT:"__FUNCTION__"="__FUNCDNAME__)

int saved=0;

void __stdcall inc(unsigned int v0){
#pragma EXPORT
	saved += v0;
}

int __stdcall retrieve(){
#pragma EXPORT
	return saved;
}

