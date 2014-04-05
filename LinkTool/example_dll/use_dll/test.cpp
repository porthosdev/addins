#include <windows.h> 
#include <stdio.h>
#include <conio.h>

typedef int (__stdcall *retVal)(int value);
typedef int (__stdcall *retVal2)(int *value);
typedef void (__stdcall *ModalForm)(void);
typedef HWND (__stdcall *NonModalForm)(void);

/*
	Visual Basic uses exception code 0xC000008F as its internal exception code. 
	You can ignore the returned float value that the exception returns. Visual Basic 
	is actually making a call to an object that returns a failure HRESULT. That 
	failure is turned into an exception that gets thrown. There is no bug in Visual 
	Basic. Instead, you need to examine the object being called, which causes the 
	exception, and determine why it is failing. 
*/

#define EXPORT comment(linker, "/EXPORT:"__FUNCTION__"="__FUNCDNAME__)


void __stdcall MyCExport(void){
#pragma EXPORT
	printf("MyCExport ran!\n");
}

void main(void){
	
	HMODULE h = LoadLibrary("project1.dll");
	if(h==0) h = LoadLibrary("./../project1.dll");

	if(h==0){
		printf("load failed press any key to exit...");
		getch();
		exit(0);
	}

	printf("Library base = %x\n", h);

	/* 
	retVal rv = 0;
	rv = (retVal)GetProcAddress(h,"retVal");
	printf("Calling rv(6) = %d\n", rv(6));

	retVal2 rv2 = 0;
	rv2 = (retVal2)GetProcAddress(h,"retVal2");
	int v = 6;
	printf("Calling rv2(6) = %d now v=%d\n", rv2(&v), v); //byref is not allowing modification in vb?
	
	ModalForm mf = 0;
	mf = (ModalForm)GetProcAddress(h,"ModalForm");
 	mf();
	*/

	NonModalForm nmf = 0;
	nmf = (NonModalForm)GetProcAddress(h,"NonModalForm");
 	HWND hwnd = nmf();

	MSG Msg;
	while(GetMessage(&Msg, hwnd, 0, 0) > 0)
	{
		if(Msg.message == WM_QUIT) break;
		TranslateMessage(&Msg);
		DispatchMessage(&Msg);
	}

	printf("press any key to exit...");
	getch();

	

}
 