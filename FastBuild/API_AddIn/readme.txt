
 CVMichael (Michael Ciurescu)
 Feb 3rd, 2004, 11:32 PM
 http://www.vbforums.com/showthread.php?277596-API-Add-In-for-Visual-Basic

API Add-In for Visual Basic

This add-in is MUCH better than the one that comes with VB because 
the search is done with the "like" operator, so you can use * or ? 
(look in MSDN for the like operator) characters to refine your search. 
Also, when you add a declaration to the list and it's using one or 
more structures, it will add those too to the list, it will add even 
structures used in structures.

For example, if you add the structure PRINTER_INFO_2 to the list, 
it will automatically add DEVMODE, and SECURITY_DESCRIPTOR structures 
because they are used in the PRINTER_INFO_2 structure, and it will 
also add the ACL structure because it is used in the 
SECURITY_DESCRIPTOR structure.

You have to compile the add-in first, then in the Add-Ins/Add-In 
Manager you can load this Add-In so you can use it.

And if you need a more complete list of declarations/types/constants, 
then you can download the Win32API_2.txt from the attached RAR files 
Win32api_2.part01.rar and Win32api_2.part02.rar (please note you have 
to download BOTH rar files and put them together using WinRAR). I had 
to do this because of the forum's limitation of 250KBytes per file.

The text file is 3.1 MBytes decompressed, and contains 
6,542 Declarations, 458 Types, and 55,566 constants 


dz mods:
 - mode to compile as standalone exe
 - syntax highlight output
 