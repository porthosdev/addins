Title: CodeHelp 1.2.2 (Tabbed IDE code editor Window) Updated
Description: This is a VB AddIn, with this AddIn VB can show the Code editor windows in tabbed style.
Much like the code editor in VS.NET IDE or Opera Browser.
As usual thanks to Paul Caton for his superb Subclassing class.
Just compile the project and then restart VB IDE to see it in action.
Tested and developed under WinXP SP 1, now also works in Win98.
_________________________________________________
1st Update May 04, 2005 ver 1.0.1
Fixed List:
- Crash on Exit if MZTools also running
- Tabstrip all over the place if Code editor is not in maximized state
_________________________________________________
2nd Update
May 06, 2005 Ver 1.2.0
- Recode the algorithm and the class structure, fix bug in rectangle calculations on large project
- Added two sub projects (in MDITest and TabWork sub folder) to help test the AddIn, as a bonus now the TabStrip also works for normal VB MDI Application.
- The close button behaviors now complies with standard button behaviours (event fired on MouseUp, added hover and pushed state indicator)
- the active tab always visible (well not always)
- active tab now also synchronize with the active code window, whether the user activates via Window Menu or Project explorer
_________________________________________________
3rd Update
May 09, 2005 Ver 1.2.1
Crash on Win98 - Fixed!!
Remove GDI leak when no tab items are present
Added Option to hide the close button
Improve button painting
_________________________________________________
4th Update
May 10, 2005 Ver 1.2.2
Fixed bug in popup menu handler, added WM_STYLECHANGED for handling caption programmatically changed
This file came from Planet-Source-Code.com...the home millions of lines of source code
You can view comments on this code/and or vote on it at: 

http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=60321&lngWId=1

The author may have retained certain copyrights to this code...please observe their request and the law by reviewing all copyright conditions at the above URL.
