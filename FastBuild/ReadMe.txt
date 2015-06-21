
modCmdOutput.base created by:
    Joacim Andersson, Brixoft Software
    http://www.brixoft.net


FastBuild 

Adds 7 features to VB6 IDE

1) Compile without the prompts for file save/overwrite file. Set path once and forget it.
2) Ability to run a post build command.
3) (!) - Execute button to launch compiled exe.
4) (C) - Compile button added to toolbar
5) (I) - Clears immediate window (default will also clear immediate window every start)
6) (F) - brings up fast build config form
7) Project -> Quick Addref menu to add ActiveX controls with search box
 
A video walkthrough of the addin is available here:

http://www.youtube.com/embed/bLfvaYNIhzk?list=UUhIoXVvn4ViA3AL4FJW8Yzw

Also shows how to hook into pre-existing IDE button events from a vb6 addin.

the 3 new sub folders in here..they used to be addins but I am going to 
compile them as standalone now and just use FastBuild to add theier menu items
to launch the external exes. none of them depend on being an addin really on
thier own and the less junk loaded in teh IDE the better