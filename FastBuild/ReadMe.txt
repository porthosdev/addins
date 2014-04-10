
modCmdOutput.base created by:
    Joacim Andersson, Brixoft Software
    http://www.brixoft.net


FastBuild 


This is an Addin for the VB6 IDE that adds a couple productivity tools.

When you goto build an exe or dll in VB6, it will prompt you for the 
file name every single time. Also, if the file already exists, it will 
make sure you really want to overwrite it..every single time. Another 
downfall to this process, is that sometimes the default build directory 
will change and you wont notice it, so the updated exe gets compiled 
somewhere unexpected. 

This plugin will allow you to manually set the path the first time, and 
from then on out, it will skip the dialogs and just automatically 
compile to the default path you set. 

It also allows you to set a post build command to run, and includes an 
execute button in the IDE to launch the compiled exe directly.

Another feature, is a new menu item under Project -> Quick Addref that 
gives you a streamlined form to add ActiveX control references. 

a video walkthrough of the addin is available here:

http://www.youtube.com/embed/bLfvaYNIhzk?list=UUhIoXVvn4ViA3AL4FJW8Yzw

Note: I removed the Compile button because it has inconsistant behavior and
can fail silently. I could work around the bug I found and test it more
but you can add your own build button manually that works right.

1) right click on the button bar and choose customize
2) commands tab, choose File in left list, then drag and drop Make <Project>
    to some where on the button bar. it will appear there now.
3) right click on the new button and you can set the name, and image as you want.


