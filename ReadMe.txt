
These are some VB6 IDE addings I have created or found useful.

CodeDB - mine - small util that has a code database and some other small tools.

LinkTool - extended version of Jim Whites LinkTool to create std dlls and 
           link in C Obj files

MemoryWindow - mine - VC like memory window for VB6

MouseWHeelFix - from MS, adds mouse wheel scrolling support to IDE.

FastBuild - mine - VB6 prompts you every single time you want to build an executable
        to give the file path. Then pops another warning if you want to overwrite
        the existing file. Also vb6 can get confused and lose which absolute directory
        to write the file to sometimes. This adding addresses all of these issues and
        lets you set a hardcoded default path to always compile to. It will be auto saved
        the first time you manually compile to a specific path, and can be changed from the
        fastBuild form anytime after that. Setting saved to the projects vbp file.

        this plugin will also allow you to set a postbuild command to run everytime the
        executable is compiled. 