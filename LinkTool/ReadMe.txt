
The vb link tool was originally created by:

   Jim White <mathimagics@yahoo.co.uk> 2006

This build has added the following commands: 
    _Export, PostBuild, Debug, AddObj, replace, wipefunc

--------------------------------------------------------

The short story:

1) rename link.exe in vb home dir to vblink.exe
2) place link.exe from ./linktool directory in vb home dir



Examples:

 example_dll 
    project is a sample of a standard dll written in vb6
    that exports some symbols, hosts a form (modal and non modal) 
    which includes an embedded activex control (ListView) list and timer.
    and has a callback to the C loader. 

    You can test it by running use_dll.exe which is a VC console app 

 c_Objfile 
    example of a vb6 application that has C functions compiled
    into it. In the IDE, it will use the C DLL, but when you compile it, the
    C Object function is linked in and it will then automatically use the
    internal version of its functions so you wouldnt have to distribute a dll.

 replace_obj
    example of using a C obj module to replace an entire vb6 module. 
    when run in IDE vb module uses the corrosponding C DLL externally, when
    compiled, the functions are linked right into the vb6 exe so you dont have 
    to distribute an extra dll.


See the pdf and link tool source for more details.

Note: if you go to build the example_dll it may say missing reference
      vbLibraryhelper.tlb. If this happens, just open references, click
      browse and re-add it. Its in the linktool directory. 

--------------------------------------------------------


'  PostBuild and AddObj support basic envirnoment variables:
'        %1                 full path and file name of target output file
'        %apppath           folder path of the project being built
'        %outname           output file name only
'        %vb                path to the vb6 installation directory where link is
'
'
'  -------------
'  LINK COMMANDS
'  -------------
'
'    EXPORT <module name> <function name list>
'
'       e.g.  Export Module1 Function1 Function2
'             Export Module2 myTest1
'             Export Module2 myTest2
'
'       The nominated functions will be exported. Function list
'           members be in form "Name1 Alias Name2", allowing a
'           function to be exported with 2 ids (handy for C
'           linking, e.g. Export Mod1 vbFunc alias vbFunc@12)
'
'       NOTE: <module name> denotes the Name Property of the
'          corresponding module, NOT its file name!
'
'    _EXPORT <function name list>
'       allows you to export raw undecorated names. use this if you are
'       linking in a C Obj file. You then use vb declare syntax on self
'       to call them.
'
'    ADDOBJ <file.obj>
'       Allows you to link in Visual Studio C obj files. (tested with VC 2008
'       VC6 should work too) Make sure functions are stdcall and in a C file (not CPP)
'
'    ENTRY <module name> <function name>
'
'       The function referenced is exported, and it is marked as the
'       DLL's entrypoint (DllMain) function.
'
'    REPLACE <module.obj> <new.obj>
'
'      this is used for swapping out modules at link time (replace a vb interface
'      with a _matching_ C++ counterpart. Actually you can also use this to replace
'      any text from the command line if you want to tweak options
'
'    WIPEFUNC <module name> <name list>
'
'       this feature allows you to remove function names from VB obj modules. Place a
'       dummy function in a module with the desired prototype, then at compile time
'       you cna replace just that function, by adding a new C obj file that contains
'       its replacement. Crudely implemented, but works.

'  ----------------------
'  MISCELLANEOUS COMMANDS
'  ----------------------
'
'    TIDY
'       VB6 DLL linking produces EXP, LIB and DEF files, which are not
'       usually needed. Include this command in the VBC file and we
'       will remove them after the DLL has been linked.
'
'    STATUS
'       Including this command tells the link tool to display the export
'       table of the new DLL after linking has been completed.
'
'    DEBUG - pops up a modal dialog allowing you to view and edit def file and
'            command line sent to real link.exe before it is executed.
'
'    PostBuild - allows you to run a command after a build is complete.
'                for complex scripts use a batch file or launch a vbs in wsh
