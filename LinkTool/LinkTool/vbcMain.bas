Attribute VB_Name = "vbcMain"
Option Explicit
Option Compare Text

'Jims original code has been extended a bit to include a couple
'more command options git repository included. thanks for the
'interesting project! - Dave Zimmer <dzzie@yahoo.com>

'http://www.xtremevbtalk.com/showthread.php?t=282796
'=========================================================
' mathimagics@yahoo.co.uk
'=========================================================
' MVBLC Link Control Tool:   Module "vbcMain"
'=======================================================
'
' Mathimagics Visual Basic Link Control tool
'
'      Jim White,  July 2006
'      Canberra, Australia
'
'=========================================================================
'
'  A simple tool for customising the linkage of VB6 ActiveX DLL's.
'
'  This is a normal standalone EXE project. Compile it as "Link.exe".
'
'  The real VB6 linker is also called Link.exe, and it is typically
'  located in the same folder as VB6.EXE (in the VB98 folder of
'  your Microsoft Visual Studio installation directory).
'
'      1. Change the real linker's name to "vbLink.exe" (make a
'         backup copy just in case!) in the VB98 folder
'
'      2. Put a copy this MVBLC version of Link.exe into the same
'         folder
'
'  MVBLC will only intervene in the normal linkage procedure if we
'  are making a DLL, and there is a VBC file (a link control command
'  file) present in the same directory as the DLL, and  with the same
'  name as the DLL.  In all other cases MVBLC operates transparently,
'  it just passes the VB6-specified link command over to the real
'  linker.
'
'  A link control file is simply a text file named [exe_name].vbc
'
'  MVBLC supports the following link control commands:
'
'      EXPORT, _Export, ENTRY, TIDY, STATUS, PostBuild, Debug, AddObj, replace
'
'  PostBuild and AddObj support basic envirnoment variables:
'        %1                 full path and file name of target output file
'        %apppath           folder path of the project being built
'        %outname           output file name only
'        %vb                path to the vb6 installation directory where link is
'
'  The Command names are not case-sensitive, but remember that all
'  LINK symbols ARE case-sensitive.  So the the names of modules and
'  functions must match exactly with those used in the project.
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
'
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
'
'  -----------
'  LINK ERRORS
'  -----------
'
'    If you specify an invalid module or function name, the (real) link
'    process will fail. This program pipes the link log to a temporary
'    file, so if the link fails, it can extract and display the error
'    messages from the real linker.
'
'    If the link fails, a message will also be displayed by the VB6 IDE,
'    with the error message "DLL Load Failed".  After the link stage
'    completes, the IDE attempts to load the DLL it just built, and if
'    the link has failed that DLL will not be available. You should
'    find and fix any errors in the .vbc file and try MAKE again.
'
'========================================================================

Const VB6FOLDER = "C:\Program Files\MicroSoft Visual Studio\VB98"
Const logfile = "c:\vbLink.log"
   
Public EXEFILE  As String  ' full pathname of exe/dll file being
Public EXENAME  As String  ' name of exe/dll being built (no ext)
Public OUTNAME  As String  ' full name of output file
Public cmdFile     As String  ' path to the current VBC file
Public defFile  As String     'path to def file generated..

Public vbCommand   As String  ' copy of original cmdline that we modify
Public orgCmdLine   As String ' original LINK command line passed in by VB6 IDE

Dim Options()   As String  ' Ccommand line tokens
Dim ObjList()   As String  ' list of project OBJ's being linked
Dim wipeList()  As String  ' array of functions to wipe for replacement..

Dim EXEPATH     As String  ' Folder containing OBJ files
Dim xList       As String  ' Export request list

Dim F           As Integer ' file unit

Public RunVisible   As Long
Dim isDLL       As Boolean
Dim hasExports  As String
Dim PostBuild   As String  ' allow for a post build command
Dim DebugMode   As Boolean ' flag for DEBUG command
Dim ShowStatus  As Boolean ' flag for STATUS command
Dim TidyFlag    As Boolean ' flag for TIDY   command
Dim NormalLink  As Boolean ' did we modify the link in any way?
Dim eMsg        As String  ' link error message
Const LANG_US = &H409

Sub Main()
   
   NormalLink = True
   orgCmdLine = Command
   vbCommand = Command   ' make a copy of the command line
  
   ' check command line for diagnostic options
   If InStr(vbCommand, "/STATUS:") Then
      frmLinkInfo.ShowStatus vbCommand
      Exit Sub
   End If
      
   If InStr(vbCommand, "/ERROR:") Then
      frmLinkInfo.ShowError vbCommand
      Exit Sub
   End If
   
   If Not FolderExists(VB6FOLDER) Then
        MsgBox "Folder not found correct path: " & VB6FOLDER
        Exit Sub
   End If
   
   If Not FileExists(VB6FOLDER & "\vbLink.exe") Then
        MsgBox "You must first copy the original link.exe to vblink.exe"
        Exit Sub
   End If
      
   If isIde() Then
   
       orgCmdLine = ReadFile(VB6FOLDER & "\lastLink.txt")
       vbCommand = orgCmdLine
       
       cmdFile = LoadCmdFile() 'this also loads some globals
       
       If GetSetting("vbLinkTool", "settings", "lastCmdProj", "") = EXENAME Then
            cmdFile = GetSetting("vbLinkTool", "settings", "lastCmdFile", "")
       End If
       
       If Len(cmdFile) > 0 And Len(vbCommand) > 0 Then DebugMode = True
       
   ElseIf Len(Command) > 0 Then
        
        writeFile App.path & "\lastLink.txt", Command
        cmdFile = LoadCmdFile() 'this also loads some globals
        
        If Len(cmdFile) > 0 Then
             SaveSetting "vbLinkTool", "settings", "lastCmdFile", cmdFile
             SaveSetting "vbLinkTool", "settings", "lastCmdProj", EXENAME
        End If
              
   End If
   
   ProcessVBC cmdFile
   
   If DebugMode Then
        vbCommand = frmDebug.DebugCommandLine()
        If Len(vbCommand) = 0 Then End
   End If

   If NormalLink Then
      Execute "VBLINK " & vbCommand, 1
      If Len(PostBuild) > 0 Then
        'MsgBox PostBuild
        On Error Resume Next
        Shell "cmd /c " & PostBuild
      End If
   Else
      If Not AryIsEmpty(wipeList) Then WipeFunctions
      RunCustomlink
   End If
   
End Sub

Private Function LoadCmdFile() As String
  
    Dim xFile   As String       ' link control file [out_name].vbc
    Dim j As Long, k As Long
    
    Options = Split(vbCommand, "/")
   '
   ' We now have:
   '   Options(0)= the LINK command + link object list
   '           1 = the /ENTRY switch
   '           2 = the /OUT   switch
   '           3, 4, etc   other switches /BASE, /SUBSYSTEM, /VERSION, /OPT etc
   '
   ' 1) get the EXEpath and EXEname from the /OUT switch.  It's always the
   '    2nd switch but we search for it just to be certain!
   '
   For k = 1 To UBound(Options)
      If Left$(Options(k), 4) = "OUT:" Then
         EXEFILE = Mid$(Options(k), 5)
         EXEFILE = Trim$(Replace(EXEFILE, """", ""))
         Exit For
      End If
   Next
   
   If InStr(1, EXEFILE, ".dll", vbTextCompare) > 0 Then isDLL = True
   
   If EXEFILE = "" Then Exit Function
   
   j = InStrRev(EXEFILE, "\")
   EXEPATH = Left$(EXEFILE, j - 1)
   OUTNAME = Mid$(EXEFILE, j + 1)
   EXENAME = Left$(OUTNAME, Len(OUTNAME) - 4)
   
   'look for the link control file in the EXEPATH folder then in the project foilder
   xFile = EXEPATH & "\" & EXENAME & ".vbc"
   If Not FileExists(xFile) Then
      xFile = CurDir & "\" & EXENAME & ".vbc"
      If Not FileExists(xFile) Then Exit Function ' no control file, normal link
   End If

   LoadCmdFile = xFile

End Function

Sub RunCustomlink()
      
   '
   ' Run the real linker via a batch file, so we can check the results
   '
   F = FreeFile
   Open "c:\vbLink.bat" For Output As #F
   Print #F, "@echo off"
   Print #F, Left$(VB6FOLDER, 2) ' ensure project drive selected  '<--added from blog post
   Print #F, "cd """ & VB6FOLDER & """"
   Print #F, "VBLINK " & vbCommand & " 1> " & logfile
   
   If Len(PostBuild) > 0 Then
        Print #F, PostBuild & " >> " & logfile
   End If
   
   Print #F, "del c:\vbLink.bat"   ' make the BAT file tidy up
   Close #F
   
   Execute "c:\vbLink.bat", 1
   
   If vbLinkError Then Exit Sub
   If TidyFlag Then Call vbTidy
   If ShowStatus Then DisplayLinkStatus

End Sub
   
Private Function ExpandPaths(ByVal sin) As String
    sin = Trim(sin)
    ExpandPaths = Replace(sin, "%1", EXEFILE)
    ExpandPaths = Replace(ExpandPaths, "%apppath", EXEPATH, , , vbTextCompare)
    ExpandPaths = Replace(ExpandPaths, "%outname", OUTNAME, , , vbTextCompare)
    ExpandPaths = Replace(ExpandPaths, "%vb", VB6FOLDER, , , vbTextCompare)
End Function

Private Sub ProcessVBC(cmdFile As String)
   
   Dim xName   As String       ' dll export  (.DEF) filename
   Dim xKey    As String       ' control file keyword
   Dim mName   As String       ' Module name
   Dim pName() As String       ' proc names to export
   Dim dName   As String       ' temp variable for decorated proc name
   Dim EntryFlag As Boolean    ' true if we find an ENTRY command
   Dim j As Long, k As Long

   If Not FileExists(cmdFile) Then Exit Sub
   
   F = FreeFile
   Open cmdFile For Input As #F
   
   Do Until EOF(F)                        ' ".vbc" Export Control File
      Line Input #F, xKey
      
      xKey = Replace(xKey, vbTab, Empty)        'ignore leading whitespace and empty lines
      xKey = Trim(xKey)
      If Len(xKey) = 0 Then GoTo NextLine
      If Left(xKey, 1) = "#" Then GoTo NextLine 'ignore comments
      
      j = InStr(xKey, ";"): If j Then xKey = Left$(xKey, j - 1)
      j = InStr(xKey, "'"): If j Then xKey = Left$(xKey, j - 1)
      
      '
      '  Commands:
      '    Export <moduleid> <function1> <function2> .....
      '    Entry  <moduleid> <function>
      '
      '    Tidy
      '
      ' 3. remove extraneous whitespace and split into tokens
      '
      xKey = Trim(xKey)
      If xKey = "" Then GoTo NextLine
      While InStr(xKey, "  ")
         xKey = Replace(xKey, "  ", " ")
      Wend
      pName = Split(xKey, " ")
      
      Select Case LCase(pName(0))
         
         Case "postbuild": PostBuild = ExpandPaths(Mid(xKey, Len("postbuild") + 1))
         Case "debug":     DebugMode = True
         Case "status":    ShowStatus = True
         Case "tidy":      TidyFlag = True
         Case "addobj":    vbCommand = ExpandPaths(Mid(xKey, Len("addobj") + 1)) & " " & vbCommand
         
         Case "wipefunc":  ' WIPEFUNC <module> <proclist>
                           If UBound(pName) <= 1 Then
                                MsgBox "VBC WipeFunc requires 2 or more arguments", vbInformation
                           Else
                                NormalLink = False
                                If UBound(pName) > 1 Then
                                   mName = pName(1)            ' module name
                                   For j = 2 To UBound(pName)
                                      dName = "?" & pName(j) & "@" & mName & "@@AAGXXZ" ' decorated name
                                      push wipeList, mName & dName
                                   Next
                                End If
                           End If
                           
         Case "replace":
                           If UBound(pName) <> 2 Then
                                MsgBox "VBC Command Replace requires 2 arguments.", vbInformation
                           Else
                                NormalLink = False
                                vbCommand = Replace(vbCommand, pName(1), pName(2), , , vbTextCompare)
                           End If
         
         Case "_export"    ' _EXPORT <proclist>  this is for using C Obj files with raw undecorated names..no module name needed..
                            If UBound(pName) > 0 Then
                               For j = 1 To UBound(pName)
                                  xList = xList & "," & pName(j)
                               Next
                            End If
                            
         Case "export"      ' EXPORT <module> <proclist>
                            If UBound(pName) > 1 Then
                               mName = pName(1)            ' module name
                               For j = 2 To UBound(pName)
                                  dName = "?" & pName(j) & "@" & mName & "@@AAGXXZ" ' decorated name
                                  xList = xList & "," & pName(j) & " = " & dName
                               Next
                            End If
               
         Case "entry"       ' ENTRY <module> <procname>
                            If UBound(pName) = 2 Then
                               mName = pName(1)            ' module name
                               dName = "?" & pName(2) & "@" & mName & "@@AAGXXZ" ' decorated name
                               xList = xList & "," & pName(2) & " = " & dName    ' we export it (for convenience only)
                               For k = 1 To UBound(Options)
                                  If Left$(Options(k), 6) = "ENTRY:" Then
                                     Options(k) = "ENTRY:" & pName(2) & " "      ' and adjust the linker
                                     vbCommand = " " & Join(Options, "/")        ' \ENTRY switch
                                     EntryFlag = True
                                     Exit For
                                  End If
                               Next
                            End If
         
         Case "adddef"      ' AddDEF <aliasname> = <name>
                            If pName(2) = "=" Then         '   we need to retrieve decorated form of <name>
                               dName = pName(3)
                               j = InStr(xList, "," & pName(3) & " = ")
                               If j Then                   '   if it is in our list
                                  dName = Mid(xList, j + Len(pName(3)) + 4)
                                  k = InStr(dName, ",")
                                  If k Then dName = Left(dName, k - 1)
                               End If
                               xList = xList & "," & pName(1) & " = " & dName
                            End If
         End Select
         
NextLine:
   Loop
   
   Close #F
   
   
   If xList = "" Then Exit Sub
   hasExports = True
   
   '
   ' For a custom link, build a DEF file and add the /DEF switch
   '     to the link command line
   '
   NormalLink = False
   If EntryFlag Then xList = xList & ",__vbaS"
   pName = Split(xList, ",")
   xName = EXEPATH & "\" & EXENAME & ".def"
   defFile = xName
   
   F = FreeFile
   Open xName For Output As #F
   
   If isDLL Then 'otherwise this would set the DLL charateristic in pe header..we can export from exes too..
        Print #F, "LIBRARY "; EXENAME
   End If
   
   Print #F, "EXPORTS"
   
   For j = 1 To UBound(pName)
      Print #F, "   " & pName(j)
   Next
   
   Close #F
   
   If Right(vbCommand, 1) <> " " Then vbCommand = vbCommand & " "
   vbCommand = vbCommand & "/DEF:""" & xName & """"

End Sub

'this is programmed horribly inefficient
Function WipeFunctions() As Boolean
   
   Dim x
   Dim F As Long, i As Long
   Dim fPath As String
   Dim data As String
   Dim pos As Long
   Dim delta As Long
   Dim newModName() As Byte
   Dim modName As String
   Dim proto As String
   
   WipeFunctions = True
   
   For Each x In wipeList
         
        pos = InStr(x, "?")
        If pos < 1 Then GoTo nextone
         
        modName = Mid(x, 1, pos - 1)
        proto = Mid(x, pos)
         
        ReDim newModName(Len(modName))
        For i = 0 To UBound(newModName)
             newModName(i) = &H61 + i
        Next
         
        fPath = EXEPATH & "\" & modName & ".obj"
        If Not FileExists(fPath) Then GoTo nextone
         
        data = ReadFile(fPath)
        F = FreeFile
        Open fPath For Binary As F
         
        delta = InStr(proto, "@") + 1
        pos = InStr(data, proto)
        If pos > 0 And delta > 1 Then
             'its going to be somethign like ?add@Module1@@AAGXXZ, lets change the module name..
             Put F, pos + delta + 1, newModName
        Else
             WipeFunctions = False
        End If
        
        Close F
        
nextone:
   Next
        
End Function

Public Sub vbTidy()

   Dim e
   Dim ext() As String
   Dim path As String
   
   ext = Array(".exp", ".lib", ".def")
   
   For Each e In ext
        path = EXEPATH & "\" & EXENAME & e
        If FileExists(path) Then Kill path
   Next
   
End Sub


Function vbLinkError() As Boolean
   
   Dim logentry As String, temp As String
   Dim i As Integer, j As Integer
   
   If Dir$(EXEFILE) <> "" Then
      If FileExists(logfile) Then Kill logfile
      If isDLL Then FixDLL
      Exit Function ' all is well
   End If
   
   vbLinkError = True
   Shell VB6FOLDER & "\Link.exe /ERROR:" & EXEFILE, 1
   Exit Function
   
BadSign:
   eMsg = "(can't open link log file - " & Err.Description & ")"
End Function


Sub DisplayLinkStatus()
   '
   ' following a successful custom DLL link, display the
   ' results (list the Export Table). To avoid interference
   ' with the IDE, we fire up an independent copy of the
   ' link tool, passing it the collected info for display
   '
   Shell VB6FOLDER & "\Link.exe /STATUS:" & EXEFILE, 1
End Sub


Function FileExists(path) As Boolean
  On Error Resume Next
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then
     If Err.Number <> 0 Then Exit Function
     FileExists = True
  End If
End Function

Function ReadFile(filename) As String 'this one should be binary safe...
  On Error GoTo hell
  If Not FileExists(filename) Then Exit Function
  F = FreeFile
  Dim b() As Byte
  Open filename For Binary As #F
  ReDim b(LOF(F) - 1)
  Get F, , b()
  Close #F
  ReadFile = StrConv(b(), vbUnicode, LANG_US)
  Exit Function
hell:   ReadFile = ""
End Function

Function writeFile(path, it) As Boolean 'this one should be binary safe...
    On Error GoTo hell
    Dim b() As Byte
    If FileExists(path) Then Kill path
    F = FreeFile
    b() = StrConv(it, vbFromUnicode, LANG_US)
    Open path For Binary As #F
    Put F, , b()
    Close F
    writeFile = True
    Exit Function
hell: writeFile = False
End Function

Public Function isIde() As Boolean
    On Error GoTo hell
    Debug.Print 1 / 0
    isIde = False
    Exit Function
hell:
    isIde = True
End Function

Function FolderExists(path) As Boolean
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbDirectory) <> "" Then FolderExists = True _
  Else FolderExists = False
End Function

Sub push(ary, value) 'this modifies parent ary object
    Dim x As Long
    On Error GoTo init
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init: ReDim ary(0): ary(0) = value
End Sub

Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
  Dim x As Long
    x = UBound(ary)
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function
