
' so after a solid day of hard labor hooking the api and working out all the bugs..i find
' vb6 has a built in event that fires to prompt for a file name and can override it.
' i never imagined they would have a hook for that built in..sooo, anyway I will include this
' module, even though it isnt used, but it might come in handy someday and I am not going to
' lose the work..
'
' live and learn!

 This is a detours style hooking library written in C that was made for use with VB6.

 You will have to write a C stub to call the original api though. 
 There is a generic example one  included that takes one argument.

 This is all complete and debugged, see sample code in the module. The hook proc is
 designed to be implemented in VB. See notes above..

 Library supports adding hooks, enable/disable hooks, and completely removing hooks
 in preparation for unload or shutdown.

