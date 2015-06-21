' Name:   VB6 Memory Window Addin
' Author: David Zimmer
' Site:   http://sandsprite.com
'
' notes:
'   might have a couple bugs, but its a decent tool and does what it needs to
'   olly.dll is open source if missing download here:
'          http://sandsprite.com/CodeStuff/olly_dll.html
'
' features:
'   view data as: hexdump, longs, long address, ascii, unicode, disasm
'   next/previous memory block
'   always on top
'   hit escape and it takes you back through the displayed address history
'   in long, long address, and disasm mode, if you ctrl + mouse over a valid address it will hyperlink it (like lazarus)
'   ability to view memory in other processes
'   ability to use expressions such as ?objptr(form1) or ?&h401000 + &h10 (VB Addin version only)
'   standalone version availabe in /debug folder

Credits:

   olly.dll is Copyright Oleh Yuschuk (GPL)
   portions copyright iDefense (GPL)
   