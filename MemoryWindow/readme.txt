' Name:   VB6 Memory Window Addin
' Author: David Zimmer
' Site:   http://sandsprite.com
'
' notes:
'   Memory viewer window like VC has for vb6 IDE. see screen shot, 
'   hitting escape brings you back through the view history (address and view type)
'   might have a couple bugs, but its a decent tool and does what it needs to
'   olly.dll is open source if missing download here:
'          http://sandsprite.com/CodeStuff/olly_dll.html
'
' features:
'   view data as: hexdump, longs, long address, ascii, unicode, disasm
'   next/previous memory block
'   always on top
'   hit escape and it takes you back through the displayed address history
'
'
' todo:
'       ctrl click on an address to goto it would be nice
'       highlight address of data in blue would be nice..
'       savemem command (binary)