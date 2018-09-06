a=split("W3School is my favorite website")
for each x in a
	MsgBox x & " "
next
MsgBox a(2)


a=split("saltillo,torreon,monclova,parras,sabinas",",")
for each x in a
	MsgBox x & " "
next
MsgBox a(2)

a=split("saltillo,torreon,monclova,parras,sabinas",",",3)
for each x in a
	MsgBox x & " "
next
MsgBox a(2)


