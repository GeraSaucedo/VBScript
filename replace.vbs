
string1 = "one two three four"
resultado = replace(string1,"two","six")
MsgBox resultado


string2 = "one two three four one two three four one two three four"
resultado = replace(string2,"two","six",1,2)
MsgBox resultado

'RELPLACE A WORD STARTING A SPECIFIC LOCATION CHARAACTER BEFORE SPECIFIED POSITION ARE REMOVED

string2 = "one two three four one two three four one two three four"
resultado = replace(string2,"two","six",8)
MsgBox resultado