Dim nombre, length
nombre = InputBox("Ingresa tu nombre")
length = Len(nombre)

for i = 1 to length
	txt = Mid(nombre,i,1)
	
	if txt="a" or txt="A" or txt="e" or txt="E" or txt="i" or txt="I" or txt="o" or txt="O" or txt="u" or txt="U" Then
		counter = counter + 1
	
	End if
Next
MsgBox ("Hola " & nombre & "Tu nombre contiene " & counter & " vocales")
	

