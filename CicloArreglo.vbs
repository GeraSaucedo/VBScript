Dim ArregloNombres(4)
Dim contador

ArregloNombres(0) = "Dayanne"
ArregloNombres(1) = "Roberto"
ArregloNombres(4) = "Gerardo"
ArregloNombres(2) = "Bladimir"
ArregloNombres(3) = "Edson"

'llamar ciclo DoLoop
CicloDoLoop
CicloWhile
CicloForNext
CicloForEachNext

sub CicloDoLoop
	MsgBox "Iniciando ciclo do loop"
	contador = 0

	do 
		MsgBox ArregloNombres(contador)
		contador = contador + 1
	loop until contador = uBound(ArregloNombres) + 1

End sub

Sub CicloWhile
	MsgBox "Iniciando ciclo while"
	contador = 0
	
	while contador <= uBound(ArregloNombres)
		MsgBox ArregloNombres(contador)
		contador = contador + 1
	wend

End Sub 
	
Sub CicloForNext
	MsgBox "Iniciando ciclo For Next avanzando 2 lugares"
		
	for contador = 0 to uBound(ArregloNombres) step 2
		MsgBox ArregloNombres(contador)
	next
End sub

Sub CicloForEachNext
	MsgBox "Iniciando ciclo For Each Next"
	Dim element
	
	For each element in ArregloNombres
	MsgBox element
next
End sub
		
		
		


