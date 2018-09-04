Dim ArregloNombres(4)
Dim contador

ArregloNombres(0) = "Dayanne"
ArregloNombres(1) = "Roberto"
ArregloNombres(4) = "Gerardo"
ArregloNombres(2) = "Bladimir"
ArregloNombres(3) = "Edson"

'llamar ciclo DoLoop
CicloDoLoop

sub CicloDoLoop
	MsgBox "Iniciando ciclo do loop"
	contador = 0

	do 
		MsgBox ArregloNombres(contador)
		contador = contador + 1
	loop until contador = uBound(ArregloNombres) + 1
End sub




