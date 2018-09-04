Dim edad
		
		edad = InputBox("Ingresa tu edad")

		If edad < 18 Then
			MsgBox "Ahorita no joven usted no cumple con la mayoria de edad"
		Elseif edad<45 Then
			MsgBox "Bienvenido usted aun es adulto"
		elseif edad<60 Then
			MsgBox "Bienvenido senior, en que le podemos ayudar"
		else
			MsgBox "Me escucha, me oye, me siente?"
		End If