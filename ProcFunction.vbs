Function findArea(radio)
	const pi = 3.14
	area = pi*radio*radio
	findArea = area
End function

radio = InputBox("Ingrese el radio")

MsgBox "The area of the circle when te radius is " & radio & " is " & findArea(radio)