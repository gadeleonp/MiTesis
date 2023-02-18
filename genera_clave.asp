<%
function clave()
		 dim set1 ,set2,set3
		 dim limitesuperior, limiteinferior
		 dim cont
		 dim password
		 dim altnumber
 
		Randomize 
		 
			'5 CARACTERES DE a-z
			 limitesuperior = 97
			 limiteinferior = 122
			 
			cont = 5
			for i=1 to cont 
					altnumber = Int((limitesuperior - limiteinferior + 1) * Rnd + limiteinferior)
					set1 = set1 & chr(altnumber) 
					
			next

			'5 Caracteres especiales del caracter 32 al 64 excluyendo numeros comilla y apostrofe
			limitesuperior = 33
			limiteinferior = 64
			cont = 5
			altnumber = ""
			set2 = ""
			while len(set2) < 5
				altnumber = Int((limitesuperior - limiteinferior + 1) * Rnd + limiteinferior)
				if altnumber = 34 or altnumber = 39 or altnumber = 37 or altnumber = 42 or (altnumber >= 48 and altnumber <=62) then
				else
				set2 = set2 & chr(altnumber) 
				end if
			wend 

			'5 caracteres de 1 - 9
			 limitesuperior = 1
			 limiteinferior = 9
			cont = 5
			for i=1 to cont 
					altnumber = Int((limitesuperior - limiteinferior + 1) * Rnd + limiteinferior)
					set3 = set3 & altnumber 
			next
			password = set1 & set2 & set3
			clave = password
end function 

%>