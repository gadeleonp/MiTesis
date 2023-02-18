<?xml version="1.0"?>
<%	
	dim resultado
	resultado = request.querystring("valor")
	if resultado > 0 then
	else
	resultado = 1
	end if
	
	response.write "<documento>"
	response.write "<resultado>" 
	response.write resultado
	response.write "</resultado>"
	response.write "</documento>"
%>