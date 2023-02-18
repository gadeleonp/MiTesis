<%@ Language=VBScript %>
<HTML>
<HEAD>
<link href="./theme/Master.css" rel="stylesheet" type="text/css">
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<script language="javascript">
 function Grafica(CodPregunta) 	
   {
	  ancho = 320
	  alto= 450
	  if (screen) {
	  leftPos = (screen.width / 2) - (ancho/2)
	  }
	  catWindow = window.open('grafico.asp?pregunta='+CodPregunta,'Encuesta','toolbar=no,menubar=no,location=no,scrollbars=no,width='+ancho+'+,height='+alto+',left='+leftPos+'+,top=20')
	 // document.location = 'encuesta.asp'
	 return false
	}
</script>
<%
dim db 
dim pregunta 
dim respuesta
dim resultado
pregunta = Request.Form("pregunta")
respuesta = Request.Form("respuesta")
set db = server.CreateObject ("SQLComandos.Comandos")
%>
<BODY onload="javascript:Grafica('<%=pregunta%>')">
<%
sql = "update opciones_respuesta set total = total + 1 where pregunta = " & pregunta & " and id  = " & respuesta
resultado = db.SQLOtras("DBBAECYS",sql)
%>
</BODY>
</HTML>
