<%@ Language=VBScript %>
<%
dim db
dim rs 
set db = server.CreateObject("SQLComandos.Comandos")

sql = "select id,descripcion,numero_res from pregunta where publicar = 1 order by fecha desc"

set rs = db.SQLSelect ("DBBAECYS",sql)

if not rs.eof then
cod_pregunta = rs("id")
pregunta = rs("descripcion")
respuestas = rs("numero_res")
set rs = nothing

%>
<HTML>
<HEAD>

<link href="./styles/interior.css" rel="stylesheet" type="text/css">
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<script language="javascript">
 function Grafica(CodPregunta) 	
   {
	  ancho = 320
	  altura = 450
	  if (screen) {
	  leftPos = (screen.width / 2) - (ancho/2)
	  }
	  catWindow = window.open('grafico.asp?Pregunta='+CodPregunta,'Encuesta','toolbar=no,menubar=no,location=no,scrollbars=no,width='+ancho+'+,height='+altura+',left='+leftPos+'+,top=20')
	  
	 return false
	}
	function Ingresar(){
  	forma.submit()
	}
</script>
<BODY>
<div align="center">
<font size="2" face="Arial, Helvetica, sans-serif"><strong><%=pregunta%></strong></font> 
<br>
<form name="forma" method ="post" action="ingresodatos_encuesta.asp" >
<table border="0" cellpadding="2" cellspacing="2">
<tr>
<th>
<input name="pregunta" type="hidden" value="<%=cod_pregunta%>">
</th>
</tr>
<%
sql = "select id,descripcion from opciones_respuesta where pregunta = "& cod_pregunta
set rs  = db.SQLSelect("DBBAECYS",sql)
if not rs.eof then
%>
<% for cont= 1 to cint(respuestas) 

%>
<tr>
<td align="left" class="celdaBgBlue05">
<input type="radio" value="<%=rs("id")%>" checked name="respuesta"></td>
<td align="left" class="celdaBgBlue10"   > <%=rs("descripcion")%> </td>
</tr>
<%
rs.MoveNext
next 
end if
set rs = nothing
%>
</table>
<br>
<br>

<table align="center" border="0" cellpadding="2" cellspacing="2">

<tr>
<td align="left">
<input name="aceptar" type=Button class="celdaBgBlue05" onclick="javascript:Ingresar()" value="Aceptar">
</td>
<td>
<input type=Button class="celdaBgBlue05"  value="Resultado" name="Resultado" onclick="javascript:Grafica('<%=cod_pregunta%>')">
</td>
<td>
<input type=Button class="celdaBgBlue05"  value="Cerrar" name="aceptar" onclick="javascript:window.close()">
</td>
</tr>
</table>
</form>
</div>
</BODY>
</HTML>
<%
end if
%>
