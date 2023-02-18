<%@ Language=VBScript %>
<%
dim db
dim noticia,correlativo
dim rs
set db = server.CreateObject ("SQLComandos.Comandos")
noticia = Request.QueryString ("noticia")
correlativo = Request.QueryString ("correlativo")
folder = Request.QueryString("folder")
%>


<HTML>
<HEAD>
<link href="../theme/Master.css" rel="stylesheet" type="text/css">
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<script language="javascript">
 function Destino(Opcion) {
 if (Opcion == "1") {
    navegacion.correlativo.value = "1"
  
	}
 if (Opcion == "2") {
    navegacion.correlativo.value = "2"
	}
 if (Opcion == "3") {
    navegacion.correlativo.value = "3"
	}
 if (Opcion == "4") {
    navegacion.correlativo.value = "4"
	}

 return true
 }
 
</script>
<BODY bgcolor="#CCCCCC">
<div align="center">
<table>
<tbody>
<tr>
<td><a href="./listadonoticia.asp" target="_top">Regresar</a></td>
<%
sql ="select id,correlativo,enlace from documento_noticia where noticia = "& noticia & " order by correlativo"
set rs = db.SQLSelect("DBBAECYS",sql)
if not rs.eof then
while rs.EOF <> true 
path = "../Noticias/"
path = path + folder +"/"
path = path + rs("enlace")
%>
<td><a href="<%=path%>" target="noticia"><%= rs("correlativo")%></a></td>
<%
rs.MoveNext 
wend
end if
set rs = nothing
%>
</tr>
</tbody>
</table>
<form name="navegacion" method="get" action="./ver_noticia.asp" >
<input name="noticia" type = "hidden" value="<%=noticia%>">
<input name="folder" type = "hidden" value="<%=folder%>">
<input name="correlativo" type = "hidden" value="">
<input type="submit" name="Primero" value="Primero" onclick="return Destino('1')">
<input type="submit" name="Siguiente" value="Siguiente" onclick="return Destino('2')">
<input type="submit" name="Anterior" value="Anterior" onclick="return Destino('3')">
<input type="submit" name="Ultimo" value="Ultimo" onclick="return Destino('4')">
</form>



</div>
</BODY>
</HTML>