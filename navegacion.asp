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
<link href="./styles/interior.css" rel="stylesheet" type="text/css">
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<script language="javascript">
function Mostrar(Pagina) {

document.parentWindow.location = Pagina

}
</script>

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
<BODY>
<div align="center"  > <a href="./principal.asp"  target="_top">Regresar</a>
  <%
sql ="select id,correlativo,enlace from documento_noticia where noticia = "& noticia & " order by correlativo"
set rs = db.SQLSelect("DBBAECYS",sql)
if not rs.eof then
while rs.EOF <> true 
path = "./Noticias/"
path = path + folder +"/"
path = path + rs("enlace")
%>
  &nbsp; <a href="<%=path%>" target="principal"><font size = "2" color="#0033FF">[<%= rs("correlativo")%>]</font></a> 
  <%
rs.MoveNext 
wend
end if
%>
</div>
</BODY>
</HTML>
<%
set rs = nothing

%>
