<%
dim usuario
usuario = session("id")


if isempty(usuario) then
Response.Redirect("login.asp")
end if
session("destforo") = "reporte_foros.asp"
%>
<HTML>
<HEAD>
<title>Foros Creados</title>
<link href="./theme/Master.css" rel="stylesheet" type="text/css">
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<!--
<style type="text/css">

table {
	border-top: thin groove #CCCCCC;
	border-right: thin groove;
	border-bottom: thin groove;
	border-left: thin groove;
}
</style>
-->
</HEAD>
<script language="javascript">
function ForosPublicados() {
 
 publicar.action = 'foro.asp'
 publicar.submit()
 
}
function NuevoForo() {
  pagina = 'nuevoforo.asp'
  document.location = pagina
}
function Eliminar(foro) {
valor = confirm("Eliminar este foro.\nEsta accin no sera reversible, Desea Continuar?")
	if (valor == true) 
	{
	parent.principal.location = "./eliminar_foro.asp?foro="+foro
	}
}
function Actualizar(foro) {
var destino;
var str;
var buscado;

 var s = navigator.appVersion;
  ss = s.substr(22,1);
  
  if (ss.indexOf("6")==0)            
  {
	destino = "nuevoforo2.asp"
  }
  else 
  {
	destino = "nuevoforo.asp"
  }
 document.publicar1.action= destino
 document.publicar1.foro.value=foro
 document.publicar1.accion.value = '1'
 document.publicar1.submit()
}

function agregar_documento(foro) {
	
	Enlace=prompt("Ingrese el texto que se mostrara como enlace para el documento.\nDeje en blanco si desea que aparezca la palabra Documento como enlace.","");
		if (Enlace!=null) {
		  ancho = 600
			  if (screen) {
				  leftPos = (screen.width / 2) - (ancho/2)
			  }
			PreviewWindow = window.open('agregar_documento.asp?foro='+foro+'&enlace='+Enlace,'AgregarDoc','toolbar=no,menubar=no,location=no,scrollbars=yes,width='+ancho+'+,height=250,left='+leftPos+'+,top=20')	
		}
		if (Enlace=="") {
		Enlace="Documento"
		  ancho = 700
			  if (screen) {
				  leftPos = (screen.width / 2) - (ancho/2)
			  }
			PreviewWindow = window.open('agregar_documento.asp?foro='+foro+'&enlace='+Enlace,'AgregarDoc','toolbar=no,menubar=no,location=no,scrollbars=yes,width='+ancho+'+,height=350,left='+leftPos+'+,top=20')	
		}
}
function eliminar_documento(foro) {
		  ancho = 700
			  if (screen) {
				  leftPos = (screen.width / 2) - (ancho/2)
			  }
		PreviewWindow = window.open('eliminar_documento.asp?foro='+foro,'EliminarDoc','toolbar=no,menubar=no,location=no,scrollbars=yes,width='+ancho+'+,height=350,left='+leftPos+'+,top=20')	
}

function RevisarSeleccion() {
	len = document.publicar.elements.length;
	var i=0;
	for( i=0; i<len; i++) {
		if (document.publicar.elements[i].name=='publicar') {
			if (document.publicar.elements[i].checked==true) 
			{
			   return true
			}
		}
	}
	alert("Error: No se selecciono ningun foro para publicar.")
	return false
}
 
</script>
<BODY onload="javascript:{if(parent.frames[0]&&parent.frames['opciones'].Go)parent.frames['opciones'].Go()}" >

<%
dim db
dim rs
dim rs2
set db = server.CreateObject ("SQLComandos.Comandos")
sql = "select f.activo,f.id,f.titulo,f.descripcion,day(f.fecha_creacion) dia, month(f.fecha_creacion) mes, year(f.fecha_creacion) anio,c.descripcion categoria from foro f,categoria c where f.autor = "& usuario
sql = sql & " and c.id = f.categoria order by f.fecha_creacion desc"
set rs = db.SQLSelect("DBBAECYS",sql)
if not rs.EOF then
%>
<table border="0">
  <tr> 
		  <td><a href="javascript:ForosPublicados()">Foros Publicados</a></td>
  </tr>
</table>
<div align="center">
<form name="publicar1" action="" method="get" >
<input name="foro" type="hidden">
<input name="accion" type="hidden">
</form>
<form name="publicar" action="eliminar_foro.asp" method="post" onSubmit="return RevisarSeleccion()">
<input name="accion" type="hidden" value="1">
<table width="766" border="0" align="center" cellPadding="2" cellSpacing="2" >
  <tr noWrap> 

        <th class="celdaBgBlue">-</th>
        <th class="celdaBgBlue">Foros Creados</th>
    <%
	cont = 0
while rs.EOF <> true 
fecha = rs("dia") &"-"& rs("mes")  &"-"& rs("anio")
categoria = rs("categoria")

  if (cont mod 2) = 0  then
   color = "celdaBgBlue05"
   else
   color = "celdaBgBlue10"
   end if
   cont = cont + 1
%>
  <tr class= "celdabgblueleft">
    <td ><input name="publicar" type="checkbox" value="<%=rs("id")%>" <%if rs("activo") <> "0" then response.write "checked" end if%>></td>
    <td >
    <img alt="Eliminar foro" src="imagenes/delete_forum.gif" onclick="javascript:Eliminar('<%=rs("id")%>')" width="15" height="15">
    &nbsp;
    <img src="imagenes/edit_forum.gif" alt="Actualizar foro" onclick="javascript:Actualizar('<%=rs("id")%>')" width="15" height="15">
    &nbsp;
    <img src="imagenes/btn_documento2.bmp" alt="Agregar Documento" onclick="javascript:agregar_documento('<%=rs("id")%>')" width="15" height="15">
	&nbsp;
    <img src="imagenes/btn_eliminadoc2.bmp" alt="Eliminar Documento" onclick="javascript:eliminar_documento('<%=rs("id")%>')" width="15" height="15">    
    </td>
    </tr>
  <tr class= "celdabgblue10"> 
    <td>&nbsp; </td>
      <td><strong><font size="1" face="Arial, Helvetica, sans-serif">Titulo</font></strong>: 
        <strong> <font size="2" face="Arial, Helvetica, sans-serif"><%=rs("titulo")%></font></strong> &nbsp;&nbsp;<strong><font size="1" face="Arial, Helvetica, sans-serif">Fecha Publicacion:</font></strong>&nbsp;&nbsp;<%= fecha%> 
		<font  size="1" face="Arial, Helvetica, sans-serif"><strong> Categoria: </strong></font>&nbsp; 
		<font  size="2" face="Arial, Helvetica, sans-serif"><strong> <%= categoria%> </strong></font>&nbsp; 
        <hr>

    <%= rs("descripcion")%> 
  </tr>
  <tr>
  <td class="celdabgblue05">
  </td>
        <td class="celdabgblue05" height="18"> 
          <%
  
  sql = "select enlace,archivo from documento where foro = "& rs("id") & " and respuesta is null " 
	set rs2 = db.SQLSelect("DBBAECYS",sql)
  if not rs2.eof then
  
  %>
  	<hr>
    <strong><font size="1" face="Arial, Helvetica, sans-serif">Documentos: </font></strong>
          <%
    while rs2.EOF <> true
	%>
 <strong><font size="2" face="Arial, Helvetica, sans-serif"><a href="foros/<%=rs2("archivo")%>"><%=rs2("enlace")%></a> </font></strong>
        <% 
     rs2.MoveNext 
    wend
  end if
	set rs2 = nothing
  %>
  <td>
  </tr>
  <%
rs.movenext 
wend 
else
%>
  <h2>No hay Foros Creados</h2>
  <%
end if
set rs = nothing
%>
</table>
<%
set db = nothing
%>
<br>
<input name="foro" value="" type="hidden">
<input name="accion" value ="" type="hidden">
<input name="aceptar" type="submit" class="celdaBgBlue05" value="Publicar">
</form>
</div>
</BODY>
</HTML>