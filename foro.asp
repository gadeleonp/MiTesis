<%@ Language=VBScript %>
<%
dim usuario
usuario = ""
usuario = session("id")


if isempty(usuario) then
	Response.Redirect("login.asp")
end if
session("destforo") = "foro.asp"
%>
<%
dim db
dim rs
dim rs2
set db = server.CreateObject("SQLComandos.Comandos")

%>
<html>
<head>
<title>Foros Publicados</title>
<link href="./theme/Master.css" rel="stylesheet" type="text/css">
<!--
<style type="text/css">

.Presentacion1 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: medium;
	font-style: normal;
	font-weight: bold;
	font-variant: normal;
	text-transform: none;
	color: #0000CC;
	border: medium dotted #000099;
}
Presentacion2 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: medium;
	font-style: normal;
	line-height: normal;
	font-weight: bold;
	font-variant: normal;
	text-transform: none;
	color: #003366;
}
</style>
-->
</head>
<script language="JavaScript">
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
forma2.action= destino
 forma2.foro.value=foro
 forma2.accion.value = '1'
 forma2.submit()
}
function Eliminar(foro) {
valor = confirm("Eliminar este foro.\nEsta accin no sera reversible, Desea Continuar?")
	if (valor == true) 
	{
	parent.principal.location = "./eliminar_foro.asp?foro="+foro
	}
}


 function Visitar(foro) {
 forma2.action='desarrollo_foro.asp'
 forma2.foro.value=foro
 forma2.submit()
 }
 
</script>
<body onload="javascript:{if(parent.frames[0]&&parent.frames['opciones'].Go)parent.frames['opciones'].Go()}">
<form name="forma2" action="" method="get">
<input name="foro" type="hidden" value="">
<input name="accion" type="hidden" value="">
</form>
<table border="0">
  <tr> 
          <td ><a href="reporte_foros.asp" >Foros Creados</a></td>
  </tr>
</table>
<br>
<table width="766" border="0" cellPadding="2" cellSpacing="2" align="center" >
  <tr noWrap > 
    <th class="celdaBgBlue">Foros Publicados</th>
  </tr>
  	  <%
		Sql = "select f.id,f.titulo,day(f.fecha_creacion) dia,month(f.fecha_creacion) mes,year(f.fecha_creacion) anio,f.autor,f.activo,f.descripcion, cat.descripcion categoria from foro f, categoria cat where f.activo = 1 and cat.id = f.categoria and f.estado = 1 and f.autor = "& session("id")
		Sql = Sql & " order by cat.descripcion, f.titulo"
		set rs = db.SQLSelect("DBBAECYS",sql)
		if not rs.eof then
		while rs.eof<> true
		foro = rs("id")
		titulo = rs("titulo")
		fecha = rs("dia")&"-"&rs("mes")&"-"&rs("anio")
		descripcion = rs("descripcion")
		categoria = rs("categoria")
		sql = "select count(id) total from respuesta where foro = "& foro
		set rs2 = db.SQLSelect("DBBAECYS",sql)
		if not rs2.eof then
		total = rs2("total")
		end if
		set rs2 = nothing
		
		Sql = "select max(fecha) FechaF from respuesta where foro =" & rs("id")
		set rs2 = db.SQLSelect("DBBAECYS",sql)
		if not rs2.eof then
		fecha_ultima = rs2("FechaF")
		end if
		set rs2 = nothing
		%>
   <tr class="celdabgblueleft">
    <td > &nbsp; <img alt="Eliminar foro" src="imagenes/delete_forum.gif" onclick="javascript:Eliminar('<%=rs("id")%>')" width="15" height="15"> 
      &nbsp; <img src="imagenes/edit_forum.gif" alt="Actualizar foro" onclick="javascript:Actualizar('<%=rs("id")%>')" width="15" height="15"> 
<!--      &nbsp; <img alt="Ir al foro" onclick="javascript:Visitar('<%=rs("id")%>')" src="imagenes/responder_foro.gif" WIDTH="15" HEIGHT="15"> -->

      
	</tr>
	<tr>
	<td class="celdaBgBlue10">
 <font  size="1" face="Arial, Helvetica, sans-serif"><strong>Titulo:</strong></font> &nbsp;
 <font  size="2" face="Arial, Helvetica, sans-serif"><strong><%= titulo%></strong></font>
 <font  size="1" face="Arial, Helvetica, sans-serif"><strong>Fecha Creacion:</strong></font>&nbsp; 
<font  size="2" face="Arial, Helvetica, sans-serif"><strong> <%= Fecha %></strong></font>&nbsp; 
 <font size="1" face="Arial, Helvetica, sans-serif"><strong>Total Visitas:</strong></font> &nbsp;
<font size="2" face="Arial, Helvetica, sans-serif"><strong> <%= total%> </strong></font>&nbsp; 
 <font  size="1" face="Arial, Helvetica, sans-serif"><strong> Ultima Visita: </strong></font>&nbsp; 
<font  size="2" face="Arial, Helvetica, sans-serif"><strong> <%= fecha_ultima%> </strong></font>&nbsp; <br>
<font  size="1" face="Arial, Helvetica, sans-serif"><strong> Categoria: </strong></font>&nbsp; 
<font  size="2" face="Arial, Helvetica, sans-serif"><strong> <%= categoria%> </strong></font>&nbsp; 

      <hr>
	  <font color="#333333" size="2" face="Arial, Helvetica, sans-serif"><%=rs("descripcion")%></font>
	</td>
	</tr>
	<tr>
	<td class="celdaBgBlue05">
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
  
    </td>
    </tr>
  	<%
	rs.movenext 
	wend
	set rs = nothing
	%>
	<%else%>
    <td><font color="#333333" size="2" face="Arial, Helvetica, sans-serif">No ha ingresado ningun foro</font></td>
    <%end if%>
</table>
</body>
</html>
<%
set db  = nothing
%>