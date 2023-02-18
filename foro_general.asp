<%@ Language=VBScript %>
<%
dim usuario
usuario = session("id")
session("destforo") = "foro_general.asp"
session("forogeneral") = 1
%>
<%
if isempty(session("id")) then
	Response.Redirect("login.asp")
end if

dim db
dim rs
dim rs2
set db = server.CreateObject("SQLComandos.Comandos")

%>
<HTML>
<HEAD>
<link href="./theme/Master.css" rel="stylesheet" type="text/css">
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<SCRIPT LANGUAGE="javascript">

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
 
function Seleccion(Cat) {
     form.submit()
     
   return true
  }

</SCRIPT>
<%
dim categoria

Categoria = Request.Form("categoria")
%>
<BODY onload="javascript:{if(parent.frames[0]&&parent.frames['opciones'].Go)parent.frames['opciones'].Go()}" >
<form name="forma2" action="" method="get">
<input name="foro" type="hidden" value="">
</form>
<form name="form" action ="foro_general.asp" method="post">
<div align="center">  
  <table border="0" cellspacing="0">
  <tbody>
    <tr><td>
		<strong><font size="2" face="Arial, Helvetica, sans-serif">Seleccione  una Categoria</font></strong> 
	  </td></tr>
  <tr>
  <td align="center">
  <select name="Categoria" size="1"  onchange="return Seleccion(this)" >
   
        <%
sql = "select id,descripcion from categoria where padre is null"

set rs = db.SQLSelect("DBBAECYS",sql)
if not rs.EOF then

%>
      <%while rs.eof <> true %>
		 <option value="<%=rs("id")%>" <% if rs("id") = Categoria then Response.Write "selected" end if%>>-<%=rs("descripcion")%>-</option>
			  <%
			sql2 = "select id,descripcion from categoria where padre = " & rs("id")
			set rs2 = db.SQLSelect("DBBAECYS",sql2)
			if not rs2.EOF then
					%>
						<%
					while rs2.EOF <> true
							%>
							 <option value="<%=rs2("id")%>" <% if rs("id") = Categoria then Response.Write "selected" end if%> >--<%=rs2("descripcion")%>--</option>
							 <%
							rs2.MoveNext 
					wend 
					%>
					<%
			end if
			set rs2 = nothing
			%>
			  <%
		rs.movenext
		wend
		end if
		set rs = nothing
%>
      </select>

  </td></tr>
  </tbody>
  </table>
  
</div>
</form>
<%


sql = " select f.id idforo,f.titulo titulo ,f.descripcion descripcion,f.fecha_creacion,c.descripcion categoria, f.autor, a.id, a.nombre, a.apellido "
sql = sql & " from foro f, usuario a, categoria c "
sql = sql & " where a.id = f.autor "
sql = Sql & " and c.id = f.categoria "
sql = Sql & " and f.estado = 1 "
sql = Sql & " and f.activo = 1 "

if not isempty(Categoria) and Categoria <>"0" then
sql = sql & " and categoria = "& categoria
end if

sql = Sql & " order by c.descripcion, f.fecha_creacion  asc "

set rs = db.SQLSelect("DBBAECYS",sql)
if not rs.EOF then
%>
<TABLE WIDTH=102% BORDER=0 CELLSPACING=2 CELLPADDING=2>
<%
 while rs.EOF <> true
 titulo = rs("titulo")
 descripcion = rs("descripcion")
 
    sql = "select count(id) Total from respuesta where foro = "& rs("idforo")
	set rs2 = db.SQLSelect("DBBAECYS",sql)
    if not rs2.EOF then
    Total = rs2("total")
    end if
	set rs2 = nothing
    
  
		 Sql = "select max(fecha) FechaF from respuesta where foro =" & rs("idforo")
		set rs2 = db.SQLSelect("DBBAECYS",sql)
		if not rs2.eof then
		fecha_ultima = rs2("FechaF")
		end if
		set rs2 = nothing
 %>
 	   <tr class="celdaBgBlueLeft" >
 	   
    <td  align="left"> 
      <div align="left">
        <%if rs("autor") = session("id") then%>
		<!--
    &nbsp; <img alt="Eliminar foro" src="imagenes/delete_forum.gif" onclick="javascript:Eliminar('<%=rs("idforo")%>')" width="15" height="15"> 
    &nbsp; <img src="imagenes/edit_forum.gif" alt="Actualizar foro" onclick="javascript:Actualizar('<%=rs("idforo")%>')" width="15" height="15"> 
	-->
    <%end if%>
      &nbsp; <img alt="Ir al foro" onclick="javascript:Visitar('<%=rs("idforo")%>')" src="imagenes/responder_foro.gif" WIDTH="15" HEIGHT="15">	</div>    </tr>
	<tr class="celdaBgBlue10">
	<td>
 <font  size="1" face="Arial, Helvetica, sans-serif"><strong>Titulo:</strong></font> &nbsp;
 <font size="2" face="Arial, Helvetica, sans-serif"><strong><%= titulo%></strong></font>
 <font size="1" face="Arial, Helvetica, sans-serif"><strong>Fecha Creacion:</strong></font>&nbsp; 
<font  size="2" face="Arial, Helvetica, sans-serif"><strong> <%= Fecha %></strong></font>&nbsp; 
 <font size="1" face="Arial, Helvetica, sans-serif"><strong>Total Visitas:</strong></font> &nbsp;
<font size="2" face="Arial, Helvetica, sans-serif"><strong> <%= total%> </strong></font>&nbsp; 
 <font  size="1" face="Arial, Helvetica, sans-serif"><strong> Ultima Visita: </strong></font>&nbsp; 
<font size="2" face="Arial, Helvetica, sans-serif"><strong> <%= fecha_ultima%> </strong></font>&nbsp; 
<br>
 <font size="1" face="Arial, Helvetica, sans-serif"><strong> Autor: </strong></font>&nbsp; 
<font size="2" face="Arial, Helvetica, sans-serif"><strong> <%= rs("nombre")%> &nbsp;<%= rs("apellido")%> </strong></font>&nbsp; 
<font size="1" face="Arial, Helvetica, sans-serif"><strong> Categoria: </strong></font>&nbsp; 
<font size="2" face="Arial, Helvetica, sans-serif"><strong> <%= rs("categoria")%> </strong></font>&nbsp; 

      <hr>
	  <font size="2" face="Arial, Helvetica, sans-serif"><%=rs("descripcion")%></font>
	  <br>
        <%
  
  sql = "select enlace,archivo from documento where foro = "& rs("idforo") & " and respuesta is null " 
  
  set rs2 = db.SQLSelect("DBBAECYS",sql)
  if not rs2.eof then
  
  %>
  	<hr>
    <strong><font size="1" face="Arial, Helvetica, sans-serif">Documentos: </font></strong>
          <%
    while rs2.EOF <> true
	%>
 <strong><font size="2" face="Arial, Helvetica, sans-serif"><a href="foros/<%=rs2("archivo")%>" target="newpage"><%=rs2("enlace")%></a> </font></strong>
        <% 
     rs2.MoveNext 
    wend
  end if
  set rs2=nothing
  %>
    </td>
    </tr>


 <%
    rs.movenext
 wend
 %>
</TABLE>
 <%
else
%>
<h2 align="center" >No Existen Foros para esta Categoria</h2>
<%
end if


set db = nothing
%>
</BODY>
</HTML>
