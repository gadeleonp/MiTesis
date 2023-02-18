<%@ Language=VBScript %>
<%


if isempty(session("id")) then
  Response.Redirect("Login.asp")
end if
session("destforo") = "desarrollo_foro.asp"

dim db
dim rs
dim rs2
set db = server.CreateObject("SQLComandos.Comandos")

dim foro 
foro = Request.querystring("foro")


%>
<html>
<head>
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
<script language="javascript">
function Correo2(autor) {
   forma.opcion.value = '1'
   forma.autor.value = autor
   forma.target='Preview'
   forma.submit()
}
function Correo(autor) {
pagina = 'formulario_correo.asp?autor='+autor+'&opcion=1'	
		
		  ancho = 700
		  altura = 500
			  if (screen) {
				  leftPos = (screen.width / 2) - (ancho/2)
			  }
			PreviewWindow = window.open(pagina,'Preview','toolbar=no,menubar=no,location=no,scrollbars=yes,width='+ancho+'+,height='+altura+'left='+leftPos+'+,top=20')	
	   forma.opcion.value = '1'
	   forma.autor.value = autor
	   forma.target='Preview'
	   forma.submit()
		
}

/**/
function Editar(respuesta,foro) {
	var destino;
	var str;
	var buscado;
	
	 var s = navigator.appVersion;
	  ss = s.substr(22,1);
	  
	  if (ss.indexOf("6")==0)            
	  {
		destino = "responder_tema2.asp"
      }
	  else 
	  {
		destino = "responder_tema.asp"
	  }
	forma2.action= destino
forma2.respuesta.value= respuesta
	 forma2.foro.value=foro
	 forma2.accion.value = '1'
	 forma2.submit()
}

/**/

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

function Responder_Tema(foro) 
{
var destino;
var str;
var buscado;

 var s = navigator.appVersion;
  ss = s.substr(22,1);
  
  if (ss.indexOf("6")==0)            
  {
	destino = "responder_tema2.asp"
  }
  else 
  {
	destino = "responder_tema.asp"
  }
 destino = destino + '?respuesta=1&foro='+foro+'&tipor=1'
 document.location = destino
}
function Regresar()
{ 
 pagina = 'foro_general.asp'
 document.location=pagina
}
function agregar_documento(respuesta,foro) {
	
	Enlace=prompt("Ingrese el texto que se mostrara como enlace para el documento.\nDeje en blanco si desea que aparezca la palabra Documento como enlace.","");
		if (Enlace!=null) {
		  ancho = 600
			  if (screen) {
				  leftPos = (screen.width / 2) - (ancho/2)
			  }
			PreviewWindow = window.open('agregar_documento.asp?respuesta='+respuesta+'&enlace='+Enlace+'&foro='+foro+'&tipo=2','AgregarDoc','toolbar=no,menubar=no,location=no,scrollbars=yes,width='+ancho+'+,height=250,left='+leftPos+'+,top=20')	
		}
		if (Enlace=="") {
		Enlace="Documento"
		  ancho = 700
			  if (screen) {
				  leftPos = (screen.width / 2) - (ancho/2)
			  }
			PreviewWindow = window.open('agregar_documento.asp?respuesta='+respuesta+'&enlace='+Enlace+'&foro='+foro+'&tipo=2','AgregarDoc','toolbar=no,menubar=no,location=no,scrollbars=yes,width='+ancho+'+,height=350,left='+leftPos+'+,top=20')	
		}
}
function eliminar_documento(respuesta,foro) {
		  ancho = 700
			  if (screen) {
				  leftPos = (screen.width / 2) - (ancho/2)
			  }
		PreviewWindow = window.open('eliminar_documento.asp?respuesta='+respuesta+'&foro='+foro+'&tipo=2','EliminarDoc','toolbar=no,menubar=no,location=no,scrollbars=yes,width='+ancho+'+,height=350,left='+leftPos+'+,top=20')	
}
</script>
<BODY onload="javascript:{if(parent.frames[0]&&parent.frames['opciones'].Go)parent.frames['opciones'].Go()}" >

<table width="766" border="0" cellPadding="2" cellSpacing="2" >
  <tr noWrap class="celdaBgBlue"> 
    <th width="730"> Foro </th>
  </tr>
  	  <%
		Sql = "select f.id,f.titulo,day(f.fecha_creacion) dia,month(f.fecha_creacion) mes,year(f.fecha_creacion) anio, f.autor,f.activo,f.descripcion, u.nombre, u.apellido "
		Sql = Sql & " from foro f, usuario u " 
		Sql = sql & " where f.activo = 1 " 
		sql = Sql & " and f.autor = u.id"
		sql = sql & " and f.id = " & foro
        
        'Response.Write sql
        'Response.end 

		set rs = db.SQLSelect("DBBAECYS",sql)
		if not rs.eof then
		while rs.eof<> true
		foro = rs("id")
		titulo = rs("titulo")
		fecha = rs("dia")&"-"&rs("mes")&"-"&rs("anio")
		descripcion = rs("descripcion")
		
		sql ="select count(id) contador from respuesta where foro = "&foro
		set rs2 = db.SQLSelect("DBBAECYS",sql)
		if not rs2.EOF then
		 total = rs2("contador")
		end if
		set rs2=nothing
		
		sql ="select max(fecha) fecha from respuesta where foro = "&foro
		set rs2 = db.SQLSelect("DBBAECYS",sql)
		if not rs2.EOF then
		 fecha_ultima = rs2("fecha")
		end if
		set rs2=nothing
		
	%>	
   <tr class="celdaBgBlueLeft">
    <td >
      <div align="left">
        <%
    if rs("autor") = session("id") then
    %>
	<!--
	&nbsp; <img alt="Eliminar foro" src="imagenes/delete_forum.gif" onclick="javascript:Eliminar('<%=rs("id")%>')" width="15" height="15"> 
	-->
    &nbsp; <img src="imagenes/edit_forum.gif" alt="Actualizar foro" onclick="javascript:Actualizar('<%=rs("id")%>')" width="15" height="15"> 
	<%
	else%>
	&nbsp; <img alt="Mandar Correo al autor" onclick="javascript:Correo('<%=rs("autor")%>')" src="imagenes/correo.gif" WIDTH="15" HEIGHT="15">     
	<%
	end if
	%>
    &nbsp; <img alt="Regresar" onclick="javascript:Regresar('<%=rs("id")%>')" src="imagenes/regresar_foro.gif" WIDTH="15" HEIGHT="15"> 
    &nbsp; <img alt="Responder tema" onclick="javascript:Responder_Tema('<%=foro%>')" src="imagenes/icon_reply_topic.gif" WIDTH="15" HEIGHT="15"> 
	
	  </div></td>
  </tr>
	<tr>
	<td>
<font color="#000066" size="2" face="Arial, Helvetica, sans-serif"><strong>Autor:</strong></font> &nbsp;
<font color="#0000FF" size="2" face="Arial, Helvetica, sans-serif"><strong><%= rs("nombre")%>&nbsp;<%= rs("apellido")%></strong></font><br>
 <font color="#000066" size="2" face="Arial, Helvetica, sans-serif"><strong>Titulo:</strong></font> &nbsp;
 <font color="#0000FF" size="2" face="Arial, Helvetica, sans-serif"><strong><%= titulo%></strong></font>
 <font color="#000066" size="2" face="Arial, Helvetica, sans-serif"><strong>Fecha Creacion:</strong></font>&nbsp; 
<font color="#0000FF" size="2" face="Arial, Helvetica, sans-serif"><strong> <%=fecha%></strong></font>&nbsp; 
 <font color="#000066" size="2" face="Arial, Helvetica, sans-serif"><strong>Total Visitas:</strong></font> &nbsp;
<font color="#0000FF" size="2" face="Arial, Helvetica, sans-serif"><strong> <%= total%> </strong></font>&nbsp; 
 <font color="#000066" size="2" face="Arial, Helvetica, sans-serif"><strong> Ultima Visita: </strong></font>&nbsp; 
<font color="#0000FF" size="2" face="Arial, Helvetica, sans-serif"><strong> <%= fecha_ultima%> </strong></font>&nbsp; 
      <hr>
	  <font color="#333333" size="2" face="Arial, Helvetica, sans-serif"><%=rs("descripcion")%></font>

    </td>
    </tr>
    <tr>
    <td>
  	<%
  	sql = "select enlace,archivo from documento where foro = " & foro & " and respuesta is null " 
	set rs2 = db.SQLSelect("DBBAECYS",sql)
	if not rs2.EOF then
  	%>
  		<hr>
    <strong><font size="1" face="Arial, Helvetica, sans-serif">Documentos: </font></strong>
          <%
    while rs2.EOF <> true
	%>
 <strong><font size="2" face="Arial, Helvetica, sans-serif"><a href="download.asp?file=<%=rs2("archivo")%>&tipo=1"><%=rs2("enlace")%></a> </font></strong>
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
	set rs = nothing

	%>
	
	<%
	sql = "select r.id, r.titulo,r.descripcion,r.fecha,u.nombre, u.apellido,r.autor "
	sql = sql & " from respuesta r, usuario u "
	sql = Sql & " where u.id = r.autor " 
	sql = Sql & " and r.foro = "& foro
	sql = sql & " order by r.fecha " 
	
		set rs = db.SQLSelect("DBBAECYS",sql)
		if not rs.EOF then


	while rs.eof <> true 
	%>
	<tr class="celdabgblueLeft">
    <td >
      <div align="left">
        <%if rs("autor") = session("id") then%>
		<!--
	&nbsp; <img alt="Eliminar respuesta" src="imagenes/delete_forum.gif" onclick="javascript:Eliminar('<%=rs("id")%>')" width="15" height="15"> 
	-->
    &nbsp; <img alt="Actualizar respuesta" src="imagenes/edit_forum.gif"  onclick="javascript:Editar('<%=rs("id")%>','<%=foro%>')" width="15" height="15"> 
   &nbsp; <img src="imagenes/btn_documento2.bmp" alt="Agregar Documento" onclick="javascript:agregar_documento('<%=rs("id")%>','<%=foro%>')" width="15" height="15">
	&nbsp;<img src="imagenes/btn_eliminadoc2.bmp" alt="Eliminar Documento" onclick="javascript:eliminar_documento('<%=rs("id")%>','<%=foro%>')" width="15" height="15">    
	<%else%>
	&nbsp; <img alt="Mandar Correo al autor" onclick="javascript:Correo('<%=rs("autor")%>')" src="imagenes/correo.gif" WIDTH="15" HEIGHT="15"> 
    <%end if%>
    &nbsp; <img alt="Responder tema" onclick="javascript:Responder_Tema('<%=foro%>')" src="imagenes/icon_reply_topic.gif" WIDTH="15" HEIGHT="15"> 
	
	  </div></td>
	</tr>
<tr>
	<td>
<font color="#000066" size="2" face="Arial, Helvetica, sans-serif"><strong>Autor:</strong></font>&nbsp;
<font color="#0000FF" size="2" face="Arial, Helvetica, sans-serif"><strong><%= rs("nombre")%>&nbsp;<%= rs("apellido")%></strong></font><br>
 <font color="#000066" size="2" face="Arial, Helvetica, sans-serif"><strong>Titulo:</strong></font> &nbsp;
 <font color="#0000FF" size="2" face="Arial, Helvetica, sans-serif"><strong><%= rs("titulo")%></strong></font>
 <font color="#000066" size="2" face="Arial, Helvetica, sans-serif"><strong>Fecha Respueta:</strong></font>&nbsp; 
<font color="#0000FF" size="2" face="Arial, Helvetica, sans-serif"><strong> <%= rs("fecha") %></strong></font>&nbsp; 
      <hr>
	  <font color="#333333" size="2" face="Arial, Helvetica, sans-serif"><%=rs("descripcion")%></font>
	<br>	
	        <%
  
  sql = "select enlace,archivo from documento where respuesta = "& rs("id")
		set rs2 = db.SQLSelect("DBBAECYS",sql)
		if not rs2.EOF then
  
  %>

    <strong><font size="1" face="Arial, Helvetica, sans-serif">Documentos: </font></strong>
          <%
    while rs2.EOF <> true
	%>
<!-- <strong><font size="2" face="Arial, Helvetica, sans-serif"><a href="foros/<%=rs2("archivo")%>" target="newpage"><%=rs2("enlace")%></a> </font></strong>-->
	<strong><font size="2" face="Arial, Helvetica, sans-serif"><a href="download.asp?file=<%=rs2("archivo")%>&tipo=1" ><%=rs2("enlace")%></a> </font></strong>
        <% 
     rs2.MoveNext 
    wend
  end if
set rs2=nothing
  %>
  	<hr>
    </td>
    </tr>
    <%
	rs.MoveNext 
	wend 
	else
	end if
	%>
	
	
	<%else%>
    <td><font color="#333333" size="2" face="Arial, Helvetica, sans-serif">No ha ingresado ningun foro</font></td>
    <%end if%>
</table>
<form name="forma2" action="" method="get">
<input name="foro" type="hidden" value="">
<input name="accion" type="hidden" value="">
	<input name="respuesta" type="hidden" value="">	
</form>
<form name="forma" action="formulario_correo.asp" method="get" target="">
	<input name="autor" type="hidden" value="<%=autor%>">
	<input name="opcion" type="hidden" value="">

</form>
</body>
</html>
<%
set rs = nothing
set db = nothing
%>