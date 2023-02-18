<%@ Language=VBScript %>
<%
if isempty (session("administrador")) then
 Response.Redirect("Login.asp")
end if
%>

<HTML>
<HEAD>
<link href="../theme/Master.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY>
<h2 align = "center" >Noticias Ingresadas</h2>
<form name="noticia" method="post" action="mantnoticias.asp">
  <table width="85%" border="0.5" align = "center" cellpadding="2" cellspacing="2">
    <tr class="celdaBgBlue">
    <th width="54%" nowrap> 
        <div align="center">Titulo Noticia</div>
    </th>
      <th width="25%" nowrap> 
        <div align="center">Autor</div>
    </th>
      <th width="17%" nowrap> 
        <div align="center">Fecha Publicacion</div>
    </th>
    <th width="17%" nowrap> 
        <div align="center">Folder</div>
    </th>
    <th width="17%" nowrap> 
        <div align="center">No. Documentos</div>
    </th>
    <th width="17%" nowrap> 
        <div align="center">Enalce</div>
    </th>
  </tr>
<%
dim db 
dim rs
dim sql
dim rs2
 set db = server.CreateObject("SQLComandos.Comandos")
 
 
 
 Sql = "select id,titulo, resumen, autor, fecha_publicacion,folder_contenedor from noticia"
 set rs = db.SQLSelect("DBBAECYS",Sql)
  if not rs.eof then
  cont = 0
    while rs.eof <> true
  fecha = rs("fecha_publicacion") 
  fecha = left(fecha,8)
	if (cont mod 2) = 0 then
	color = "celdaBgBlue05"
	else
	color = "celdaBgBlue10"
	end if
	cont = cont + 1
	%>
	
    <tr class="<%=color%>">  
      <td width="4%" nowrap> 
        <div align="center"><%=rs("titulo")%></div>
      </td>
      <td width="54%" nowrap> 
        <div align="center"><%=rs("Autor")%></div>
        
      </td>
      <td width="25%" nowrap>
<div align="center"><%=fecha%></div>
</td>
      <td width="25%" nowrap>
<div align="center"><%=rs("folder_contenedor")%></div>
</td>
<%
sql = "select count(correlativo) correlativo from documento_noticia where noticia = " & rs("id")
set rs2 = db.SQLSelect("DBBAECYS",sql)
if not rs2.EOF  then
%>
<td width="25%" nowrap>
<div align="center"><%=rs2("correlativo")%></div>
</td>
<%

set rs2 = nothing
end if
sql = "select id,enlace,correlativo from documento_noticia where correlativo = 1 and noticia = " & rs("id")

set rs2 = db.SQLSelect("DBBAECYS",sql)
if not rs2.EOF then
%>
      <td width="17%" nowrap><a href="./ver_noticia.asp?folder=<%=rs("folder_contenedor")%>&noticia=<%=rs("id")%>">Ver Noticia.</a></td>
  </tr>

<%
set rs2=nothing
end if
rs.movenext
wend 
end if

%>
</table>
<br>
<table align = "center">
<tbody>
<tr>
<td><input type="submit" name="Regresar" value="Regresar"></td>
</tr>
</tbody>
</table>
</form>
</BODY>
</HTML>
