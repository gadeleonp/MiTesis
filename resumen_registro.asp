<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
dim db
dim sql
set db = server.CreateObject("SQLComandos.Comandos")

sql = "select u.id,u.nombre,u.apellido,u.direccion,u.telefono,u.mail,u.login,u.carnet,u.colegiado,u.tipo,u.arc_resume,u.sexo,p.nombre nacionalidad from usuario u,pais p where u.id = "& session("id") & " and u.pais = p.id"

set rs = db.SQLSelect("DBBAECYS",sql)

if not rs.eof then
login = rs("login")
nombre = rs("nombre")
apellido = rs("apellido")
direccion = rs("direccion")
telefono = rs("telefono")
mail = rs("mail")
carnet = rs("carnet")
colegiado = rs("colegiado")
tipo = rs("tipo")
curri = rs("arc_resume")
sexo = rs("sexo")
nacionalidad = rs("nacionalidad")
end if
%>
<html>
<head>
<link href="./theme/Master.css" rel="stylesheet" type="text/css">
<title>Resumen Datos Registro</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<script language="javascript">
function redireccion() {
destino = "./principal.asp?ver=0&path=opciones_datos.asp"
window.parent.location = destino
}
</script>
<body onload="javascript:{if(parent.frames[0]&&parent.frames['opciones'].Go)parent.frames['opciones'].Go()}">
<div align="center">
 <h1>Datos Personales</h1>
<table border = "0" cellpadding="2" cellspacing="2">
    <tbody>
      <tr> 
        <th width="140" nowrap class="celdaBgBlueLeft"> <div align="right"><b>Nombre</div></th>
        <th width="389" class="celdaBgBlue05"><div align="left"><%=nombre%>&nbsp;<%=apellido%> </div></th>
      </tr>
      <tr> 
        <th width="140" nowrap class="celdaBgBlueLeft"> <div align="right"><b>Usuario: </b></div></th>
        <th width="389" class="celdaBgBlue05"> 
		
		  <div align="left"><%=login%> 
          </div></th>
      </tr>
      <tr> 
        <th width="140" nowrap class="celdaBgBlueLeft"> <div align="right"><b>Direccion: </b></div></th>
        <th width="389" class="celdaBgBlue05"> 
		  <div align="left"><%=direccion%> 
		            </div>
            </th>

      </tr>
      <tr> 
        <th width="140" nowrap class="celdaBgBlueLeft"> <div align="right"><b>Telefono: </b></div></th>
        <th width="389" class="celdaBgBlue05"> 
		  <div align="left"><%=telefono%> 
          </div></th>
      </tr>
      <tr> 
        <th width="140" nowrap class="celdaBgBlueLeft"> <div align="right"><b>Correo Electronico: </b></div></th>
        <th width="389" class="celdaBgBlue05"><div align="left"><%=mail%>
          </div>
        </td>
      </tr>
	        <%
	if cint(tipo) = 1 then
%>
      <tr> 
        <th width="140" nowrap class="celdaBgBlueLeft"> <div align="right"><b>No. Carnet: </b></div></th>
        <th width="389" class="celdaBgBlue05"> <div align="left"><%=carnet%> 
</div></th>
      </tr>
      <%
	else
%>
      <tr> 
        <th width="140" nowrap class="celdaBgBlueLeft"> <div align="right"><b>No. Colegiado: </b></div></th>
        <th width="389" class="celdaBgBlue05"> <div align="left"><%=colegiado%> 
</div></th>
      </tr>
      <%
	end if
%>
<%
set rs = nothing
sql ="select u.nombre,p.nombre pais from universidad u, usuario us, pais p where p.id = u.pais and us.universidad = u.id and us.id = "& session("id")  


set rs = db.SQLSelect("DBBAECYS",sql)
if not rs.eof then
%>
      <tr> 
        <th nowrap class="celdaBgBlueLeft"><div align="right"><strong>Sexo:</strong></div></th>
        <th class="celdaBgBlue05"> <div align="left">
            <%
		if sexo = "1" then
		response.write "Masculino"
		end if
		if sexo = "2" then
		response.write "Femenino"
		end if
		%>
</div></th>
      </tr>
      <tr> 
        <th nowrap class="celdaBgBlueLeft"><div align="right"><strong>Nacionalidad</strong></div></th>
        <th class="celdaBgBlue05"><div align="left"><%=nacionalidad%></div></th>
      </tr>
      <tr> 
        <th nowrap class="celdaBgBlueLeft"><div align="right"><strong> Universidad </strong></div></th>
        <th class="celdaBgBlue05"> <div align="left"><%=rs("nombre")%> de <%=rs("pais")%></div></th>
      </tr>
      <%
else
%>
      <tr> 
        <th nowrap class="celdaBgBlueLeft"><div align="right"><strong> Universidad </strong></div></th>
        <th class="celdaBgBlue05"> <div align="left"><font color="#0033FF">Informacion no disponible </font></div></th>
      </tr>
      <%
end if
set rs = nothing %>
    </tbody>
  </table>
<!-- -->

<h2>Areas Seleccionadas</h2>

<%

usuario = session("id")
dim rs2
dim rs3

sql = "select id,descripcion from categoria where padre is null"
set rs = db.SQLSelect("DBBAECYS",sql)
if not rs.EOF then
%>
    <table border="0" cellpadding="2" cellspacing="2" width="450">
      <tbody> 
      <%while rs.eof <> true %>
      <tr class="celdaBgBlue"> 
        <th align="center"><%=rs("descripcion")%></th>
      </tr>

      <%
	sql2 = "select id,descripcion from categoria where padre = " & rs("id")
	set rs2 = db.SQLSelect("DBBAECYS",sql2)
	if not rs2.EOF then
	%>
	 <tr>
     <td align="center">
	<table border="0" cellpadding="2" cellspacing="2" width="450">
		<%
	while rs2.EOF <> true
	%>
        <tr class="celdaBgBlue10"> 

          <td><b><%=rs2("descripcion")%></b></td>
     </tr>
     <%
	sql3 = "select id,descripcion,explicacion from area where categoria = " & rs2("id")  & " and id in ( select area from conocimiento where usuario = " & usuario & ")"
	
	set rs3 = db.SQLSelect("DBBAECYS",sql3)
	if not rs3.EOF then
	%>
	 <tr>
     <td align="center">
	<table border="0" cellpadding="2" cellspacing="2" width="450">
		<%
	while rs3.EOF <> true
	%>
        <tr class="celdaBgBlue05"> 
		  <td></td>
		  <td><%=rs3("descripcion")%></td>
		  <% if len(rs3("explicacion")) > 0 then 
		  %>
			<td><img src="./iconos/help.gif" alt="<%=rs3("explicacion")%>" WIDTH="16" HEIGHT="16"></td>
		  <%else%>
		  <td>&nbsp;&nbsp;&nbsp;&nbsp;</td>
		  <%end if%>
     </tr>
     <%
	rs3.MoveNext 
	wend 
	%>
     </table>
     </td>
</tr>
	<%
	else
	%>
	<tr><td>Area No Seleccionada</td></tr>
	<%
	end if
	set rs3 = nothing
	%>
     <%
	rs2.MoveNext 
	wend 
	%>
     </table>
     </td>
</tr>
	<%
	end if
set rs2 = nothing
	%>
      <%
rs.MoveNext 
wend
%>
      </tbody> 
    </table>
<%
end if
%>
<br>
<!-- -->
  <table align="center" border="0" cellpadding="2" cellspacing="2">
  <tr>
  <td>
<input name="Opciones" type="submit" class="celdaBgBlue05" onclick="javascript:redireccion()" value="Opciones">
</td>
</tr>
</table>
</div>
</body>
</html>