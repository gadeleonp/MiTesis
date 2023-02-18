<%@ Language=VBScript %>
<%
dim usuario
usuario = session("id")
dim db
dim rs
dim rs2
dim rs3
dim rs4
set db = server.CreateObject ("SQLComandos.Comandos")
sql = "select id,descripcion from categoria where padre is null"
%>
<html>
<head>
<link href="./theme/Master.css" rel="stylesheet" type="text/css">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</head>
<script language="JavaScript" type="text/JavaScript">
function regresar() {
 form.action='actualiza_preferencias.asp'
 form.submit()
 return true
}
</script>
<body onload="javascript:{if(parent.frames[0]&&parent.frames['opciones'].Go)parent.frames['opciones'].Go()}">
<div align="center">
<h2>Selecciona las areas de su conocimiento.</h2>
<form name="form" method="post" action="ingresar_conocimiento.asp">
<%
set rs = db.SQLSelect("DBBAECYS",sql)
if not rs.eof then
%>
    <table  class="celdaBgBlue10" border="0" cellpadding="2" cellspacing="2">
      <tbody> 
		  <%while rs.eof <> true %>
		  <tr class="celdaBgBlue"> 
			<th align="center"><b><font ><%=rs("descripcion")%></font></b></th>
		  </tr>
	
		  <%
		sql2 = "select id,descripcion from categoria where padre = " & rs("id")
		
		set rs2 = db.SQLSelect("DBBAECYS",sql2)
		
		if not rs2.eof then
		%>
		 <tr>
		 <td align="center">
		<table  border="0" cellpadding="2" cellspacing="2">
			<%
		while rs2.EOF <> true
		%>
			<tr> 
	
			  <td><b><%=rs2("descripcion")%></b></td>
		 </tr>
		 <%
		sql3 = "select id,descripcion,explicacion from area where categoria = " & rs2("id")
		
		set rs3=db.SQLSelect("DBBAECYS",sql3)
		if not rs3.eof then
		%>
		 <tr>
		 <td align="center">
		<table border="0" cellpadding="2" cellspacing="2">
			<%
		while rs3.EOF <> true
		ingresado = 0
		sql4 = "select area from conocimiento where usuario = "& usuario &" and area = "&rs3("id")
		set rs4=db.SQLSelect("DBBAECYS",sql4)
		if not rs4.eof then
		ingresado = 1
		end if
	
		set rs4 = nothing
	
		
		
		%>
			<tr class="celdaBgBlueLeft" > 
			  <td><input name="area" type="checkbox" value="<%=rs3("id")%>"<%if ingresado > 0 then response.write " checked " end if%>></td>
			  <td width="350"><%=rs3("descripcion")%></td>
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
		end if
		set rs3= nothing
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
		set rs2= nothing
		%>
		  <%
	rs.MoveNext 
	wend
	%>
      </tbody> 
    </table>
<%
end if
set rs= nothing
set db = nothing
%>
<br>
<input name="usuario" type="hidden" value="<%=usuario%>">
<table>
<tr>
<td><input name="Regresar"  type="button" class="celdaBgBlue05" onclick="return regresar()" value="<- Regresar"></td>
<td width="17"></td>
<td><input name="Aceptar" type="submit" class="celdaBgBlue05" value="Continuar ->"></td>
</tr>
</table>
</form>
</div>
</body>
</html>
