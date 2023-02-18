<%@ Language=VBScript %>
<%
dim rs
dim area
dim usuario
dim pos
set db = server.createobject ("SQLComandos.Comandos")
area = ""
usuario = ""
usuario = session("id")
area = Request.Form("area")

	sql = "delete from conocimiento where area not in ("&area&") and usuario = "& usuario
	resultado = db.SQLOtras("DBBAECYS",sql)


area = area + ","



While Len(area) > 0
	pos = InStr(1, area, ",", 1)
	area_in = Left(area, pos - 1)
	

	sql = "select area from conocimiento where usuario = "& usuario & " and area = "& area_in
	
	set rs = db.SQLSelect("DBBAECYS",sql)

	if rs.EOF then
	
		sql =  "insert into conocimiento (area,usuario) values (" & area_in & "," & usuario & ")"  
	
		resultado = db.SQLOtras("DBBAECYS",sql)
	end if
    If Len(area) > (pos + 1) Then
    
        area = Right(area, Len(area) - (pos + 1))
        
    Else
    
        area = ""
        
    End If
    
set rs = nothing
Wend

%>
<HTML>
<HEAD>
<link href="./theme/Master.css" rel="stylesheet" type="text/css">
<script language="javascript">
function regresar() {
 Conocimiento.action='ingreso_conocimiento.asp'
 Conocimiento.submit()
}
</script>
</HEAD>
<BODY onload="javascript:{if(parent.frames[0]&&parent.frames['opciones'].Go)parent.frames['opciones'].Go()}">
<div align="center">
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
		  <td><input name="area" type="checkbox" checked value="<%=rs3("id")%>"></td>
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
<form name="Conocimiento" method="post" action="resumen_registro.asp">
<input name="usuario" type="hidden" value="<%=usuario%>">
<table border="0" cellpadding="2" cellspacing="2">
<tr>
<td><input name="Regresar" type="button" value="<--Regresar" class="celdaBgBlue05" onclick="return regresar()"></td>
<td width="17"></td>
<td><input name="Continuar" type="submit" value="Continuar-->" class="celdaBgBlue05"></td>
</tr>
</table>
</form>
</div>
</BODY>
</HTML>
<%
set rs = nothing
set db = nothing
%>