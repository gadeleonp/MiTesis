<%@ Language=VBScript %>
<%
dim db
dim rs
dim sql



set db = server.CreateObject ("SQLComandos.Comandos")

sql = "select id,descripcion from categoria order by descripcion"

set rs = db.SQLSelect("DBBAECYS",sql)
%>
<HTML>
<HEAD>
<link href="../theme/Master.css" rel="stylesheet" type="text/css">
</HEAD>
<script>
function Regresar() {
  form1.action = "mant_areas.asp"
return true
}
</script>
<BODY>
<div align="center">
  <h2>Ingreso Area </h2>
  <form name="form1" method="post" action="./ingresar_area.asp">
    <table width="43%" border="0" cellpadding="2" cellspacing="2">
      <tr>
        <td width="32%" class="celdaBgBlueLeft"><font face="Arial, Helvetica, sans-serif" size="2">Descricpcion Area</font></td>
        <td width="68%" class="celdaBgBlue05"> 
          <div align="left">
            <input type="text" name="area" size="50" maxlength="50">    
          </div></td>
  </tr>
  <tr>
        <td width="32%" height="32" class="celdaBgBlueLeft"><font size="2">Categoria Padre</font></td>
        <td width="68%" height="32" class="celdaBgBlue05"> 
          <div align="left">
            <select name="categoria" size="1">
              <option value="0" selected >--Ninguno--</option>
              <%
			  
			  
          while rs.EOF <> true 
		  %>
		      <option value="<%=rs("id")%>" ><%=rs("descripcion")%></option>
		      <%
		  rs.MoveNext 
		  wend

		  set rs = nothing

		  %>
            </select>
          </div></td>
  </tr>
  <tr>
	<td class="celdaBgBlueLeft">Explicacion del Area:</td>
	<td class="celdaBgBlue05">
          <div align="left">
            <textarea name="explicacion" cols="33" rows="3" ></textarea>
          </div></td>
  </tr>
</table>
    <br>
    <table border="0" cellpadding="2" cellspacing="2">
<td>
      <input name="Aceptar" type="submit" class="celdaBgBlue05" value="Aceptar">
	  </td>
	  <td width="17">
	  </td>
	  <td>
      <input name="Cancelar" type="submit" class="celdaBgBlue05" onclick="return Regresar()" value="Cancelar">
	  </td>
  </table>

  </form>

<%
sql = "select a.descripcion area,c.descripcion categoria, a.explicacion explicacion from area a, categoria c where a.categoria  = c.id order by categoria"
set rs = db.SQLSelect("DBBAECYS",sql)
%>
<h2>Areas Ingresadas</h2>
<table border ="0" cellpadding="2" cellspacing="2">
 <tbody>
   <tr class="celdaBgBlue">
    <th>Area</th>
	  <th>Categoria</th>
	  <th>Explicacion</th>
	 </tr>
   <%while rs.eof <> true %>
   <tr class="celdaBgBlue05">
	 <td><%=rs("area")%></td>
	 <td><%=rs("Categoria")%></td>
	 <td><textarea name="des_expli" cols="33" rows="3"><%=rs("explicacion")%></textarea></td>	 
   </tr>
   <%
   rs.movenext
   wend 
   %>
 </tbody>
</table>
</div>
</BODY>
<%
		set rs = nothing
		set db = nothing
%>
</HTML>
