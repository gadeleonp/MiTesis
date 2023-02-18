<%@ Language=VBScript %>
<%
dim db
dim rs
set db = server.CreateObject ("SQLComandos.Comandos")
sql = "select id,descripcion from categoria order by descripcion"
set rs = db.SQLSelect("DBBAECYS",sql)
%>
<HTML>
<HEAD>
<link href="../theme/Master.css" rel="stylesheet" type="text/css">
<script>
function Regresar() {
  form1.action = "mant_areas.asp"
return true
}
</script>

</HEAD>
<BODY>
<div align="center">
  <h2>Ingreso Categoria </h2>
  <form name="form1" method="post" action="./ingresar_categoria.asp">
    <table width="43%" border="0" cellpadding="2" cellspacing="2">
      <tr>
        <td width="32%" class="celdaBgBlueLeft"><font size="2" face="Arial, Helvetica, sans-serif" class="celdaBgBlueLeft">Descricpcion 
          Categoria</font></td>
        <td width="68%" class="celdaBgBlue05"> 
          <input type="text" name="categoria" size="50" maxlength="50">    </td>
  </tr>
  <tr>
        <td width="32%" height="32" class="celdaBgBlueLeft"><font size="2">Categoria Padre</font></td>
        <td width="68%" height="32" class="celdaBgBlue05"> 
          <select name="padre" size="1">
          <option value="0" selected >--Ninguno--</option>
          <%if rs.EOF <> true then
          while rs.EOF <> true 
		  %>
		  <option value="<%=rs("id")%>" ><%=ucase (rs("descripcion"))%></option>
		  <%
		  rs.MoveNext 
		  wend
		  end if
		set rs = nothing
		set db = nothing
		  %>
      </select>    </td>
  </tr>
</table><br>

<table border="0" cellpadding="2" cellspacing="2">
<tr>
<td>
      <input name="Aceptar" type="submit" class="celdaBgBlue05" value="Aceptar">
	  </td>
	  <td width="17"></td>
	  <td>
      <input name="Cancelar" type="submit" class="celdaBgBlue05" onclick="return Regresar()" value="Cancelar">
	  </td>
</tr>	  
 </table>

  </form>
</div>
</BODY>
</HTML>
