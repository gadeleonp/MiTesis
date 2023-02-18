<%@ Language=VBScript %>
<%
dim db
dim rs
set db = server.CreateObject ("SQLComandos.Comandos")
%>
<HTML>
<HEAD>
<link href="../theme/Master.css" rel="stylesheet" type="text/css">
<title>Datos de la Universidad</title>
</HEAD>
<BODY>
<script language="javascript">
function Regresar() {
 form1.action = "mant_universidades.asp"
 form1.submit
 return true
 }
</script>
<h2 align="center">Datos Universidad</h2>
 <form name="form1" method="post" action="ingreso_universidad.asp">
  <table align="center" width="75%" border="0" cellpadding="2" cellspacing="2">
    <tr> 
      <th width="29%" nowrap class="celdaBgBlueLeft"><div align="left">Nombre de la Universidad</div></th>
      <td width="71%" class="celdaBgBlue05"> <div align="left">
        <input type="text" name="nombre"> 
      </div></td>
    </tr>
    <tr> 
      <th width="29%" nowrap class="celdaBgBlueLeft"><div align="left">Pais</div></th>
      <td width="71%" class="celdaBgBlue05"> <div align="left">
        <select name="pais" size="1" >
            <option value ="0" selected>--Selecciona una Opcion--</option>
            <%
    
    sql = "select id,nombre from pais order by nombre"
    set rs = db.SQLSelect("DBBAECYS",sql)
    while rs.eof <> true
    %>
            <option value ="<%=rs("id")%>" ><%=rs("nombre")%></option>
            <%
    rs.movenext
    wend
    set rs = nothing
    %>
        </select> 
      </div></td>
    </tr>
    <tr> 
      <th width="29%" nowrap class="celdaBgBlueLeft"><div align="left">Direccion</div></th>
      <td width="71%" class="celdaBgBlue05"> <div align="left">
        <textarea name="direccion" cols="33" rows="3"></textarea>
      </div></td>
    </tr>
    <tr> 
      <th width="29%" nowrap class="celdaBgBlueLeft"><div align="left">Telefono</div></th>
      <td width="71%" class="celdaBgBlue05"><div align="left">
        <input name="telefono" type="text">
      </div></td>
    </tr>
    <tr> 
      <th width="29%" nowrap class="celdaBgBlueLeft"><div align="left">Sitio en Internet</div></th>
      <td width="71%" class="celdaBgBlue05"><div align="left">
        <input name="sitio" type="text">
      </div></td>
    </tr>
    <tr> 
      <th width="29%" nowrap class="celdaBgBlueLeft"><div align="left">Correo Electronico</div></th>
      <td width="71%" class="celdaBgBlue05"><div align="left">
        <input type="text" name="email"> 
      </div></td>
    </tr>
  </table>
<table align="center">
  <tbody>
  <tr>
    <td><input name="Aceptar" type="submit" class="celdaBgBlue05" value="Aceptar"></td>
	<td width="17"></td>
	    <td><input name="Cancelar" type="submit" class="celdaBgBlue05" onclick="return Regresar()" value="Cancelar"></td>
  </tr>
  </tbody>
</table>
</form>
</BODY>
</HTML>
