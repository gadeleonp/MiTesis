<%@ Language=VBScript %>
<%
if isempty (session("administrador")) then
 Response.Redirect("Login.asp")
end if
dim db
dim rs

set db = server.CreateObject ("SQLComandos.Comandos")
dim universidad
dim resultado
universidad = Request.Form("universidad")

%>
<HTML>
<HEAD>
<link href="../theme/Master.css" rel="stylesheet" type="text/css">
<title>Datos de la Universidad</title>
</HEAD>
<script language="javascript">
function Regresar() {
 form1.action = "listado_universidad.asp"
 form1.submit
 return true
 }
</script>
<BODY>
<%
sql = "select id,nombre,direccion,telefono,sitio,email,pais from universidad where id = " & universidad
set rs = db.sqlselect("DBBAECYS",sql)
if not rs.eof  then
id = rs("id")
nombre = rs("nombre")
direccion = rs("direccion")
telefono = rs("telefono")
sitio = rs("sitio")
email = rs("email")
pais = rs("pais")
end if
set rs = nothing
%>

<h2 align="center">Actualizar datos universidad</h2>
 <form name="form1" method="post" action="actualiza_universidad.asp">
<table width="75%" border="0" align="center" cellpadding="2" cellspacing="2">
  <tr>
    <th width="29%" nowrap class="celdaBgBlueLeft">Nombre de la Universidad</th>
    <td width="71%" class="celdaBgBlue05">
     
        <div align="left">
          <input type="text" name="nombre" value="<%=nombre%>">
  
      </div></td>
  </tr>
  <tr>
    <th width="29%" nowrap class="celdaBgBlueLeft">Pais</th>
    <td width="71%" class="celdaBgBlue05">
      <div align="left">
        <select name="pais" size="1" >
        <option value ="0" selected>--Selecciona una Opcion--</option>
        <%
    
    sql = "select id,nombre from pais order by nombre"
    set rs = db.sqlselect("DBBAECYS",sql)
    while rs.eof <> true
    if rs("id") = pais then
    %>
        <option value ="<%=rs("id")%>" selected><%=rs("nombre")%></option>
        <%
    else
    %>
        
    <option value ="<%=rs("id")%>" ><%=rs("nombre")%></option>
        <%
    end if
    rs.movenext
    wend
    set rs = nothing
    %>
        </select>
      </div></td>
  </tr>
  <tr>
    <th width="29%" nowrap class="celdaBgBlueLeft">Direccion</th>
      <td width="71%" class="celdaBgBlue05">
        <div align="left">
          <textarea name="direccion" cols="33" rows="3"><%=direccion%></textarea>
        </div></td>
  </tr>
  <tr>
    <th width="29%" nowrap class="celdaBgBlueLeft">Telefono</th>
    <td width="71%" class="celdaBgBlue05"><div align="left">
      <input name="telefono" type="text" value="<%=telefono%>">
    </div></td>
  </tr>
  <tr>
    <th width="29%" nowrap class="celdaBgBlueLeft">Sitio en Internet</th>
      <td width="71%" class="celdaBgBlue05"><div align="left">
        <input name="sitio" type="text" value="<%=sitio%>">
      </div></td>
  </tr>
  <tr>
    <th width="29%" nowrap class="celdaBgBlueLeft">Correo Electronico</th>
      <td width="71%" class="celdaBgBlue05"><div align="left">
        <input type="text" name="email" value="<%=email%>" >
      </div></td>
  </tr>
</table>
<table align="center">
  <tbody>
  <tr>
    <td><input name="Aceptar" type="submit" class="celdaBgBlue05" value="Aceptar"></td>
    <td><input type="hidden" name="id" value="<%=id%>"></td>
	    <td><input name="Cancelar" type="submit" class="celdaBgBlue05" onclick="return Regresar()" value="Cancelar"></td>
  </tr>
  </tbody>
</table>
</form>
</BODY>
</HTML>
