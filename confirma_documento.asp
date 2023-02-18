<%@ Language=VBScript %>
<%
dim db
dim rs 
dim noticia
dim enlace
dim folder

set db = server.CreateObject("SQLComandos.Comandos")

	noticia = Request.QueryString("noticia")
	enlace = Request.QueryString("enlace")
		Sql = "select folder_contenedor from noticia where id = " & noticia
		

		set rs = db.SQLSelect("DBBAECYS",sql)
		
		if rs.eof = false then
			folder = rs("folder_contenedor")
		else
		Response.Write "Error en la consulta" +Sql
		Response.End 
		end if
		
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<script language="javascript">
	function Menu(Form,opcion) {
	if (opcion == '1') {
	//regresar al menu
	Form.action = "mantnoticias.asp"
	}
	if (opcion == '2') {
	//actualizar noticia
	Form.action = "actualizar_noticia.asp"
	}
	if (opcion == '3') {
	//actualizar noticia
	Form.action = "agregar_documento.asp"
	}
	if (opcion == '4') {
	//actualizar noticia
	Form.action = "eliminar_documento.asp"
	}

   return true
}
	</script>
<BODY>
<div align="center">
<%
	sql = "select id,enlace,correlativo from documento_noticia where noticia = " & noticia &" and enlace <> '"& enlace & "'"
	set rs = db.SQLSelect("DBBAECYS",sql)
	if rs.eof = true then
	else
	
	dim rs2 
	
	
	
	sql = "select titulo,fecha_publicacion,autor,resumen,folder_contenedor from noticia where id = " & noticia
	Response.Write sql
	
	set rs2 = db.SQLSelect("DBBAECYS",sql)
	if rs2.eof = true then
	else
	
	titulo = rs2("titulo")
	autor = rs2("autor")
	resumen = rs2("resumen")
	fecha = rs2("fecha_publicacion")
%>
  <h2>DATOS DE LA NOTICIA</h2>
  <table width="43%" border="0" cellpadding="0" cellspacing="0">
    <tr> 
      <th width="14%" nowrap> <div align="right">TITULO</div></th>
      <td width="86%"><%=titulo%></td>
    </tr>
    <tr> 
      <th width="14%" nowrap> <div align="right">AUTOR</div></th>
      <td width="86%"><%=autor%></td>
    </tr>
    <tr> 
      <th width="14%" nowrap> <div align="right">RESUMEN</div></th>
      <td width="86%"> <textarea name="resumen" readonly rows="4" cols="50"><%=resumen%></textarea> 
      </td>
    </tr>
    <tr> 
      <th width="14%"> <div align="right">Fecha</div></th>
      <td width="86%"><%=fecha%></td>
    </tr>
  </table>

<%
set rs2 = nothing
end if%>
<h2>DOCUMENTOS DE LA NOTICIA</h2>
  <table width="44%" border="0" cellpadding="2" cellspacing="2">
    <tr> 
      <th nowrap> 
        <div align="center">NOMBRE DOCUMENTO</div>
      </th>
        <th nowrap> 
        <div align="center">NUMERO CORRELATIVO</div>
      </th>
    
<%	
	while rs.eof <> true 
	
	
	documento = rs("enlace")
	correlativo = rs("correlativo")		
	rs.movenext 
%>

    <tr> 
      <td>
        <div align="center"><%=documento%> </div>
      </td>
        <td>
        <div align="center"><%=correlativo%> </div>
      </td>
    
    </tr>

<% 
 wend
 %>
   </table>
 <%
end if
  rs.close
%>
  <h2>DOCUMENTO AGREGADO CON EXITO</h2>

<TABLE border ="0" cellpadding="2" cellspacing="2">
	<TR>
		<TD>DOCUMENTO: </TD>	
		<TD><%=enlace%></TD>
	</TR> 
	
</TABLE>
<form name="resumen" method="post" action="mantnoticias.asp">
				<input name="noticia" type = "hidden" value="<%=noticia%>">
				<input name="folder" type = "hidden" value="<%=folder%>">
			<table>
				<tr>
					<td><input type="submit" name="Aceptar" value="Aceptar" onclick = "return Menu(resumen,'1')"></td>
					<td><input type="submit" name="Actualizar" value="Actualizar" onclick = "return Menu(resumen,'2')"></td>
					<td><input type="submit" name="Cancelar" value="Cancelar" onclick = "return Menu(resumen,'4')"></td>					
					<td><input type="submit" name="Agregar" value="Agregar Documento" onclick = "return Menu(resumen,'3')"></td>
				</tr>
			</table>
		</form>
</div>		
</BODY>
</HTML>
