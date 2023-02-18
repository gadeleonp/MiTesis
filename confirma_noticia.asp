<%@ Language=VBScript %>

<%

	Dim Fecha 
	dim rs2,rs
	dim db
	
	set db = server.CreateObject("SQLComandos.Comandos")
	
	Fecha = Request.Form("fecha")
	Sql = "select id,titulo,resumen, autor, fecha_publicacion, folder_contenedor from noticia where fecha_publicacion = '" & Fecha & "'"
	
	set rs = db.SQLSelect("DBBAECYS",sql)
	if  rs.EOF = true  then
	else
	    Titulo = rs("titulo")
		Descripcion = rs("resumen")
		autor = rs("autor")
		noticia = rs("id")
		folder = rs("folder_contenedor")
	end if
		set rs = nothing
	
	Sql = "select enlace,imagen_principal,imagen_secundaria from documento_noticia where noticia = " & noticia
	set rs2 = db.SQLSelect("DBBAECYS",sql)
	if  not rs2.eof then
	    arc_noticia = rs2("enlace")		
		arc_img1 = rs2("imagen_principal")
		arc_img2 = rs2("imagen_secundaria")
		set rs2 = nothing
	end if

%>
<html>
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

   return true
}
	</script>
	<body>
	<h3 align="center">Notica Ingresada Correctamente</h3>
	<div align="center">
	
  <table width="34%" border="0" cellpadding="2" cellspacing="2">
    <tr> 
      <th width="36%" nowrap>Titulo de la Noticia</th>
      <td width="64%"><%=titulo%></td>
    </tr>
    <tr> 
      <th width="36%" nowrap>Archivo Initical</th>
      <td width="64%"><%=arc_noticia%></td>
    </tr>
    <tr> 
      <th width="36%" nowrap>Autor: </th>
      <td width="64%"><%=Autor%></td>
    </tr>
    <tr> 
      <th width="36%" nowrap>Fecha de Creacion: </th>
      <td width="64%"><%=Fecha%></td>
    </tr>
    <tr> 
      <th width="36%" nowrap>Imagen Principal: </th>
      <td width="64%"> 
        <%if len(arc_img1) = 0 then response.write("Imagen no ingresada.") else response.write(arc_img1) end if 
		  %>
      </td>
    </tr>
    <tr> 
      <th width="36%">Imagen Secundaria: </th>
      <td width="64%"> 
        <%if len(arc_img2) = 0 then response.write("Imagen no ingresada.") else response.write(arc_img2) end if 
			%>
      </td>
    </tr>
  </table>
			<form name="resumen" method="post" action="mantnoticias.asp">
				<input name="noticia" type = "hidden" value="<%=noticia%>">
				<input name="folder" type = "hidden" value="<%=folder%>">
			<table>
				<tr>
					<td><input type="submit" name="Aceptar" value="Aceptar" onclick = "return Menu(resumen,'1')"></td>
					<td><input type="submit" name="Actualizar" value="Actualizar" onclick = "return Menu(resumen,'2')"></td>
					<td><input type="submit" name="Agregar" value="Agregar Documento" onclick = "return Menu(resumen,'3')"></td>
				</tr>
			</table>
		</form>
</div>
</body>
</html>