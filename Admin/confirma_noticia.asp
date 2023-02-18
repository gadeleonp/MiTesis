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
<head>
<link href="../theme/Master.css" rel="stylesheet" type="text/css">
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
</head>
	<body>
	<h1 align="center">Noticia Ingresada Correctamente</h1>
	<div align="center">
	
  <table width="34%" border="0" cellpadding="2" cellspacing="2">
    <tr> 
      <th width="36%" nowrap class="celdaBgBlueLeft">Titulo </th>
      <td width="64%" class="celdaBgBlue05"><%=titulo%></td>
    </tr>
    <tr> 
      <th width="36%" nowrap class="celdaBgBlueLeft">Descripción</th>
      <td width="64%" class="celdaBgBlue05"><textarea cols="25" rows="4" readonly><%=Descripcion%></textarea></td>
    </tr>

    <tr> 
      <th width="36%" nowrap class="celdaBgBlueLeft">Archivo Inicial</th>
      <td width="64%" class="celdaBgBlue05"><%=arc_noticia%></td>
    </tr>
    <tr> 
      <th width="36%" nowrap class="celdaBgBlueLeft">Autor: </th>
      <td width="64%" class="celdaBgBlue05"><%=Autor%></td>
    </tr>
    <tr> 
      <th width="36%" nowrap class="celdaBgBlueLeft">Fecha de Creacion: </th>
      <td width="64%" class="celdaBgBlue05"><%=Fecha%></td>
    </tr>
    <tr> 
      <th width="36%" nowrap class="celdaBgBlueLeft">Imagen Principal: </th>
      <td width="64%" class="celdaBgBlue05"> 
        <%if len(arc_img1) = 0 then response.write("Imagen no ingresada.") else response.write(arc_img1) end if 
		  %>      </td>
    </tr>
    <tr> 
      <th width="36%" nowrap class="celdaBgBlueLeft">Imagen Secundaria: </th>
      <td width="64%" class="celdaBgBlue05"> 
        <%if len(arc_img2) = 0 then response.write("Imagen no ingresada.") else response.write(arc_img2) end if 
			%>      </td>
    </tr>
  </table>
			<form name="resumen" method="post" action="mantnoticias.asp">
				<input name="noticia" type = "hidden" value="<%=noticia%>">
				<input name="folder" type = "hidden" value="<%=folder%>">
			<table>
				<tr>
					<td><input name="Actualizar" type="submit" class="celdaBgBlue05" onclick = "return Menu(resumen,'2')" value="Aceptar"></td>
					<!--<td><input type="submit" name="Agregar" value="Agregar Documento" onclick = "return Menu(resumen,'3')"></td>-->
				</tr>
			</table>
		</form>
</div>
</body>
</html>