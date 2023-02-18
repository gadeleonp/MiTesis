<%@ Language=VBScript %>
<%
if isempty (session("administrador")) then
 %>
<html>
<body onload="javascript:cerrarventana()" >
<script language="javascript">
function cerrarventana() {
opener.window.location = "Login.asp"
window.close()
}
</script>
</body>
</html>
<%
end if
dim db
dim noticia
dim sql
dim rs
dim folder
dim folder_imagen

noticia = ""
noticia =Request.QueryString ("noticia")

	if len(noticia) = 0 then
		noticia = Request.Form("noticia")
	end if

	set db = server.CreateObject ("SQLComandos.Comandos")
	sql = "select folder_contenedor,folder_imagenes from noticia where id = " & Noticia
	set rs = db.SQLSelect("DBBAECYS",sql)
	if not rs.eof  then
		folder = rs("folder_contenedor")
		folder_imagen = rs("folder_imagenes")
	end if
	set rs = nothing

%>
<html>
<head>
<link href="../theme/Master.css" rel="stylesheet" type="text/css">
<title>
Agregar Imagen
</title>
</head>
<script language="javascript">
function Cambiar(Form) {
Form.action = "mantnoticias.asp"
return true
}

function NoVacio(Texto){
if (Texto.value.length == 0 ) {
	return false
}
return true
}

function ValidaForma(Form)
{
	if ( !NoVacio(Form.txtimagen1) && Form.action == "ingresa_imagen.asp" )
	 {
		alert("Debe Seleccionar un archivo.")
		
		return false
	}  
	
	if (ingdocumento.folder_img.value.length == 0) 
	{
	
		alert ("Error: No especifico un folder para almacenar las imagenes.")
		return false
	}
	
	return true
}


function cerrarventana() {
	
	   window.close()
 	}


</script>
<body>

<h2 align="center">Ingresar Nuevo Imagen a la Notica</h2>
<form name="ingdocumento" onsubmit="return ValidaForma(this)"  enctype="multipart/form-data" method="post" action="ingresa_imagen.asp">
<input name="noticia" type = "hidden" value="<%= noticia%>">
<input name="folder" type = "hidden" value="<%= folder%>">
<table align = "center" border="0" cellpadding="2" cellspacing="2">	
      <th nowrap> <div align="left">Folder de Imagenes</div></th>
      <th nowrap>&nbsp;</th>
      <td nowrap> <div align="left"> 
          <input NAME="folder_img" id="folder_img" TYPE="text" value="<%=folder_imagen%>">
        </div></td>
    </tr>
</table>	
<br>
  <table align="center" width="42%" border="0" cellpadding="2" cellspacing="2">
    <tr> 
      <th nowrap> <div align="left">Imagen: </div></th>
      <th nowrap>&nbsp;</th>
      <td> <div align="left"> 
          <input NAME="txtimagen1" id="1" TYPE="file" >
        </div></td>
    </tr>
    <tr> 
      <th nowrap> <div align="left">Imagen: </div></th>
      <th nowrap>&nbsp;</th>
      <td> <div align="left"> 
          <input NAME="txtimagen2" id="2" TYPE="file" >
        </div></td>
    </tr>
    <tr> 
      <th nowrap> <div align="left">Imagen: </div></th>
      <th nowrap>&nbsp;</th>
      <td> <div align="left"> 
          <input NAME="txtimagen3" id="3"  TYPE="file" >
        </div></td>
    </tr>
    <tr> 
      <th nowrap> <div align="left">Imagen: </div></th>
      <th nowrap>&nbsp;</th>
      <td> <div align="left"> 
          <input NAME="txtimagen4" id="4"  TYPE="file" >
        </div></td>
    </tr>
    <tr> 
      <th nowrap> <div align="left">Imagen: </div></th>
      <th nowrap>&nbsp;</th>
      <td> <div align="left"> 
          <input NAME="txtimagen5" id="5" TYPE="file" >
        </div></td>
    </tr>
  </table>
  <br>
<table width="14%" border="0" align="center">
  <tr>
    <td width="41%"> 
      <input type="submit" name="Ingresar" value="Ingresar">
    </td>
    <td width="59%"> 
      <input type="button" name="Cancelar" value="Cancelar" onClick="javascript:cerrarventana()">
    </td>
  </tr>
</table>
      </form>
</body>      
</html>
      