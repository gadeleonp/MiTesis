<%@ Language=VBScript %>
<%
if isempty (session("administrador")) then
 Response.Redirect("Login.asp")
end if
%>
<HTML>

<HEAD>
<title>Ingresar Notica</title>
<link href="../theme/Master.css" rel="stylesheet" type="text/css">
<BODY>
<script language="javascript">
 function Cambiar(form) {
   form.action = "mantnoticias.asp"
   return true
 }
 
function Longitud(id,Texto,longitud) {
	if (Texto.value.length >  longitud)
	{
	alert("La Longitud del "+ id +" no puede ser mayor de " +longitud+ " caracteres")
	return false
	}
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
if (Form.action == "ingresar_noticia2.asp" ) {
  if (!NoVacio(Form.txttitulo)) {
  alert("Debe Ingresar un Titulo\n")
  Form.txttitulo.focus()
  return false
  }

  if (!NoVacio(Form.txtdescripcion) ){
  alert("Debe Ingresar una descripcion.\n")
  Form.txtdescripcion.focus()
  return false
  }
  
 if (!NoVacio(Form.txtfile)) {
  alert("Debe Seleccionar un archivo de noticia.")
  Form.txtfile.focus()
  return false
 }  
  
  if (!Longitud('Titulo',Form.txttitulo,50)) {
  Form.txttitulo.focus()
  return false
  }
  
  if (!Longitud('Descripcion',Form.txtdescripcion,400)) {
  Form.txtdescripcion.select()
  Form.txtdescripcion.focus()
  return false
  }
 

 }
 return true
}



 
</script>
<h1 align="center">Ingresar Noticia</h1>
<form onsubmit="return ValidaForma(this)"name="ingnoticia" enctype="multipart/form-data"   method="post" action="ingresar_noticia2.asp">
  <table align="center" width="55%" border="0" cellpadding="2" cellspacing="2">
    <tr> 
      <th nowrap class="celdaBgBlueLeft"> 
      <div align="right" class="celdaBgBlueLeft">
        <div align="center">Titulo:</div>
      </div>      </th>
      <td class="celdaBgBlue05"> 
        <div align="left"> 
          <input name="txttitulo" size="60" maxlength="200" >
        </div>
      </td>
    </tr>
    <tr> 
      <th nowrap class="celdaBgBlueLeft"> 
      <div align="right" class="celdaBgBlueLeft">
        <div align="center">Descripcion:</div>
      </div>      </th>
      <td class="celdaBgBlue05"> 
        <div align="left"><TEXTAREA cols=50 name="txtdescripcion" rows=4></TEXTAREA>
        </div>
      </td>
    </tr>
        <tr> 
      <th nowrap class="celdaBgBlueLeft"> 
        <div align="center">Autor:</div></th>
      <td class="celdaBgBlue05"> 
        <div align="left"> 
          <input name="txtAutor" size="30" maxlength="100" >
        </div>
      </td>
    </tr>
    <tr> 
      <th nowrap class="celdaBgBlueLeft"> 
        <div align="center">Archivo de la Notica:</div></th>
      <td class="celdaBgBlue05"> 
        <div align="left"> 
                  <input name="txtfile" type="file" >
        </div>
      </td>
    </tr>
     <tr> 
      <th nowrap class="celdaBgBlueLeft"> 
        <div align="center">Imagen Principal:</div></th>
      <td class="celdaBgBlue05"> 
        <div align="left"> <input NAME="txtimagen1" TYPE="file" ></div>
      </td>
    </tr>
	<tr> 
      <th nowrap class="celdaBgBlueLeft"> 
        <div align="center">Imagen Secundaria:</div></th>
      <td class="celdaBgBlue05"> 
        <div align="left"> <input NAME="txtimagen2a" TYPE="file" ></div>
      </td>
    </tr>
  </table>
  <br>
<table width="14%" border="0" align="center">
  <tr>
    <td width="41%"> 
      <input name="Ingresar" type="submit" class="button" value="Ingresar">
    </td>
    <td width="59%"> 
      <input type="submit" name="Cancelar" value="Cancelar" onClick="return Cambiar(ingnoticia)">
    </td>
  </tr>
</table>
</form>

</BODY>
</HTML>
