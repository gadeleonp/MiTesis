<%@ Language=VBScript %>
<%
if isempty (session("administrador")) then
 Response.Redirect("Login.asp")
end if
%>
<%
dim noticia,folder
noticia = Request.Form("noticia")
folder = Request.Form("folder")

if (noticia =null or len(noticia) = 0) then
 noticia = Request.QueryString ("noticia")

end if

if (folder =null or len(folder) = 0) then
 folder = Request.QueryString ("folder")

end if

'Response.Write folder +" - " +noticia


%>
<link href="../theme/Master.css" rel="stylesheet" type="text/css">
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
	if ( !NoVacio(Form.txtfile) && Form.action == "ingresar_documento.asp" )
	 {
		alert("Debe Seleccionar un archivo de noticia.")
		Form.txtfile.focus()
		return false
	}  
	return true
}


</script>
<html>
<body>

<h1 align="center">Ingresar Nuevo Documento a la Notica</h1>
<form onsubmit="return ValidaForma(this)"name="ingdocumento" enctype="multipart/form-data"   method="post" action="ingresar_documento.asp">
<input name="noticia" type = "hidden" value="<%= noticia%>">
<input name="folder" type = "hidden" value="<%= folder%>">

  <table width="55%" border="0" align="center" cellpadding="2" cellspacing="2">
  
    <tr>
    <th nowrap class="celdaBgBlueLeft"> 
            <div align="right">Archivo de la Notica:</div>   </th>
      <td> 
        <div align="left" class="celdaBgBlue05" > 
                  <input name="txtfile" type="file" >
        </div>
      </td>
    </tr>
    <!--
     <tr> 
      <td> 
        <div align="right">Archivo de Imagen Principal:</div>
      </td>
      <td> 
        <div align="left"> <input NAME="txtimagen1" TYPE="file" ></div>
      </td>
    </tr>
	<tr> 
      <td> 
        <div align="right">Archivo de Imagen Secundaria:</div>
      </td>
      <td> 
        <div align="left"> <input NAME="txtimagen2a" TYPE="file" ></div>
      </td>
    </tr>
    -->
  </table>
  <br>
<table width="14%" border="0" align="center">
  <tr>
    <td width="41%"> 
      <input name="Ingresar" type="submit" class="celdaBgBlue05" value="Ingresar">
    </td>
	<td width="17"></td>
    <td width="59%"> 
      <input name="Cancelar" type="submit" class="celdaBgBlue05" onClick="return Cambiar(ingdocumento)" value="Cancelar">
    </td>
  </tr>
</table>
</form>
</body>      
</html>
      