<%@ Language=VBScript %>
<%
if isempty (session("administrador")) then
 Response.Redirect("Login.asp")
end if
dim db
dim noticia,documento
noticia = Request.querystring ("noticia")
documento = Request.QueryString ("documento")
dim rs
set db = server.CreateObject ("SQLComandos.Comandos")
sql = "select id,enlace from documento_noticia where id = " & documento & " and noticia = " & noticia

set rs = db.SQLSelect("DBBAECYS",sql)

if not rs.EOF  then
  enlace = rs("enlace")
 end if

%>
<html>
<head >
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
	if ( !NoVacio(Form.txtfile) && Form.action == "actualizadatos_documento.asp" )
	 {
		alert("Debe Seleccionar un archivo.")
		Form.txtfile.focus()
		return false
	}  
	return true
}


</script>
</head>
<body>

<h1 align="center">Ingresar Nuevo Documento a la Notica</h1>
<form onsubmit="return ValidaForma(this)"name="ingdocumento" enctype="multipart/form-data"   method="post" action="actualizadatos_documento.asp">
<input name="noticia" type = "hidden" value="<%= noticia%>">
<input name="documento" type = "hidden" value="<%= documento%>">

  <table align="center"  border="0" cellpadding="2" cellspacing="2">
  
    <tr>
		<th nowrap class="celdaBgBlueLeft">
      <div align="right">Documento: </div>      </th>
    <td> 
    <table border = "0" cellpadding="2" cellspacing="2">
    <tr>
    <td class="celdaBgBlue05">Actual:</td><td class="celdaBgBlue10"><input name="enlace" type ="text" value="<%=enlace%>" readonly></td>
    </tr>
		<tr>
          <td class="celdaBgBlue05">
			 <div align="right">Cambiar:</div>
		  </td>
		   <td class="celdaBgBlue10"> 
				<div align="left"> 
                  <input name="txtfile" type="file" >
				</div>
		   </td>
        <tr>
        
    </table>
   </td>
      
    </tr>
     <tr> 
      <th nowrap class="celdaBgBlueLeft"> 
       <div align="right">Archivo de Imagen Principal:</div>      </th>
      <td> 
        <div align="left"> <input NAME="txtimagen1" TYPE="file" ></div>
      </td>
    </tr>
	<tr> 
      <th nowrap class="celdaBgBlueLeft"> 
      <div align="right">Archivo de Imagen Secundaria:</div>      </th>
      <td> 
        <div align="left"> <input NAME="txtimagen2a" TYPE="file" ></div>
      </td>
    </tr>
  </table>
  <br>
<table width="14%" border="0" align="center">
  <tr>
    <td width="41%"> 
      <input name="Ingresar" type="submit" class="celdaBgBlue05" value="Ingresar">
    </td>
    <td width="59%"> 
      <input name="Cancelar" type="submit" class="celdaBgBlue05" onClick="return Cambiar(ingdocumento)" value="Cancelar">
    </td>
  </tr>
</table>
</form>
</body>      
</html>
      