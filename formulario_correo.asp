<%@ Language=VBScript %>
<%
if isempty(session("id")) then
%>
<html>
<head>
<link href="./theme/Master.css" rel="stylesheet" type="text/css">
<title>Correo Electrónico</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
	<script language="javascript">
	 function cerrarventana() {
	   opener.window.location = 'Login.asp'
	   window.close
	 }
	</script>
<body onload="javascript:cerrarventana()">
</body>
</html>
<%
end if
dim opcion
dim autor
	autor = ""
	opcion = "0"
	
	autor = Request.Form("autor")
	opcion = Request.Form("opcion")
	Response.Write opcion &" - " &autor
	

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<link href="./theme/Master.css" rel="stylesheet" type="text/css">
<title>Correo Electrónico</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<script language="JavaScript">
function Cancelar() {
   window.close()
}
function Previsualizar(forma) {
desc = forma.contenido.value
	  ancho = 700
	  if (screen) {
	  
	  leftPos = (screen.width / 2) - (ancho/2)
	  }
PreviewWindow = window.open('previewforo.asp?desc='+desc,'Preview','toolbar=no,menubar=no,location=no,scrollbars=yes,width='+ancho+'+,height=350,left='+leftPos+'+,top=20')	
}
function Validar(forma) {
 if (forma.titulo.value.length > 50) {
 alert("Excedio el tamaño del titulo")
 forma.titulo.focus
 return false
 }
 if (forma.titulo.value.length < 1) {
 alert("No ingreso el titulo")
 forma.titulo.focus
 return false
 }
 
 if (forma.contenido.value.length < 1) {
 alert ("No ingreso el contenido del correo.")
 forma.descripcion.focus
 return false
 }
}

function Search(str1,str2)
{
  var s = str2.indexOf(str1);
  return(s);
}

function ReplaceDemo(s)
{
  var r, re;
  re = /\r\n/i;
  r = s.replace(re, "<br>");
  return(r);
}

function Verificar(forma) {
var re;
   re = "\r\n";
  
 if (Search (re,forma.contenido.value)!= "-1" )
  {
  
	forma.contenido.value = ReplaceDemo(forma.contenido.value) 
  }
  
  return false
}

function Edicion(forma,valor) {
	texto = forma.contenido.value 	
	if (valor == "1" ) { //Negrilla
	forma.contenido.value =  forma.contenido.value + "<b></b>"
	}
	if (valor == "2") { //SubRayar
	forma.contenido.value = forma.contenido.value + "<u></u>"
    }
	if (valor == "3") { //Italica
	forma.contenido.value = forma.contenido.value + "<em></em>"
    }
	if (valor == "4") { //Italica
	forma.contenido.value = forma.contenido.value + "<hr>"
    }
	if (valor == "5") { //Izquierda
	forma.contenido.value = forma.contenido.value + "<div align=left></div>"
    }
	if (valor == "6") { //Centro
	forma.contenido.value = forma.contenido.value + "<div align=center></div>"
    }
	if (valor == "7") { //Italica
	forma.contenido.value = forma.contenido.value + "<div align=derecha></div>"
    }
	if (valor == "8") { //Correo
	Enlace=prompt("Ingrese el texto que se mostrara como enlace para el correo.\nDeje en blanco si desea que el correo aparezca como enlace.","");
		if (Enlace!=null) {
			WebSite = prompt("Ingrese su direccion de correo.","mailto:");
			if (WebSite!=null) {
				if (Enlace=="") {
					Texto='<a href="'+WebSite+'" target="Nueva">'+WebSite+'</a>'
					
				} else {
					Texto='<a href="'+WebSite+'" target="Nueva">'+Enlace+'</a>'
					
				}
			}
		}
	forma.contenido.value = forma.contenido.value + Texto
	}
	if (valor == "9") { //Web site
	Enlace=prompt("Ingrese el texto que se mostrara como enlace.\nDeje en blanco si desea que la pagina aparezca com enlace.","");
		if (Enlace!=null) {
			WebSite = prompt("Ingrese la direccion URL.","http://");
						
			if (WebSite!=null) {
				if (Enlace=="") {
					Texto='<a href="'+WebSite+'" target="Nueva">'+WebSite+'</a>'
					
				} else {
					Texto='<a href="'+WebSite+'" target="Nueva">'+Enlace+'</a>'
					
				}
			}
		}
	forma.contenido.value = forma.contenido.value + Texto
    }
	
	forma.contenido.focus = true
	forma.contenido.select    

return true
}

function ClickDerecho() {
	if ((event.button == 2) || (event.button == 3)) 
	{
		alert("Esta opción no esta disponible.")
	}
}
</script>
<body onmousedown2="return ClickDerecho()">
<div align="center">
<h1>Correo Electrónico</h1>
<form name="form" action="enviar_correo2.asp" enctype="multipart/form-data" method="post" >
<table  align="center" width="0%" border="0" cellspacing="2" cellpadding="2">
  <tr>
    <td width="68"></td>
	<td width="640">
	<div align="left">
	<input class="celdaBgBlue05" name="enviar" type="submit" value="Enviar"  tabindex=1 style="WIDTH: 5em" >&nbsp;
	<input class="celdaBgBlue05" name="cancelar" type="button" style="WIDTH: 5em" tabIndex=2 size=40  value ="Cancelar" onclick="javascript:Cancelar()">&nbsp;
	<input class="celdaBgBlue05" name="Previ" type="button" style="WIDTH: 5em"  tabindex=3 value="Pre-visualizar" onclick="javascript:Previsualizar(form)">&nbsp;	
	</div>
     </td>
  </tr>
  <tr>
  <%
  if opcion <> "1" then
	
  %>
  <th class="celdaBgBlueLeft">Para: </th>
  <td>
  <div align="left">
	<input name="destino" type="text" style="WIDTH: 40em" tabIndex=4 size=40  ></td>
	</div>
  </tr>
  <tr>
    <th class="celdaBgBlueLeft">CC:</th><td>
	<div align="left">
	<input name="copia" type="text" style="WIDTH: 20em" tabIndex=5 size=20  ></td>
	</div>
  </tr>
  <%
  end if
  %>
  <tr>
	<th class="celdaBgBlueLeft">Tema: </th><td>
		<div align="left">
	<input name="tema" type="text"style="WIDTH: 40em" tabIndex=6 size=40   ></td>
	</div>
  </tr>
  <tr>
  <th class="celdaBgBlueLeft">Archivo</th><td>
  <div align="left">
  <input name="txtarchivo" type="file"  tabindex=7 alt="Archivo..."></td>
  </div>
  <tr>
  <tr>
  <th class="celdaBgBlueLeft">Imagen de fondo</th><td>
  <div align="left">
  <input name="txtImage" type="file"  tabindex=7 id="Imagen">
  </div>
  </td>
    </tr>

  <tr>
    <td></td>
	<td>
		<div align="left">
	<img src="imagenes/icon_editor_bold.gif" alt="Negrilla" onclick="return Edicion(form,'1')" WIDTH="23" HEIGHT="22"> 
          <img src="imagenes/icon_editor_underline.gif" alt="Subrayar" onclick="return Edicion(form,'2')" WIDTH="23" HEIGHT="22"> 
          <img src="imagenes/icon_editor_italicize.gif" alt="Italica" onclick="return Edicion(form,'3')" WIDTH="23" HEIGHT="22"> 
          <img src="imagenes/icon_editor_hr.gif" alt="Linea" onclick="return Edicion(form,'4')" WIDTH="23" HEIGHT="22"> 
          <img src="imagenes/icon_editor_left.gif" alt="Izquierda" onclick="return Edicion(form,'5')" WIDTH="23" HEIGHT="22"> 
          <img src="imagenes/icon_editor_center.gif" alt="Centrar" onclick="return Edicion(form,'6')" WIDTH="23" HEIGHT="22"> 
          <img src="imagenes/icon_editor_right.gif" alt="Derecha" onclick="return Edicion(form,'7')" WIDTH="23" HEIGHT="22"> 
          <img src="imagenes/icon_editor_email.gif" alt="Correo" onclick="return Edicion(form,'8')" WIDTH="23" HEIGHT="22"> 
          <img src="imagenes/icon_editor_url.gif" alt="WebSite" onclick="return Edicion(form,'9')" WIDTH="23" HEIGHT="22"> 
		  </div>
		  </td>
		  
  </tr>
  <tr>
    <td></td>
		<td>
		<div align="left">
		<textarea name="contenido" cols="50" rows="4"  tabindex=8 style="WIDTH: 40em" onkeypress ="Verificar(form)" ></textarea>
		</div>
		</td>

  </tr>
  <tr>
    <td></td>
	<td align="left">
		<div align="left">
			<input class="celdaBgBlue05" name="enviar" type="submit" value="Enviar" style="WIDTH:5em" tabindex=9 >&nbsp;
			<input  class="celdaBgBlue05" name="cancelar" type="button" style="WIDTH: 5em" tabIndex=10 size=40  value ="Cancelar" onclick="javascript:Cancelar()">&nbsp;
			<input class="celdaBgBlue05" name="Previ" type="button" style="WIDTH: 5em"  tabindex=11 value="Pre-visualizar" onclick="javascript:Previsualizar(form)">
		</div>
	</td>
  </tr>
 </table>
			<input name="autor" type="hidden" value="<%=autor%>">
			<input name="opcion" type="hidden" value="<%=opcion%>">
	</form>
	</div>
</body>
</html>
