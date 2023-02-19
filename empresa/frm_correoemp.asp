<%@ Language=VBScript %>
<%
dim user
user = session("UsEmpresa")
if isempty(user) then
  Response.Redirect("../Login.asp")
end if

dim mail

mail = "" 
mail =  request.form("mail")
if isempty(mail) or isnull(mail) then
  mail = session("UsEmpresa")
  error = 1
end if
%>
<HTML>
<head>
<title>Correo Electrónico</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../theme/Master.css" rel="stylesheet" type="text/css">
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
				forma.contenido.value = forma.contenido.value + Texto				
			}
		}

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
				forma.contenido.value = forma.contenido.value + Texto
			}
		}

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
<body onmousedown2="return ClickDerecho()" onload="javascript:{if(parent.frames[0]&&parent.frames['opciones'].Go)parent.frames['opciones'].Go()}">
<div align="center">
<h3>Correo Electrónico</h3>
<form name="form" action="enviar_correo2.asp" enctype="multipart/form-data" method="post" >
    <table  align="center" width="0%" border="0" cellspacing="2" cellpadding="0">
      <tr> 
        <th nowrap class="celdaBgBlueLeft">CC:</th>
		<td width="5"></td>
        <td class="celdaBgBlue05"><input name="copia" type="text" style="WIDTH: 20em" tabIndex=5 size=20  ></td>
      </tr>
      <tr> 
        <th nowrap class="celdaBgBlueLeft">Tema: </th>
		<td width="5"></td>
        <td class="celdaBgBlue05"><input name="tema" type="text"style="WIDTH: 40em" tabIndex=6 size=40   ></td>
      </tr>
      <tr> 
        <th nowrap class="celdaBgBlueLeft"></th>
  	    <td width="5"></td>
        <td class="celdaBgBlue05"><div align="left"><img src="../imagenes/icon_editor_bold.gif" alt="Negrilla" onclick="return Edicion(form,'1')" WIDTH="23" HEIGHT="22"> 
            <img src="../imagenes/icon_editor_underline.gif" alt="Subrayar" onclick="return Edicion(form,'2')" WIDTH="23" HEIGHT="22"> 
            <img src="../imagenes/icon_editor_italicize.gif" alt="Italica" onclick="return Edicion(form,'3')" WIDTH="23" HEIGHT="22"> 
            <img src="../imagenes/icon_editor_hr.gif" alt="Linea" onclick="return Edicion(form,'4')" WIDTH="23" HEIGHT="22"> 
            <img src="../imagenes/icon_editor_left.gif" alt="Izquierda" onclick="return Edicion(form,'5')" WIDTH="23" HEIGHT="22"> 
            <img src="../imagenes/icon_editor_center.gif" alt="Centrar" onclick="return Edicion(form,'6')" WIDTH="23" HEIGHT="22"> 
            <img src="../imagenes/icon_editor_right.gif" alt="Derecha" onclick="return Edicion(form,'7')" WIDTH="23" HEIGHT="22"> 
            <img src="../imagenes/icon_editor_email.gif" alt="Correo" onclick="return Edicion(form,'8')" WIDTH="23" HEIGHT="22"> 
            <img src="../imagenes/icon_editor_url.gif" alt="WebSite" onclick="return Edicion(form,'9')" WIDTH="23" HEIGHT="22"> 
        </div></td>
      </tr>
      <tr> 
        <th nowrap class="celdaBgBlueLeft">Archivo</th>
		<td width="5"></td>
        <td class="celdaBgBlue05"><div align="left">
          <input name="txtarchivo" type="file"  tabindex=7 alt="Archivo...">
        </div></td>
      <tr> 
      <tr> 
        <th nowrap class="celdaBgBlueLeft">&nbsp;&nbsp;Imagen de fondo&nbsp;&nbsp;</th>
		<td width="5"></td>
        <td class="celdaBgBlue05"><div align="left">
          <input name="txtImage" type="file"  tabindex=7 id="Imagen">
        </div></td>
      <tr> 
      <tr> 
        <th nowrap class="celdaBgBlueLeft"></th>
		<td width="5"></td>
        <td class="celdaBgBlue05"><textarea name="contenido" cols="50" rows="4"  tabindex=8 style="WIDTH: 40em" onkeypress ="Verificar(form)" ></textarea></td>
      </tr>
      <tr> 
        <th nowrap class="celdaBgBlueLeft"></th>
		<td width="5"></td>
        <td class="celdaBgBlue05"> <input name="enviar" type="submit" class="celdaBgBlue10" style="WIDTH:5em" tabindex=9 value="Enviar" > 
          &nbsp; <input name="cancelar" type="button" class="celdaBgBlue10" style="WIDTH: 5em" tabIndex=10 onclick="javascript:document.location='opciones_empresa.asp'"  value ="Cancelar" size=40> 
          &nbsp; <input name="Previ" type="button" class="celdaBgBlue10" style="WIDTH: 5em"  tabindex=11 onclick="javascript:Previsualizar(form)" value="Pre-visualizar"> 
        </td>
      </tr>
    </table>
<input name="autor" type="hidden" value="<%=autor%>">
<input name="opcion" type="hidden" value="<%=opcion%>">
<input name="lstus" type="hidden" value ="<%=mail%>">
</form>
</div>
</body>
</HTML>
