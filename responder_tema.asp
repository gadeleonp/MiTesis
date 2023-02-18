<%@ Language=VBScript %>
<%
if isempty(session("id")) then
Response.Redirect("login.asp")
end if
%>

<%
dim db
dim rs
dim sql
set db = server.CreateObject("SqlComandos.comandos")

 
dim foro
dim respuesta
respuesta = Request.QueryString("respuesta")
tipo = Request.QueryString("tipor")
foro = Request.QueryString("foro")
accion = request.querystring("accion")

if not isempty(respuesta) and tipo = "1" then
else
    archivo = request.querystring("archivo")
	agregar = request.querystring("descripcion")
	enlace = request.querystring("enlace")

	if not isempty(archivo) then
		archivo = agregar & " <a href=" &  """"& "foros/"&archivo & """"& ">"&enlace&"</a>"
	end if

		sql = "select id,foro,titulo,autor,descripcion from respuesta where foro = "& foro &" and id = " & respuesta
'		response.write respuesta & "<-r foro-> " & foro & "tipo= " & tipo & "archivo =" & archivo & "agregar = " & agregar &  "enlace = " & enlace
'		response.write "<br>" & sql
'		response.end
		set rs = db.SQLSelect("DBBAECYS",sql)
		if not rs.EOF then
			titulo = rs("titulo")
			descripcion = rs("descripcion")
		end if
			set rs = nothing

	
end if

%>

<html>
<head>
<link href="./theme/Master.css" rel="stylesheet" type="text/css">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</head>
<script language="JavaScript">
function Cambiar() {
 window.location = 'opciones_datos.asp'
}
function Previsualizar(forma) {
desc = forma.descripcion.value
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
 
 if (forma.descripcion.value.length > 400) {
 alert("Excedio el tamaño de la descripcion.")
 forma.descripcion.focus
 return false
 }
 if (forma.descripcion.value.length < 1) {
 alert ("No ingreso descripcion del foro")
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
if (forma.descripcion.value.length > 400) {
  alert("Excedio el tamaño del campo")
  forma.descripcion.focus
  forma.descripcion.disabled
 }
   re = "\r\n";
  
 if (Search (re,forma.descripcion.value)!= "-1" )
  {
  
	forma.descripcion.value = ReplaceDemo(forma.descripcion.value) 
  }
  
  return false
}

function Edicion(forma,valor) {
	texto = forma.descripcion.value 	
	if (valor == "1" ) { //Negrilla
	forma.descripcion.value =  forma.descripcion.value + "<b></b>"
	}
	if (valor == "2") { //SubRayar
	forma.descripcion.value = forma.descripcion.value + "<u></u>"
    }
	if (valor == "3") { //Italica
	forma.descripcion.value = forma.descripcion.value + "<em></em>"
    }
	if (valor == "4") { //Italica
	forma.descripcion.value = forma.descripcion.value + "<hr>"
    }
	if (valor == "5") { //Izquierda
	forma.descripcion.value = forma.descripcion.value + "<div align=left></div>"
    }
	if (valor == "6") { //Centro
	forma.descripcion.value = forma.descripcion.value + "<div align=center></div>"
    }
	if (valor == "7") { //Italica
	forma.descripcion.value = forma.descripcion.value + "<div align=derecha></div>"
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
			forma.descripcion.value = forma.descripcion.value + Texto	
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
			forma.descripcion.value = forma.descripcion.value + Texto				
			}
		}

    }
	
	forma.descripcion.focus = true
	forma.descripcion.select    

return true
}

</script>
<BODY onload="javascript:{if(parent.frames[0]&&parent.frames['opciones'].Go)parent.frames['opciones'].Go()}" >
<form name="form" target="_self" action="ingreso_respuesta.asp" method="post" onsubmit="return Validar(this)">
<div align="center">
  <table width="75" border="0" cellpadding="2" cellspacing="2">
    <tr>
      <td>
	  	  <table width="420" height="14" border="0" align="center">
		  <tr>
		    <td width="35" height="15" class="celdaBgBlueLeft">Titulo</td>
		    <td width="369">
				<input name="titulo" type="text" class="celdaBgBlue05" value ="<%=titulo%>" size="50" maxlength="50">
			</td>
		  </tr>
		 </table> 
	  </td>
    </tr>
    <tr>
        <td> <img src="imagenes/icon_editor_bold.gif" alt="Negrilla" onclick="return Edicion(form,'1')" WIDTH="23" HEIGHT="22"> 
          <img src="imagenes/icon_editor_underline.gif" alt="Subrayar" onclick="return Edicion(form,'2')" WIDTH="23" HEIGHT="22"> 
          <img src="imagenes/icon_editor_italicize.gif" alt="Italica" onclick="return Edicion(form,'3')" WIDTH="23" HEIGHT="22"> 
          <img src="imagenes/icon_editor_hr.gif" alt="Linea" onclick="return Edicion(form,'4')" WIDTH="23" HEIGHT="22"> 
          <img src="imagenes/icon_editor_left.gif" alt="Izquierda" onclick="return Edicion(form,'5')" WIDTH="23" HEIGHT="22"> 
          <img src="imagenes/icon_editor_center.gif" alt="Centrar" onclick="return Edicion(form,'6')" WIDTH="23" HEIGHT="22"> 
          <img src="imagenes/icon_editor_right.gif" alt="Derecha" onclick="return Edicion(form,'7')" WIDTH="23" HEIGHT="22"> 
          <img src="imagenes/icon_editor_email.gif" alt="Correo" onclick="return Edicion(form,'8')" WIDTH="23" HEIGHT="22"> 
          <img src="imagenes/icon_editor_url.gif" alt="WebSite" onclick="return Edicion(form,'9')" WIDTH="23" HEIGHT="22"> 
    </tr>
    <tr>
      <td>
	    <table width="281">
  <tr>
    
    <td>
		<textarea name="descripcion" onkeypress ="Verificar(form)" cols="40" rows="10" wrap="virtual"><%=descripcion%></textarea>
	</td>
  </tr>
  </table>
	    <p class="celdaBgBlue05">&nbsp;</p></td>
    </tr>
  </table>
<br>
<script language="javascript">
function VerForos(opcion) {
if (opcion == '1') {
  pagina = "reporte_foros.asp"
  }
  if (opcion == '2') {
  pagina = "foro.asp"
  }
  document.location = pagina
}

</script>
<input name="foro" value="<%=foro%>" type="hidden">
<input name="Accion" type="hidden"  value="<%=accion%>">
<input name="respuesta" type="hidden"  value="<%=respuesta%>">
<table width="75" border="0">
  <tr>
    <td><input name="Aceptar" type="submit" class="celdaBgBlue05" value="Aceptar"></td>
    <td><input name="Cancelar" type="button" class="celdaBgBlue05" value="Cancelar" onclick="javascript:Cambiar()"></td>
    <td><input name="Previ" type="button" class="celdaBgBlue05" onclick="javascript:Previsualizar(form)" value="Pre-visualizar"></td>

  </tr>
</table>

  </div>
</form>
</body>
</html>
