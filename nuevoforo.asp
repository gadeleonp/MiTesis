<%@ Language=VBScript %>

<%
dim usuario
usuario = ""
usuario = session("id")

if isempty(usuario) then
	Response.Redirect("login.asp")
end if
%>
<html>
<head>
<link href="./theme/Master.css" rel="stylesheet" type="text/css">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</head>
<script language="JavaScript">
function Previsualizar(forma) {
desc = forma.descripcion.value
	  ancho = 700
	  if (screen) {
	  
	  leftPos = (screen.width / 2) - (ancho/2)
	  }
	//PreviewWindow = window.open('previewforo.asp?Noticia='+Noticia,'Preview','toolbar=no,menubar=no,location=no,scrollbars=yes,width='+ancho+'+,height=350,left='+leftPos+'+,top=20')
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
 
 if (forma.descripcion.value.length > 500) {
 alert("Excedio el tamaño de la descripcion.")
 forma.descripcion.focus
 return false
 }
 if (forma.descripcion.value.length < 1) {
 alert ("No ingreso descripcion del foro")
 forma.descripcion.focus
 return false
 }

 if (forma.Categoria.value == "0" ) {
 alert ("No ha seleccionado una categoria para el foro")
 form.Categoria.select
 forma.Categoria.focus
 
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
<%
dim db
dim rs 
set db = server.CreateObject ("SQLComandos.Comandos")
		
dim foro
dim descripcion
dim titulo
dim categoria
dim accion

foro = ""
accion = ""
descripcion = ""
titulo = ""
categoria= ""


foro = Request.querystring("foro")
accion = request.querystring("accion") 
dim rs2
dim rs3

%>
<BODY onload="javascript:{if(parent.frames[0]&&parent.frames['opciones'].Go)parent.frames['opciones'].Go()}" >
<form name="form" target="_self" action="<%
if isempty(accion) then
response.write "ingreso_foro.asp"
else
response.write "actualizar_foro.asp"
end if
%>" method="post" onsubmit="return Validar(this)">

<%
	if isempty(foro) then

	else
		sql = "select id,titulo,autor,descripcion,fecha_creacion,categoria from foro where id = "& foro &" and autor = " & usuario
		set rs = db.SQLSelect("DBBAECYS",sql)
		if not rs.EOF then
			categoria = rs("categoria")
		end if
		set rs = nothing
	end if
	if isempty(foro) then

	else
		sql = "select id,titulo,autor,descripcion,fecha_creacion,categoria from foro where id = "& foro &" and autor = " & usuario
		set rs = db.SQLSelect("DBBAECYS",sql)
		if not rs.EOF then
			titulo = rs("titulo")
			descripcion = rs("descripcion")
		end if
			set rs = nothing
	end if
%>
			
		
			
			

  <div align="center">
  <table width="75" border="0" cellpadding="2" cellspacing="2">
      <tr> 
        <td>Categoria &nbsp;

	<select name="Categoria" size="1" >
        <option value="0" selected>-----Sin Clasificar-----</option>
        
        <%


sql = "select id,descripcion from categoria where padre is null"

set rs = db.SQLSelect("DBBAECYS",sql)
if not rs.EOF then
%>
      <%while rs.eof <> true %>
    
    <option value="<%=rs("id")%>" <%if rs("id") = categoria then Response.Write "selected" end if%>>-<%=rs("descripcion")%>-</option>
    
      

      <%
	sql2 = "select id,descripcion from categoria where padre = " & rs("id")
set rs2 = db.SQLSelect("DBBAECYS",sql2)
	if not rs2.EOF then
	%>
	 
		<%
	while rs2.EOF <> true
	%>
     <option value="<%=rs2("id")%>" <%if rs2("id") = categoria then Response.Write "selected" end if%> >--<%=rs2("descripcion")%>--</option>
     
     <%
	rs2.MoveNext 
	wend 
	%>
	<%
	end if
		set rs2 = nothing
	%>
      <%
rs.MoveNext 
wend
%>


<%
end if
set rs = nothing
%>

        
          
      </select>
	</td>
  </tr>
    <tr>
      <td>
	  	  <table width="420" height="14" border="0" align="center" cellpadding="2" cellspacing="2">
		  <tr>
		    <td width="35" height="15">Titulo</td>
		    <td width="369">
				<input name="titulo" size="50" type="text" maxlength="50" value ="<%=titulo%>">
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
	  </td>
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
function Cambiar() {
 window.location = 'opciones_datos.asp'
}
</script>
<table width="75" border="0">
  <tr>
  <!--   <td><input name="Foros" type="button" class="celdaBgBlue05" onclick="javascript:VerForos('1')" value="Ver Foros"></td>-->
  <!--  <td><input name="Foros" type="button" class="celdaBgBlue05" onclick="javascript:VerForos('2')" value="Foros Publicados"></td> -->
    <td><input name="Aceptar" type="submit" class="celdaBgBlue05" value="Aceptar"></td>
    <td><input name="Cancelar" type="button" class="celdaBgBlue05" onclick="javascript:Cambiar()" value="Cancelar"></td>
    <td><input name="Previ" type="button" class="celdaBgBlue05" onclick="javascript:Previsualizar(form)" value="Pre-visualizar"></td>
  </tr>
</table>

  </div>
  <input name="foro" type="hidden" value = "<%=foro%>">
</form>
</body>
</html>
<%
set db = nothing
%>