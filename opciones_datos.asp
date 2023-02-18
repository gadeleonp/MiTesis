<%@ Language=VBScript %>
<%
if isempty(session("id")) then
    Response.Redirect("login.asp")
end if
%>
<HTML>
<HEAD>
<link href="./styles/interior.css" rel="stylesheet" type="text/css">
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<script language="javascript">
function NuevoForo() 
{
var destino;
var str;
var buscado;

 var s = navigator.appVersion;
  ss = s.substr(22,1);
  
  if (ss.indexOf("6")==0)            
  {
	destino = "nuevoforo2.asp"
  }
  else 
  {
	destino = "nuevoforo.asp"
  }
	document.forma.action=destino;
	document.forma.submit();
}
 function Curriculum() 
	{
	  ancho = 700
	  if (screen) {
	  
	  leftPos = (screen.width / 2) - (ancho/2)
	  }
	  catWindow = window.open('ingreso_curriculum.asp','Curriculum','toolbar=no,menubar=no,location=no,scrollbars=yes,width='+ancho+'+,height=350,left='+leftPos+'+,top=20')
	}	
	function Desactivar() 
	{
	window.location = "./eliminar_cuenta.asp";
	}
</script>
<BODY onload="javascript:{if(parent.frames[0]&&parent.frames['opciones'].Go)parent.frames['opciones'].Go()}">
<div align="center">
<table >
  <tr>
	<tH class="celdaBgBlue">REGISTRO</tH>
  </tr>	

  <tr>
	<td class="celdaBgBlue05"><a href="actualizar_registro.asp" >Actualizar Datos Personales</a></td>
  </tr>
  <tr>
    <td class="celdaBgBlue05"><a href="actualizar_password.asp" >Cambiar Contrase&ntilde;a </a></td>
  </tr>	
  <tr>
	<td class="celdaBgBlue05"><a href="ingreso_conocimiento.asp" >Actualizar Preferencias</a></td>
  </tr>	
  <tr>
	<td class="celdaBgBlue05"><a href="javascript:Curriculum()" >Actualizar Curriculum</a></td>
  </tr>
    <tr>
	<td class="celdaBgBlue05"><a href="eliminar_cuenta.asp" >Desactivar Cuenta</a></td>
  </tr>
    <tr>
	<tH class="celdaBgBlue">FORO</tH>
  </tr>	

  <tr>
	  <td class="celdaBgBlue05"><a href="javascript:NuevoForo();" >Nuevo Foro</a></td>
  </tr>
    <tr>
	  <td class="celdaBgBlue05"><a href="foro.asp" >Mis Foros</a></td>
  </tr>
  <tr>
	  <td class="celdaBgBlue05"><a href="foro_general.asp" >Todos los Foros</a></td>
  </tr>
  <tr>
	<td class="celdaBgBlue05"><a href="salir.asp" >Terminar Sesion</a></td>
  </tr>
</table>
</div>
<form name="forma" action="" method="post"></form>
</BODY>
</HTML>
