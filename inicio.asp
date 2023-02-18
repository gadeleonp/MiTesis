<%@ Language=VBScript %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>:::...  Asociaci&oacute;n de Estudiantes de Ciencias y Sistemas ...:::</title>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>
<script language="javascript">
function Destino(Opcion) 
{
if (Opcion == "1") {
destino = "./principal.asp?ver=0&path=ingreso_conocimiento.asp"
window.parent.location = destino
}
if (Opcion == "2") {
destino = "./principal.asp?ver=0&path=Login.asp?menu=1"
window.parent.location = destino
}
if (Opcion == "3") {
destino = "./principal.asp?ver=0&path=noticias.asp?opcion=1"
window.parent.location = destino
}

if (Opcion == "4") {
destino = "./admin/Login.asp"
window.parent.location  = destino
//parent.principal.location = destino
}
if (Opcion == "5") {
destino = "./principal.asp?ver=0&path=Login.asp?menu=2"
window.parent.location = destino
}
if (Opcion == "6") {
destino = "./principal.asp?ver=0&path=actualizar_registro.asp"
window.parent.location = destino
}
if (Opcion == "7") {
destino = "./principal.asp?ver=0&path=busqueda.asp"
window.parent.location = destino
}

if (Opcion == "8") {
 Encuesta()
}
if (Opcion == "9") {
destino = "./principal.asp?ver=0&path=opciones_datos.asp"
window.parent.location = destino
}

if (Opcion == "10") 
   {
		destino = "./principal.asp?ver=0&path=ingdatospersonales.asp"
		window.parent.location = destino
	}

if (Opcion == "12") 
   {
		destino = "./principal.asp"
		window.parent.location = destino
	}

}


function Ingresar() {
if (verificar(document.FORM1) == true) {
FORM1.submit()
}
}
 function verificar(form) {

   if (form.txtlogin.value.length == 0) {
   alert("No ingreso su login")
   form.txtlogin.select()
   form.txtlogin.focus()
   return false
   }
   if (form.txtpassword.value.length ==0) {
   alert("No ingreso su contraseña")
   form.txtlogin.select()
   form.txtlogin.focus()
   return false
   }
 
  return true
 }


 function Curriculum() 
	{
	  ancho = 700
	  if (screen) {
	  
	  leftPos = (screen.width / 2) - (ancho/2)
	  }
	  catWindow = window.open('ingreso_curriculum.asp','Curriculum','toolbar=no,menubar=no,location=no,scrollbars=yes,width='+ancho+'+,height=350,left='+leftPos+'+,top=20')
	}	
	

function DesidEditor()
{
var destino;
var str;
var buscado;

 var s = navigator.appVersion;
  ss = s.substr(22,1);
  
  if (ss.indexOf("6")==0)            
  {
	destino = "./principal.asp?ver=0&path=nuevoforo2.asp"
  }
  else 
  {
	destino = "./principal.asp?ver=0&path=nuevoforo.asp"
  }
window.parent.location = destino

}

function Encuesta() 	
   {
	  ancho = 320
	  altura = 450
	  if (screen) {
	  leftPos = (screen.width / 2) - (ancho/2)
	  }
	  catWindow = window.open('encuesta.asp','Encuesta','toolbar=no,menubar=no,location=no,scrollbars=no,width='+ancho+'+,height='+altura+',left='+leftPos+'+,top=20')
	  
	 return false
	}
</script>

<link href="styles/interior.css" rel="stylesheet" type="text/css">
</head>
<body onLoad="MM_preloadImages('images/loginDown.gif')">
<table width="365" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="791" colspan="2"><img src="images/logo01.gif" width="365" height="79"></td>
  </tr>
  <tr>
    <td rowspan="2" valign="top"><img src="images/logo02.gif" width="44" height="29"></td>
    <td valign="top"><img src="images/logo03.gif" width="321" height="14"></td>
  </tr>
  <tr>
    <td><div align="center"><strong><font color="#000099" size="1">&nbsp;</font></strong></div></td>
  </tr>
</table>
<table width="100%" height="24" border="0" align="center" cellpadding="0" cellspacing="0" background="images/bar02.gif">
  <tr>
    <td width="2"><img src="images/bar01.gif" width="2" height="24"></td>
	<td width="776"><div align="center"><font size="2">
			    <a href="javascript:Destino('12')">Ingresar</a>&nbsp; 
	    :&nbsp; <a href="javascript:Destino('4')">Administraci&oacute;n</a>&nbsp; 
		:&nbsp; <a href="javascript:Destino('5')">Foros</a>&nbsp; 
		:&nbsp; <a href="javascript:Destino('8')">Encuestas</a>&nbsp; 
		:&nbsp; <a href="javascript:Destino('10')">Registro</a>
		</font></div></td>
    <td width="2" align="right"><img src="images/bar03.gif" width="3" height="24"></td>
  </tr>
</table>

<table width="790" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="30%" valign="top">
	<table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><img src="images/spacer.gif" width="10" height="10"></td>
        </tr>
        <tr>
          <td><table width="132" height="110" border="0" align="center" cellpadding="0" cellspacing="0" background="images/bar04.gif">
              <tr>
                <td height="24"><table width="132" height="24" border="0" align="center" cellpadding="0" cellspacing="0" background="images/bar02.gif">
                    <tr>
                      <td width="9"><img src="images/bar01.gif" width="2" height="24"></td>
                      <td width="776"><div align="left"><strong><font size="2">Login</font></strong></div></td>
                      <td width="10" align="right"><img src="images/bar03.gif" width="3" height="24"></td>
                    </tr>
                </table></td>
              </tr>
              <tr>
                <td height="92" align="center"><form name="FORM1" method="post" action="./verifica_usuario.asp">
                    <table width="95%"  border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td align="left"><strong><font color="#000099" size="2">Usuario:</font></strong></td>
                      </tr>
                      <tr>
                        <td align="center"><input name="txtlogin" type="text" id="txtlogin" size="15"></td>
                      </tr>
                      <tr>
                        <td align="left"><font color="#000099" size="2"><strong>Clave:</strong></font></td>
                      </tr>
                      <tr>
                        <td align="center"><input name="txtpassword" type="password" size="15"></td>
                      </tr>
                      <tr>
                        <td><a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('login','','images/loginDown.gif',1)"><img src="images/loginUp.gif"  onClick="javascript:Ingresar();" alt="Ingresar como usuario registrado" name="login" width="123" height="18" border="0"></a></td>
                      </tr>
                    </table>
                </form></td>
              </tr>
              <tr>
                <td><img src="images/bar05.gif" width="132" height="3"></td>
              </tr>
          </table></td>
        </tr>
        <tr>
          <td><img src="images/spacer.gif" width="10" height="10"></td>
        </tr>
        <tr>
          <td><table width="200" height="80" border="0" align="center" cellpadding="0" cellspacing="0" background="images/bar04.gif">
              <tr>
                <td height="24">
				<table width="200" height="24" border="0" align="center" cellpadding="0" cellspacing="0" background="images/bar02.gif">
                    <tr>
                      <td width="10"><img src="images/bar01.gif" width="2" height="24"></td>
                      <td width="776"><div align="left"><strong><font size="2">B&uacute;squedas</font></strong></div></td>
                      <td width="10" align="right"><img src="images/bar03.gif" width="3" height="24"></td>
                    </tr>
                </table></td>
              </tr>
              <tr>
                <td height="53" align="center"><form name="form1" method="post" action="">
                    <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="35" align="left"><strong><img src="images/lupa.gif" width="35" height="33"></strong></td>
                        <td align="left"><input name="palabra" type="text" id="palabra2" size="20"></td>
                      </tr>
                      <tr align="center">
                        <td colspan="2"><input name="Submit" type="submit" class="inputBottom" value=" B U S C A R "></td>
                      </tr>
                    </table>
                </form></td>
              </tr>
              <tr>
                <td><img src="images/bar05.gif" width="200" height="3"></td>
              </tr>
          </table></td>
        </tr>
        <tr>
          <td><img src="images/spacer.gif" width="10" height="10"></td>
        </tr>
        <tr>
          
        </tr>
        <tr>
          <td>
		  	<table height="80" width="150"  border="0" align="center" cellpadding="0" cellspacing="0" background="images/bar04.gif">
              <tr>
                <td >
				<table width="275" height="24" border="0" align="center" cellpadding="0" cellspacing="0" background="images/bar02.gif">
                    <tr>
                      <td width="9"><img src="images/bar01.gif" width="2" height="2"></td>
                      <td ><div align="center"><strong><font size="2">Enlaces de Interes</font></strong></div></td>
                      <td width="10" align="right"><img src="images/bar03.gif" width="3" height="2"></td>
                    </tr>
                </table></td>
              </tr>
              <tr>
                <td height="53" align="center">
					  <iframe  id="datamain" src="listado_url.asp"  width=275 height=100 align="middle" marginwidth=0 marginheight=0 hspace=0 vspace=0 frameborder=0 scrolling=yes >
					  </iframe>
				</td>
              </tr>
              <tr>
                <td><img src="images/bar05.gif" width="275" height="3"></td>
              </tr>
          </table>
		  
		  
		  </td>
        </tr>
        <tr>
          <td><img src="images/spacer.gif" width="10" height="10"></td>
        </tr>
		<!--
        <tr>
          <td>
		  <table width="200" height="80" border="0" align="center" cellpadding="0" cellspacing="0" >
              <tr>
                <td height="24">
				<table width="200" height="24" border="0" align="center" cellpadding="0" cellspacing="0" background="images/bar02.gif">
                    <tr>
                      <td width="9"><img src="images/bar01.gif" width="2" height="24"></td>
                      <td width="776"><div align="left"><strong><font size="2">Proyectos</font></strong></div></td>
                      <td width="10" align="right"><img src="images/bar03.gif" width="3" height="24"></td>
                    </tr>
                </table></td>
              </tr>
              <tr>
                <td height="53" align="center">
				<table width="100%" height="75"  border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td valign="top"><div align="justify"><a href="#">a</a></div></td>
                    </tr>
                </table></td>
              </tr>
              <tr>
                <td><img src="images/bar05.gif" width="200 height="3"></td>
              </tr>
          </table></td>
        </tr>
		-->
        <tr>
          <td><img src="images/spacer.gif" width="10" height="10"></td>
        </tr>
    </table></td>
    <td align="center" valign="top"><p>&nbsp;</p>
<!-- ********************************************************   QUITARME  *********************************	-->
    <iframe  id="mainnoticias" src="./noticias.asp?opcion=1"  width=700 height=450 align="middle" marginwidth=0 marginheight=0 hspace=0 vspace=0 frameborder=0 scrolling=yes >
	</iframe>
    </td>
  </tr>
  <tr>
  <td>&nbsp;</td>
  <td>	 <font size="1" color="#0000FF" face="Verdana, Arial, Helvetica, sans-serif">&copy; Copy Right 2004 - Asociaci&oacute;n de Estudiantes de Ciencias y Sistemas, USAC </font></div></td>
  </tr>
</table>
</body>
</html>
