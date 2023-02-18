<%@ Language=VBScript %>

<html>
<head>
<!--<link href="./theme/Master.css" rel="stylesheet" type="text/css">-->
<link href="./styles/interior.css" rel="stylesheet" type="text/css">
<%
dim menu
menu = ""
menu = request.QueryString("menu")

if not isempty(session("id")) then
	if menu = "1" then
    Response.Redirect("opciones_datos.asp")
	end if
	if menu="2" then
    Response.Redirect("foro.asp")
	end if
end if

dim Error
Error = ""
Error = Request.Form ("Error")


%>
<style type="text/css">
<!--
boton {
	background-color: #ffffff;
	border-top: thin none #ffffff;
	border-right: thin none;
	border-bottom: thin none;
	border-left: thin none;
	font-family: Arial, Helvetica, sans-serif;
	font-style: normal;
	line-height: normal;
	font-weight: bold;
	font-variant: normal;
	text-decoration: underline;
}
.style1 {
	color: #333333;
	font-weight: bold;
}
.style3 {color: #333333}
-->
</style>
</head>
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
 
 function Destino() 
   {
		destino = "./principal.asp?ver=0&path=ingdatospersonales.asp"
		window.parent.location = destino
	}
</script>
<body>
<div align="center">
<h2>
<%
Response.Write(Error)
if error = "1" then
Response.Write(error)
end if
%>
</h2>

<form onsubmit="return verificar(this)" id="FORM1" name="FORM" action="./verifica_usuario.asp" method="post">
	<br>
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
                      <td width="776"><div align="left"><strong><font size="2">Log&iacute;n</font></strong></div></td>
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
                </form>
	<br>

	<input  class="celdaBgBlue05" name="entrar" value="Login" type="submit">
  </form>
  <a href="javascript:Destino()">REGISTRARSE</a>
</div>
</body>
</html>
