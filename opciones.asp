<html>
<head>
<!--<link href="./theme/Master.css" rel="stylesheet" type="text/css">-->
<link href="./style/interior.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<script language="javascript">

 function Curriculum() 
	{
	  ancho = 700
	  if (screen) {
	  
	  leftPos = (screen.width / 2) - (ancho/2)
	  }
	  catWindow = window.open('ingreso_curriculum.asp','Curriculum','toolbar=no,menubar=no,location=no,scrollbars=yes,width='+ancho+'+,height=350,left='+leftPos+'+,top=20')
	}	
	


function Destino(Opcion) {
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

if (Opcion == "11") {
destino = "./inicio.asp"
window.parent.location = destino
}

}
function DesidEditor()
{
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
<script language="JavaScript">
	var NoOffFirstLineMenus=7;			// Number of first level items
	var LowBgColor='#FFFFFF';			// Background color when mouse is not over
	var LowSubBgColor='#FFFFFF';			// Background color when mouse is not over on subs
	var HighBgColor='white';			// Background color when mouse is over
	var HighSubBgColor='white';			// Background color when mouse is over on subs
	var FontLowColor='#000066';			// Font color when mouse is not over
	var FontSubLowColor='#000066';			// Font color subs when mouse is not over
	var FontHighColor='#990000';			// Font color when mouse is over
	var FontSubHighColor='#990000';			// Font color subs when mouse is over
	var BorderColor='#78A0AF';			// Border color
	var BorderSubColor='#78A0AF';			// Border color for subs
	var BorderWidth=1;				// Border width
	var BorderBtwnElmnts=1;			// Border between elements 1 or 0
	var FontFamily="Verdana, Arial, Helvetica, sans-serif"	// Font family menu items
	var FontSize="8";				// Font size menu items
	var FontBold=1;				// Bold menu items 1 or 0
	var FontItalic=0;				// Italic menu items 1 or 0
	var MenuTextCentered='left';			// Item text position 'left', 'center' or 'right'
	var MenuCentered='left';			// Menu horizontal position 'left', 'center' or 'right'
	var MenuVerticalCentered='top';		// Menu vertical position 'top', 'middle','bottom' or static
	var ChildOverlap=.1;				// horizontal overlap child/ parent
	var ChildVerticalOverlap=.1;			// vertical overlap child/ parent
	var StartTop=80;				// Menu offset x coordinate
	var StartLeft=1;				// Menu offset y coordinate
	var VerCorrect=0;				// Multiple frames y correction
	var HorCorrect=0;				// Multiple frames x correction
	var LeftPaddng=2;				// Left padding
	var TopPaddng=2;				// Top padding
	var FirstLineHorizontal=0;			// SET TO 1 FOR HORIZONTAL MENU, 0 FOR VERTICAL
	var MenuFramesVertical=1;			// Frames in cols or rows 1 or 0
	var DissapearDelay=40;			// delay before menu folds in
	var TakeOverBgColor=0;			// Menu frame takes over background color subitem frame
	var FirstLineFrame='opciones';			// Frame where first level appears
	var SecLineFrame='principal';			// Frame where sub levels appear
	var DocTargetFrame='principal';			// Frame where target documents appear
	var TargetLoc='';				// span id for relative positioning
	var HideTop=0;				// Hide first level when loading new document 1 or 0
	var MenuWrap=1;				// enables/ disables menu wrap 1 or 0
	var RightToLeft=0;				// enables/ disables right to left unfold 1 or 0
	var UnfoldsOnClick=0;			// Level 1 unfolds onclick/ onmouseover
	var WebMasterCheck=0;			// menu tree checking on or off 1 or 0
	var ShowArrow=1;				// Uses arrow gifs when 1
	var KeepHilite=1;				// Keep selected path highligthed
	var Arrws=['tri.gif',5,10,'tridown.gif',10,5,'trileft.gif',5,10];	// Arrow source, width and height

function BeforeStart(){return}
function AfterBuild(){return}
function BeforeFirstOpen(){return}
function AfterCloseAll(){return}
// Menu tree
//	MenuX=new Array(Text to show, Link, background image (optional), number of sub elements, height, width);
//	For rollover images set "Text to show" to:  "rollover:Image1.jpg:Image2.jpg"
//rollover:tri.gif:tridown.gif
//Menu1=new Array("Login","javascript:Destino('2')","",0,20,150);

Menu1=new Array("Inicio","javascript:Destino('11')","",0,20,150);
<%
if isempty(session("id")) then
%>
Menu2=new Array("Login","javascript:Destino('9')","",0,20,150);
<%
 else
%>
Menu2=new Array("Opciones","javascript:Destino('9')","",2,20,150);
	Menu2_1=new Array("Datos","javascript:Destino('9')","",3,20,50);	
		Menu2_1_1=new Array("Actualizar Datos","javascript:Destino('6')","",0,20,150);	
		Menu2_1_2=new Array("Actualizar Preferencias","javascript:Destino('1')","",0,20,150);		
		Menu2_1_3=new Array("Actualizar Curriculum","javascript:Curriculum()","",0,20,150);
	Menu2_2=new Array("Foro","","",3,20,50);
		Menu2_2_1=new Array("Nuevo Foro","javascript:DesidEditor()","",0,20,115);
		Menu2_2_2=new Array("Todos los Foros","./foro_general.asp","",0,20,115);
		Menu2_2_3=new Array("Mis Foros","./foro.asp","",0,20,115);				


<%	
end if
%>
//Menu2=new Array("rollover:./iconos/noticias1.jpg:./iconos/noticias2.jpg","javascript:Destino('3')","",2);
Menu3=new Array("Noticias","javascript:Destino('3')","",2);
	Menu3_1=new Array("Noticias Actuales","./noticias.asp?Opcion=1","",0,20,125);	
	Menu3_2=new Array("Todas las noticias","./noticias.asp?Opcion=2","",0,20,125);
//Menu3=new Array("rollover:./iconos/foro1.jpg:./iconos/foro2.jpg","javascript:Destino('5')","",2);
Menu4=new Array("Foro","javascript:Destino('5')","",0);
//Menu4=new Array("Registro","javascript:Destino('1')","",0);
//Menu5=new Array("rollover:./iconos/registro1.jpg:./iconos/registro2.jpg","javascript:Destino('1')","",0);
//	Menu5_1=new Array("Nuevo Usuario","javascript:Destino('1')","",0,20,150);
Menu5=new Array("Buqueda","javascript:Destino('7')","",0);
Menu6=new Array("Encuesta","javascript:Destino('8')","",0);	
Menu7=new Array("Administracion","javascript:Destino('4')","",0);	


</script>

<body >
<font color="#336699"></font> 
<br>
<br>
<script type='text/javascript'>
function Go(){return}
</script>
<TABLE BORDER="0" cellpadding="0" cellspacing="0" background="images/bar02.gif">
<TR>
<TD width="500"><font color="#000099">OPCIONES</font></TD>
</TR>
<TR>
<td>&nbsp;</td>
<TD width="500">
<script type='text/javascript' src='menu_com.js'></script>
</TD>
</TR>
</TABLE>
</body>




</body>
</html>
