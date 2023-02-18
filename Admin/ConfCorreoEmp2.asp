<%@LANGUAGE="VBSCRIPT" %>
<%
if isempty (session("administrador")) then
 Response.Redirect("Login.asp")
end if
%>
<%
dim db
dim sql
dim rs

set db = server.CreateObject("SqlComandos.Comandos")

 sql = "select subject,saludo,TituloAutorResumen,Publicar,TipoLetra,Negrilla,cuenta from ConfCorreo where descripcion = 'Empresa'"

  set rs = db.SQLSelect("DBBAECYS",Sql)
  if not rs.eof then
  	  tema = rs("subject")
	  saludo = rs("saludo")
	  TituloAutorResumen =  rs("TituloAutorREsumen")
	  publicar = rs("publicar")
	  tipoletra = rs("tipoLetra")
	  Negrilla = rs("Negrilla")
	  cuenta = trim(rs("cuenta"))
  end if
  set rs = nothing
  set db = nothing
  
   
  if len(publicar) > 0 and not isnumeric(publicar) then
    publicar = 2
  end if

  

function buscar(CadenaBusqueda,CaracterBusqueda)
	MiPos = Instr(1, CadenaBusqueda, CaracterBusqueda)
	if isnumeric(MiPos) then
	  regreso = MiPos
	else
	  regreso = -1
	end if
	buscar = Regreso
end function
%>
<HTML xmlns:mytb
	  xmlns:mymenu
	  xmlns:t ="urn:schemas-microsoft-com:time" >
<head>
<?IMPORT namespace="mytb" implementation="toolbar.htc">
<?IMPORT namespace="mymenu" implementation="menu.htc">
<?IMPORT namespace="t" implementation="#default#time2">
<style>
		.time 		    { behavior: url(#default#time2) }
</style>
<TITLE>Configuracion Correos Empresa</TITLE>

<HTA:APPLICATION ID="oHTA"
     APPLICATIONNAME="HTML Editor"
     BORDER="thick"
     BORDERSTYLE="raised"
     CAPTION="yes"
     ICON="netpad.ico"
     MAXIMIZEBUTTON="yes"
     MINIMIZEBUTTON="yes"
     SHOWINTASKBAR="yes"
     SINGLEINSTANCE="yes"
     SYSMENU="yes"
     VERSION="1.0"
     WINDOWSTATE="normal"/> 

<script language="javascript">
  function validacion_email(entered, alertbox){
	with (entered){
		apos=value.indexOf("@");
		dotpos=value.lastIndexOf(".");
		lastpos=value.length-1;
		if (apos<1 || dotpos-apos<2 || lastpos-dotpos>3 || lastpos-dotpos<2) {
			if (alertbox) {
				alert(alertbox);
			} 
			select();
			focus();
			return false;
		}
		else {
			return true;
		}
	}
}

function verificar(formaingreso) 
  {
		if(validacion_email(formaingreso.correo,'No es una direccion de correo valida')==false){
				return false;
		}
		 return true
  }
  
 function Cambiar() 
 {
   document.location = 'configuracion_correos.asp'
  }
  
  function Configuracion() 
	{
  <% 
  if publicar = 1 then
  %>
  document.form1.chkenviacorreo.checked = true
  <% 
  end if
 
  if buscar(TituloAutorResumen,"t") > 0 then
  %>
  document.form1.titulo.checked = true
  <%
  end if
  if buscar(TituloAutorResumen,"a")  > 0  then
  %>
  document.form1.autor.checked = true
  <%
  end if
  if buscar(TituloAutorResumen,"r")  > 0 then  
  %>
  document.form1.resumen.checked = true
  <%
  end if
  if isnumeric(TipoLetra) then
   %>
  	document.form1.TipoLetra.value= '<%=TipoLetra%>'
   <%end if
   
   if isnumeric(Negrilla) then
   %>
  	document.form1.Negrilla.checked	= '<%=Negrilla%>'
   <% end if%>		
    }
</script>
<script>
window.onload=doInit
//sInitColor is a global variable. It holds the value of the selected color in the color dialog box when it displays
var sInitColor = null;
//sPersistValue holds the value of the saved innerHTML 
var sPersistValue
function doInit(){
 	for (i=0; i<document.all.length; i++)
		//ensure that all document elements except the content editable DIV are unselectable
    	document.all(i).unselectable = "on";
    	oDiv.unselectable = "off";
		//clear any text in the Document window and set the focus
		oDiv.innerHTML="";
		oDiv.focus();	
		//call routines to populate the drop-down list boxes on the toolbar
    	getSystemFonts();
 		getBlockFormats();
		window.resizeTo(730,560);
		goContext(); 
}

//This function works for all of the command identifiers used in this page
function callFormatting(sFormatString){
	document.execCommand(sFormatString);
}

//Fonts routines
//getSystemFonts uses the dialog helper object to return an array of all of the fonts on the user's system, then populates a drop-down listbox in the toolbar with the array elements
function getSystemFonts(){
	var a=dlgHelper.fonts.count;
	var fArray = new Array();
	var oDropDown = oToolBar.createDropDownListAt("4");
	oDropDown.setAttribute("id","FontNameList");
	for (i = 1;i < dlgHelper.fonts.count;i++){ 
		fArray[i] = dlgHelper.fonts(i);
		var aOptions = oDropDown.getOptions();	
		var oOption = document.createElement("OPTION");
		aOptions.add(oOption);	
		oOption.text = fArray[i];
		oOption.Value = i;
	} 
	//attaching the onchange event is necessary in order to detect when a user changes the value in the drop-down listbox
		oDropDown.setAttribute("onchange",ChangeFont);
}

//changeFontSize detects the value of the item in the drop-down listbox and applies the value to the font of the selected text
function changeFontSize(){
	var sSelected=oToolBar.getItem(6).getOptions().item(oToolBar.getItem(6).getAttribute("selectedIndex"));
   	document.execCommand("FontSize", false, sSelected.value);
}

//changeFont detects the value of the item in the drop-down listbox and applies the value to the font of the selected text
function ChangeFont(){	
	var sSelected=oToolBar.getItem(4).getOptions().item(oToolBar.getItem(4).getAttribute("selectedIndex"));
	document.execCommand("FontName", false, sSelected.text);
}

//BlockFormats routines
//getBlockFormats uses the dialog helper object to return an array of all of the block formats on the user's system, then populates a drop-down listbox in the toolbar with the array elements
function getBlockFormats(){
	var a=dlgHelper.blockFormats.count;
	var fArray = new Array();
	var oDropDown = oToolBar.createDropDownListAt("5");
	oDropDown.setAttribute("id","FormatList");
	for (i = 1;i < dlgHelper.blockFormats.count;i++)
	{ 
		fArray[i] = dlgHelper.blockFormats(i);
		var aOptions = oDropDown.getOptions();	
		var oOption = document.createElement("OPTION");
		aOptions.add(oOption);	
		oOption.text = fArray[i];
		oOption.Value = i;
	} 
	//attach the onchange event
	oDropDown.setAttribute("onchange",ChangeFormat);
}

//ChangeFormat detects the value of the item in the drop-down listbox and applies the value to the font of the selected text
function ChangeFormat(){
	var sSelected=oToolBar.getItem(5).getOptions().item(oToolBar.getItem(5).getAttribute("selectedIndex"));
	document.execCommand("FormatBlock", false, sSelected.text);
} 


//callColorDlg uses the dialog helper object's ChooseColorDlg method to open the color dialog box, then changes the font or back color of the selected text
function callColorDlg(sColorType){

if (sInitColor == null) 
	//display color dialog box
	var sColor = dlgHelper.ChooseColorDlg();
else
	var sColor = dlgHelper.ChooseColorDlg(sInitColor);
	//change decimal to hex
	sColor = sColor.toString(16);
	//add extra zeroes if hex number is less than 6 digits
if (sColor.length < 6) {
  	var sTempString = "000000".substring(0,6-sColor.length);
  	sColor = sTempString.concat(sColor);
}
	//change color of the selected text
	document.execCommand(sColorType, false, sColor);
	sInitColor = sColor;
	oDiv.focus();
}

//VerticalMode changes the orientation of the text from left to right to top to bottom
 function VerticalMode(){
 	if (oDiv.style.writingMode == 'tb-rl')
    	oDiv.style.writingMode = 'lr-tb';
  	else
    	oDiv.style.writingMode = 'tb-rl';
}

//SaveDocument uses the common dialog box object to display the save as dialog, then writes a textstream object from the value of the div's innerHTML property
function SaveDocument(){
//Setting CancelError to true and using try/catch allows the user to click cancel on the save as dialog without causing a script error
  	cDialog.CancelError=true;
  	try{
  		cDialog.Filter="HTM Files (*.htm)|*.htm|Text Files (*.txt)|*.txt"
  		cDialog.ShowSave();
  		var fso = new ActiveXObject("Scripting.FileSystemObject");
  		var f = fso.CreateTextFile(cDialog.filename,  true);
  		f.write(oDiv.innerHTML);
  		f.Close();
  		sPersistValue=oDiv.innerHTML;}
  	catch(e){
  		var sCancel="true";
  		return sCancel;}
	oDiv.focus();	
  }

//LoadDocument uses the common dialog box object to display the open dialog box, then reads the file and displays its contents in the div
function LoadDocument(){
//Setting CancelError to true and using try/catch allows the user to click cancel on the save as dialog without causing a script error
   	cDialog.CancelError=true;
   	try{
   		var answer = checkForSave();
   		//The user has clicked yes in the modal dialog box called in the checkForSave function
   		if (answer) {var sCancel = SaveDocument();
   			//The user has clicked cancel in the save as dialog box; exit function
   			if (sCancel) return; 
   			cDialog.Filter="HTM Files (*.htm)|*.htm|Text Files (*.txt)|*.txt"
	    	cDialog.ShowOpen();
   			var ForReading = 1;
   			var fso = new ActiveXObject("Scripting.FileSystemObject");
   			var f = fso.OpenTextFile(cDialog.filename, ForReading);
   			var r = f.ReadAll();
   			f.close();
   			oDiv.innerHTML=r;
   			//This variable is used in the checkForSave function to see if there is new content in the div 
   			sPersistValue=oDiv.innerHTML;
			
   		}
   		//The user has clicked no in the modal dialog box called in the checkForSave function
  		if (answer==false)
   			{cDialog.Filter="HTM Files (*.htm)|*.htm|Text Files (*.txt)|*.txt"
  			cDialog.ShowOpen();
   			var ForReading = 1;
  			var fso = new ActiveXObject("Scripting.FileSystemObject");
   			var f = fso.OpenTextFile(cDialog.filename, ForReading);
  			var r = f.ReadAll();
  			f.close();
  			oDiv.innerHTML=r;
  			sPersistValue=oDiv.innerHTML;
			}
  		oDiv.focus();
		}
   	catch(e){
   		var sCancel="true";
   		return sCancel;}
} 

//NewDocument creates a new "document" by clearing the content of the div. If there is any new content in the div, the user is asked whether or not to save
function NewDocument(){
  	var answer = checkForSave();
   	if (answer) {var sCancel = SaveDocument();
	 	if (sCancel) return;
		oDiv.innerHTML="";}
    if (answer==false)
        oDiv.innerHTML="";
	oDiv.focus();
}

//This function checks to see if the div has new text, then displays a modal dialog box if appropriate
function checkForSave()
{ 
	if ((oDiv.innerHTML!=sPersistValue)&&(oDiv.innerHTML !=""))
	 	var checkSave=showModalDialog('dcheckForSave.htm','','dialogHeight:125px;dialogWidth:250px;scroll:off');
	else
	  	var checkSave=false;
  	return checkSave;
}

//this function is used to call other functions when the user clicks on a menu item. These are the same functions that are called by the toolbar buttons.
function CallMenuFunction(){
var menuChoice = event.result;
switch(menuChoice){
     	case "open":	
			LoadDocument();
			break;
		case "new":
            NewDocument();
		    break;
		case "save":
			SaveDocument();
			break;
		case "exit":
			 window.close();
			break;
		case "cut":
			callFormatting('Cut');
			break;
        case "copy":
			callFormatting('Copy');
		    break;
        case "paste":
			callFormatting('Paste');
            break;
		case "bold":
			callFormatting('Bold');
            break;
		case "underline":
			callFormatting('Underline');
            break;
		case "italic":
			callFormatting('Italic');
            break;
		case "fontColor":
			callColorDlg('ForeColor');
            break;
		case "highlight":
			callColorDlg('BackColor');
            break;
		 case "about":
            goContext(); 
            break;
        default:
			break;
			}
}

//These functions create and display the splash screen and are used when the application is launched (called by onInit function) as well as when the user clicks help/about on the menu
var oPopup = window.createPopup()
function goContext()
{
/*  var oPopupBody = oPopup.document.body;

  oPopupBody.innerHTML = oContext.innerHTML;
  oPopup.show(175, 125, 400, 300, document.body);
  document.body.onmousedown = oPopup.hide;
*/
}

 </script>

<link href="../theme/Master.css" rel="stylesheet" type="text/css">
</HEAD>

<BODY STYLE="overflow:hidden; margin:0px" >
<!--Create the Dialog Helper Object-->
 <OBJECT ID=dlgHelper CLASSID="clsid:3050f819-98b5-11cf-bb82-00aa00bdce0b" WIDTH="0px" HEIGHT="0px"></OBJECT>

<!-- Create the licensing object for the common dialog activex control -->
<OBJECT CLASSID="clsid:5220cb21-c88d-11cf-b347-00aa00a28331">
    <PARAM NAME="LPKPath" VALUE="comdlg.lpk">
</OBJECT>

<!--Create the Common Dialog Box activex control-->
<OBJECT ID="cDialog" WIDTH="0px" HEIGHT="0px"
    CLASSID="CLSID:F9043C85-F6F2-101A-A3C9-08002B2F49FB"
    CODEBASE="http://activex.microsoft.com/controls/vb5/comdlg32.cab">
</OBJECT>
<t:animate dur="5" onend="oPopup.hide();"/>
<DIV ID="oContext" STYLE="display:none">
<DIV STYLE="position:absolute; top:0; left:0; width:400px; height:300px; border:1px solid black; background:#eeeeee; "> 
<DIV STYLE="padding:20px; background:white; border-bottom:5px solid #cccccc">
<img src="htmledlogo.gif" align="absmiddle"></div>
<DIV STYLE="padding:20px; font-size:8pt; line-height:1.5em; font-family:verdana; color:black;">This HTML Editor is created entirely with DHTML technologies and is provided as an example for developers using Internet Explorer 6.
<br>
<br>
<br>
<br><center>
&#169;2001 Microsoft Corporation. All rights reserved. </center>
</DIV>
</DIV>
</DIV>




<DIV  ID="oContainer" STYLE="background-color:threedface; border:1px solid #cccccc; position:relative; height:80%; top:0;">


<mymenu:menu id="File" backcolor="threedface" onsubmenu_click="CallMenuFunction();">File
<mymenu:menu id="new" backcolor="threedface">New</mymenu:menu>
<mymenu:menu id="open" backcolor="threedface">Open</mymenu:menu>
<mymenu:menu id="save" backcolor="threedface">Save</mymenu:menu>
<mymenu:menu id="exit" backcolor="threedface">Exit</mymenu:menu>
</mymenu:menu>
<mymenu:menu id="Edit" backcolor="threedface" onsubmenu_click="CallMenuFunction();">Edit
<mymenu:menu id="cut" backcolor="threedface">Cut Ctrl+X</mymenu:menu>
<mymenu:menu id="copy" backcolor="threedface">Copy Ctrl+C</mymenu:menu>
<mymenu:menu id="paste" backcolor="threedface">Paste Ctrl+V</mymenu:menu>
</mymenu:menu>
<mymenu:menu id="Format" backcolor="threedface" onsubmenu_click="CallMenuFunction();">Format
<mymenu:menu id="bold" backcolor="threedface">Bold</mymenu:menu>
<mymenu:menu id="italic" backcolor="threedface">Italic</mymenu:menu>
<mymenu:menu id="underline" backcolor="threedface">Underline</mymenu:menu>
<mymenu:menu id="fontColor" backcolor="threedface">Font Color</mymenu:menu>
<mymenu:menu id="highlight" backcolor="threedface">HighLight</mymenu:menu>
</mymenu:menu>
<mymenu:menu id="Help" backcolor="threedface" onsubmenu_click="CallMenuFunction();">Help
<mymenu:menu id="about" backcolor="threedface">About</mymenu:menu>
</mymenu:menu>
 <nobr>
 

<mytb:TOOLBAR STYLE="display:inline; width:100%" ID="oToolBar" 
>
<!--New document button-->
<mytb:TOOLBARBUTTON  IMAGEURL="UI_new.gif"
title="New Document"
onclick="NewDocument();"/>

<!--Open document button-->
<mytb:TOOLBARBUTTON  IMAGEURL="UI_open.gif"
title="Open Document"
onclick="LoadDocument();" />
<!--Save document button-->

<mytb:TOOLBARBUTTON  IMAGEURL="UI_save.gif"
title="Save Document"
onclick="SaveDocument();"/>

<mytb:TOOLBARSEPARATOR />

	
 <!-- Font size list--> 
<mytb:TOOLBARDROPDOWNLIST id="oFontSize" onchange="changeFontSize();">
  <option value=1>1
  <option value=2>2
  <option value=3>3
  <option value=4 selected>4
  <option value=5 >5
  <option value=6>6
  <option value=7>7

</mytb:TOOLBARDROPDOWNLIST>

<mytb:TOOLBARSEPARATOR/>

<!--Bold button-->
<mytb:TOOLBARBUTTON IMAGEURL="UI_bold.gif" 
title="Bold"
onclick="callFormatting('Bold');"/>

<!--Italic button--> 
<mytb:TOOLBARBUTTON IMAGEURL="UI_italic.gif"
title="Italic" 
onclick="callFormatting('Italic');"/>

<!--Underline button-->  
<mytb:TOOLBARBUTTON IMAGEURL="UI_underline.gif"
title="Underline" 
onclick="callFormatting('Underline');"/>

<!--Strikethrough button-->  
<mytb:TOOLBARBUTTON IMAGEURL="UI_form-strike.gif" 
title="StrikeThrough"
onclick="callFormatting('StrikeThrough');"/>
 
<mytb:TOOLBARSEPARATOR/>

<!--Superscript button-->  
<mytb:TOOLBARBUTTON IMAGEURL="UI_superscript.gif"
title="SuperScript"
onclick="callFormatting('SuperScript');"/>

<!--Subscript button -->  
<mytb:TOOLBARBUTTON IMAGEURL="UI_subscript.gif" 
title="SubScript"
onclick="callFormatting('SubScript');"/>

</mytb:TOOLBAR>

</nobr>
<nobr>

<mytb:TOOLBAR STYLE="display:inline; width:100%" ID="oToolBar2"

>

<!--Cut button--> 
<mytb:TOOLBARBUTTON  IMAGEURL="UI_tool-cut.gif"
title="Cut"
onclick="callFormatting('Cut');"/>

<!--Copy button--> 
<mytb:TOOLBARBUTTON  IMAGEURL="UI_form-copy.gif"
title="Copy"
onclick="callFormatting('Copy');"/>

<!--Paste button--> 
<mytb:TOOLBARBUTTON  IMAGEURL="UI_paste.gif"
title="Paste"
onclick="callFormatting('Paste');"/>

<mytb:TOOLBARSEPARATOR/>

<!--Undo button--> 
<mytb:TOOLBARBUTTON  IMAGEURL="UI_undo.gif"
title="Undo"
onclick="callFormatting('Undo');"/>

<!--Redo button--> 
<mytb:TOOLBARBUTTON  IMAGEURL="UI_redo.gif"
title="Redo"
onclick="callFormatting('Redo');"/>

<mytb:TOOLBARSEPARATOR/>

<!--Numbered list button--> 
<mytb:TOOLBARBUTTON  IMAGEURL="UI_numberlist.gif"
title="InsertOrderedList"
onclick="callFormatting('InsertOrderedList');"/>

<!--Bulleted list button--> 
<mytb:TOOLBARBUTTON  IMAGEURL="UI_bulletlist.gif"
title="InsertUnorderedList"
onclick="callFormatting('InsertUnorderedList');"/>

<!--Outdent button--> 
<mytb:TOOLBARBUTTON  IMAGEURL="UI_outdent.gif"
title="Outdent"
onclick="callFormatting('Outdent');"/>

<!--Indent button-->
<mytb:TOOLBARBUTTON  IMAGEURL="UI_indent.gif"
title="Indent"
onclick="callFormatting('Indent');"/>

<mytb:TOOLBARSEPARATOR/>

<!--Left alignment button-->
<mytb:TOOLBARBUTTON  IMAGEURL="UI_leftalign.gif"
title="JustifyLeft"
onclick="callFormatting('JustifyLeft');"/>

<!--Center alignment button-->
<mytb:TOOLBARBUTTON  IMAGEURL="UI_centeralign.gif"
title="JustifyCenter"
onclick="callFormatting('JustifyCenter');"/>

<!--Right alignment button-->
<mytb:TOOLBARBUTTON  IMAGEURL="UI_rightalign.gif"
title="JustifyRight"
onclick="callFormatting('JustifyRight');"/>

<mytb:TOOLBARSEPARATOR/>

<!--Vertical alignment button-->
<mytb:TOOLBARBUTTON  IMAGEURL="UI_tool_vertical.gif"
title="Vertical Alignment"
onclick="VerticalMode();"/>

<mytb:TOOLBARSEPARATOR/>

<!--Font Color button--> 
<mytb:TOOLBARBUTTON IMAGEURL="UI_fontcolor.gif"
title="Font Color" 
onclick="callColorDlg('ForeColor');"/>

<!--Background color button --> 
<mytb:TOOLBARBUTTON  IMAGEURL="UI_backcolor.gif"
title="HighLight"
onclick="callColorDlg('BackColor');"/>

</mytb:TOOLBAR> 
</nobr>

<!--The content editable div-->
<DIV ID=oDiv CONTENTEDITABLE STYLE="height:80%;background-color:white; overflow:auto; width:100%;"> 
</DIV></DIV>
<form name="Form" method="post" action="IngConfCorreoNoti.asp" onsubmit="return verificar(this)">
<table border="0" align="center" cellpadding="2" cellspacing="0" class="celdaBgBlue05">
  <tr class="celdaBgBlue">
    <td>&nbsp;</td>
  </tr>
  <tr class="celdaBgBlue05">
    <th nowrap class="celdaBgBlue10"><div align="left">
      <input type="checkbox" name="chkenviacorreo" value="1" tabindex="1">
      Permitir a las empresas enviar correos.</div>
    </th>
  </tr>
  <tr>
    <th nowrap class="celdaBgBlue10" ><div align="left" class="celdaBgBlue05">Tema:
     </div></th>
   </tr>
  <tr>
    <td class="celdaBgBlue10"><div align="left">
      <input name="tema" type="text" size="100" maxlength="250" tabindex="2" value="<%=tema%>">
    </div></td>
    
  </tr>
  <tr>
    <th class="celdaBgBlue10"><div align="left" class="celdaBgBlue05">Saludo:</div></th>
 
  </tr>
  <tr>
    <td class="celdaBgBlue10"><textarea name="saludo" cols="75" rows="3" tabindex="3"><%=Saludo%></textarea></td>
 
  </tr>
    <tr>
    <th class="celdaBgBlue05"><div align="left">Cuenta de Correo para empresas: </div></th>
  </tr>
  <tr>
    <td class="celdaBgBlue10"><div align="left">
      <input type="text" name="correo" size="50" value="<%=cuenta%>" tabindex="7">
    </div></td>
   </tr>
</table>
<table border="0" align="center" cellpadding="2" cellspacing="2">
  <tr>
    <td><input name="Aceptar" type="submit" class="celdaBgBlue05" value="Aceptar" tabindex="10" ></td>
    <td width="17"></td>
    <td><input name="Cancelar" type="button" class="celdaBgBlue05" value="Cancelar" tabindex="11" onClick="javascript:Cambiar()"></td>
  </tr>
</table>
<input type="hidden" value="Empresa" name="tipocorreo" >

<input name="data" value="" type="hidden">

</form>
</BODY>
</HTML>