<%@ Language=VBScript %>
<%
dim db
dim rs
set db = server.CreateObject ("SQLComandos.Comandos")
	
	session.Abandon()
		
Dim pais

pais = ""
nombre = ""
apellido = ""
login = ""
password = ""
direccion = ""
telefono = ""
email = ""
tipo = ""
documento = ""
universidad = ""
pais = ""
documento = ""
tipo_id = ""
error = ""

error = Request.Form("error")

nombre = Request.Form("nombre")
apellido = Request.Form("apellido")
login = Request.Form("login")
password = Request.Form("password")
direccion = Request.Form("direccion")
telefono = Request.Form("telefono")
email = Request.Form("email")
tipo = Request.Form("tipo_id")
documento = Request.Form("documento")
universidad = Request.Form("universidad")
pais = Request.Form("pais")
documento = Request.Form("documento")
tipo_id = Request.Form("tipo_id")
%>
<HTML>
<HEAD>
<!--<link href="./theme/Master.css" rel="stylesheet" type="text/css">-->
<link href="./styles/interior.css" rel="stylesheet" type="text/css">
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<title>Ingreso Datos Personales</title>
<script language="javascript">
<!--#INCLUDE FILE="listado_registro.asp"-->
</script>
<script>
function Visualizar()
{

		if (document.FORM.correo.checked == true) 
		{
		document.FORM.correo_noticias.disabled = true
		document.FORM.correo_foros.disabled = true
		document.FORM.correo_contactos.disabled = true

		document.FORM.correo_noticias.checked = false
		document.FORM.correo_foros.checked = false
		document.FORM.correo_contactos.checked = false

		
		}
		else
		{
		document.FORM.correo_noticias.disabled = false
		document.FORM.correo_foros.disabled = false
		document.FORM.correo_contactos.disabled = false

		document.FORM.correo_noticias.checked = false
		document.FORM.correo_foros.checked = false
		document.FORM.correo_contactos.checked = false

		}
}




function ShowUniversidades(cboArea, cbosubarea, sel) 
	{

		if (sel == 0)
		{
			return false;
		}

		var sel_count = 0;
	 		for (var i=0; i < cboArea.options.length; i++)
				if (cboArea.options[i].selected)
				sel_count++;
		if (sel != 0)
 			var opts = eval("la" + sel);
	 	if (sel != 0 && sel_count == 1 && opts.length > 1) 
		{	
 			while (opts.length < cbosubarea.options.length) 
			{
   				cbosubarea.options[(cbosubarea.options.length - 1)] = null;
	 		}

 			for (var i=0; i < opts.length; i++) 
			{
   				eval("cbosubarea.options[i] = new Option" + opts[i]);
 			}
		}
		else 
		{
 			while (cbosubarea.options.length)
   				cbosubarea.options[(cbosubarea.options.length - 1)] = null;
   			eval("cbosubarea.options[0] = new Option('-------------', '-1', true, true)");
		}
	}

function ShowCarreras(cboArea, cbosubarea, sel) 
	{

		if (sel == 0)
		{
			return false;
		}

		var sel_count = 0;
	 		for (var i=0; i < cboArea.options.length; i++)
				if (cboArea.options[i].selected)
				sel_count++;
		if (sel != 0)
 			var opts = eval("lb" + sel);
	 	if (sel != 0 && sel_count == 1 && opts.length > 1) 
		{	
 			while (opts.length < cbosubarea.options.length) 
			{
   				cbosubarea.options[(cbosubarea.options.length - 1)] = null;
	 		}

 			for (var i=0; i < opts.length; i++) 
			{
   				eval("cbosubarea.options[i] = new Option" + opts[i]);
 			}
		}
		else 
		{
 			while (cbosubarea.options.length)
   				cbosubarea.options[(cbosubarea.options.length - 1)] = null;
   			eval("cbosubarea.options[0] = new Option('-------------', '-1', true, true)");
		}
	}


</script>
<script language="javascript">
  function ingreso_pais(formaingreso) {
   formaingreso.action = "ingdatospersonales.asp"
   formaingreso.submit()
   
     return true
  }
 
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
  function validacion_password(entered, min, max, alertbox){
	with (entered){
		if ((value.length<min) || (value.length>max)){
			if (alertbox!="") {
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
  if (formaingreso.action == "ingreso_datos.asp") {
 			if (formaingreso.nombre.value.length == 0) {
			formaingreso.nombre.focus()
  			formaingreso.nombre.select()
			alert("No ingreso su nombre")
			return false
			  }
			if (formaingreso.apellido.value.length == 0) {
			formaingreso.apellido.focus()
			formaingreso.apellido.select()
			  alert("No ingreso su apellido")
			  return false
			}
			if (formaingreso.login.value.length == 0) {
				formaingreso.login.focus()
				formaingreso.login.select()
				alert("No ingreso su login")
				return false
			}

			   if(validacion_email(formaingreso.email,'No es una direccion de correo valida')==false){
				return false;
			}
		
				
				if (formaingreso.documento.value.length == 0) 
				{
				  formaingreso.documento.focus()
				  formaingreso.documento.select()
				  alert("No ingreso su documento de identificación.")
				  return false
				}

	}
	  return true
  }
</script>
<BODY onLoad="javascript:Visualizar()">
<div align="center"> 
<%
if len(error) > 0 or error <> "" then
%>
<h2>El login: "<%=login%>" ya fue ingresado previamente. <br>Por favor vuelve a intentarlo</h2>
<%
end if
%>
  <FORM onsubmit="return verificar(this)" id="FORM" name="FORM" action="ingreso_datos.asp" method="post">
  <!---->
   <table width="95%"  border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td><table width="95%"  border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td width="3%" background="images/bar02.gif"><img src="images/bar01.gif" width="2" height="24"></td>
              <td width="39%" background="images/bar02.gif"><div align="center"><strong><font color="#000099">Ingrese Sus Datos Personales</font></strong></div></td>
              <td width="34%"><img src="images/degra_pic.gif" width="202" height="24"></td>
              <td width="24%" rowspan="3"><img src="images/keyboa_mouse.jpg" width="142" height="191"></td>
            </tr>
            <tr>
              <td colspan="3">&nbsp;</td>
            </tr>
            <tr>
              <td colspan="3" valign="top">
  <!---->
    <TABLE WIDTH="52%" BORDER="0" cellpadding="2" cellspacing="2" >
      <TR class="reglon01"> 
        <th width="11%" align="right" nowrap ><div align="right">Nombre:</div></th>
        <TD width="38%"> <div align="left">
          <INPUT id="nombre" name="nombre" type="text" maxlength="30" size="30" value ="<%=nombre%>"> 
        </div></TD>
      </tr>
      <tr class="reglon02"> 
        <th width="26%" align="right" nowrap class="celdaBgBlueLeft" ><div align="right">Apellido:</div></th>
        <TD width="25%" class="celdaBgBlue05"> <div align="left">
          <INPUT id="apellido" name="apellido" type="text" size="30" maxlength="30" value ="<%=apellido%>"> 
        </div></TD>
      </TR>
      <TR class="reglon01"> 
        <th width="11%" align="right" nowrap class="celdaBgBlueLeft" ><div align="right">Login:</div></th>
        <TD width="38%" class="celdaBgBlue05"> <div align="left">
          <INPUT id="login" name="login" type="text"> 
        </div></TD>
      </TR class="reglon02">
	  <!--
      <TR> 
        <th width="11%" align="right" nowrap class="celdaBgBlueLeft" ><div align="right">Password: </div></th>
        <TD width="38%" class="celdaBgBlue05"> <div align="left">
          <input id=¨"password" type="password" name="password"> 
        </div></TD>
      </TR>
      <TR> 
        <th width="11%" align="right" nowrap class="celdaBgBlueLeft" ><div align="right">Reingreso Password: 
        </div></th>
        <TD width="38%" class="celdaBgBlue05"> <div align="left">
          <INPUT id=¨"password2" type="password" name="password2">
        </div></TD>
      </TR>
	  -->
      <TR class="reglon01"> 
        <th height="28" align="right" nowrap class="celdaBgBlueLeft" ><div align="right">Sexo</div></th>
        <TD height="28" class="celdaBgBlue05"> <div align="left">
          <select name="sexo">
              <option value="1" selected>Masculino</option>
              <option value="2">Femenino</option>
            </select>
        </div></TD>
      </TR>
      <TR class="reglon02"> 
        <th width="11%" height="100" align="right" nowrap class="celdaBgBlueLeft" ><div align="right">Direccion: 
        </div></th>
        <TD width="38%" height="100" class="celdaBgBlue05"> <div align="left">
          <textarea  name="direccion" cols="33" rows="3" ><%=direccion%></textarea> 
        </div></TD>
      </TR>
      <tr class="reglon01"> 
        <th nowrap class="celdaBgBlueLeft"><div align="right">Pais Nacionalidad: </div></th>
        <td class="celdaBgBlue05"> <div align="left">
          <SELECT size=1 id=nacionalidad name="nacionalidad" >
              <%
		sql = "select id,nombre from pais order by nombre"
		set rs = db.SQLSelect("DBBAECYS",sql)
		if not rs.eof then
		while rs.eof <> true
		pais_num = cint(pais)
		cod_pais = cint(rs("id"))
		if pais_num = cod_pais  then
		%>
              <OPTION value="<%=rs("id")%>" selected><%=rs("nombre")%></OPTION>
              <%
		else
		%>
              <OPTION value="<%=rs("id")%>"><%=rs("nombre")%></OPTION>
              <%
		end if
		rs.movenext
		wend 
		end if
		set rs = nothing	
		%>
            </SELECT> 
        </div></td>
      </tr >
      <TR class="reglon02">  
        <th width="11%" align="right" nowrap class="celdaBgBlueLeft" ><div align="right">Datos Universidad: 
        </div></th>
        <TD width="38%" class="celdaBgBlue05"> <div align="left">
          <table class="reglon02" border="0" cellpadding="2" cellspacing="2">
              <tr > 
                <td align="rigth" ><div align="right">Pais </div></td>
                <td align="left"> <SELECT size=1 id=pais name="pais" onchange="ShowUniversidades(document.FORM.pais, document.FORM.universidad, document.FORM.pais.options[document.FORM.pais.selectedIndex].value)" >
                    <%
		sql = "select id,nombre from pais order by nombre"
		set rs = db.SQLSelect("DBBAECYS",sql)
		if not rs.eof then
		while rs.eof <> true
		pais_num = cint(pais)
		cod_pais = cint(rs("id"))
		if pais_num = cod_pais  then
		%>
                    <OPTION value="<%=rs("id")%>" selected><%=rs("nombre")%></OPTION>
                    <%
		else
		%>
                    <OPTION value="<%=rs("id")%>"><%=rs("nombre")%></OPTION>
                    <%
		end if
		rs.movenext
		wend 
		end if
		set rs = nothing
		%>
                  </SELECT> </td>
              </tr>
              <tr> 
                <TD align="right" width="11%" ><div align="right">Universidad</div></TD>
                <TD align="left"><SELECT size=1 id=universidad name="universidad" onchange="ShowCarreras(document.FORM.universidad, document.FORM.carrera, document.FORM.universidad.options[document.FORM.universidad.selectedIndex].value)" >
                    <SCRIPT>ShowUniversidades(document.FORM.pais, document.FORM.universidad, document.FORM.pais.options[document.FORM.pais.selectedIndex].value);
		</SCRIPT>
                  </SELECT> </TD>
              </tr>
              <tr>
                <TD align="right" ><div align="right">Carrera</div></TD>
                <td align="left"><SELECT size=1 id=carrera name="carrera" >
                    <SCRIPT>ShowCarreras(document.FORM.universidad, document.FORM.carrera, document.FORM.universidad.options[document.FORM.universidad.selectedIndex].value);
					</SCRIPT>
                  </SELECT>
				  </td>
              </tr>
			  <tr>
			  <td align="left">Especifique:</td><td align="left"><input name="carrera_esp" type="text" maxlength="100" size="25"></td>
			  </tr>
            </table>
        </div></TD>
      </TR>
      <TR class="reglon01"> 
        <th width="11%" align="right" nowrap class="celdaBgBlueLeft" ><div align="right">Telefono: </div></th>
        <TD width="38%" class="celdaBgBlue05"> <div align="left">
          <INPUT id=¨"telefono" type="text" name="telefono" value ="<%=telefono%>">
        </div></TD>
      </TR>
      <TR class="reglon02"> 
        <th width="11%" align="right" nowrap class="celdaBgBlueLeft" ><div align="right">Correo Electronico: 
        </div></th>
        <TD width="38%" class="celdaBgBlue05"> <div align="left">
          <INPUT id=¨"email" type="text" name="email" value ="<%=email%>">
        </div></TD>
      </TR>
      <TR class="reglon01"> 
        <th width="11%" align="right" nowrap class="celdaBgBlueLeft" ><div align="right">Deseo Recibir 
        Correos: </div></th>
        <TD width="38%" class="celdaBgBlue05"> <div align="left">
            <INPUT name="correo" type="checkbox" id=¨"correo" value ="NO" checked onClick="javascript:Visualizar()">
          NO. 
          <INPUT id=¨"correo_foros" type="checkbox" name="correo_foros" value ="F">
          Foros. 
          <INPUT id=¨"correo_noticias" type="checkbox" name="correo_noticias" value ="N">
          Noticias. 
          <INPUT id=¨"correo_contactos" type="checkbox" name="correo_contactos" value ="C">
          Contactos. </div></TD>
      </TR>
      <TR class="reglon02"> 
        <th width="11%" height="67" align="right" nowrap class="celdaBgBlueLeft" ><div align="right">Documento</div></th>
        <TD width="38%" class="celdaBgBlue05"> <div align="left">
          <table border="0" width="408" >
              <tr> 
                <td> <table class="reglon02" border="0">
                    <tr> 
                      <td> 
                        <%tipo = cint(tipo_id) %>
                        <INPUT id=¨"tipo_id" type="radio" name="tipo_id" <%if tipo = 1 then Response.Write "checked" end if%> value="1" > 
                      </td>
                      <td> Carnet Universitario: </td>
                    <tr> 
                      <td> <INPUT id=¨"tipo_id" type="radio" name="tipo_id" <%if tipo = 2 then Response.Write "checked" end if%> value = "2" > 
                      </td>
                      <td> No. Colegiado </td>
                    </tr>
                  </table>
                <td class="reglon02"> Numero: 
                  <input id=¨"documento" type="text" name="documento" value ="<%=documento%>"> 
                </td>
              </tr>
            </table>
        </div></TD>
      </TR>
    </TABLE>
	<!---->
	</td>
            </tr>
          </table>            </td>
        </tr>
    </table>
	<!---->
	
	<br>
	<INPUT name="aceptar" type="submit" class="inputBottom" id="Aceptar" value="Registrar">
</FORM>
</div>
<h3><b><font color="#0000FF">* Datos obligatorios</font></b> </h3>
</BODY>
</HTML>

<%
set db = nothing
%>