<%@ Language=VBScript %>
<%
dim db
if isempty(session("id")) then
    Response.Redirect("Login.asp")
end if
id = session("id")
dim id
dim rs
set db = server.CreateObject ("SQLComandos.Comandos")
%>
<HTML>
<HEAD>
<link href="./styles/interior.css" rel="stylesheet" type="text/css">
</HEAD>
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
//			if (document.FORM.correo.value == "SI")
		//	{
				document.FORM.correo_noticias.disabled = false
				document.FORM.correo_foros.disabled = false
				document.FORM.correo_contactos.disabled = false
		
				document.FORM.correo_noticias.checked = false
				document.FORM.correo_foros.checked = false
				document.FORM.correo_contactos.checked = false
			//}
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
</script>
<%


sql = "select u.sexo,u.nombre,u.login,u.apellido,u.mail,u.direccion,u.telefono,u.tipo,u.carnet,u.colegiado,u.pais,u.tipo_mail,u.recibir_mail,uv.id coduv ,uv.nombre universidad,uv.id univid, uv.pais paisuv from usuario u, universidad uv where u.id = " & id & " and u.universidad = uv.id "

set rs = db.SQLSelect ("DBBAECYS",sql)


if not rs.EOF then
nombre = rs("nombre")
apellido = rs("apellido")
direccion = rs("direccion")
telefono = rs("telefono")
mail = rs("mail")
login = rs("login")
tipo_id = rs("tipo")
pais = rs("pais")
paisuv = rs("paisuv")
universidad = rs("universidad")
coduniversidad = rs("coduv")
sexo = rs("sexo")
tipo_mail = rs("tipo_mail")
recibir_mail = rs("recibir_mail")
univid = rs("univid")

		if cint(tipo_id) = 1 then 
		documento = rs("carnet")
		else
		documento = rs("colegiado")
		end if
		set rs = nothing
end if
 pos1 = instr(1,tipo_mail,"F") 
 pos2 = instr(1,tipo_mail,"N") 
 pos3 = instr(1,tipo_mail,"C") 
 posf =  pos1 + pos2 +pos3 

%>
<script language="javascript">
	function inicio()
	{
	document.FORM.universidad.selectedIndex = <%=univid%>
	//alert('<%=univid%>'+" "+document.FORM.universidad.selectedIndex+"#"+document.FORM.universidad.value+ "#"	)
	}
</script>
<BODY onload="javascript:{if(parent.frames[0]&&parent.frames['opciones'].Go)parent.frames['opciones'].Go()}; inicio();<%if  posf = 0 then%>Visualizar();<%end if%>" >


<div align="center">
<h2>Actualizar Datos</h2>
<FORM  id="FORM" name="FORM" action="actualizar_datosregistro.asp" method="post">
    <TABLE WIDTH="52%" BORDER="0" align ="center" cellpadding="2" cellspacing="2" >
      <TR class="reglon01"> 
        <th width="11%" align="right" nowrap class="celdaBgBlueLeft" >Nombre:</th>
        <TD width="38%" class="celdaBgBlue05"> <div align="left">
          <INPUT id="nombre" name="nombre" type="text" maxlength="30" size="30" value ="<%=nombre%>">        
        </div></TD>
      </tr>
      <tr class="reglon02"> 
        <th width="26%" align="right" nowrap class="celdaBgBlueLeft" >Apellido:</th>
        <TD width="25%" class="celdaBgBlue05"> <div align="left">
          <INPUT id="apellido" name="apellido" type="text" size="30" maxlength="30" value ="<%=apellido%>">        
        </div></TD>
      </TR>
      <TR class="reglon01"> 
        <th width="11%" align="right" nowrap class="celdaBgBlueLeft" >Login:</th>
        <TD width="38%" class="celdaBgBlue05"> <div align="left">
          <INPUT id="login" name="login" type="text" value="<%=login%>">        
        </div></TD>
      </TR>
	  <!--
      <TR> 
        <th width="11%" align="right" nowrap class="celdaBgBlueLeft" >Password: </th>
        <TD width="38%" class="celdaBgBlue05"> <div align="left">
          <input id=¨"password" type="password" name="password">        
        </div></TD>
      </TR>
      <TR> 
        <th width="11%" align="right" nowrap class="celdaBgBlueLeft" >Reingreso Password: </th>
        <TD width="38%" class="celdaBgBlue05"> <div align="left">
          <INPUT id=¨"password2" type="password" name="password2">
        </div></TD>
      </TR>
	  -->
      <TR class="reglon02"> 
        <th height="28" align="right" nowrap class="celdaBgBlueLeft" >Sexo</th>
        <TD height="28" class="celdaBgBlue05">
          <div align="left">
            <select name="sexo">
              <option value="1" <%if sexo = "1" then Response.Write ("selected") end if%> >Masculino</option>
              <option value="2" <%if sexo = "2" then Response.Write ("selected") end if%> >Femenino</option>
            </select>
        </div></TD>
      </TR>
      <TR class="reglon01"> 
        <th width="11%" height="74" align="right" nowrap class="celdaBgBlueLeft" >Direccion: </th>
        <TD width="38%" height="74" class="celdaBgBlue05"> <div align="left">
          <textarea  name="direccion" cols="33" rows="3" ><%=direccion%></textarea>
        </div></TD>
      </TR>
      <tr class="reglon02"> 
        <th nowrap class="celdaBgBlueLeft"><div align="right"><font size="-1">Pais de origen: </font></div></th>
        <td class="celdaBgBlue05"> <div align="left">
          <SELECT size=1 id=nacionalidad name="nacionalidad" >
              <%
		sql = "select id,nombre from pais order by nombre"
		set rs = db.SQLSelect ("DBBAECYS",sql)
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
      </tr>
      <TR class="reglon01"> 
        <th width="11%" align="right" nowrap class="celdaBgBlueLeft" >Datos Universidad: </th>
        <TD width="38%" class="celdaBgBlue05"> <div align="left">
          <table class="reglon01" border="0" cellpadding="2" cellspacing="2">
              <tr> 
                <td class="celdaBgBlue10">Pais </td>
                <td> <SELECT size=1 id=pais name="pais" onchange="ShowUniversidades(document.FORM.pais, document.FORM.universidad, document.FORM.pais.options[document.FORM.pais.selectedIndex].value)" >
                    <%
		sql = "select id,nombre from pais order by nombre"
		
		set rs=db.SQLSelect ("DBBAECYS", sql)
		if not rs.eof  then
		while rs.eof <> true
		pais_num = cint(paisuv)
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
				  <script >inicio();alert("holamundo")</script>
			    </td>
              </tr>
              <tr> 
                <TD width="11%" align="right" class="celdaBgBlue10" >Universidad: </TD>
                <td><SELECT size=1 id=universidad name="universidad" >
                    <SCRIPT>ShowUniversidades(document.FORM.pais, document.FORM.universidad, document.FORM.pais.options[document.FORM.pais.selectedIndex].value);
		</SCRIPT>
                  </SELECT> </td>
              </tr>
          </table>
        </div></TD>
      </TR>
      <TR class="reglon02"> 
        <th width="11%" align="right" nowrap class="celdaBgBlueLeft" >Telefono: </th>
        <TD width="38%" class="celdaBgBlue05"> <div align="left">
          <INPUT id=¨"telefono" type="text" name="telefono" value ="<%=telefono%>">
        </div></TD>
      </TR>
      <TR class="reglon01"> 
        <th width="11%" align="right" nowrap class="celdaBgBlueLeft" >Correo Electronico: </th>
        <TD width="38%" class="celdaBgBlue05"> <div align="left">
          <INPUT id=¨"email" type="text" name="email" value ="<%=mail%>">
        </div></TD>
      </TR>
      <TR class="reglon02"> 
        <th width="11%" align="right" nowrap class="celdaBgBlueLeft" >Deseo Recibir Correos: </th>
        <TD width="38%" class="celdaBgBlue05"> <div align="left">
          <INPUT name="correo" type="checkbox" id=¨"correo" value ="NO" <%if recibir_mail = "NO" then%>checked<%end if%> onClick="javascript:Visualizar(this.value)">
          NO. 
          <INPUT id=¨"correo_foros" type="checkbox" name="correo_foros" <%if Instr(1, tipo_mail, "F") <>0 then %>checked<%end if%> value ="F">
          Foros. 
          <INPUT id=¨"correo_noticias" type="checkbox" name="correo_noticias"  <%if Instr(1, tipo_mail, "N") <>0 then%>checked<%end if%> value ="N">
          Noticias. 
          <INPUT id=¨"correo_contactos" type="checkbox" name="correo_contactos"  <%if Instr(1, tipo_mail, "C") <>0 then%>checked<%end if%> value ="C">
          Contactos. </div></TD>
      </TR>
      <TR class="reglon01"> 
        <th width="11%" height="67" align="right" nowrap class="celdaBgBlueLeft" >Documento</th>
        <TD width="38%" class="celdaBgBlue05"> <div align="left">
          <table width="408" border="0" class="reglon01" >
              <tr> 
                <td>
			    <table border="0" class="reglon01">  
			    <tr>
			    <td class="celdaBgBlue10">
                    <div align="left">
                        <%tipo = cint(tipo_id) %>
				    
                    <INPUT id=¨"tipo_id" type="radio" name="tipo_id" <%if tipo = 1 then Response.Write "checked" end if%> value="1" >				 
                      </div></td> 
				   <td>
                    Carnet Universitario: 
				    </td>
				    <tr>
				    <td class="celdaBgBlue10">
                      <div align="left">
                        <INPUT id=¨"tipo_id" type="radio" name="tipo_id" <%if tipo = 2 then Response.Write "checked" end if%> value = "2" >				
                      </div></td>
				  <td>
                  No. Colegiado </td>
				  </tr>
				  </table>
			    <td> Numero: 
                  <input id=¨"documento" type="text" name="documento" value ="<%=documento%>"> 
                </td>
              </tr>
          </table>
        </div></TD>
      </TR>
    </TABLE>
	<INPUT id="id" type="hidden" name="id" value="<%=id%>">
	<br>
	<INPUT name="actualizar" type="submit" class="celdaBgBlue05" id="actualizar" value="Actualizar">
</FORM>
</div>
</BODY>
</HTML>
