<%@ Language=VBScript %>
<%
dim user
user = session("UsEmpresa")
if isempty(user) then
  Response.Redirect("../Login.asp")
end if
dim db
dim rs
set db = server.CreateObject("SQLComandos.Comandos")
%>
<HTML>
<HEAD>
<link href="../theme/Master.css" rel="stylesheet" type="text/css">
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<script language="javascript">
<!--#INCLUDE FILE="./listado_universidades.asp"-->
</script>
<script>
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
	function DisableArea()	{

	FORM.Area.disabled
	FORM.Area.clearAttributes
	FORM.Area.value = "FF"

	}
</script>

<BODY>

<div align="center">
  <h2>Buscar Usuarios</h2>
  <form id="FORM" name="FORM" method="post" action="resultados_busqueda.asp">
  <table border="0" align="center" cellpadding="2"  cellspacing="2">
	<tbody>
	<tr>
		  <td width="39%" class="celdaBgBlueLeft">Area: 
            <input type="checkbox" name="checkbox" value="checkbox" onclick="javascript:DisableArea()">
          No usar</td>
		<td class="celdaBgBlue05">
			<select name="Area" size="3" multiple >
			<option value="AA">---Todas---</option>			
			<%
			sql = "select id,descripcion from area "
			set rs = db.SQLSelect("DBBAECYS",sql)
			if not rs.EOF  then
			while rs.EOF <> true 
			%>
			<option value="<%=rs("id")%>" ><%=rs("descripcion")%></option>			
			<%
			rs.MoveNext 
			wend 
			
			end if
			set rs = nothing
			%>
		  </select>		</td>
	<tr>
	<tr>
		  <td class="celdaBgBlueLeft">Sexo</td>
		<td class="celdaBgBlue05">
			<select name="sexo" >
				<option  value="AA" type="radio" selected>--Ambos--</option>			
				<option  value="1" type="radio" >Masculino</option>
				<option  value="2" type="radio" >Femenino</option>				
		  </select>		</td>
	<tr>
	<tr>
		<td class="celdaBgBlueLeft">Pais: </td>
		<td class="celdaBgBlue05">
			<select name="paisorigen" >
				<option value="AA" selected>---Todos---</option>
		<%
		sql = "select id,nombre from pais order by nombre"
		set rs = db.SQLSelect("DBBAECYS",sql)
		if not rs.eof  then
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
		  </select>		</td>
	<tr>
	<tr>
		  <td class="celdaBgBlueLeft">Universidad: 
		  <br>
            <input type="checkbox" name="UsarUni" value="1" checked>
          No usar</td>
		<TD width="61%" class="celdaBgBlue05"> <table border="0" cellpadding="2" cellspacing="2">
            <tr> 
              <td>Pais </td>
              <td> <SELECT size=1 id=pais name="pais" onchange="ShowUniversidades(document.FORM.pais, document.FORM.universidad, document.FORM.pais.options[document.FORM.pais.selectedIndex].value)" >
                           
                  <%
		sql = "select id,nombre from pais order by nombre"
		set rs = db.sqlselect("DBBAECYS",sql)
		if not rs.eof  then
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
              <TD align="right" width="11%" >Universidad: </TD>
              <td><SELECT size=1 id=universidad name="universidad" >
                  <SCRIPT>ShowUniversidades(document.FORM.pais, document.FORM.universidad, document.FORM.pais.options[document.FORM.pais.selectedIndex].value);
		</SCRIPT>
                </SELECT> </td>
            </tr>
          </table></TD>
	<tr>	
	<tr>
		<td class="celdaBgBlueLeft">Mostrar unicamente <br> 
		  con Curriculum publicado: </td>
		<td class="celdaBgBlue05">
			<input value="1" name="curriculum" type="checkbox" >&nbsp;&nbsp;Si		</td>
	<tr>	
	</tbody>
  </table>
  <br>
  <table border="0" cellpadding="2" cellspacing="2">
	<tbody>  
	<tr>
	<td>
  		<input name="Buscar" type="Submit" class="celdaBgBlue05" value="Buscar" >
	</td>	
	<td width="17">

	</td>		
	<td>
  		<input name="Cancelar" type="Submit" class="celdaBgBlue05" value="Cancelar" >
	</td>				
	</tr>		
	</tbody>		
	</table>		
	<input value="1" type="hidden" name="pagina">
	<input value="0" type="hidden" name="criterio">
  </form>
</div>

</BODY>
</HTML>
<%
set db = nothing
%>