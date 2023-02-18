<%@ Language=VBScript %>
<%
if isempty (session("administrador")) then
 Response.Redirect("Login.asp")
end if
%>

<HTML>
<HEAD>
<link href="../theme/Master.css" rel="stylesheet" type="text/css">

<script language="javascript">
function Criterio(Crit) {
 noticia.action = "mantnoticia.asp"
 noticia.criterio.value = Crit
 noticia.submit 
 
 }

function CambiaDestino() {
window.location = "mantnoticias.asp"
return true
}

function CheckAll(checked) {
	len = document.Frmnoticia.elements.length;
	var i=0;
	for( i=0; i<len; i++) {
		if (document.Frmnoticia.elements[i].name=='noticia') {
			document.Frmnoticia.elements[i].checked=checked;
		}
	}
}
function RevisarSeleccion() {
	len = document.Frmnoticia.elements.length;
	var i=0;
	for( i=0; i<len; i++) {
		if (document.Frmnoticia.elements[i].name=='noticia') {
			if (document.Frmnoticia.elements[i].checked==true) 
			{
			   return true
			}
		}
	}
	alert("No se selecciono ninguna noticia para eliminar.")
	return false
}


</script>
</HEAD>
<BODY>
<h2 align = "center" >Eliminación de Noticias</h2>
<form name="Frmnoticia" method="post" action="elimina_noticia.asp" onSubmit="return RevisarSeleccion()">
  <table width="82%" border="0" align = "center" cellpadding="2" cellspacing="2">
    <tr class="celdaBgBlue"> 
      <th width="4%" nowrap> <div align="center">-</div></th>
      <th width="20%" nowrap> <div align="center"><a href="eliminarnoticia.asp?criterio=1">Titulo 
          Noticia</a></div></th>
      <th width="20%" nowrap> <div align="center"><a href="">Archivo Inicial</a></div></th>
      <th width="20%" nowrap> <div align="center"><a href="">Resumen</a></div></th>
      <th width="24%" nowrap> <div align="center"><a href="javascript:Criterio('2')">Autor</a></div></th>
      <th width="12%" nowrap> <div align="center"><a href="">Fecha Publicacion</a></div></th>
      <th width="12%" nowrap> <div align="center"><a href="">Numero de documentos</a></div></th>
      <th width="12%" nowrap> <div align="center">Folder Contenedor</div></th>
    </tr>
    <%
dim db
dim rs
dim sql
dim criterio
set db = server.CreateObject("SQLComandos.Comandos")
 
 criterio = ""
 criterio = Request.Form("criterio")
 
 'Response.Write "criterio =" & criterio
 
 Sql = "select id,titulo, resumen,  autor, fecha_publicacion,folder_contenedor from noticia"
 
	if len(criterio) > 0 then
		sql = sql & " order by "
		if criterio = "1" then
			Sql = Sql & " titulo "
		end if
	else 
	end if
  set rs=db.SQLSelect("DBBAECYS",sql)
  if rs.eof = true then
  else
  dim rs2
  cont = 0
  while rs.eof <> true
  
  sql = "select id,enlace,correlativo from documento_noticia where correlativo = 1 and noticia = " & rs("id")
  set rs2=db.SQLSelect("DBBAECYS",sql)
  
  fecha = rs("fecha_publicacion") 
  fecha = left(fecha,8)
	if (cont mod 2) = 0 then
	color = "celdaBgBlue05"
	else
	color = "celdaBgBlue10"
	end if
	cont = cont + 1
	%>
	
    <tr class="<%=color%>">      <td width="4%"> <input type="checkbox" name="noticia" value="<%=rs("id")%>"></td>
      <td width="20%" nowrap> <div align="center"><%=rs("titulo")%></div></td>
      <%
      if rs2.eof <> true then
      %>
      <td width="20%" nowrap> <div align="center"><%=rs2("enlace")%> </div></td>
      <%
      set rs2= nothing
      end if
      
      %>
      <td width="20%"> <div align="center"> 
          <textarea name="resumen" readonly ><%=rs("resumen")%></textarea>
        </div></td>
      <td width="24%" nowrap> <div align="center"><%=rs("Autor")%></div></td>
      <td width="12%" nowrap> <div align="center"><%=fecha%></div></td>
      <%
      sql = "select count(correlativo) correlativo from documento_noticia where noticia = " & rs("id")
      set rs2=db.SQLSelect("DBBAECYS",sql)
      if rs2.EOF <> true then
      NumCorrelativo = rs2("correlativo")
      %>
      <td width="12%"> <div align="center"><%=NumCorrelativo%></div></td>
      <%
      rs2.Close 
      end if
      %>
      <td width="12%" nowrap> <div align="center"><%=rs("folder_contenedor")%></div></td>
    </tr>
    <%
rs.movenext
wend 
end if

%>
  </table>
<br>
<table align = "center">
<tbody>
<tr>
<td><input type="submit" name="Eliminar" value="Eliminar"></td>
      <td> <font color="#FFFFFF">----</font> </td>
<td><input type="button" name="Regresar" value="Regresar" onclick="javascript:CambiaDestino()"></td>
</tr>
</tbody>
</table>
<input id = "criterio" name="criterio" value = "" type = "hidden">
</form>
</BODY>
</HTML>
