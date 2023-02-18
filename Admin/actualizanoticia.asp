<%@ Language=VBScript %>
<%
if isempty (session("administrador")) then
 Response.Redirect("Login.asp")
end if
dim db
set db = server.CreateObject("SQLComandos.Comandos")

%>

<HTML>
<HEAD>
<link href="../theme/Master.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style1 {color: #FFFFFF}
-->
</style>
</HEAD>
<script language="javascript">

function CambiaDestino() {
window.location = "mantnoticias.asp"
return true
}
function RevisarSeleccion() {
	len = document.Actnoticia.elements.length;
	var i=0;
	for( i=0; i<len; i++) {
		if (document.Actnoticia.elements[i].name=='noticia') {
//		alert("Valor: "+i+document.Actnoticia.elements[i].checked)
			if (document.Actnoticia.elements[i].checked==true) 
			{
			   return true
			}
		}
	}
	alert("No se selecciono ninguna noticia para actualizar.")
	return false
}
</script>
<BODY>
<h2 align = "center" >Mantenimiento de Noticias</h2>
<form name="Actnoticia" method="post" action="actualizar_noticia.asp" onSubmit="return RevisarSeleccion()">
  <table width="82%" border="0" align = "center" cellpadding="2" cellspacing="2">
    <tr class="celdaBgBlue"> 
      <th width="4%" nowrap> <div align="center" class="style1">-</div></th>
      <th width="20%" nowrap> <a href="eliminarnoticia.asp?criterio=1"><font color="#FFFFFF">Titulo Noticia</font></a></th>
      <th width="20%" nowrap> <div align="center" class="style1"><a href="" class="style1"><font color="#FFFFFF">Pagina Inicial</font></a></div></th>
      <th width="20%" nowrap> <div align="center" class="style1"><a href="" class="style1"><font color="#FFFFFF">Resumen</font></a></div></th>
      <th width="24%" nowrap> <div align="center" class="style1"><a href="javascript:Criterio('2')"><font color="#FFFFFF">Autor</font></a></div></th>
      <th width="12%" nowrap> <div align="center" class="style1"><a href=""><font color="#FFFFFF">Numero de Documentos de la Noticia</font></a></div></th>
      <th width="12%" nowrap> <div align="center" class="style1"><a href=""><font color="#FFFFFF">Fecha Publicacion</font></a></div></th>
    </tr>
    <%
dim rs
dim sql
dim criterio
 
 criterio = ""
 criterio = Request.Form("criterio")
 
' Response.Write "criterio =" & criterio
 
 
 Sql = "select id,titulo, resumen, autor, fecha_publicacion from noticia"
 
 
	
	if len(criterio) > 0 then
		sql = sql & " order by "
		if criterio = "1" then
			Sql = Sql & " titulo "
		end if
	else 
	end if
 'Response.Write sql
 set rs = db.SQLSelect("DBBAECYS",sql)
  if rs.eof = true then
  else
  
  dim rs2
  
  cont = 0
  while rs.eof <> true
  sql = "select id,enlace,correlativo from documento_noticia where correlativo = 1 and noticia = " & rs("id")
  
  set rs2 = db.SQLSelect("DBBAECYS",sql)
  
  fecha = rs("fecha_publicacion") 
  fecha = left(fecha,8)
	
	if (cont mod 2) = 0 then
	color = "celdaBgBlue05"
	else
	color = "celdaBgBlue10"
	end if
	cont = cont + 1
	%>
	
    <tr class="<%=color%>">
      <td width="4%" nowrap aling = "center"> <input type="radio" name="noticia" value="<%=rs("id")%>"> 
      </td>
      <td  nowrap> <div align="center"><%=rs("titulo")%></div></td>
      <td nowrap > 
        <%
      if rs2.eof <> true then
      %>
        <div align="center"><%=rs2("enlace")%> </div>
        <%
       end if
       set rs2 = nothing
       %>
      </td>
      <td width="20%" nowrap> <div align="center"> 
          <textarea name="resumen" readonly ><%=rs("resumen")%></textarea>
        </div></td>
      <td width="24%" nowrap> <div align="center"><%=rs("Autor")%></div></td>
      <%
       sql = "select count(correlativo) total from documento_noticia where noticia = "& rs("id")
       
       set rs2 = db.SQLSelect("DBBAECYS",sql)
       
       if rs2.EOF <> true then
       numdocumentos = rs2("total")
      %>
      <td width="12%" nowrap> <div align="center"><%=numdocumentos%></div></td>
      <% 
        end if
       set rs2 = nothing
        %>
      <td width="12%" nowrap> <div align="center"><%=fecha%></div></td>
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
<td><input type="submit" name="Actualizar" value="Actualizar"></td>
      <td> <font color="#FFFFFF">----</font> </td>
<td><input type="button" name="Regresar" value="Regresar" onclick="return CambiaDestino()"></td>
</tr>
</tbody>
</table>
<input id = "criterio" name="criterio" value = "" type = "hidden">
</form>
</BODY>
</HTML>
