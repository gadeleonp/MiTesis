<%@ Language=VBScript %>
<%
dim user
user = session("UsEmpresa")
if isempty(user) then
  Response.Redirect("../Login.asp")
end if
dim db
dim area
dim paisorigen
dim universidad
dim sexo
dim curriculum
dim rs

set db = server.CreateObject("SQLComandos.Comandos")

area = ""
pasiorigen = ""
universidad = ""
sexo = ""
curriculum = ""

pagina = request.form("pagina")
paisuni = Request.Form("pais")
area = Request.Form("area")
paisorigen = Request.Form("paisorigen")
universidad = Request.Form("universidad")
sexo = Request.Form("sexo")
curriculum = Request.Form("curriculum")
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" >
<link href="../theme/Master.css" rel="stylesheet" type="text/css">
</HEAD>
<script language="JavaScript">
  function Ordenar(criterio) {
  
     criterionum = parseInt(criterio)
         
     if (criterionum == parseInt(forma.criterio.value) && (criterionum > 0) )
     {
		criterionum = criterionum + 5
		forma.criterio.value = criterionum
     }
     else
     {
     	forma.criterio.value = criterionum
     }
     
        
    forma.action = 'resultados_busqueda.asp'
    forma.submit()
	
  }
  function MandarCorreo() {
  if (RevisarSeleccion()) {
	  forma.action = 'frm_correoemp.asp'
	  forma.submit()
	  }
  }
  function NuevaBusqueda() {
    forma.action = 'buscar_usuario.asp'
	forma.submit()
  }
  function Siguiente() {
	pagina = parseInt(forma.pagina.value)
	totalus = parseInt(forma.totalus.value)
	if ((pagina * 20) >= totalus) {
	}
	else
	{
	pagina = pagina +1 
	}
	forma.pagina.value = pagina
	forma.submit()
  }
    function Anterior() {
	pagina = parseInt(forma.pagina.value)
	if (pagina == 1) {
	pagina = 1
	}
	else {
	pagina = pagina - 1 
	}
	forma.pagina.value = pagina
	forma.submit()
  }
 function RevisarSeleccion() {
	len = document.forma.elements.length;
	var i=0;
	for( i=0; i<len; i++) {
		if (document.forma.elements[i].name=='mail') {
			if (document.forma.elements[i].checked==true) 
			{
			   return true
			}
		}
	}
	alert("Error: No se selecciono a ningun usuario.")
	return false
}
</script>
<BODY onload="javascript:{if(parent.frames[0]&&parent.frames['opciones'].Go)parent.frames['opciones'].Go()}">
<%
sql = "select u.id CodUs, u.mail MailUs, u.nombre NomUs, u.apellido ApUs, u.arc_resume CurriUs, p.nombre PaisRes, uni.id CodUni ,uni.nombre Universidad, Puni.nombre PaisUni, day(u.ultima_visita) Dia, month(u.ultima_visita) Mes, year(u.ultima_visita) Anio, u.ultima_visita UVisita "
sql = sql & " from usuario u, pais p, universidad uni,pais puni  "
sql = sql & " where uni.id = u.universidad" 
sql = sql & " and uni.pais = puni.id "
sql = sql & " and  u.pais = p.id" 
sql = sql & " and u.empresa is null "

if instr(1,area,"AA") <> 0 then
	sql = sql & " and u.id in (select distinct(c.usuario) from conocimiento c where c.area  in ( select distinct(a.id) from area a))"
else
	if not isempty(area) and len(area) > 0 and isnumeric(area) then
		sql = sql & " and u.id in (select distinct(c.usuario) from conocimiento c where c.area  in ("& area & "))"
	end if
end if

if instr(1,sexo,"AA") = 0 then
	if instr(1,sexo,"1") > 0 then
	  sql = sql & " and u.sexo = 1 "	
	else
	  sql = sql & " and u.sexo = 2 "	
	end if
end if

if not isempty(curriculum) and len(curriculum) > 0 and isnumeric(curriculum) then
  sql = sql & " and u.arc_resume is not null "
end if


if not isempty(paisorigen) and len(paisorigen) > 0 and isnumeric(paisorigen) then
  sql = sql & " and u.pais = " & paisorigen
end if

pos = instr(1,universidad,"-1") 
'Response.Write "universidad = -" &universidad &"-<br>"
usaruni = ""
usaruni = Request.Form("UsarUni")

if instr(1,usaruni,"1") > 0 and len(usaruni) > 0 and not isempty(usaruni) then
else
	if not isempty(universidad) and len(universidad) > 0 and isnumeric(universidad) and pos = 0 then
		sql = sql & " and u.universidad = " & universidad
	end if
	if not isempty(universidad) and len(universidad) > 0 and instr(1,universidad,"AA") > 0 then
		sql = sql & " and u.pais = " & paisuni
	end if
end if

criterio = ""
criterio = request.form("criterio")
'Response.Write "criterio sin cambios = -"& criterio & "-"
if isnull(criterio) and len(criterio) < "1" then
  criterio = "0"
end if
	
	
    
    
    critnum = cint(criterio)
	 Select Case  critnum
           Case 1   sql = sql & " order by u.apellido asc "
           Case 2   sql = sql & " order by p.nombre  asc "
           Case 3   sql = sql & " order by u.ultima_visita   asc  "
           case 4   sql = sql & " order by u.arc_resume asc "
		   Case 5   sql = sql & " order by uni.nombre asc "           
           Case 6   sql = sql & " order by u.apellido desc "
           Case 7   sql = sql & " order by p.nombre desc "
           Case 8   sql = sql & " order by u.ultima_visita desc"
		   case 9   sql = sql & " order by u.arc_resume desc "         
		   Case 10   sql = sql & " order by uni.nombre desc "           		     
           Case Else     
      End Select

set rs = db.SQLSelect("DBBAECYS",sql)

if not rs.EOF  then
%>
<form id="forma" name="forma" method="post" action="resultados_busqueda.asp">
<table width="0%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr class="celdaBgBlue"> 
    <th>--</th>
	<th width="17"></th>        
      <th><a class=".Ordenar" href="javascript:Ordenar('1')"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><strong>Nombre</strong></font></a></th>
	<th width="17"></th>    
    <th><a class=".Ordenar" href="javascript:Ordenar('4')"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><strong>Curriculum</strong></font></a></th>
	<th width="17"></th>        
    <th><a class=".Ordenar" href="javascript:Ordenar('2')"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><strong>Pais</strong></font></a></th>
	<th	width="17"></th>        
    <th><a class=".Ordenar" href="javascript:Ordenar('3')"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><strong>Última visita</strong></font></a></th>    
 	<th	width="17"></th>        
    <th><a class=".Ordenar" href="javascript:Ordenar('5')"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><strong>Universidad</strong></font></a></th>    
  </tr>
  
  <%
  cont = 0
  paginanum = cint(pagina)
  intervalo = 20
  
  if paginanum = 1 then
  inicio = 0
  fin = paginanum * intervalo 
  else
  inicio = (paginanum * intervalo) - (intervalo )
  fin  = (paginanum * intervalo)
  end if
  total = 0
  while rs.EOF <> true
	 total = total  + 1 
	 rs.MoveNext
  wend 
  maxpaginas = total / intervalo
  redondeadas = round(total/intervalo)
  
  if redondeadas < maxpaginas then
  maxpaginas = maxpaginas + 1
  
  end if
%>    
<h6> Pagina [<%=paginanum%>] de <%=round(maxpaginas)%></h6>
<br>
<%
rs.MoveFirst
   while rs.EOF <> true 
 
      if (cont >= inicio and cont < fin ) then
	  if  cont mod 2 = 0 then
	  	  color = "celdaBgBlue05"
 	  else
		  color = "celdaBgBlue10"
	   end if
 	 %>
	  <tr class="<%=color%>"> 
	    <td><input name="mail" type="checkbox" value="<%=rs("CodUs")%>"></td>
		<td width="17"></td>            
	    <td><%=rs("NomUs")%>&nbsp; <%=rs("ApUs")%></td>
		<td width="17"></td>                
	    <td><%
		    if isnull(rs("CurriUs")) then
			    Response.Write "No Ingreso Curriculum"
		    else 
		    %>
	    <a  class=".Download1" href="../usuarios/<%=rs("CurriUs")%>" >Curriculum</a>
	    <%
	    end if
	    %>
	    </td>
		<td width="17"></td>                
	    <td><%=rs("PaisRes")%></td>
		<td width="17"></td>                
	    <td>
		<% if not isnull(rs("dia")) then%>
		<%=rs("dia")%> / <%=rs("mes")%> / <%=rs("anio")%></td>    
		<%end if%>	
		<td width="17"></td>                
	    <td><%=rs("Universidad")%> 
	    <%if cint(rs("CodUni")) >  0 then%>
	    Campus: <%=rs("paisuni")%></td>
	    <%end if%>
	
	  </tr>
	  <%
	  end if
	  cont = cont + 1 
	  rs.movenext
	  
	  wend
  %>
</table>
<br>

<table align="center" border = 0 cellpadding="2" cellspacing="2">
	<tr>
	<td>
		<input name="Mail" type="button" class="celdaBgBlue05"  onclick="javascript:MandarCorreo()" value="Correo">
	</td>
	<td width="17">
	</td>
	<td>
		<input name="busqueda" type="button" class="celdaBgBlue05"  onclick="javascript:NuevaBusqueda()" value="Otra Busqueda">
	</td>
		<td width="17"></td>
	<td>
		<input name="busqueda" type="button" class="celdaBgBlue05"  onclick="javascript:Siguiente()" value="Siguiente">
	</td>
	<td width="17"></td>
	<td>
		<input name="busqueda" type="button" class="celdaBgBlue05"  onclick="javascript:Anterior()" value="Anterior">
	</td>

	</tr>
  </table>
	<input name="area" type="hidden" value="<%=area%>" >
	<input name="paisorigen" type="hidden" value="<%=paisorigen%>" >
	<input name="universidad" type="hidden" value="<%=universidad%>" >
	<input name="sexo" type="hidden" value="<%=sexo%>" >
	<input name="curriculum" type="hidden" value="<%=curriculum%>" >
	<input name="criterio" type="hidden" value="<%=criterio%>" >
	<input name="pagina" type="hidden" value="<%=pagina%>" >
	<input name="totalus" type="hidden" value="<%=total%>" >
	<input name="usaruni" type="hidden" value="<%=usaruni%>" >
	<input name="pais" type="hidden" value="<%=paisuni%>" >														
</form>
<%
else
%>
<div align="center">
<h2>No se encontraron registros.</h2>
	<input name="busqueda" type="button" class="celdaBgBlue05"  onclick="javascript:document.location='buscar_usuario.asp'" value="Otra Busqueda">
</div>		
<%
end if
%>

</BODY>
</HTML>

<%
set db = nothing
%>