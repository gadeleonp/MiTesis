<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<!--<link href="./theme/Master.css" rel="stylesheet" type="text/css">-->
<script language="javascript">
  function validar(forma) {
     if (forma.cadena.value.length == 0 ) 
    {
	alert("No ingreso ningún criterio de busqueda")
	  return false
    }
    
    return true
  }
</script>
<title>Pagina de Búsqueda</title>
<link href="styles/interior.css" rel="stylesheet" type="text/css">
</head>
<body background="imagen/fondo.jpg"  onload="javascript:{if(parent.frames[0]&&parent.frames['opciones'].Go)parent.frames['opciones'].Go()}" >

<%valp=request.querystring("p")%>
<% if IsEmpty( Request("cadena"))  and valp<>1 and valp<>"-1" then%>
<FORM name="forma" ACTION="busqueda.asp" METHOD="post" onsubmit="return validar(this)">
<!--
	<table align="center" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td>
				<h1 align="center"><font color="#000099">Formulario de  Búsqueda</font> </h1>
			</td>
		<tr>
			<td>
				<font size="3"><b>Introduzca su consulta:</b></font>
			</td>
		</tr>
   	     <tr>
			<td>
				<INPUT TYPE="TEXT" NAME="cadena" SIZE="60" 	MAXLENGTH="100" VALUE="">
		</td>	 
		</tr>
		<td align="center">
				<INPUT NAME="Action" TYPE="SUBMIT" class="celdaBgBlue05" VALUE="Buscar">
				&nbsp;
				<INPUT NAME="limpiar" TYPE="Reset" class="celdaBgBlue05" VALUE="Limpiar">
		</td>
	</tr>	
</table>
<!---->
		<table width="200" height="80" border="0" align="center" cellpadding="0" cellspacing="0" >
              <tr>
                <td height="24">
				<table width="600" height="24" border="0" align="center" cellpadding="0" cellspacing="0" background="images/bar02.gif">
                    <tr>
                      <td width="10"><img src="images/bar01.gif" width="2" height="24"></td>
                      <td width="776"><div align="center"><strong><font size="2">Formulario de  Búsqueda</font></strong></div></td>
                      <td width="10" align="right"><img src="images/bar03.gif" width="3" height="24"></td>
                    </tr>
                </table></td>
              </tr>
              <tr>
                <td height="53" align="center"><form name="form1" method="post" action="">
                    <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="35" align="left"><strong><img src="images/lupa.gif" width="35" height="33"></strong></td>
                        <td align="left"><input name="cadena" type="text" id="cadena" size="60" maxlength="60"></td>
                      </tr>
                      <tr align="center">
                        <td colspan="2"><input name="Action" type="submit" class="inputBottom" value=" B U S C A R ">
						
						</td>
                      </tr>
                    </table>
                </form></td>
              </tr>
              <tr>
                <td><img src="images/bar05.gif" width="600" height="3"></td>
              </tr>
          </table>
<!---->

</FORM>



<% else %>



			


<% 
			set Bus = Server.CreateObject("ixsso.Query")
			set util = Server.CreateObject("ixsso.util")
	
			Bus.CATALOG="ProyectoAecys"
			
			
			%>
		<%if valp<>1 and valp<>"-1" then%>
			<%
			session("cad")=Request("cadena")
			%>		

		<%end if%>


			<%
			
			Bus.Query = session("cad")
			Bus.SortBy = "rank[d]"
			Bus.Columns = "DocTitle, path, filename, size, write, DocCategory, Rank, Create, Characterization,vpath "
			'Bus.Query = "@all " & session("cad") & " AND NOT #path *_* AND NOT #path *downloads* AND NOT #path *images* AND NOT #filename *.class AND NOT #filename *.asa AND NOT #filename *.css AND NOT #filename *postinfo.html"
			'Bus.Query = "@all " & session("cad") & " AND NOT #path *images* AND NOT #filename *.doc AND NOT #filename *.xls AND NOT #filename *.css AND NOT #filename *.txt"			
			'oQuery.Columns = "DocAuthor, vpath, doctitle, FileName, Path, Write, Size, Rank, Create, Characterization, DocCategory"
			set RS = Bus.createrecordset("nonsequential")
			
			
			RS.PageSize=10
			ActiveQuery=True
			%>


		<%if valp=1 and request.querystring("pos")>0 then
			RS.AbsolutePosition=request.querystring("pos")	
		elseif valp="-1" and request.querystring("pos")>11 then
			RS.AbsolutePosition=request.querystring("pos")-20%>
		<%end if%>
	<td height="189">
    <p align="left">
	<%if not(RS.eof) then%>

  <h2 align="center"><font color="#000099">Resultados de la Búsqueda</font></h2>

		
  <table width="818" height="42"  border=0 align="center" cellpadding="2" cellspacing="2">
    <tr > 
      <th class="reglon02"  >No.</th>
      <th width="17" class="reglon01"></th>
      <th class="reglon02" >Nivel de Concordancia</th>
      <th width="17" class="reglon01"></th>
      <th class="reglon02" >Título del Documento</th>
      <th width="17" class="reglon01"></th>
      <th class="reglon02" >Nombre del Archivo</th>
      <th width="17" class="reglon01"></th>
      <th class="reglon02" >Fecha de Creación</th>
      <th width="17" class="reglon01"></th>
	  <th width="20"   nowrap class="reglon02">Tamaño (KBytes)</th>
    </tr>
    <%counter=0%>
    <%do while counter<10 and not (RS.eof)%>
    <%path= RS.fields("path") %>
    <%
	     
		  'cadena = server.MapPath("..")
		  cadena =  "e:\aecys\peacycod\"
		 
		  path=Mid(path, len(cadena)+1, len(path))    
		  
			

			  if  counter mod 2 = 0 then
			  			  color = "reglon01"
				else
							  color = "reglon02"
				end if
		
		   %>
    <tr class="<%=color%>"> 
      <td ><div align="center"><font size="3"><%=RS.AbsolutePosition%></font> 
        </div>
      <td width="17"><div align="center"></div></td>
      <td ><div align="center"> 
          <%
				rango = cint(RS.fields("Rank"))
				'Response.Write rango & "-"
				resultado = ((rango * 100 ) / 1000 )
				Response.Write resultado & " %"
				%>
          </div></td>
      <td width="17"></td>
      <td ><%=RS.fields("DocTitle")%></td>
      <td width="17"></td>
      <td > 
        <%
'		  		pos = instr(1,RS("vpath"),"noticias")
'				if  pos  > 0 then
				%>
        <a href="<% =path %>" target="foundfile" ><font size="3" >
        <% =RS.fields("filename") %>
        </font></a> 
        <%
'				else response.write ("Documento no disponible")
				%>
        <%
				'end if%>
      </td>
      <td width="17"></td>
      <td ><%=RS.fields("create")%></td>
      <td width="17"></td>
      <td nowrap align="center">      <%
				tamanio = clng(RS.fields("size"))
				'Response.Write rango & "-"
				kbytes = (tamanio / 1024 )
				pos = instr(1,kbytes,".")
				Response.write left(kbytes,pos+2) & " KB" 
	  %></td>
      <%counter=counter+1%>
    </tr>
    <%
'			Response.write "<b>Size:</b> " & RS("Size") & "<br>"
'			Response.write "<b>Create:</b> " & RS("Create") & "<br>"
'			Response.write "<b>Write:</b> " & RS("Write") & "<br>"
'			Response.write "<b>Characterization:</b> " & RS("Characterization") & "<br>"
'			Response.write "<b>Rank:</b> " & RS("Rank") & "<br>"
'			Response.write "<b>VPath:</b> " & RS("vpath") & "<br>" 
			%>
    <%RS.movenext%>
    <%loop%>
  </table>
<br>
        <table  align="center" border="0" cellspacing="2" cellpadding="2">
          <tr>
            <td align="center">
						<font size="2">
				        	<a href="busqueda.asp?p=-1&pos=<%=RS.AbsolutePosition%>&cad=<%=session("cad")%>">
							<!--<img src="imagenes/icon_private_remall.gif" width="23" height="22" border="0">-->
							Inicio</a>
			  </font>

            </td>
            <td align="center">
											<font size="2">
							  <a href="busqueda.asp?p=0">
							  <!--
							  				<img border="0" src="imagenes/otra_consulta.gif" width="23" height="22">
											-->
											Nueva
								</a>
						  </font>
            </td>
            <td align="center">
				        <font size="2">
								<a href="busqueda.asp?p=1&pos=<%=RS.AbsolutePosition%>&cad=<%=session("cad")%>">
								<!--
										<img border="0" src="imagenes/icon_private_add.gif" width="23" height="22">
										-->
										 Siguiente
								</a>
						</font>
            </td>
          </tr>
        </table>
              <%else%>
              <h2 align="center"><b><font size="4" color="#0000CC">No existen  coincidencias.</font></b> <br>
								<img border="0" src="imagenes/otra_consulta.gif" width="23" height="22" onClick="javascript:window.location='busqueda.asp'" alt="Volver a Consultar">
                <%end if%>
                &nbsp; </h2>
<%end if%>
</body>
</html>
