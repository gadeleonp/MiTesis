<%@ Language=VBScript %>
<%
dim db
dim rs
dim sql
dim rs2


 
 set db = server.CreateObject("SQLComandos.Comandos")
 
%>
<html>
<head>
<style type="text/css">
<!--
body {
	background-attachment: fixed;
	background-image: url(./images/w2k.jpg);
	background-repeat: no-repeat;
	background-position: center center;
}
-->
</style>
<link href="./styles/interior.css" rel="stylesheet" type="text/css">
</head>
<!-- onload="javascript:{if(parent.frames[0]&&parent.frames['opciones'].Go)parent.frames['opciones'].Go()}" -->
<body background="./images/w2k.jpg"  style="fixed" bgcolor="white" text="#000000"  >
<%
 Sql = "select id,titulo, resumen,imagen_principal, autor, fecha_publicacion,folder_contenedor from Noticia"
 Sql = Sql & " order by fecha_publicacion desc "

  
  
  set rs = db.SQLSelect("DBBAECYS",sql)
	   
  if not rs.eof  then
   cont = 1
'    while (rs.eof <> true) 'or (cont >= 3)
dim opcion
dim opcion_num

opcion = Request.QueryString("opcion")
opcion_num = cint(opcion)
	if opcion = 1 then 'Solo las ultimas tres noticias
'		while cont <= 4 and rs.eof <> true
		if rs.eof <> true then
		id = rs("id")
		  autor = rs("autor")
		  imagen_principal = rs("imagen_principal")
		  titulo = rs("titulo")
		  descripcion = rs("resumen")
		  fecha = rs("fecha_publicacion") 
		  fecha = left(fecha,9)
		  folder = rs("folder_contenedor")
		  sql = "select id,enlace from documento_noticia where correlativo = 1 and  noticia = " & rs("id") 
		  set rs2 = db.SQLSelect ("DBBAECYS", sql)
			if not rs2.eof  then
				documento = rs2("enlace")
			end if
		  set rs2 = nothing
		  cont = cont + 1
'				  while cont <= 4 and rs.eof <> true
  rs.movenext				
						  if cont < rs.recordcount and cont = 2 then
  							 id2 = rs("id")
							  autor2 = rs("autor")
							  imagen_principal2 = rs("imagen_principal")
							  titulo2 = rs("titulo")
							  descripcion2 = rs("resumen")
							  fecha2 = rs("fecha_publicacion") 
							  fecha2 = left(fecha2,9)
							  folder2 = rs("folder_contenedor")
							  sql = "select id,enlace from documento_noticia where correlativo = 1 and  noticia = " & rs("id") 
							  set rs2 = db.SQLSelect ("DBBAECYS", sql)
								if not rs2.eof  then
									documento2 = rs2("enlace")
								end if
							  set rs2 = nothing
							  cont = cont + 1
						  end if
						  rs.movenext						
						  
						    if cont < rs.recordcount and cont = 3 then
									id3 = rs("id")
							  autor3 = rs("autor")
							  imagen_principal3 = rs("imagen_principal")
							  titulo3 = rs("titulo")
							  descripcion3 = rs("resumen")
							  fecha3 = rs("fecha_publicacion") 
							  fecha3 = left(fecha3,9)
							  folder3 = rs("folder_contenedor")
							  sql = "select id,enlace from documento_noticia where correlativo = 1 and  noticia = " & rs("id") 
							  set rs2 = db.SQLSelect ("DBBAECYS", sql)
								if not rs2.eof  then
									documento3 = rs2("enlace")
								end if
							  set rs2 = nothing
							  cont = cont + 1
						  end if
						  rs.movenext						  

						    if cont < rs.recordcount and cont = 4 then
									id4 = rs("id")
							  autor4 = rs("autor")
							  imagen_principal4 = rs("imagen_principal")
							  titulo4= rs("titulo")
							  descripcion4 = rs("resumen")
							  fecha4 = rs("fecha_publicacion") 
							  fecha4 = left(fecha4,9)
							  folder4 = rs("folder_contenedor")
							  sql = "select id,enlace from documento_noticia where correlativo = 1 and  noticia = " & rs("id") 
							  set rs2 = db.SQLSelect ("DBBAECYS", sql)
								if not rs2.eof  then
									documento4 = rs2("enlace")
								end if
							  set rs2 = nothing
							  cont = cont + 1
						  end if
						  rs.movenext						  

	'			wend 
			end if
				  
		%>

<!--*************************************************************************************-->
      <div align="center">
      <table width="95%"  border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td><table width="100%" height="24" border="0" align="center" cellpadding="0" cellspacing="0" background="images/bar02.gif">
            <tr>
              <td width="2"><img src="images/bar01.gif" width="2" height="24"></td>
              <td width="776"><div align="left"><strong><font color="#000099"><%=Titulo%></font></strong></div></td>
              <td width="2" align="right"><img src="images/bar03.gif" width="3" height="24"></td>
            </tr>
          </table></td>
        </tr>
        <tr>
          <td><table width="95%"  border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td width="19%" rowspan="3"><img src="./Noticias/<%=folder%>/<%=imagen_principal%>" alt="<%=titulo%>" width="111" height="74" border="1"></td>
              <td width="2%" rowspan="3"><img src="images/spacer.gif" width="10" height="10"></td>
              <td width="79%" align="right" valign="bottom" class="fecha">Guatemala, <%=fecha%></td>
            </tr>
            <tr>
              <td align="right" valign="bottom" class="escritor">Por: <%=autor%></td>
            </tr>
            <tr>
              <td valign="bottom"><div align="justify"><font size="2"><%=titulo%>
              </font> </div>                
                <div align="justify"><font size="2"><%=descripcion%></font></div>
				</td>
				<tr>
				<td>
				<a href="ver_noticia.asp?noticia=<%=id%>&folder=<%=folder%>" ><font size="2" face="Arial, Helvetica, sans-serif" color="#0000CC">Ver Mas...</font></a> 
				</td>
				</tr>
            </tr>
          </table></td>
        </tr>
        <tr>
          <td><img src="images/spacer.gif" width="10" height="10"></td>
        </tr>
        <tr>
		
          <td>
		  
		  <table width="90%"  border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td width="33%"><table width="100%" height="24" border="0" align="center" cellpadding="0" cellspacing="0" background="images/bar02.gif">
                  <tr>
                    <td width="2"><img src="images/bar01.gif" width="2" height="24"></td>
                    <td width="776"><div align="left"><strong><font color="#000099"><%=titulo2%></font></strong></div></td>
                    <td width="2" align="right"><img src="images/bar03.gif" width="3" height="24"></td>
                  </tr>
              </table></td>
              <td width="2%"><img src="images/spacer.gif" width="10" height="10"></td>
              <td width="32%"><table width="100%" height="24" border="0" align="center" cellpadding="0" cellspacing="0" background="images/bar02.gif">
                  <tr>
                    <td width="2"><img src="images/bar01.gif" width="2" height="24"></td>
                    <td width="776"><div align="left"><strong><font color="#000099"><%=titulo3%></font></strong></div></td>
                    <td width="2" align="right"><img src="images/bar03.gif" width="3" height="24"></td>
                  </tr>
              </table></td>
              <td width="2%"><img src="images/spacer.gif" width="10" height="10"></td>
              <td width="31%"><table width="100%" height="24" border="0" align="center" cellpadding="0" cellspacing="0" background="images/bar02.gif">
                  <tr>
                    <td width="2"><img src="images/bar01.gif" width="2" height="24"></td>
                    <td width="776"><div align="left"><font color="#000099"><strong><%=titulo4%></strong></font></div></td>
                    <td width="2" align="right"><img src="images/bar03.gif" width="3" height="24"></td>
                  </tr>
              </table></td>
            </tr>
            <tr>
              <td><div align="right"><span class="fecha">Guatemala, <%=fecha2%></span></div></td>
              <td>&nbsp;</td>
              <td><div align="right"><span class="fecha">Guatemala, <%=fecha3%></span></div></td>
              <td>&nbsp;</td>
              <td><div align="right"><span class="fecha">Guatemala, <%=fecha4%></span></div></td>
            </tr>
            <tr>
              <td><div align="right"><span class="escritor">Por: <%=autor2%></span></div></td>
              <td>&nbsp;</td>
              <td><div align="right"><span class="escritor">Por: <%=autor3%></span></div></td>
              <td>&nbsp;</td>
              <td><div align="right"><span class="escritor">Por: <%=autor4%></span></div></td>
            </tr>
            <tr>
              <td align="center"><img src="./Noticias/<%=folder2%>/<%=imagen_principal2%>" alt="<%=titulo2%>"  width="111" height="74" border="1"></td>
              <td>&nbsp;</td>
              <td><img src="./Noticias/<%=folder3%>/<%=imagen_principal3%>" alt="<%=titulo3%>"  width="111" height="74" border="1"></td>
              <td>&nbsp;</td>
              <td align="center"><img src="./Noticias/<%=folder4%>/<%=imagen_principal4%>" alt="<%=titulo4%>"  width="111" height="74" border="1"></td>
            </tr>
            <tr>
              <td><div align="justify"><font size="2"><%=descripcion2%><br></font></div>				 </td>
              <td>&nbsp;</td>
              <td><div align="justify"><font size="2"><%=descripcion3%></font></div></td>
              <td>&nbsp;</td>
              <td><div align="justify"><font size="2"><%=descripcion4%></font></div></td>
            </tr>
            <tr>
              <td>&nbsp;<a href="ver_noticia.asp?noticia=<%=id2%>&folder=<%=folder2%>" ><font size="2" face="Arial, Helvetica, sans-serif" color="#0000CC">Ver Mas...</font></a></td>
              <td>&nbsp;</td>
              <td>&nbsp;<a href="ver_noticia.asp?noticia=<%=id3%>&folder=<%=folder3%>" ><font size="2" face="Arial, Helvetica, sans-serif" color="#0000CC">Ver Mas...</font></a></td>
              <td>&nbsp;</td>
              <td>&nbsp;<a href="ver_noticia.asp?noticia=<%=id4%>&folder=<%=folder4%>" ><font size="2" face="Arial, Helvetica, sans-serif" color="#0000CC">Ver Mas...</font></a></td>
            </tr>
          </table></td>
        </tr>
        <tr>
          <td><img src="images/spacer.gif" width="10" height="10"></td>

        </tr>
		  <!--		
		  <tr>
          <td>
		  

		  <table width="90%"  border="0" align="center" cellpadding="5" cellspacing="1">
            <tr>
              <td height="20" class="reglon01"><strong>Titular:</strong> Complemento del t&iacute;tulo</td>
            </tr>
            <tr>
              <td height="20" class="reglon02"><strong>Titular:</strong> Complemento del t&iacute;tulo</td>
            </tr>
            <tr>
              <td height="20" class="reglon01"><strong>Titular:</strong> Complemento del t&iacute;tulo</td>
            </tr>
            <tr>
              <td height="20" class="reglon02"><strong>Titular:</strong> Complemento del t&iacute;tulo</td>
            </tr>
            <tr>
              <td height="20" class="reglon01"><strong>Titular:</strong> Complemento del t&iacute;tulo</td>
            </tr>
            <tr>
              <td height="20" class="reglon02"><strong>Titular:</strong> Complemento del t&iacute;tulo</td>
            </tr>
          </table></td>
		  
        </tr>
		-->
      </table>
<!--**************************************************************************************-->
<%
	else
		while cont > 0 and rs.eof <> true

		  autor = rs("autor")
		  imagen_principal = rs("imagen_principal")
		  titulo = rs("titulo")
		  descripcion = rs("resumen")
		  fecha = rs("fecha_publicacion") 
		  fecha = left(fecha,9)
		  folder = rs("folder_contenedor")

		  sql = "select id,enlace from documento_noticia where correlativo = 1 and  noticia = " & rs("id") 
		  set rs2 = db.SQLSelect ("DBBAECYS",sql)
		  if not rs2.eof then
		  documento = rs2("enlace")
		  end if
		  set rs2 = nothing
		%>
		<table>
		<tr>
		<td height="127"><img src="./Noticias/<%=folder%>/<%=imagen_principal%>" alt="<%=titulo%>" width="100" height="100"> </td>
		<td><h2><%=Titulo%></h2>
		</td>
		</tr>
		</table>
		<table border = "0">
		  <tr>
		    <td> 
		      <h4><%=descripcion%></h4>
		</td>
		</tr>
		</table>
		<h3>Autor: <%=Autor%></h3>
		<h3>Fecha: <%=fecha%></h3>
		<a href="ver_noticia.asp?noticia=<%=rs("id")%>&folder=<%=folder%>" ><font size="4" face="Arial, Helvetica, sans-serif" color="#0000CC">Ver Mas...</font></a> 
		<hr>
		<%
		cont = cont + 1
		rs.movenext
		wend	
	end if
end if
set rs = nothing
%>
</body>
</html>
