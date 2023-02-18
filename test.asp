<%@ Language=VBScript %>
<%
	option explicit
	Response.Expires = 1
	Response.Buffer = 0
%>
<%
   ' Pagina en la que se muestra un listado de los cursos con sus secciones y el
  ' número de alumnos asignados
  dim objeto, resultado
  dim txtconsulta, curso
  dim ciclo
  
  ' Se crea el objeto
  set objeto = Server.CreateObject("SQLComandos.Comandos")

  txtconsulta = ""
  txtconsulta = txtconsulta & " select id, login,password from administrador "
  
  set resultado = objeto.SQLSelect ("DBBAECYS",txtconsulta)
  if not resultado.eof then ' Imprimimos los resultados
    resultado.movefirst
    response.write("<p class=texto_titulo>Reporte de alumnos asignados<br>Ciclo y A&ntilde;o Actual</p>")
    response.write("<table width=85% align=center>")
    response.write("<tr>")  
    response.write("<th>Carnet</th>")
    response.write("<th>Nombre del curso</th>")
    response.write("</tr>")
    
  for ciclo = 1 to resultado.recordcount
      response.write("<tr>")
      response.write("<td>")
      response.write(resultado.fields("id"))
      response.write("</td>")
      response.write("<td>")
      response.write(resultado.fields("password"))
      response.write("</td>")     
	response.write("<td>")
      response.write(resultado.fields("login"))
      response.write("</td>")     
            
      response.write("</tr>")
      resultado.movenext
    next
    response.write("</table>")
  end if
  dim sql
  dim hola
  sql = "insert into administrador (id,login,password) values (2,'hola','hola')"
  hola =  objeto.SQLOtras ("DBBAECYS",sql)
  if hola = -1 then
    response.write ("Error al insertar")
  end if
%>
