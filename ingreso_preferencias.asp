<%@ Language=VBScript %>

<%

login = Request.form("login")
archivo = Request.querystring("archivo")
ingresado = 0
if len(archivo) > 0 then
   ingresado = 1
end if

dim db

dim resultado
set db = server.CreateObject("SQLComandos.Comandos")
dim rs
%>
<HTML>
<HEAD>
<title>Registro Usuarios</title>
<link href="./theme/Master.css" rel="stylesheet" type="text/css">
</HEAD>
<script language="javascript">
 function CorregirDatos(id) {
   action = "actualizar_registro.asp?id="+id
  form.action = action
   return true    
 }
function prueba() {

	if (preferencias.comentario.value.length() > 5 ) {
	
	
	alert ("No puede ingresar mas de 5 caracteres")
	return false
	
	}
	return true
	}

function Curriculum(Login) 
	{
	  ancho = 700
	  if (screen) {
	  
	  leftPos = (screen.width / 2) - (ancho/2)
	  }
	  catWindow = window.open('ingreso_curriculum.asp?login='+Login,'Curriculum','toolbar=no,menubar=no,location=no,scrollbars=yes,width='+ancho+'+,height=350,left='+leftPos+'+,top=20')
	}	
		
</script>
<BODY onload="javascript:{if(parent.frames[0]&&parent.frames['opciones'].Go)parent.frames['opciones'].Go()}" >
<div align="center">
<%


if ingresado = 0 then
sql = "select u.id,u.nombre,u.apellido,u.direccion,u.telefono,u.mail,u.login,u.carnet,u.colegiado,u.tipo,u.arc_resume,u.sexo,p.nombre nacionalidad from usuario u,pais p where u.login = '"& login & "' and u.pais = p.id"
else
sql = "select u.id,u.nombre,u.apellido,u.direccion,u.telefono,u.mail,u.login,u.carnet,u.colegiado,u.tipo,u.arc_resume,u.sexo,p.nombre nacionalidad from usuario u,pais p where u.id = "& session("id") & " and u.pais = p.id"
end if
set rs = db.SQLSelect("DBBAECYS",sql)
if not rs.eof then
nombre = ""
apellido = ""
direccion = ""
telefono = ""
mail = ""
carnet = ""
colegiado = ""
tipo = ""
curri = ""
sexo = ""
nacionalidad = ""
id = rs("id")

'session("id") = id
login = rs("login")
nombre = rs("nombre")
apellido = rs("apellido")
direccion = rs("direccion")
telefono = rs("telefono")
mail = rs("mail")
carnet = rs("carnet")
colegiado = rs("colegiado")
tipo = rs("tipo")
curri = rs("arc_resume")
sexo = rs("sexo")
nacionalidad = rs("nacionalidad")

%>
  <h1>Datos Personales</h1>
  <table border = "0" cellpadding="2" cellspacing="2">
    <tbody>
      <tr> 
        <th width="140" nowrap class="celdaBgBlueLeft"> <div align="right"><b>Nombre</div></th>
        <th width="389" class="celdaBgBlue05"><div align="left"><%=nombre%>          <%response.write " " %>
          <%=apellido%> </div></th>
      </tr>
      <tr> 
        <th width="140" nowrap class="celdaBgBlueLeft"> <div align="right"><b>Usuario: </b></div></th>
        <th width="389" class="celdaBgBlue05"> 
		
		  <div align="left"><%=login%> 
          </div></th>
      </tr>
      <tr> 
        <th width="140" nowrap class="celdaBgBlueLeft"> <div align="right"><b>Direccion: </b></div></th>
        <th width="389" class="celdaBgBlue05"> 
		  <div align="left"><%=direccion%> 
		            </div>
            </th>

      </tr>
      <tr> 
        <th width="140" nowrap class="celdaBgBlueLeft"> <div align="right"><b>Telefono: </b></div></th>
        <th width="389" class="celdaBgBlue05"> 
		  <div align="left"><%=telefono%> 
          </div></th>
      </tr>
      <tr> 
        <th width="140" nowrap class="celdaBgBlueLeft"> <div align="right"><b>Correo Electronico: </b></div></th>
        <th width="389" class="celdaBgBlue05"><div align="left"><%=mail%>
          </div>
        </td>
      </tr>
      <%
	if cint(tipo) = 1 then
%>
      <tr> 
        <th width="140" nowrap class="celdaBgBlueLeft"> <div align="right"><b>No. Carnet: </b></div></th>
        <th width="389" class="celdaBgBlue05"> <div align="left"><%=carnet%> 
</div></th>
      </tr>
      <%
	else
%>
      <tr> 
        <th width="140" nowrap class="celdaBgBlueLeft"> <div align="right"><b>No. Colegiado: </b></div></th>
        <th width="389" class="celdaBgBlue05"> <div align="left"><%=colegiado%> 
</div></th>
      </tr>
      <%
	end if
end if

%>
      <%
set rs = nothing
if ingresado = 0 then
sql ="select u.nombre,p.nombre pais from universidad u, usuario us, pais p where p.id = u.pais and us.universidad = u.id and us.login = '"& login & "'"
else
sql ="select u.nombre,p.nombre pais from universidad u, usuario us, pais p where p.id = u.pais and us.universidad = u.id and us.id = "& session("id")  
end if

set rs = db.SQLSelect("DBBAECYS",sql)
if not rs.eof then
%>
      <tr> 
        <th nowrap class="celdaBgBlueLeft"><div align="right"><strong>Sexo:</strong></div></th>
        <th class="celdaBgBlue05"> <div align="left">
            <%
		if sexo = "1" then
		response.write "Masculino"
		end if
		if sexo = "2" then
		response.write "Femenino"
		end if
		%>
</div></th>
      </tr>
      <tr> 
        <th nowrap class="celdaBgBlueLeft"><div align="right"><strong>Nacionalidad</strong></div></th>
        <th class="celdaBgBlue05"><div align="left"><%=nacionalidad%></div></th>
      </tr>
      <tr> 
        <th nowrap class="celdaBgBlueLeft"><div align="right"><strong> Universidad </strong></div></th>
        <th class="celdaBgBlue05"> <div align="left"><%=rs("nombre")%> de <%=rs("pais")%></div></th>
      </tr>
      <%
else
%>
      <tr> 
        <th nowrap class="celdaBgBlueLeft"><div align="right"><strong> Universidad </strong></div></th>
        <th class="celdaBgBlue05"> <div align="left"><font color="#0033FF">Informacion no disponible </font></div></th>
      </tr>
      <%
end if
set rs = nothing %>
    </tbody>
  </table>
  <!--
  <h2>Publicar su curriculum.</h2>
  <table border = "0" cellpadding="2" cellspacing="2">
    <tr> 
      <th nowrap class="celdaBgBlueLeft"><div align="right">Curriculum:</div></th>
      <%
	  if len(curri) > 0 then
	  %>
      <th class="celdaBgBlue05"><font face="Times New Roman, Times, serif" color="#000099"><b><%=curri%></b></font></th>
      <%
	  else
	  %>
      <td><font face="Times New Roman, Times, serif"><b><font color="#000099">Aún 
        no ha sido ingresado por el usuario.</font></b></font></td>
      <%
	  end if
	  %>
    </tr>
    <tbody>
      <tr> 
        <th nowrap class="celdaBgBlueLeft"> <div align="right">Archivo: </div></th>
        <th class="celdaBgBlue05"> <a href="javascript:Curriculum('<%=id%>')"><font face="Arial, Helvetica, sans-serif" size="2"><b><font color="#000099"> 
          <%if len(curri) > 0 then %>
          Actualizar Curriculum 
          <% else %>
          Ingresar Curriculm 
          <% end if %>
        </font></b></font></a> </th>
      </tr>
    </tbody>
  </table>
  -->
</div>
	 <div align="center">
	 <form name="form" action="enviar_clave.asp" method="post">
	 <input name = "id" type="hidden" value = "<%=id%>">
	 <table border="0" cellpadding="2" cellspacing="2">
	 <tr>
	 <td>
  <input name ="corregir" type="submit" class="celdaBgBlue05" onclick="return CorregirDatos('<%=id%>')" value="<- Corregir">
  </td>
  <td width="17"></td>
    <td>
  <input name ="Continuar" type="submit" class="celdaBgBlue05" value="Continuar ->" >
</td>
  </tr>
	  </table>
	  <input name="mail" type="hidden" value="<%=mail%>" >
  </form>
  </div>
</BODY>
</HTML>
