<%@ Language=VBScript %>
<%
dim db
dim sql
dim rs
dim stream

On Error Resume Next


function QuitarComa(CadenaQuitar) 
   Cadena = left(CadenaQuitar,len(CadenaQuitar)-1)
   QuitarComa = Cadena
end function

function buscar(CadenaBusqueda,CaracterBusqueda)
	MiPos = Instr(1, CadenaBusqueda, CaracterBusqueda)
	if isnumeric(MiPos) then
	  regreso = MiPos
	else
	  regreso = -1
	end if
	buscar = Regreso
end function

    
set db = server.createObject("SqlComandos.Comandos")

'Esta consulta generá la lista de usuarios que desean recibir correos sobre las noticias que se publiquen


sql = "select mail,nombre,apellido,recibir_mail, tipo_mail from usuario where recibir_mail = 'SI' and tipo_mail like '%N%' "
set rs = db.SqlSelect("dbbaecys",sql)
correo = ""
if not rs.eof then
	while not rs.eof 
	 correo = correo & rs("mail") & ","
	  nombre = rs("nombre")
	  apellido = rs("apellido")
	  recibir_mail = rs("recibir_mail")
	  tipo_mail = rs("tipo_mail")
	  rs.movenext
	wend
end if

set rs = nothing

Correo = QuitarComa(Correo)


'Armar el contendio del correo electrónico

	sql = "select subject,saludo,cuenta,TituloAutorResumen,publicar,TipoLetra,Negrilla from ConfCorreo where Descripcion = 'Noticia'"
	
	set rs2 = db.SqlSelect("DBBAECYS",sql)
	
	if not rs2.eof then
		subject = rs2("subject")
		saludo = rs2("saludo")
		cuenta = rs2("cuenta")
		TituloAutorResumen = rs2("TituloAutorResumen")
		Publicar = rs2("publicar")
		TipoLetra = rs2("TipoLetra")
		Negrilla = rs2("Negrilla")
	end if
	set rs2 = nothing
	
	if Publicar = "1" then
	
	'Armar el contenido de la noticia
	
	Noticia = Request.queryString("Noticia")
	
	sql = "select autor,titulo,resumen from noticia where id = " & Noticia
	'Response.write sql
	
	set rs3 = db.SQLSelect("DBBAECYS",sql)
	
	if not rs3.eof then
	  autor = rs3("autor")
	  titulo = rs3("titulo")
	  resumen = rs3("resumen")
	end if
	set rs3 = nothing
	'Cuerpo del Correo electrónico
	
	BodyMail = "<html><head><title>Correo de Noticias</title><meta http-equiv=Content-Type >" & _
	"</head><body bgcolor=#33CCFF><div align=center ><h1>Correo de Noticias</h1></div>" & _
	"<p>" & saludo & "</p>" 
	if buscar(TituloAutorResumen,"t") > 0 then
	BodyMail = BodyMail & "<h2>" &Titulo & "</h2>"  
	end if
	if buscar(TituloAutorResumen,"a") > 0 then
	BodyMail = BodyMail & "<h4>Autor:  " &Autor & "</h4>"  
	end if	 
	if buscar(TituloAutorResumen,"r") > 0 then
	BodyMail = BodyMail & "<h3>" &Resumen & "</h3>"  
	end if	 	 
	
	BodyMail = BodyMail & "</body></html>"
 	response.write BodyMail 
	end if
	set db = nothing
	
	'''''''''''''''''''''''''''''''''''''''''''''''
	''' COMPONENTE QUE ENVIA CORREOS''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''
	dim oMail
	set oMail = server.CreateObject("CDONTS.NewMail")
	
	oMail.MailFormat = 0
	oMail.BodyFormat = 0
	oMail.Body = BodyMail
	oMail.To = correo
	oMail.Subject = subject
	oMail.Value("Reply-To") = cuenta	
	oMail.From = cuenta
	oMail.Send 



	If Err.Number <> 0 Then
		Response.write ("No se pudo enviar el mensaje, el siguiente error ocurrio: " & Err.Description)
	Else
		Response.write ("El correo html fue enviado exitosamente. ")
	End If
%>
<script language="javascript">
  function Regresar() {
    history.back()
  }
</script>
<div align="center">
<input name="Aceptar" value="Aceptar" type="button" onclick="javascript:Regresar()">
</div>