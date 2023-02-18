<%@ Language=VBScript %>
<%
	option explicit
  Response.Buffer = False
  Response.Expires = 0
  Response.CacheControl = "Private"
%>
<%
if isempty (session("administrador")) then
 Response.Redirect("Login.asp")
end if
%>
<%
dim objupload
dim TempFolderName
dim FinalFolderName
dim nombrefolder
dim RootFolderName
On Error Resume Next

dim it
dim arc_noticia
dim arc_img1
dim arc_img2

dim ctype_noticia
dim ctype_img1
dim ctype_img2

dim db
dim sql
dim resultados

dim folder,noticia,titulo,descripcion,autor
dim ingresado1,ingresado2,ingresado3

ingresado1 = 0
ingresado2 = 0
ingresado3 = 0

it = 0
arc_notica = ""
arc_img1 = ""
arc_img2 = ""
ctype_noticia = ""
ctype_img1 = ""
ctype_img2 = ""

Set objUpload = Server.CreateObject("Dundas.Upload.2")
objupload.UseVirtualDir = true

nombrefolder = "Transicion"
TempFolderName =  "/aecys/paecycod/Noticias/" & nombrefolder
objupload.DirectoryCreate TempFolderName,false
objupload.UseUniqueNames = false
objupload.Save TempFolderName


	if InStr(1,objupload.Files.item("txtfile").TagName,"txtfile",1) <> 0 then	
	if InStr(1,objUpload.Files.Item("txtfile").ContentType,"octet-stream",1) = 0 then
		If  InStr(1,objUpload.Files.Item("txtfile").ContentType,"image",1)  = 0 Then		
			arc_noticia = objupload.GetFileName (objupload.Files.Item("txtfile").OriginalPath)
			ctype_noticia  = objupload.Files.Item("txtfile").ContentType
	'		Response.Write "NOticia ingresada"
		end if
	end if	
	end if

	

	if InStr(1,objupload.Files.item("txtimagen1").TagName,"txtimagen1",1) <> 0 then	
	if InStr(1,objUpload.Files.Item("txtimagen1").ContentType,"octet-stream",1) = 0 then	
		If InStr(1,objUpload.Files.Item("txtimagen1").ContentType,"image",1)  <> 0 Then				
			arc_img1 = objupload.GetFileName (objupload.Files.Item("txtimagen1").OriginalPath)
			ctype_img1= objupload.Files.Item("txtimagen1").ContentType
	'		Response.Write "Imagen 1 Ingresada"			
		end if
	end if
	end if


	
	if InStr(1,objupload.Files.item("txtimagen2a").TagName,"txtimagen2a",1) <> 0 then	
	if InStr(1,objUpload.Files.item("txtimagen2a").ContentType,"octet-stream",1) = 0 then
		If InStr(1,objUpload.Files.Item("txtimagen2a").ContentType,"image",1)  <> 0 Then				
			arc_img2 = objupload.GetFileName (objupload.Files.Item("txtimagen2a").OriginalPath)
			ctype_img2= objupload.Files.Item("txtimagen2a").ContentType
	'		Response.Write "Imagen 2 ingresada"			
		end if
	end if
	end if

folder = objUpload.Form("folder")
noticia = objupload.Form("noticia")
titulo = objupload.Form("txttitulo")
descripcion = objupload.Form("txtdescripcion")
autor = objupload.Form("txtautor")


'Response.Write len(arc_noticia)
	if len(arc_noticia) > 0 then
			'If InStr(1,ctype_noticia,"text/htm",1) =  0 Then
				FinalFolderName = "/Noticias/" & folder &"/"
				FinalFolderName = FinalFolderName & arc_noticia
				objupload.Files.Item("txtfile").Move FinalFolderName,false
				ingresado1 = 1				
			'else
			'    objupload.Files.Item ("txtfile").Delete()
			'    Response.Write "Archivo Eliminado por no ser html " & arc_noticia & "<br>"
			'    ingresado1 = 0
			'end if
	end if

	
	if len(arc_img1) >  0 then
		If InStr(1,ctype_img1,"image",1) <> 0 Then
			FinalFolderName = RootFolderName & "/Noticias/" & folder &"/"
			FinalFolderName = FinalFolderName & arc_img1
			objupload.Files.Item ("txtimagen1").Move FinalFolderName ,false
			ingresado2 = 1
		else
		    objupload.Files.Item("txtimagen1").Delete()
		    ingresado2 = 0
		    Response.Write "Archivo Eliminado por no ser Imagen " & arc_img1 & "<br>"
		end if
	end if
	
	if len(arc_img2) > 0 then
		If InStr(1,ctype_img2,"image",1) <> 0 Then
			ingresado3 = 1
			FinalFolderName = RootFolderName & "/Noticias/" & folder &"/"
			FinalFolderName = FinalFolderName & arc_img2
			objupload.Files.Item ("txtimagen2a").Move FinalFolderName ,false
		else
			objupload.Files.Item("txtimagen2a").Delete()								
			ingresado3 = 0
			Response.Write "Archivo Eliminado por no ser Imagen " & arc_img2 & "<br>"
		end if
	end if

set db = server.CreateObject ("SQLComandos.Comandos")


Sql = "update noticia set titulo = '" & titulo & "', autor = '" & autor & "', resumen = '"& descripcion & "'"


if ingresado2 = 1 then
	Sql = Sql & ", imagen_principal = '" & arc_img1 & "'"
end if


if ingresado3 = 1 then
	if ingresado2 = 1 then
		Sql = Sql & ","
	end if
Sql = Sql & " imagen_secundaria = '" & arc_img2 & "'"
end if

Sql = Sql & " where id = " & noticia

'Response.Write SQL
'Response.Write "ingresado2 = "& ingresado2 &"<br>"

resultados = db.SQLOtras("DBBAECYS",sql)




		if ingresado1 = 1 then
			sql = "update documento_noticia set "
			Sql = Sql & " enlace = '" & arc_noticia & "'"
			Sql = Sql & " where noticia = " & noticia 
			resultados = db.SQLOtras("DBBAECYS",sql)
		end if




objupload.DirectoryDelete TempFolderName,true



'release resources



Set objUpload = Nothing
%>
<html>
<script language="javascript">
 function Continuar(Forma) {
Forma.submit()
 return true
 }
</script>
<body onload ="return Continuar(redireccion)">
<form name="redireccion" action="actualizar_noticia.asp" method = "post">
  <input name="noticia" type ="hidden" value ="<%=noticia%>">
</form>
</body>
</html>