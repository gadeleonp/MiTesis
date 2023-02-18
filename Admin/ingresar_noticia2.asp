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

'Datos de la noticia
dim titulo
dim descripcion
dim fecha
dim autor
dim arc_noticia
dim arc_img1
dim arc_img2
dim noticia
'Datos de la noticia
dim nombrefolder
dim TempFolder
dim maximo

dim objUpload		
dim strMessage		
dim strPath		


 dim db
 set db = server.CreateObject ("SQLComandos.Comandos")

dim rs 
dim resultado
dim sql
sql = "select max(id) Maximo from noticia "

set rs = db.SQLSelect("DBBAECYS",sql)

	if not rs.EOF then
			maximo = rs("Maximo")
		if len(maximo) = 0 then
			maximo = "1"
		end if
	else
		maximo = "1"
	end if
set rs = nothing

set objUpload = Server.CreateObject("Dundas.Upload.2")

objUpload.UseVirtualDir = true


nombrefolder = "News"& maximo 
TempFolder = "/aecys/paecycod/Noticias/" & nombrefolder
fecha = now




objUpload.DirectoryCreate TempFolder,false

objUpload.UseUniqueNames = false
objUpload.Save TempFolder

dim cont
dim objFormItem
dim item				

cont = 1
				for each objFormItem in objUpload.Form
					select case objFormItem
						case "txttitulo"
							titulo = objFormItem.value
						case "txtdescripcion"
							descripcion = objFormItem.value
						case "txtAutor"
							autor = objFormItem.value
					end select
				next


				

	it = 0
	arc_noticia = ""
	arc_img1 = ""
	arc_img2 = ""
dim it	


for each item in objUpload.Files
	'Response.Write item.tagname
		if len(item.tagname) = 7 then
		arc_noticia = objUpload.GetFileName (item.Path)
		If InStr(1,objUpload.Files.Item("txtfile").ContentType,"octet-stream",1)  <> 0 or InStr(1,objUpload.Files.Item("txtfile").ContentType,"image",1)  <> 0 Then		
				Response.Write objUpload.Files.Item("txtfile").Size
				objUpload.FileDelete (TempFolder+"/"+arc_noticia)
				arc_noticia = "" 
		end if
	end if
	
	if len(item.tagname) = 10  then
	arc_img1 = objUpload.GetFileName (item.Path)
		If InStr(1,objUpload.Files.item("txtimagen1").ContentType,"octet-stream",1) <> 0 or  InStr(1,objUpload.Files.Item("txtimagen1").ContentType,"image",1)  = 0 Then				
				objUpload.FileDelete (TempFolder+"/"+arc_img1)
				arc_img1 = "" 
		end if
	end if
	
	if len(item.tagname) = 11 then
	arc_img2 = objUpload.GetFileName (item.Path)
		If InStr(1,objUpload.Files.Item("txtimagen2a").ContentType,"octet-stream",1) <> 0 or InStr(1,objUpload.Files.Item("txtimagen2").ContentType,"image",1)  = 0 Then				
			objUpload.FileDelete (TempFolder+"/"+arc_img2)		
			arc_img2 = "" 
		end if
	end if
next



	if len(arc_noticia) = 0 then
	%>
		<html>
		<head >
		<link href="./iconos/Master.css" rel="stylesheet" type="text/css">
		</head>
	<body >
	<h2>Error: No se ingreso la noticia por no ser un tipo valido.</h2>
	</body>
	</html>
		<%
	else

		Sql = "insert into noticia (titulo,resumen,fecha_publicacion,autor,folder_contenedor,imagen_principal,imagen_secundaria) values ('" & titulo & "','" & descripcion & "','" & fecha  &  "','" & autor & "','"& nombrefolder & "','" & arc_img1 & "','" & arc_img2 & "')"
		'Response.Write sql
		
		resultado = db.SQLOtras("DBBAECYS",sql)
		if (resultado = -1) then
		 ' Response.Write ("Error en la insercion de la noticia : ("+sql+")")
		end if
					
		 
		sql = "select id from noticia where fecha_publicacion = '" & fecha & "'"
		'Response.Write sql
	
		set rs = db.SQLSelect("DBBAECYS",sql)
	
		if not rs.EOF then
		noticia = rs("id")
		set rs = nothing
	
		'Response.Write "noticia = " & noticia &"<br>"
		Sql = "insert into documento_noticia (noticia,correlativo,enlace,imagen_principal,imagen_secundaria) values ("& noticia &",1,'" & arc_noticia & "','" & arc_img1 & "','" & arc_img2 & "')"
		
		resultado = db.SQLOtras("DBBAECYS",sql)
		 if (resultado = -1) then
		   Response.Write ("Error en la inserción del documento")
		 end if
		end if
		

	
	
	
 


'Response.Redirect "confirma_noticia.asp?fecha="&fecha

%>
<html>
<script language="javascript">
 function Redirecciona(Form) {
Form.submit()
 return true 
 }
 
</script>
<body onload="return Redirecciona(redireccion)">
<form name="redireccion" action="confirma_noticia.asp" method="post">
<input id="fecha" name="fecha" type ="hidden" value="<%=fecha%>" >
</form>
</body>
</html>
<%
	end if
	set objUpload = nothing
%>