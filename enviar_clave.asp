<%@ Language=VBScript %>
<%

On Error Resume Next

dim id 
dim destino
dim password
dim db
dim rs

id = Request.Form("id")
destino = ""
destino = Request.Form("mail")

set db = server.CreateObject("SqlComandos.Comandos")

sql = "select password from usuario where id = " & id

'Response.Write sql
set rs = db.SQLSelect("DBBAECYS",sql)

if not rs.eof then
   password = rs("password")
end if

set rs = nothing
set db = nothing


HTMLBody = " <!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"" "  & _
" ""http://www.w3.org/TR/html4/loose.dtd"" > "& _
"<html>"& _
"<head>" & _
"<title>Untitled Document</title>"& _
"<meta http-equiv=""Content-Type"" content=""text/html""; charset=iso-8859-1"" >"& _
"<style type=""text/css""> "& _
"<!--" & _
".style1 {color: #000066}" & _
".style2 {color: #0033FF}" & _
".style3 {color: #333399}" & _
".style4 {color: #003399}" & _
"-->" & _
"</style> " & _
"</head>" & _
"<body>" & _
"<h1 class=""style1"">Bienvenido a la Asociación de Estudiantes de Ciencias y Sistemas.</h1>" & _
"<h3 class=""style3"">Su contraseña es: </h3><h2 class=""style2"">"&password&"</h2> " & _
"<h3 class=""style4"">Por favor ingrese la contraseña en la pagina de Registro y cámbiela inmediatamente." & _
 "Atte, Web Master AECYS" & _
"</h3>" & _ 
"</body>" & _

"</html>" 


dim oMail
  
  set oMail = server.createobject("CDONTS.NewMail")

  oMail.From = "aecys@ing.usac.edu.gt"
  oMail.To = destino
  oMail.BodyFormat = 0
  oMail.MailFormat = 0
  oMail.Value("Reply-To") = "aecys@ing.usac.edu.gt"
  oMail.Subject = "Nuevo Registro"
  oMail.Body = HtmlBody
  
  oMail.Send 
  set oMail = nothing
'Response.Write  HTMLBody
'Response.End 


'Response.redirect("Notifica_envio.asp?mail="&destino)


If Err.Number <> 0 Then
Response.Write Err.Description
Response.End 
End If


%>
<html>
<script>
  function Redireccion() {
  document.Form.submit()
  return true;
  }
</script>
<body onload="javascript:Redireccion()">
<form name="Form" method="post" action="Notifica_envio.asp">
<input name="mail" value="<%=destino%>" type="hidden">
<input name="id" value="<%=id%>" type="hidden">
</form>
</body>
</html>