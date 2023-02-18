<%@ Language=VBScript %>
<%
	option explicit
  Response.Buffer = False
  Response.Expires = 0
  Response.CacheControl = "Private"
%>
<%
dim db
dim rs


dim correlativo
dim folder
dim sql
dim enlace
dim path
dim noticia

set db = server.CreateObject ("SQLComandos.Comandos")
noticia = Request.QueryString ("noticia")
folder = Request.QueryString("folder")
correlativo = ""
correlativo = Request.QueryString ("correlativo")

if len(correlativo) = 0 then
	correlativo = 1
end if

sql = "select id,correlativo,enlace from documento_noticia where correlativo = "& correlativo & " and noticia = " & noticia
set rs = db.SQLSelect ("DBBAECYS",sql)

if not rs.eof then
	enlace = rs("enlace")
	correlativo = ("correlativo")
end if
path = "./Noticias/"
path = path + folder + "/"
path = path + enlace
%>

<html>
<script language="javascript">
function Navegar() 
	{
	  ancho = 200
	  if (screen) {
	  
	  leftPos = (screen.width / 2) - (ancho/2)
	  }
	  catWindow = window.open('<%response.write "navegacion.asp?noticia="&noticia&"&correlativo="& correlativo & "&folder="& folder %>','Imagenes','toolbar=no,menubar=no,location=no,scrollbars=yes,width='+ancho+'+,height=5,left='+leftPos+'+,top=20')
	}
 
 function Redireccion() {
 window.parent.location = "principal.asp?ver=1&dnoticia=<%=noticia%>&correlativo=<%=correlativo%>&folder=<%=folder%>&path=<%=path%>"
 
 

 
 return true
 }
</script>
<body onload="return Redireccion()">

</body>
</html>
