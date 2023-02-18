
<%
dim db
dim rs
set db = server.CreateObject ("SQLComandos.Comandos")
dim correlativo

noticia = Request.QueryString ("noticia")
folder = Request.QueryString("folder")
correlativo = ""
correlativo = Request.QueryString ("correlativo")

if len(correlativo) = 0 then
	correlativo = 1
end if

sql = "select id,correlativo,enlace from documento_noticia where correlativo = "& correlativo & " and noticia = " & noticia

 set rs = db.SQLSelect("DBBAECYS",Sql)
  if not rs.eof then
	enlace = rs("enlace")
	correlativo = rs("correlativo")
end if
set rs = nothing
path = "../Noticias/"
path = path + folder + "/"
path = path + enlace
%>
<html>
<head>
<title><%=path%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<frameset rows="570,50" cols="*" frameborder="NO" border="0" framespacing="0"> 
  <frame name="noticia" scrolling="yes" noresize src="<%=path%>" >
  <frame name="navegacion"  scrolling="no" src="navegacion.asp?noticia=<%=noticia%>&correlativo=<%=correlativo%>&folder=<%=folder%>">
  </frameset>
<noframes><body bgcolor="#FFFFFF" text="#000000">
</body></noframes>
</html>