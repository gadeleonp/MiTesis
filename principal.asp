<%@ Language=VBScript %>
<%
	option explicit
  Response.Buffer = False
  Response.Expires = 0
  Response.CacheControl = "Private"
%>
<%

dim ver
dim noticia
dim correlativo
dim folder
dim path

ver = ""
noticia = ""
correlativo = ""
folder = ""
path = ""
ver = Request.Item("ver")
noticia = Request.Item("dnoticia")
correlativo = Request.Item("correlativo")
folder = Request.Item("folder")
path = Request.Item("path")

%>
<html>
<head>
<title>Asociacion de Estudiantes de Ciencias y Sistemas</title>
</head>

<frameset rows="100,*" cols="*" frameborder="NO" border="0" framespacing="0">
  <frame name="titulo" scrolling="no" noresize src="titulo.asp" >
  <frameset cols="150,*" frameborder="NO" border="0" framespacing="0">
    <frame name="opciones" noresize scrolling="NO" src="opciones.asp">
    <%if  InStr(1,ver,"1",1) <> 0 Then %>
    <frameset rows = "*,35" frameborder="0" border="No" framespacing="0">
      <frame name="principal" scrolling="yes" noresize src="<%=path%>" >
      <frame name="navegacion"  scrolling="NO" src="navegacion.asp?noticia=<%=noticia%>&correlativo=<%=correlativo%>&folder=<%=folder%>">
    </frameset>
    <%else
  if InStr(1,ver,"0",1) <> 0 Then
  %>
    <frameset rows = "*,0" frameborder="0" border="No" framespacing="0">
      <frame name="principal" src="<%=path%>">
      <frame name="navegacion" scrolling="NO" src="navegacion.asp">
    </frameset>
    <%else%>
    <frameset rows = "*,0" frameborder="0" border="No" framespacing="0">
      <frame name="principal" src="noticias.asp?opcion=1">
      <frame name="navegacion" scrolling="NO" src="navegacion.asp">
    </frameset>
    <%end if%>
    <%end if%>
  </frameset>
</frameset>
<noframes><body bgcolor="#FFFFFF" text="#000000">
</body></noframes>
</html>
