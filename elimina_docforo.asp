<%@ Language=VBScript %>
<%
dim usuario
usuario = session("id")
if isempty(usuario) then
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<script language="javascript">
  function cerrarventana() {
   window.close()
  }
</script>
<BODY onload="javascript:cerrarventana()">
</BODY>
</HTML>
<%   
end if
dim db
dim resultado
dim objupload
dim documento
dim rs

documento = Request.Form("documento")
tipo = request.form("tipo")
set db = server.CreateObject("SQLComandos.Comandos")

sql = "select foro,archivo from documento where id = "& documento
set rs = db.SQLSelect("DBBAECYS",sql)
if not rs.EOF then
		archivo =  rs("archivo")
		foro = rs("foro")
end if
set rs = nothing
	
	RootFolderName = Server.MapPath(".") & "\"
	Path = RootFolderName & "Foros" &"\" & archivo
	
set objupload = server.CreateObject ("dundas.upload.2")


objupload.FileDelete (Path)

set objupload = nothing

sql = "delete from documento where id = "& documento


resultado = db.SQLOtras("DBBAECYS",sql)
set db = nothing
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<script language="javascript">
  function cerrarventana() {
  if (forma.tipo.value=='2') {
  //alert(forma.tipo.value+'si ingreso')
  forma.action = 'desarrollo_foro.asp'
  forma.submit()
  }
  else
  {
    opener.document.location = "reporte_foros.asp"
  }
	if (window && !window.closed) 
	{
		   window.close()
	 }
   }
</script>
<BODY onload="javascript:cerrarventana()">
<form name="forma" action="" method="get" target="principal">
<input name="foro" type="hidden" value="<%=foro%>">
<input name="tipo" type="hidden" value="<%=tipo%>">
</form>
</BODY>
</HTML>
