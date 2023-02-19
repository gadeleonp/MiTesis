<%
dim busqueda
busqueda = Request.QueryString("busqueda")

set rs2 = server.createobject("adodb.recordset")
sql = "select id,nombre from pais"
set rs = db.sqlselect("DBBAECYS",sql)
if not rs.eof then
while rs.eof <> true
sql = "select id,nombre from universidad where pais =  " & rs("id")
set rs2 = db.sqlselect("DBBAECYS", sql)
%>
	var la<%=rs("id")%>=new Array("('---Todas las de <%=rs("nombre")%>---','AA',false,true)" 
	<%
	if rs2.EOF <> true then
	while rs2.eof <> true%> ,"('<%=rs2("nombre")%>','<%=rs2("id")%>')" <% 
	rs2.movenext 
	wend  
	end if %> );
<%
set rs2 = nothing
rs.movenext
wend
end if
set rs = nothing
%>