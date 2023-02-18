<%
dim rs2
sql = "select id,nombre from pais"
set rs  = db.SQLSelect("DBBAECYS",sql)
if  not rs.eof then
while rs.eof <> true
sql = "select id,nombre from universidad where pais =  " & rs("id")
set rs2 = db.SQLSelect("DBBAECYS",sql)
%>
	var la<%=rs("id")%>=new Array("('---Universidades <%=rs("nombre")%>---','0',false,true)" 
	<%
	if not rs2.EOF then
		while rs2.eof <> true%> ,"('<%=rs2("nombre")%>','<%=rs2("id")%>')" <% 
		rs2.movenext 
		wend  
	end if %> );
<%
set rs2=nothing
rs.movenext
wend
end if
set rs = nothing

sql = "select id,nombre from universidad"
set rs  = db.SQLSelect("DBBAECYS",sql)
if  not rs.eof then
while rs.eof <> true
sql = "select id,nombre from carrera where universidad =  " & rs("id")
set rs2 = db.SQLSelect("DBBAECYS",sql)
%>
	var lb<%=rs("id")%>=new Array("('---Carreras <%=rs("nombre")%>---','0',false,true)" 
	<%
	if not rs2.EOF then
		while rs2.eof <> true%> ,"('<%=rs2("nombre")%>','<%=rs2("id")%>')" <% 
		rs2.movenext 
		wend  
	end if %> );
<%
set rs2=nothing
rs.movenext
wend
end if
set rs = nothing


%>