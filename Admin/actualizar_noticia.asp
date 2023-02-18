<%@ Language=VBScript %>
<%
if isempty (session("administrador")) then
 Response.Redirect("Login.asp")
end if
dim db 


Dim noticia
Dim sql
Dim rs
Dim rs2

dim titulo,resumen,enlace,imagen_principal,imagen_secundaria,folder,autor


noticia = ""
noticia = Request.Form("noticia")

set db = server.CreateObject("SQLComandos.Comandos")

if len(noticia > 0) then
	Sql = "select id,autor,titulo,resumen, folder_contenedor,imagen_principal,imagen_secundaria from noticia where id =" & noticia
	
 set rs = db.SQLSelect("DBBAECYS",sql)
  if rs.eof = true then
  else
		titulo = rs("titulo")
		resumen = rs("resumen")
		autor = rs("autor")		
	    folder = rs("folder_contenedor")	
		imagen_principal = rs("imagen_principal")
		imagen_secundaria = rs("imagen_secundaria")	    
	    
	  sql = "select id,enlace,correlativo,imagen_principal,imagen_secundaria from documento_noticia where correlativo = 1 and noticia = " & rs ("id")  
	  
	    
		set rs2 = db.SQLSelect("DBBAECYS",sql)
		if rs2.eof = true then
		else	  
	    enlace = rs2("enlace")

	    end if
	  
	end if
	set rs = nothing
	
end if

%>
<HTML>
<HEAD>
<link href="../theme/Master.css" rel="stylesheet" type="text/css">
<script language="javascript">
function Imagen2(Noticia) {
winName = "nueva"
		features  = "toolbar=no";      // yes|no 
		features += ",location=no";    // yes|no 
		features += ",directories=no"; // yes|no 
		//features += ",status=" + (showStatus?"yes":"no");  // yes|no 
		features += ",menubar=no";     // yes|no 
		//features += ",scrollbars=" + (isViewer?"yes":"no");   // auto|yes|no 
		//features += ",resizable=" + (isViewer?"yes":"no");   // yes|no 
		//features += ",dependent";      // close the parent, close the popup, omit if you want otherwise 
		features += ",height=300";
		features += ",width=300";
		features += ",left=400";
		features += ",top=100";
		//winName = winName.replace(/[^a-z]/gi,"_");
		return window.open('agregar_imagen.asp?Noticia='+Noticia,winName,features);
	} 

	function Imagen(Noticia) 
	{
	  ancho = 700
	  if (screen) {
	  
	  leftPos = (screen.width / 2) - (ancho/2)
	  }
	  catWindow = window.open('agregar_imagen.asp?Noticia='+Noticia,'Imagenes','toolbar=no,menubar=no,location=no,scrollbars=yes,width='+ancho+'+,height=350,left='+leftPos+'+,top=20')
	}	


function CambiaDestino(Opcion) {

if (Opcion == "1") {
actualizar.action = "mantnoticias.asp"
}

if (Opcion == "2") {
actualizar.action = "agregar_documento.asp?noticia=<%=noticia%>&folder=<%=folder%>"
}

if (Opcion == "3") {
actualizar.action = "enviacorreonoticia.asp?noticia=<%=noticia%>&folder=<%=folder%>"
}

return true
}
</script>
</HEAD>
<BODY><br>
<h2 align="center">Actualización de Noticias</h2><br>
<form name="actualizar" enctype="multipart/form-data"  action="actualizar_datosnoticia3.asp" method="post">
	<input name="noticia" type ="hidden" value="<%=noticia%>">
	<input name="folder" type ="hidden" value="<%=folder%>">
  <table width="65%" border="0" align="center" cellpadding="2" cellspacing="2">
    <tr> 
      <th nowrap class="celdaBgBlueLeft"> <div align="center">Titulo:</div></th>
      <td class="celdaBgBlue05"> <div align="left"> 
          <input name="txttitulo" size="60" maxlength="200" value="<%=titulo%>">
      </div></td>
    </tr>
    <tr> 
      <th nowrap class="celdaBgBlueLeft"> <div align="center">Descripcion:</div></th>
      <td class="celdaBgBlue05"> <div align="left">
          <TEXTAREA cols=50 name="txtdescripcion" rows=4><%=resumen%></TEXTAREA>
      </div></td>
    </tr>
    <tr> 
      <th nowrap class="celdaBgBlueLeft"> <div align="center">Autor:</div></th>
      <td class="celdaBgBlue05"> <div align="left"> 
          <input name="txtAutor" size="30" maxlength="100" value="<%=autor%>" >
      </div></td>
    </tr>
	    <tr> 
      <th nowrap class="celdaBgBlueLeft"> <div align="center">Folder</div></th>
      <td class="celdaBgBlue05"> <div align="left"> "/Noticias/<%=folder%>" </div></td>
    </tr>
    <tr> 
      <th nowrap class="celdaBgBlueLeft"> <div align="center">Documento Inicial:</div></th>
      <td class="celdaBgBlue05"> <div align="left"> 
          <table width="92%" border="0" cellpadding="2" cellspacing="2">
            <tr> 
              <td> <div align="left">Actual: </div></td>
              <td> <div align="left"> 
                  <input name="arc_actual" type ="text" value="<%=enlace%>" size="50" readonly >
                </div></td>
            </tr>

            <tr> 
              <td>Cambiar</td>
              <td> <div align="left"> 
                  <input name ="txtfile"  TYPE="file" >
                </div></td>
            </tr>
          </table>
      </div></td>
    </tr>

    <tr> 
      <th nowrap class="celdaBgBlueLeft"> <div align="center">Archivo de Imagen Principal:</div></th>
      <td class="celdaBgBlue05"> <div align="left"> 
          <table width="97%" border="0" cellpadding="2" cellspacing="2">
            <tr> 
              <td>Actual: </td>
              <td> 
                <div align="left">
                  <%
			if len(imagen_principal) = 0 then
			Response.Write "Imagen No Ingresada"
			else
			%>
                  <input name="arc_img1" type ="text" value="<%= imagen_principal%>" size="50" readonly > 
                  <%
			end if
			%>
                </div></td>
            </tr>
            <tr> 
              <td>Cambiar</td>
              <td> <div align="left">
                <input NAME="txtimagen1" TYPE="file" > 
              </div></td>
            </tr>
          </table>
      </div></td>
    </tr>
    <tr> 
      <th nowrap class="celdaBgBlueLeft"> <div align="center">Archivo de Imagen Secundaria:</div></th>
      <td class="celdaBgBlue05"> <div align="left"> 
          <table width="75%" border="0" cellpadding="2" cellspacing="2">
            <tr> 
              <td>Actual:</td>
              <td> 
                <div align="left">
                  <%
			  if len(imagen_secundaria) = 0 then
				Response.Write "Imagen No Ingresada"
				else 
			  %>
                  <input name="arc_img2a" type ="text" value="<%= imagen_secundaria%>" size="50" readonly> 
                  <% end if %>
                </div></td>
            </tr>
            <tr> 
              <td> Cambiar </td>
              <td> <div align="left">
                <input NAME="txtimagen2a" TYPE="file" > 
              </div></td>
            </tr>
          </table>
      </div></td>
    </tr>
    <tr> 
      <th nowrap class="celdaBgBlueLeft"><div align="center">Imagenes:</div></th>
      <td class="celdaBgBlue05"><a href="javascript:Imagen('<%=Noticia%>')">Mantenimiento de Imagenes</a>
      <td> </tr>
  </table>
<br>
<table align = "center">
<tbody>

<tr>
      <td>
		<input name="Aceptar" type="submit" class="celdaBgBlue05" value="Actualizar"></td>
      <td width="17"> <font color="#FFFFFF">----</font> </td>
      <td>
		<input name="Cancelar" type="submit" class="celdaBgBlue05" onclick = "return CambiaDestino('1')" value="Salir..."></td>
      <td width="17"> <font color="#FFFFFF">----</font> </td>
      <td>
		<input name="Agregar" type="submit" class="celdaBgBlue05" onclick = "return CambiaDestino('2')" value="(+) Documento"></td>      
      <td width="17"> <font color="#FFFFFF">----</font> </td>
      <td>
		<input name="mail" type="submit" class="celdaBgBlue05" onclick = "return CambiaDestino('3')" value="Enviar Correo"></td>      
</tbody>
</table>
<br>
 <table width="75%" border="0" align = "center" cellpadding="2" cellspacing="2">
    <tr class="celdaBgBlue">
      <th width="32%">Nombre</th>
      <th width="13%">Actualizar</th>
      <th width="14%">Eliminar</th>
      <th width="18%">Posicion Correlativa</th>
      <th width="23%">Cambiar Posicion</th>
    </tr>
	<%
	sql = "select max(correlativo) maximo from documento_noticia where noticia = " & noticia
	
	set rs = db.SQLSelect("DBBAECYS",sql)
	if rs.eof = true then
	else
	maximo = rs("maximo")
	end if
	set rs = nothing
	sql = "select id,correlativo,enlace from documento_noticia where noticia = " & noticia
	Sql = Sql & " order by correlativo"
	
set rs = db.SQLSelect("DBBAECYS",sql)
  if rs.eof = true then
  else
	cont = 0
	while rs.EOF <> true 
	enlace = rs("enlace")
	id = rs("id")
	correlativo = rs("correlativo")

	if (cont mod 2) = 0 then
	color = "celdaBgBlue05"
	else
	color = "celdaBgBlue10"
	end if
	cont = cont + 1
	%>
	
    <tr class="<%=color%>">
      <td width="32%" height="28">
	  <div align="center"><%=enlace%></div></td>
      <td width="13%" height="28"> 
        <div align="center"><a href="actualizar_documento.asp?documento=<%= id %>&noticia=<%=noticia%>"><IMG alt="Actualizar Documento" border=0 height=25 src="./iconos/refresh.gif" width="25"></a></div>
      </td>
      <td width="14%" height="28"> 
        <div align="center"><a href="eliminar_documento.asp?documento=<%= id %>&noticia=<%=noticia%>"><IMG alt="Eliminar Documento" border=0 height=18 src="./iconos/delete.png" width="18"></a></div>
      </td>
      <td width="18%" height="28">
	  <div align="center">
	  	  <%=correlativo%>
	    </div>
	  </td>
	  
      <td width="23%" height="28"> 
        <table align="center">
      <tr>
        <td>
        <%
        if correlativo = 1 then
        else
        %>
        
        <div align="center"><a href="cambiar_posicion.asp?correlativo=<%= correlativo%>&direccion=1&documento=<%= id%>&noticia=<%=noticia%>"><IMG alt="Mover Documento Arriba" border=0 height=18 src="./iconos/moveprevious1.gif" width="18"></a></div>
        <%
        end if
        %>
        </td>
        <td>
        <%if correlativo = maximo then
        else
        %>
        
        
        <div align="center"><a href="cambiar_posicion.asp?correlativo=<%= correlativo%>&direccion=2&documento=<%= id%>&noticia=<%=noticia%>"><IMG alt="Mover Documento Abajo" border=0 height=18 src="./iconos/movenext1.gif" width="18"></a></div>
        <%end if%>
        </td>    
        
       </tr>
      </table>
      </td>
    </tr>
    <%
    rs.MoveNext 
    wend 
    set rs = nothing
    end if
    
    %>
  </table>
</form>


</BODY>
</HTML>
