<%@ Language=VBScript %>
<%
if isempty (session("administrador")) then
 Response.Redirect("Login.asp")
end if
%>
<%
dim db
dim sql
dim rs

set db = server.CreateObject("SqlComandos.Comandos")

 sql = "select subject,saludo,TituloAutorResumen,Publicar,TipoLetra,Negrilla,cuenta from ConfCorreo where descripcion = 'Empresa'"

  set rs = db.SQLSelect("DBBAECYS",Sql)
  if not rs.eof then
  	  tema = rs("subject")
	  saludo = rs("saludo")
	  TituloAutorResumen =  rs("TituloAutorREsumen")
	  publicar = rs("publicar")
	  tipoletra = rs("tipoLetra")
	  Negrilla = rs("Negrilla")
	  cuenta = trim(rs("cuenta"))
  end if
  set rs = nothing
  set db = nothing
  
   
  if len(publicar) > 0 and not isnumeric(publicar) then
    publicar = 2
  end if

  

function buscar(CadenaBusqueda,CaracterBusqueda)
	MiPos = Instr(1, CadenaBusqueda, CaracterBusqueda)
	if isnumeric(MiPos) then
	  regreso = MiPos
	else
	  regreso = -1
	end if
	buscar = Regreso
end function
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" >
<link href="../theme/Master.css" rel="stylesheet" type="text/css">
<script language="javascript">
  function validacion_email(entered, alertbox){
	with (entered){
		apos=value.indexOf("@");
		dotpos=value.lastIndexOf(".");
		lastpos=value.length-1;
		if (apos<1 || dotpos-apos<2 || lastpos-dotpos>3 || lastpos-dotpos<2) {
			if (alertbox) {
				alert(alertbox);
			} 
			select();
			focus();
			return false;
		}
		else {
			return true;
		}
	}
}

function verificar(formaingreso) 
  {
		if(validacion_email(formaingreso.correo,'No es una direccion de correo valida')==false){
				return false;
		}
		 return true
  }
  
 function Cambiar() 
 {
   document.location = 'configuracion_correos.asp'
  }
  
  function Configuracion() 
	{
  <% 
  if publicar = 1 then
  %>
  document.form1.chkenviacorreo.checked = true
  <% 
  end if
 
  if buscar(TituloAutorResumen,"t") > 0 then
  %>
  document.form1.titulo.checked = true
  <%
  end if
  if buscar(TituloAutorResumen,"a")  > 0  then
  %>
  document.form1.autor.checked = true
  <%
  end if
  if buscar(TituloAutorResumen,"r")  > 0 then  
  %>
  document.form1.resumen.checked = true
  <%
  end if
  if isnumeric(TipoLetra) then
   %>
  	document.form1.TipoLetra.value= '<%=TipoLetra%>'
   <%end if
   
   if isnumeric(Negrilla) then
   %>
  	document.form1.Negrilla.checked	= '<%=Negrilla%>'
   <% end if%>		
    }
</script>
</HEAD>
<BODY  onload="javascript:Configuracion()">
<form action="IngConfCorreoNoti.asp" method="post" name="form1" onsubmit="return verificar(this)">
<table border="0" align="center" cellpadding="2" cellspacing="0" class="celdaBgBlue05">
  <tr class="celdaBgBlue">
    <td>&nbsp;</td>
  </tr>
  <tr class="celdaBgBlue05">
    <th nowrap class="celdaBgBlue10"><div align="left">
      <input type="checkbox" name="chkenviacorreo" value="1" tabindex="1">
      Permitir a las empresas enviar correos.</div>
    </th>
  </tr>
  <tr>
    <th nowrap class="celdaBgBlue10" ><div align="left" class="celdaBgBlue05">Tema:
     </div></th>
   </tr>
  <tr>
    <td class="celdaBgBlue10"><div align="left">
      <input name="tema" type="text" size="100" maxlength="250" tabindex="2" value="<%=tema%>">
    </div></td>
    
  </tr>
  <tr>
    <th class="celdaBgBlue10"><div align="left" class="celdaBgBlue05">Saludo:</div></th>
 
  </tr>
  <tr>
    <td class="celdaBgBlue10"><textarea name="saludo" cols="75" rows="3" tabindex="3"><%=Saludo%></textarea></td>
 
  </tr>
    <tr>
    <th class="celdaBgBlue05"><div align="left">Cuenta de Correo para empresas: </div></th>
  </tr>
  <tr>
    <td class="celdaBgBlue10"><div align="left">
      <input type="text" name="correo" size="50" value="<%=cuenta%>" tabindex="7">
    </div></td>
   </tr>
</table>
<br>
<br>
<table border="0" align="center" cellpadding="2" cellspacing="2">
  <tr>
    <td><input name="Aceptar" type="submit" class="celdaBgBlue05" value="Aceptar" tabindex="10" ></td>
    <td width="17"></td>
    <td><input name="Cancelar" type="button" class="celdaBgBlue05" value="Cancelar" tabindex="11" onClick="javascript:Cambiar()"></td>
  </tr>
</table>
<input type="hidden" value="Empresa" name="tipocorreo" >
</form>
</BODY>
</HTML>
