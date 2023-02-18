<%@ Language=VBScript %>
<HTML>
<HEAD>
<link href="../theme/Master.css" rel="stylesheet" type="text/css">
<TITLE>INGRESO PAIS</TITLE>
</HEAD>
<script language="javascript">
 function Regresar() {
 window.location ="mant_paises.asp"
 
 }
	function validar(form) 
	{
		if (form.PAIS.value.length == 0 )
		{
		alert("Error: No ingreso el pais")
		return false
		} 
		return true 
	}	
 
</script>
<BODY>
<DIV ALIGN="CENTER">
<FORM NAME="FORM" METHOD="POST" ACTION="ingreso_pais.asp" onsubmit="return validar(this)">
    <TABLE width="252" border="0" cellpadding="2" cellspacing="2">
      <TR> 
        <TD> 
          <div align="center" class="celdaBgBlue">NOMBRE </div>
        </TD>
      </TR>
      <TR> 
        <TD> 
          <div align="center"> 
            <INPUT NAME="PAIS" TYPE="TEXT" size="40">
          </div>
        </TD>
      </TR>
    </TABLE>
<br>
<INPUT NAME="Aceptar" TYPE="SUBMIT" VALUE="Aceptar">
<INPUT NAME="Cancelar" TYPE="button" VALUE="Cancelar" onclick="javascript:Regresar()" >
</FORM>
</DIV>
</BODY>
</HTML>
