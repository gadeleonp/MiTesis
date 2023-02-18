<%@ Language=VBScript %>

<%session.Abandon%>

<HTML>
<HEAD>
</HEAD>
<script language="javascript">
function Redireccion() {
window.parent.location="principal.asp"
}

</script>
<BODY onload="return Redireccion()">

</BODY>
</HTML>
