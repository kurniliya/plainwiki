<%@ Page Language="VB" EnableSessionState=False %>
<script language="VB" runat="Server">
Dim ScriptEngineMinorVersion As Byte
Dim ScriptEngineMajorVersion As Byte


</script>

<!-- #INCLUDE FILE="../ow/owpreamble.aspx" -->
<!--   // -->
<!-- #INCLUDE FILE="../ow/owconfig_default.aspx" -->
<!--   // -->
<html>
<head>
        <title>Clear all Caches</title>
        <link rel="stylesheet" type="text/css" href="../ow/css/ow.css" />
</head>
<body>
<h2>Clear all caches</h2>
<p>
<%
If (Request.Form("submitted") = "yes") And (Request.Form("password") = gAdminPassword) Then
	Application("ow__ow.xsl") = ""
	Application("ow__owrss09.xsl") = ""
	Application("ow__owrss091.xsl") = ""
	Application("ow__owrss10.xsl") = ""
	Application("ow__owrss10aggr.xsl") = ""
	Application("ow__owrss10export.xsl") = ""
	Application("ow__owrssscriptingnews.xsl") = ""
	Application("ow__owsearchrss10export.xsl") = ""
	Application("ow__xmldisplay.xsl") = ""
	
	Response.Write("Caches cleared!")
	
Else
	%>
<p>
Click the button to clear all caches
</p>
<form name="theform" method="POST">
 <input type="hidden" name="submitted" value="yes">
 Admin password (required): <input type="password" name="password">
 <br /><input type="submit" name="doit" value="Clear Caches">
</form>
<%	
End If
%>
</p>
</body>
</html>
