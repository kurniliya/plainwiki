<%@ Page Language="VB" EnableSessionState=False %>
<%@ Import namespace="ADODB" %>
<script language="VB" runat="Server">
Dim ScriptEngineMinorVersion As Byte
Dim ScriptEngineMajorVersion As Byte
Dim ct As Short

'	// DECLARES AND INITS //
Dim oConn As ADODB.Connection
Dim tConn As ADODB.Connection
Dim oRs As ADODB.Recordset
Dim tRs As ADODB.Recordset
Dim FSO As Scripting.FileSystemObject
Dim interwiki_Path As String
Dim ctFalt As Object
Dim oSQL As String
Dim aName As String
Dim csvFile As Scripting.ITextStream
Dim interwiki_Folder As String
Dim aURL As String
Dim interwiki_Filename As String
Dim tSQL As String
'	Jet uses '.' internally as an identifier for table names,
'	such as database.table. Jet eliminates ambiguity by mapping '#'
'	as the delimiter for external files.
Dim interwiki_JetFilename As String


</script>
<%Response.BUFFER = True%>

<!-- #INCLUDE FILE="../ow/owpreamble.aspx" -->
<!--   // -->
<!-- #INCLUDE FILE="../ow/owconfig_default.aspx" -->
<!--   // -->
<!-- #INCLUDE FILE="../ow/owado.aspx" -->
<!--   // -->

<%FSO = New Scripting.FileSystemObject
oConn = New ADODB.Connection
oRs = New ADODB.Recordset
oConn.Open(OPENWIKI_DB)
tConn = New ADODB.Connection
tRs = New ADODB.Recordset
interwiki_Folder = Server.MapPath("/cgi-bin") & "\"
interwiki_Filename = "interwiki.csv"
interwiki_JetFilename = "interwiki#csv"
interwiki_Path = interwiki_Folder & interwiki_Filename
ct = 0

'	// Set up the CSV file
csvFile = FSO.CreateTextFile(interwiki_Path, True)
csvFile.WriteLine(("""wik_name"",""wik_url"""))
csvFile.Close()
'UPGRADE_NOTE: Object csvFile may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
csvFile = Nothing

'	// Open the textfile driver
tConn.Open("Driver={Microsoft Text Driver (*.txt; *.csv)};" & "Dbq=" & interwiki_Folder & "/;" & "Extensions=asc,csv,tab,txt")

%>
<html>
<head>
	<title>Export InterWiki to csv Page</title>
	<link rel="stylesheet" type="text/css" href="../ow/css/ow.css" />
</head>
<body>
<p>
<%
'	// Initial report
Response.Write("<h2>InterWiki Export to CSV File</h2><hr />")
If (Request.Form("submitted") = "yes") And (Request.Form("password") = gAdminPassword) Then
	Response.Write("Creating " & interwiki_Path & "...<br />Please wait. working..")
	Response.Flush()
	
	'	// Open a recordset of the InterWikis    
	oSQL = "SELECT * FROM [openwiki_interwikis] ORDER BY wik_name ASC;"
	oRs.Open(oSQL, oConn, adOpenForwardOnly)
	
	'	// Loop through, and write to interwiki_Filename
	Do While Not oRs.EOF
		aName = IIF(IsDBNull(oRs.Fields.Item("wik_name").Value), Nothing, oRs.Fields.Item("wik_name").Value)
		aURL = IIF(IsDBNull(oRs.Fields.Item("wik_url").Value), Nothing, oRs.Fields.Item("wik_url").Value)
		tSQL = "INSERT INTO [" & interwiki_JetFilename & "] ([wik_name], [wik_url]) VALUES ('" & aName & "', '" & aURL & "');"
		tConn.Execute(tSQL)
		Response.Write(".")
		Response.Flush()
		ct = ct + 1
		oRs.MoveNext()
	Loop 
	oRs.Close()
	
	
	'	// Final report
	Response.Write("<br />Done! " & ct & " entries successfully written.")
	Response.Write("<br />CSV file: <a href='/cgi-bin/" & interwiki_Filename & "'>" & interwiki_Filename & "</a>")
	Response.Write("<hr /><small>Export Facility by <a href='mailto:openwiking@gmail.com'>OpenWikiNG team</a> &copy;2004 GPL license</small>")
Else
	%>
<p>
Click the button to export the InterWiki links from database to csv file
</p>
<form name="theform" method="POST">
 <input type="hidden" name="submitted" value="yes">
 Admin password (required): <input type="password" name="password">
 <br /><input type="submit" name="doit" value="Export">
</form>
<%	
End If
%>
</p>
</body>
</html>
<%
'	// TIDY UP
'UPGRADE_NOTE: Object tConn may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
tConn = Nothing
'UPGRADE_NOTE: Object oConn may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
oConn = Nothing
'UPGRADE_NOTE: Object FSO may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
FSO = Nothing
'UPGRADE_NOTE: Object tRs may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
tRs = Nothing
'UPGRADE_NOTE: Object oRs may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
oRs = Nothing
%>
