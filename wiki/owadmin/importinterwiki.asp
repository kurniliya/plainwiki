<%@ Language=VBScript EnableSessionState=False %>
<% RESPONSE.BUFFER=TRUE %>

<!-- #include file="../ow/owpreamble.asp" //-->
<!-- #include file="../ow/owconfig_default.asp" //-->
<!-- #include file="../ow/owado.asp" //-->

<%
'        // DECLARES AND INITS //
Dim oConn,tConn,oRs,tRs,oSQL,tSQL,FSO,csvFile,aName,aURL,interwiki_Folder,interwiki_Filename,interwiki_Path,ct
Dim Success
'	Jet uses '.' internally as an identifier for table names,
'	such as database.table. Jet eliminates ambiguity by mapping '#'
'	as the delimiter for external files.
Dim interwiki_JetFilename

Set FSO = Server.CreateObject("Scripting.FileSystemObject")
Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.Recordset")
oConn.Open OPENWIKI_DB
Set tConn = Server.CreateObject("ADODB.Connection")
Set tRS = Server.CreateObject("ADODB.Recordset")
interwiki_Folder=Server.MapPath("/cgi-bin") & "\"
interwiki_Filename="interwiki.csv"
interwiki_JetFilename = "interwiki#csv"
interwiki_Path=interwiki_Folder & interwiki_Filename
ct=0
'        // Open the textfile driver
tConn.Open "Driver={Microsoft Text Driver (*.txt; *.csv)};" & _
    "Dbq=" & interwiki_Folder & ";" & _
    "Extensions=asc,csv,tab,txt;" &_
    "Format=CSVDelimited;" &_
    "ColNameHeader = true;"

%>
<html>
<head>
 <title>Import InterWiki from csv Page</title>
 <link rel="stylesheet" type="text/css" href="../ow/css/ow.css" />
</head>
<body>
<p>
<%
'        // Initial report
Response.Write("<h2>InterWiki Import from CSV File</h2><hr />")
If (Request.Form("submitted")="yes") AND (Request.Form("password")=gAdminPassword) then
    Response.Write("Importing from " & interwiki_Path & "...<br />Please wait. working..")
    Response.Flush

    Success=False
    'On Error Resume Next
    If FSO.FileExists(interwiki_Path) then
            '        // Open a recordset of the InterWikis
            tSQL="SELECT DISTINCT * FROM [" & interwiki_JetFilename & "];"
            tRS.Open tSQL, tConn, adOpenForwardOnly
            oConn.BeginTrans()
            If OPENWIKI_DB_SYNTAX = DB_SQLSERVER and OPENWIKI_DB_SYNTAX <> DB_ACCESS then
				oConn.Execute("DELETE openwiki_interwikis;")
			else
	            oConn.Execute("DELETE * FROM [openwiki_interwikis];")
			end if
            '        // Loop through, and write to interwiki_Filename
            Do While Not tRS.EOF
                    aName=tRs.Fields("wik_name")
                    aURL=tRs.Fields("wik_url")
                    oSQL="INSERT INTO [openwiki_interwikis] ([wik_name], [wik_url]) VALUES ('" &_
                    aName & "','" & aURL & "');"
                    oConn.Execute(oSQL)
                    Response.Write(".")
                    Response.Flush
                    ct=ct+1
                    tRs.MoveNext
            Loop
            tRS.Close
            If oConn.Errors.Count=0 then
                    oConn.CommitTrans()
                    Success=True
            Else
            		Response.Write(oConn.Errors.Item(0))
                    oConn.RollbackTrans()
            End If
    End If
    If Success=True then
            '        // Final report
            Response.Write("<br />Success! " & ct & " entries successfully read.")
            Response.Write("<hr /><small>Import Facility by <a href='mailto:openwiking@gmail.com'>OpenWikiNG team</a> &copy;2004 GPL license</small>")
    Else
            Response.Write("<br />Sorry - there was an error!  Quitting.")
            Response.Write("<hr /><small>Import Facility by <a href='mailto:openwiking@gmail.com'>OpenWikiNG team</a> &copy;2004 GPL license</small>")
    End If
Else
%>
<p>
Click the button to export the InterWiki links from database to csv file
</p>
<form name="theform" method="POST">
 <input type="hidden" name="submitted" value="yes">
 Admin password (required): <input type="password" name="password">
 <br /><input type="submit" name="doit" value="Import">
</form>
<%
End If
%>
</p>
</body>
</html>
<%
'        // TIDY UP
set tConn = Nothing
set oConn = Nothing
set FSO = Nothing
set tRs = Nothing
set oRs = Nothing
%>