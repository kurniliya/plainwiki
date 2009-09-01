<%@ Page Language="VB" EnableSessionState=False %>
<%@ Import namespace="ADODB" %>
<script language="VB" runat="Server">
Dim ScriptEngineMajorVersion As Byte
Dim ScriptEngineMinorVersion As Byte


Dim gDaysToKeep As Byte
Dim rs As ADODB.Recordset
Dim vText As Date
Dim v As Vector
Dim q As String


Dim vFSO As Scripting.FileSystemObject
Dim vPagename As String
Dim vInvalidPathCharactersPattern As String

' delete deprecated attachments
Dim vPath As String


Sub adoDMLQuery(ByRef pQuery As String)
	Dim conn As ADODB.Connection
	conn = New ADODB.Connection
	conn.Open(OPENWIKI_DB)
	conn.Execute(pQuery)
	'UPGRADE_NOTE: Object conn may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
	conn = Nothing
End Sub

</script>

<!-- #INCLUDE FILE="../ow/owpreamble.aspx" -->
<!--   // -->
<!-- #INCLUDE FILE="../ow/owconfig_default.aspx" -->
<!--   // -->
<!-- #INCLUDE FILE="../ow/owvector.aspx" -->
<!--   // -->
<!-- #INCLUDE FILE="../ow/owado.aspx" -->
<!--   // -->
<!-- #INCLUDE FILE="../ow/owattach.aspx" -->
<!--   // -->
<!-- #INCLUDE FILE="../ow/owregexp.aspx" -->
<!--   // -->

<%
'
' This script deletes deprecated pages and attachments.
'
%>
<html>
<head>
        <title>Delete deprecated pages and attachments</title>
        <link rel="stylesheet" type="text/css" href="../ow/css/ow.css" />
</head>
<body>
<h2>Delete deprecated pages and attachments</h2>
<p>
<%

' HttpContext.Current.Response.Write("<p>Look in the script!</p>")

' RUN AT YOUR OWN RISK !!!
' ALSO DELETES DEPRECATED ATTACHMENTS
' MAKE SURE YOU'VE SET THE VARIABLE OPENWIKI_DB CORRECTLY IN YOUR CONFIG FILE
'
' COMMENT THE NEXT LINE AND REFRESH THE PAGE
' HttpContext.Current.Response.End

' If a page is marked as deprecated, but was last modified less than
' <gDaysToKeep> days, then keep the page and/or attachment. Otherwise
' delete it.
gDaysToKeep = OPENWIKI_DAYSTOKEEP_DEPRECATED
If (Request.Form("submitted") = "yes") And (Request.Form("password") = gAdminPassword) Then
	q = "SELECT wrv_name, wrv_timestamp, wrv_text FROM openwiki_revisions WHERE wrv_current = 1"
	v = New Vector
	rs = New ADODB.Recordset
	rs.Open(q, OPENWIKI_DB, adOpenForwardOnly)
	Do While Not rs.EOF
		'UPGRADE_NOTE: Date operands have a different behavior in arithmetical operations. Copy this link in your browser for more: ms-its:C:\Soft\Dev\ASP to ASP.NET Migration Assistant\AspToAspNet.chm::/1023.htm
		If IIF(IsDBNull(rs.Fields.Item("wrv_timestamp").Value), Nothing, rs.Fields.Item("wrv_timestamp").Value) < (System.Date.FromOADate(Now().ToOADate - gDaysToKeep)) Then
			vText = IIF(IsDBNull(rs.Fields.Item("wrv_text").Value), Nothing, rs.Fields.Item("wrv_text").Value)
			If Len(CStr(vText)) >= 11 Then
				If Left(CStr(vText), 11) = "#DEPRECATED" Then
					Response.Write("<p>" & IIF(IsDBNull(rs.Fields.Item("wrv_name").Value), Nothing, rs.Fields.Item("wrv_name").Value) & "</p>")
					v.Push("" & IIF(IsDBNull(rs.Fields.Item("wrv_name").Value), Nothing, rs.Fields.Item("wrv_name").Value))
				End If
			End If
		End If
		rs.MoveNext()
	Loop 
	'UPGRADE_NOTE: Object rs may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
	rs = Nothing
	
	vInvalidPathCharactersPattern = "[" & Chr(0) & Chr(9) & Chr(10) & Chr(11) & Chr(12) & Chr(13) & Chr(32) & Chr(38) & "" & Chr(42) & Chr(44) & Chr(58) & Chr(60) & Chr(62) & Chr(63) & Chr(160) & Chr(34) & "]"
	vFSO = New Scripting.FileSystemObject
	
	' delete pages and their attachments
	Do While Not v.IsEmpty
		vPagename = v.Pop
		adoDMLQuery("DELETE FROM openwiki_revisions WHERE wrv_name = '" & Replace(vPagename, "'", "''") & "'")
		adoDMLQuery("DELETE FROM openwiki_attachments_log WHERE ath_wrv_name = '" & Replace(vPagename, "'", "''") & "'")
		adoDMLQuery("DELETE FROM openwiki_attachments WHERE att_wrv_name = '" & Replace(vPagename, "'", "''") & "'")
		
		If Not m(OPENWIKI_UPLOADDIR & vPagename & "/", vInvalidPathCharactersPattern, 0, 1) Then
			If vFSO.FolderExists(Server.MapPath(OPENWIKI_UPLOADDIR & vPagename & "/")) Then
				vFSO.DeleteFolder((Server.MapPath(OPENWIKI_UPLOADDIR & vPagename & "/")))
			End If
		Else
			Response.Write("<p>Folder was not deleted (if exists) due to special characters in a page name. Look in the script!</p>")
		End If
	Loop 
	'UPGRADE_NOTE: Object v may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
	v = Nothing
	
	adoDMLQuery("DELETE FROM openwiki_pages WHERE NOT EXISTS (SELECT 'x' FROM openwiki_revisions WHERE wrv_name = wpg_name)")
	q = "SELECT att_wrv_name, att_filename, att_timestamp FROM openwiki_attachments WHERE att_deprecated = 1"
	v = New Vector
	rs = New ADODB.Recordset
	'rs.Open q, OPENWIKI_DB, adOpenForwardOnly
	rs.Open(q, OPENWIKI_DB, adOpenKeyset, adLockOptimistic, adCmdText)
	Do While Not rs.EOF
		'UPGRADE_NOTE: Date operands have a different behavior in arithmetical operations. Copy this link in your browser for more: ms-its:C:\Soft\Dev\ASP to ASP.NET Migration Assistant\AspToAspNet.chm::/1023.htm
		If IIF(IsDBNull(rs.Fields.Item("att_timestamp").Value), Nothing, rs.Fields.Item("att_timestamp").Value) < (System.Date.FromOADate(Now().ToOADate - gDaysToKeep)) Then
			vPath = HttpContext.Current.Server.MapPath(OPENWIKI_UPLOADDIR & IIF(IsDBNull(rs.Fields.Item("att_wrv_name").Value), Nothing, rs.Fields.Item("att_wrv_name").Value) & "/" & IIF(IsDBNull(rs.Fields.Item("att_filename").Value), Nothing, rs.Fields.Item("att_filename").Value))
			If vFSO.FileExists(vPath) Then
				Response.Write(vPath & "<br />")
				vFSO.DeleteFile((vPath))
				rs.Delete()
			End If
		End If
		rs.MoveNext()
	Loop 
	'UPGRADE_NOTE: Object rs may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
	rs = Nothing
	
	Response.Write("<p />Done!")
Else
	%>
<p>
Click the button to erase deprecated pages and attachments
</p>
<form name="theform" method="POST">
 <input type="hidden" name="submitted" value="yes">
 Admin password (required): <input type="password" name="password">
 <br /><input type="submit" name="doit" value="Delete">
</form>
<%	
End If

%>
</p>
</body>
</html>