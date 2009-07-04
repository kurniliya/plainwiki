<script language="VB" runat="Server">

<%--'________________________________________________Function FormatDateISO8601(pTimestamp)--%>
<%--Dim vTemp As Object--%>


Dim gEquation As Object

</script>

'
' ---------------------------------------------------------------------------
' Copyright(c) 2000-2002, Laurens Pit
' All rights reserved.
'
' Redistribution and use in source and binary forms, with or without
' modification, are permitted provided that the following conditions
' are met:
'
'   * Redistributions of source code must retain the above copyright
'     notice, this list of conditions and the following disclaimer.
'   * Redistributions in binary form must reproduce the above
'     copyright notice, this list of conditions and the following
'     disclaimer in the documentation and/or other materials provided
'     with the distribution.
'   * Neither the name of OpenWiki nor the names of its contributors
'     may be used to endorse or promote products derived from this
'     software without specific prior written permission.
'
' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
' "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
' LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS
' FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE
' REGENTS OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT,
' INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING,
' BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
' LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
' CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT
' LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN
' ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
' POSSIBILITY OF SUCH DAMAGE.
'
' ---------------------------------------------------------------------------
'      $Source: /usr/local/cvsroot/openwiki/dist/owbase/ow/owdb.asp,v $
'    $Revision: 1.6 $
'      $Author: pit $
' ---------------------------------------------------------------------------
'

Class OpenWikiNamespace
Private vConn As ADODB.Connection
	Private vRS As ADODB.Recordset
	Private vQuery As Object
	Private vIndexSchemes As Object
	Private vCachedPages As Scripting.Dictionary
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1061.asp'
	Private Sub Class_Initialize_Renamed()
		Dim OPENWIKI_DB As Object
		Dim cCacheXML As Object
		Dim cWikiLinks As Object
		Dim cAllowAttachments As Object
		If OPENWIKI_DB = "" Then
			cAllowAttachments = 0
			cWikiLinks = 0
			cCacheXML = 0
		Else
			vConn = New ADODB.Connection
			vConn.Open(OPENWIKI_DB)
			vRS = New ADODB.Recordset
		End If
		vIndexSchemes = New IndexSchemes
		vCachedPages = New Scripting.Dictionary
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1061.asp'
	Private Sub Class_Terminate_Renamed()
		On Error Resume Next
		vConn.Close()
		'UPGRADE_NOTE: Object vConn may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
		vConn = Nothing
		'UPGRADE_NOTE: Object vRS may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
		vRS = Nothing
		'UPGRADE_NOTE: Object vIndexSchemes may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
		vIndexSchemes = Nothing
		'UPGRADE_NOTE: Object vCachedPages may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
		vCachedPages = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	Sub BeginTrans(ByRef pConn As Object)
		Dim DB_MYSQL As Object
		Dim OPENWIKI_DB_SYNTAX As Object
		If OPENWIKI_DB_SYNTAX <> DB_MYSQL Then
			pConn.BeginTrans()
		End If
	End Sub
	
	Sub CommitTrans(ByRef pConn As Object)
		Dim DB_MYSQL As Object
		Dim OPENWIKI_DB_SYNTAX As Object
		If OPENWIKI_DB_SYNTAX <> DB_MYSQL Then
			pConn.CommitTrans()
		End If
	End Sub
	
	Sub RollbackTrans(ByRef pConn As Object)
		Dim DB_MYSQL As Object
		Dim OPENWIKI_DB_SYNTAX As Object
		If OPENWIKI_DB_SYNTAX <> DB_MYSQL Then
			pConn.RollbackTrans()
		End If
	End Sub
	
	Private Function CreatePageKey(ByRef pPagename As Object, ByRef pRevision As Object, ByRef pIncludeText As Object, ByRef pIncludeAllChangeRecords As Object) As Object
		CreatePageKey = pRevision & "_" & pIncludeText & "_" & pIncludeAllChangeRecords & "_" & pPagename
	End Function
	
	Private Function GetCachedPage(ByRef pPagename As Object, ByRef pRevision As Object, ByRef pIncludeText As Object, ByRef pIncludeAllChangeRecords As Object) As Object
		Dim vKey As Object
		vKey = CreatePageKey(pPagename, pRevision, pIncludeText, pIncludeAllChangeRecords)
		If vCachedPages.Exists(vKey) Then
			GetCachedPage = vCachedPages.Item(vKey)
		Else
			'UPGRADE_NOTE: Object GetCachedPage may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
			GetCachedPage = Nothing
		End If
	End Function
	
	Private Sub SetCachedPage(ByRef pPagename As Object, ByRef pRevision As Object, ByRef pIncludeText As Object, ByRef pIncludeAllChangeRecords As Object, ByRef vPage As Object)
		Dim vKey As Object
		vKey = CreatePageKey(pPagename, pRevision, pIncludeText, pIncludeAllChangeRecords)
		vCachedPages.Add(vKey, vPage)
	End Sub
	
	Public Function GetIndexSchemes() As Object
		GetIndexSchemes = vIndexSchemes
	End Function
	
	Function GetPageAndAttachments(ByRef pPagename As Object, ByRef pRevision As Object, ByRef pIncludeText As Object, ByRef pIncludeAllChangeRecords As Object) As Object
		Dim cAllowAttachments As Object
		Dim vPage As Object
		vPage = GetCachedPage(pPagename, pRevision, pIncludeText, pIncludeAllChangeRecords)
		'UPGRADE_WARNING: TypeName has a new behavior. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1041.asp'
		If TypeName(vPage) = TypeName(Nothing) Then
			vPage = GetPage(pPagename, pRevision, pIncludeText, False)
			If cAllowAttachments Then
				Call GetAttachments(vPage, pRevision, pIncludeAllChangeRecords)
			End If
		ElseIf cAllowAttachments And Not vPage.AttachmentsLoaded Then 
			Call GetAttachments(vPage, pRevision, pIncludeAllChangeRecords)
		End If
		GetPageAndAttachments = vPage
	End Function
	
	Function GetPage(ByRef pPagename As Object, ByRef pRevision As Object, ByRef pIncludeText As Object, ByRef pIncludeAllChangeRecords As Object) As Object
		Dim OPENWIKI_RCNAME As Object
		Dim cWikiLinks As Object
		Dim vPage, vChange As Object
		If cWikiLinks = 0 Then
			GetPage = New WikiPage
			GetPage.AddChange()
			GetPage.Name = "FrontPage"
			GetPage.text = "Please provide a value for {{{OPENWIKI_DB}}} in your owconfig.asp file."
			Exit Function
		End If
		vPage = GetCachedPage(pPagename, pRevision, pIncludeText, pIncludeAllChangeRecords)
		'UPGRADE_WARNING: TypeName has a new behavior. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1041.asp'
		If TypeName(vPage) = TypeName(Nothing) Then
			'Response.Write("LOAD PAGE: " & pPageName & "<br />")
			vPage = New WikiPage
			If pIncludeText Then
				vQuery = "SELECT * "
			Else
				vQuery = "SELECT wpg_name, wpg_changes, wpg_lastminor, wpg_lastmajor, wrv_revision, wrv_status, wrv_timestamp, wrv_minoredit, wrv_by, wrv_byalias, wrv_comment "
			End If
			vQuery = vQuery & " FROM openwiki_pages, openwiki_revisions WHERE wpg_name = '" & Replace(pPagename, "'", "''") & "' AND wrv_name = wpg_name"
			If pRevision > 0 Then
				vQuery = vQuery & " AND wrv_revision = " & pRevision
			ElseIf pIncludeAllChangeRecords Then 
				vQuery = vQuery & " ORDER BY wrv_revision DESC"
			Else
				vQuery = vQuery & " AND wrv_current = 1"
			End If
			
			On Error Resume Next
			vRS.Open(vQuery, vConn, ADODB.CursorTypeEnum.adOpenForwardOnly)
			If Err.Number <> 0 Then
				If Err.Number = -2147467259 Then
					HttpContext.Current.Response.Write("<h2>Error:</h2>")
					HttpContext.Current.Response.Write("Cannot find the data sources or the data sources are locked by another application.")
					HttpContext.Current.Response.Write("Make sure you've set the constant <code><b>OPENWIKI_DB</b></code> correctly in your config file, pointing it to your data sources.<br /><br /><br />")
				Else
					HttpContext.Current.Response.Write(Err.Number & "<br />" & Err.Description)
				End If
				HttpContext.Current.Response.End()
			End If
			On Error GoTo 0
			
			If vRS.EOF Then
				If pRevision = 0 Then
					vPage.Name = pPagename
					vPage.AddChange()
				Else
					' TODO: addMessage("Revision " & pRevision & " not available (showing current version instead)"
					vRS.Close()
					GetPage = GetPage(pPagename, 0, pIncludeText, pIncludeAllChangeRecords)
					Exit Function
				End If
			Else
				vPage.Name = IIF(IsDBNull(vRS.Fields.Item("wpg_name").Value), Nothing, vRS.Fields.Item("wpg_name").Value)
				vPage.Changes = CShort(IIF(IsDBNull(vRS.Fields.Item("wpg_changes").Value), Nothing, vRS.Fields.Item("wpg_changes").Value))
				vPage.LastMinor = CShort(IIF(IsDBNull(vRS.Fields.Item("wpg_lastminor").Value), Nothing, vRS.Fields.Item("wpg_lastminor").Value))
				vPage.LastMajor = CShort(IIF(IsDBNull(vRS.Fields.Item("wpg_lastmajor").Value), Nothing, vRS.Fields.Item("wpg_lastmajor").Value))
				If pIncludeText Then
					vPage.text = IIF(IsDBNull(vRS.Fields.Item("wrv_text").Value), Nothing, vRS.Fields.Item("wrv_text").Value)
				End If
				If CShort(IIF(IsDBNull(vRS.Fields.Item("wpg_lastminor").Value), Nothing, vRS.Fields.Item("wpg_lastminor").Value)) = CShort(IIF(IsDBNull(vRS.Fields.Item("wrv_revision").Value), Nothing, vRS.Fields.Item("wrv_revision").Value)) Then
					' wrv_current = 1
					' vPage.Revision = vRS("wrv_revision") ??? ---> No! Because of the xsl script.
					vPage.Revision = 0
				ElseIf pRevision > 0 Then 
					vPage.Revision = pRevision
				End If
				Do While Not vRS.EOF
					vChange = vPage.AddChange
					vChange.Revision = CShort(IIF(IsDBNull(vRS.Fields.Item("wrv_revision").Value), Nothing, vRS.Fields.Item("wrv_revision").Value))
					vChange.Status = CShort(IIF(IsDBNull(vRS.Fields.Item("wrv_status").Value), Nothing, vRS.Fields.Item("wrv_status").Value))
					vChange.MinorEdit = CShort(IIF(IsDBNull(vRS.Fields.Item("wrv_minoredit").Value), Nothing, vRS.Fields.Item("wrv_minoredit").Value))
					vChange.Timestamp = IIF(IsDBNull(vRS.Fields.Item("wrv_timestamp").Value), Nothing, vRS.Fields.Item("wrv_timestamp").Value)
					vChange.By = IIF(IsDBNull(vRS.Fields.Item("wrv_by").Value), Nothing, vRS.Fields.Item("wrv_by").Value)
					vChange.ByAlias = IIF(IsDBNull(vRS.Fields.Item("wrv_byalias").Value), Nothing, vRS.Fields.Item("wrv_byalias").Value)
					vChange.Comment = IIF(IsDBNull(vRS.Fields.Item("wrv_comment").Value), Nothing, vRS.Fields.Item("wrv_comment").Value)
					vRS.MoveNext()
				Loop 
			End If
			vRS.Close()
			
			' TODO: move this out of this method
			' If this is the RecentChanges page, then force the presence of the
			' <RecentChanges> element in the page.
			If vPage.Name = OPENWIKI_RCNAME Then
				'UPGRADE_NOTE: Global Sub/Function s is not accessible
				vPage.text = s(vPage.text, "\<RecentChanges\>", "<RecentChangesLong>", True, True)
				'UPGRADE_NOTE: Global Sub/Function m is not accessible
				If Not m(vPage.text, "\<RecentChangesLong\>", True, True) Then
					vPage.text = vPage.text & "<RecentChangesLong>"
				End If
			End If
			
			Call SetCachedPage(pPagename, pRevision, pIncludeText, pIncludeAllChangeRecords, vPage)
		End If
		
		GetPage = vPage
	End Function
	
	Function GetPageCount() As Object
		vQuery = "SELECT COUNT(*) FROM openwiki_pages"
		vRS.Open(vQuery, vConn, ADODB.CursorTypeEnum.adOpenForwardOnly)
		GetPageCount = CShort(IIF(IsDBNull(vRS.Fields.Item(0).Value), Nothing, vRS.Fields.Item(0).Value))
		vRS.Close()
	End Function
	
	Function GetRevisionsCount() As Object
		vQuery = "SELECT COUNT(*) FROM openwiki_revisions"
		vRS.Open(vQuery, vConn, ADODB.CursorTypeEnum.adOpenForwardOnly)
		GetRevisionsCount = CShort(IIF(IsDBNull(vRS.Fields.Item(0).Value), Nothing, vRS.Fields.Item(0).Value))
		vRS.Close()
	End Function
	
	Function ToXML(ByRef pXmlStr As Object) As Object
		Dim cEmbeddedMode As Object
		Dim cAllowAttachments As Object
		Dim PrettyWikiLink As Object
		Dim OPENWIKI_FRONTPAGE As Object
		Dim OPENWIKI_TITLE As Object
		Dim OPENWIKI_ICONPATH As Object
		Dim OPENWIKI_IMAGEPATH As Object
		Dim gScriptName As Object
		Dim gServerRoot As Object
		Dim gAction As Object
		Dim OPENWIKI_ENCODING As Object
		Dim OPENWIKI_NAMESPACE As Object
		Dim OPENWIKI_XMLVERSION As Object
		Dim cReadOnly As Object
		Dim cUseRecaptcha As Object
		Dim OPENWIKI_PROTECTEDPAGES As Object
		Dim gPage As Object
		Dim gEditPassword As Object
		Dim vProtection As Object
		If cReadOnly Then
			vProtection = "readonly"
			'UPGRADE_NOTE: Global Sub/Function m is not accessible
		ElseIf gEditPassword <> "" And m(gPage, OPENWIKI_PROTECTEDPAGES, False, False) Then 
			vProtection = "password"
		ElseIf cUseRecaptcha Then 
			vProtection = "captcha"
		Else
			vProtection = "none"
		End If
		'UPGRADE_NOTE: Global Sub/Function PCDATAEncode is not accessible
		'UPGRADE_NOTE: Global Sub/Function CDATAEncode is not accessible
		ToXML = "<ow:wiki version='" & OPENWIKI_XMLVERSION & "' xmlns:ow='" & OPENWIKI_NAMESPACE & "' encoding='" & OPENWIKI_ENCODING & "' mode='" & gAction & "'>" & "<ow:useragent>" & PCDATAEncode(HttpContext.Current.Request.ServerVariables("HTTP_USER_AGENT")) & "</ow:useragent>" & "<ow:location>" & PCDATAEncode(gServerRoot) & "</ow:location>" & "<ow:scriptname>" & PCDATAEncode(gScriptName) & "</ow:scriptname>" & "<ow:imagepath>" & PCDATAEncode(OPENWIKI_IMAGEPATH) & "</ow:imagepath>" & "<ow:iconpath>" & PCDATAEncode(OPENWIKI_ICONPATH) & "</ow:iconpath>" & "<ow:about>" & PCDATAEncode(gServerRoot & gScriptName & "?" & HttpContext.Current.Request.ServerVariables("QUERY_STRING")) & "</ow:about>" & "<ow:protection>" & vProtection & "</ow:protection>" & "<ow:title>" & PCDATAEncode(OPENWIKI_TITLE) & "</ow:title>" & "<ow:frontpage name='" & CDATAEncode(OPENWIKI_FRONTPAGE) & "' href='" & gScriptName & "?" & HttpContext.Current.Server.URLEncode(OPENWIKI_FRONTPAGE) & "'>" & PCDATAEncode(PrettyWikiLink(OPENWIKI_FRONTPAGE)) & "</ow:frontpage>"
		If cEmbeddedMode = 0 Then
			If cAllowAttachments = 1 Then
				ToXML = ToXML & "<ow:allowattachments/>"
			End If
			If CStr(HttpContext.Current.Request("redirect")) <> "" Then
				'UPGRADE_NOTE: Global Sub/Function URLDecode is not accessible
				'UPGRADE_NOTE: Global Sub/Function PCDATAEncode is not accessible
				'UPGRADE_NOTE: Global Sub/Function CDATAEncode is not accessible
				ToXML = ToXML & "<ow:redirectedfrom name='" & CDATAEncode(URLDecode(HttpContext.Current.Request("redirect"))) & "'>" & PCDATAEncode(PrettyWikiLink(URLDecode(HttpContext.Current.Request("redirect")))) & "</ow:redirectedfrom>"
			End If
			'UPGRADE_NOTE: Global Sub/Function GetCookieTrail is not accessible
			'UPGRADE_NOTE: Global Sub/Function getUserPreferences is not accessible
			ToXML = ToXML & getUserPreferences() & GetCookieTrail()
		End If
		ToXML = ToXML & pXmlStr & "</ow:wiki>"
	End Function
	
	Private Function isValidDocument(ByRef pText As Object) As Object
		Dim vXslDoc As Object
		Dim MSXML_VERSION As Object
		Dim Wikify As Object
		On Error Resume Next
		Dim vXmlDoc As MSXML2.FreeThreadedDOMDocument
		Dim vXmlStr As Object
		vXmlStr = "<ow:wiki xmlns:ow='x'>" & Wikify(pText) & "</ow:wiki>"
		If MSXML_VERSION <> 3 Then
'UPGRADE_NOTE: The 'Msxml2.FreeThreadedDOMDocument." & MSXML_VERSION & ".0' object is not registered in the migration machine. Copy this link in your browser for more: ms-its:C:\Soft\Dev\ASP to ASP.NET Migration Assistant\AspToAspNet.chm::/1016.htm
			vXmlDoc = HttpContext.Current.Server.CreateObject("Msxml2.FreeThreadedDOMDocument." & MSXML_VERSION & ".0")
			vXslDoc.ResolveExternals = True
			vXslDoc.setProperty("AllowXsltScript", True)
		Else
			vXmlDoc = New MSXML2.FreeThreadedDOMDocument
		End If
		vXmlDoc.async = False
		If vXmlDoc.loadXML(vXmlStr) Then
			isValidDocument = True
		Else
			isValidDocument = False
			HttpContext.Current.Response.Write("<h1>Error occured</h1>")
			HttpContext.Current.Response.Write("<b>Your input did not validate to well-formed and valid Wiki format.<br />")
			HttpContext.Current.Response.Write("Please go back and correct. The XML output attempt was:</b><br /><br />")
			HttpContext.Current.Response.Write("<pre>" & vbCrLf & HttpContext.Current.Server.HTMLEncode(vXmlStr) & vbCrLf & "</pre>" & vbCrLf)
		End If
	End Function
	
	Function SavePage(ByRef pRevision As Object, ByRef pMinorEdit As Object, ByRef pComment As Object, ByRef pText As Object) As Object
		Dim OPENWIKI_DAYSTOKEEP As Object
		Dim gPage As Object
		Dim OPENWIKI_DB As Object
		Dim GetRemoteHost As Object
		Dim vReplacedTS, vBy, vHost, vRevision, vStatus, vUserAgent, vByAlias, vRevsDeleted As Object
		
		pText = pText & ""
		If Not isValidDocument(pText) Then
			SavePage = False
			HttpContext.Current.Response.End()
		End If
		
		vHost = GetRemoteHost()
		vUserAgent = HttpContext.Current.Request.ServerVariables("HTTP_USER_AGENT")
		'UPGRADE_NOTE: Global Sub/Function GetRemoteUser is not accessible
		vBy = GetRemoteUser()
		If vBy = "" Then
			vBy = vHost
		End If
		'UPGRADE_NOTE: Global Sub/Function GetRemoteAlias is not accessible
		vByAlias = GetRemoteAlias()
		
		Dim conn As ADODB.Connection
		conn = New ADODB.Connection
		conn.Open(OPENWIKI_DB)
		BeginTrans((conn))
		vQuery = "SELECT * FROM openwiki_revisions WHERE wrv_name = '" & Replace(gPage, "'", "''") & "' AND wrv_current = 1"
		vRS.Open(vQuery, conn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
		If vRS.EOF Then
			If Trim(pText) = "" Then
				RollbackTrans((conn))
				conn.Close()
				'UPGRADE_NOTE: Object conn may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
				conn = Nothing
				SavePage = True
				Exit Function
			End If
			vRevision = 1
			vStatus = 1 ' new
		ElseIf IIF(IsDBNull(vRS.Fields.Item("wrv_text").Value), Nothing, vRS.Fields.Item("wrv_text").Value) = pText Then 
			RollbackTrans((conn))
			conn.Close()
			'UPGRADE_NOTE: Object conn may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
			conn = Nothing
			SavePage = True
			Exit Function
		Else
			If (CShort(IIF(IsDBNull(vRS.Fields.Item("wrv_revision").Value), Nothing, vRS.Fields.Item("wrv_revision").Value)) <> (pRevision - 1)) Then
				If ((IIF(IsDBNull(vRS.Fields.Item("wrv_by").Value), Nothing, vRS.Fields.Item("wrv_by").Value) <> vBy) Or (IIF(IsDBNull(vRS.Fields.Item("wrv_host").Value), Nothing, vRS.Fields.Item("wrv_host").Value) <> vHost) Or (IIF(IsDBNull(vRS.Fields.Item("wrv_agent").Value), Nothing, vRS.Fields.Item("wrv_agent").Value) <> vUserAgent)) Then
					RollbackTrans((conn))
					conn.Close()
					'UPGRADE_NOTE: Object conn may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
					conn = Nothing
					SavePage = False
					Exit Function
				End If
			End If
			vRevision = CShort(IIF(IsDBNull(vRS.Fields.Item("wrv_revision").Value), Nothing, vRS.Fields.Item("wrv_revision").Value)) + 1
			If ((IIF(IsDBNull(vRS.Fields.Item("wrv_by").Value), Nothing, vRS.Fields.Item("wrv_by").Value) = vBy) And (IIF(IsDBNull(vRS.Fields.Item("wrv_host").Value), Nothing, vRS.Fields.Item("wrv_host").Value) = vHost) And (IIF(IsDBNull(vRS.Fields.Item("wrv_agent").Value), Nothing, vRS.Fields.Item("wrv_agent").Value) = vUserAgent)) Then
				vStatus = CShort(IIF(IsDBNull(vRS.Fields.Item("wrv_status").Value), Nothing, vRS.Fields.Item("wrv_status").Value))
			Else
				vStatus = 2 ' updated
			End If
		End If
		
		If InStr(pText, "#DEPRECATED") = 1 Then
			vStatus = 3 ' deleted
		ElseIf vStatus = 3 Then 
			vStatus = 2 ' updated
		End If
		
		If vRS.EOF Then
			vQuery = "INSERT INTO openwiki_pages (wpg_name, wpg_lastminor, wpg_changes, wpg_lastmajor) VALUES " & "('" & Replace(gPage, "'", "''") & "'," & vRevision & " ,1 ," & vRevision & ")"
			conn.Execute(vQuery)
		Else
			vQuery = "UPDATE openwiki_pages " & "SET wpg_changes = wpg_changes + 1" & ",   wpg_lastminor = " & vRevision
			If pMinorEdit = 0 Then
				vQuery = vQuery & ", wpg_lastmajor = " & vRevision
			End If
			vQuery = vQuery & " WHERE wpg_name = '" & Replace(gPage, "'", "''") & "'"
			conn.Execute(vQuery)
			
			vQuery = "UPDATE openwiki_revisions SET wrv_current = 0 WHERE wrv_name = '" & Replace(gPage, "'", "''") & "' AND wrv_current = 1"
			conn.Execute(vQuery)
		End If
		vRS.Close()
		
		vRS.Open("openwiki_revisions", conn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdTable)
		vRS.AddNew()
		vRS.Fields("wrv_name").Value = gPage
		vRS.Fields("wrv_revision").Value = vRevision
		vRS.Fields("wrv_current").Value = 1
		vRS.Fields("wrv_status").Value = vStatus
		vRS.Fields("wrv_timestamp").Value = Now
		vRS.Fields("wrv_minoredit").Value = pMinorEdit
		vRS.Fields("wrv_host").Value = vHost
		vRS.Fields("wrv_agent").Value = vUserAgent
		vRS.Fields("wrv_by").Value = vBy
		vRS.Fields("wrv_byalias").Value = vByAlias
		vRS.Fields("wrv_comment").Value = pComment
		vRS.Fields("wrv_text").Value = pText
		vRS.Update()
		vRS.Close()
		
		' delete old revisions
		vQuery = "SELECT wrv_revision, wrv_timestamp FROM openwiki_revisions WHERE wrv_name = '" & Replace(gPage, "'", "''") & "' ORDER BY wrv_revision DESC"
		vRS.Open(vQuery, conn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
		If Not vRS.EOF Then
			' this is the current revision
			vRS.MoveNext()
			If Not vRS.EOF Then
				vReplacedTS = IIF(IsDBNull(vRS.Fields.Item("wrv_timestamp").Value), Nothing, vRS.Fields.Item("wrv_timestamp").Value)
				' keep at least one old revision
				vRS.MoveNext()
				Do While Not vRS.EOF
					' check the timestamp of revision that replaced this revision
					'UPGRADE_NOTE: Date operands have a different behavior in arithmetical operations
					If vReplacedTS < (System.Date.FromOADate(Now.ToOADate - CDate(OPENWIKI_DAYSTOKEEP).ToOADate)) Then
						vQuery = "DELETE FROM openwiki_revisions WHERE wrv_name = '" & Replace(gPage, "'", "''") & "' AND wrv_revision <= " & CShort(IIF(IsDBNull(vRS.Fields.Item("wrv_revision").Value), Nothing, vRS.Fields.Item("wrv_revision").Value))
						conn.Execute(vQuery)
						vRS.Close()
						vQuery = "SELECT COUNT(*) FROM openwiki_revisions WHERE wrv_name = '" & Replace(gPage, "'", "''") & "'"
						vRS.Open(vQuery, conn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
						vQuery = "UPDATE openwiki_pages SET wpg_changes = " & CShort(IIF(IsDBNull(vRS.Fields.Item(0).Value), Nothing, vRS.Fields.Item(0).Value)) & " WHERE wpg_name = '" & Replace(gPage, "'", "''") & "'"
						conn.Execute(vQuery)
						Exit Do
					Else
						vReplacedTS = IIF(IsDBNull(vRS.Fields.Item("wrv_timestamp").Value), Nothing, vRS.Fields.Item("wrv_timestamp").Value)
					End If
					vRS.MoveNext()
				Loop 
			End If
		End If
		vRS.Close()
		
		' throw out the bath and the bathwater. TODO: keep the bath
		ClearDocumentCache((conn))
		
		CommitTrans((conn))
		conn.Close()
		
		'UPGRADE_NOTE: Object conn may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
		conn = Nothing
		
		SavePage = True
	End Function
	
	
	' returns the name of the file as you should save it
	' pStatus : 0 = normal, 1 = hidden, 2 = deprecated
	Function SaveAttachmentMetaData(ByRef pFilename As Object, ByRef pFilesize As Object, ByRef pAddLink As Object, ByRef pHidden As Object, ByRef pComment As Object) As Object
		Dim DB_MYSQL As Object
		Dim OPENWIKI_DB_SYNTAX As Object
		Dim gPage As Object
		Dim GetRemoteHost As Object
		Dim vFilename, vPageRevision, vBy, vHost, vUserAgent, vByAlias, vFileRevision, vPos As Object
		
		pFilename = Replace(pFilename, " ", "_")
		
		If pHidden = "" Then
			pHidden = 0
		End If
		
		vHost = GetRemoteHost()
		vUserAgent = HttpContext.Current.Request.ServerVariables("HTTP_USER_AGENT")
		'UPGRADE_NOTE: Global Sub/Function GetRemoteUser is not accessible
		vBy = GetRemoteUser()
		If vBy = "" Then
			vBy = vHost
		End If
		'UPGRADE_NOTE: Global Sub/Function GetRemoteAlias is not accessible
		vByAlias = GetRemoteAlias()
		
		vQuery = "SELECT wpg_lastminor FROM openwiki_pages WHERE wpg_name = '" & Replace(gPage, "'", "''") & "'"
		vRS.Open(vQuery, vConn, ADODB.CursorTypeEnum.adOpenForwardOnly)
		If vRS.EOF Then
			vPageRevision = 1 ' page doesn't exist yet
		Else
			vPageRevision = CShort(IIF(IsDBNull(vRS.Fields.Item(0).Value), Nothing, vRS.Fields.Item(0).Value))
		End If
		vRS.Close()
		vQuery = "SELECT MAX(att_revision) FROM openwiki_attachments WHERE att_wrv_name = '" & Replace(gPage, "'", "''") & "' AND att_name = '" & Replace(pFilename, "'", "''") & "'"
		vRS.Open(vQuery, vConn, ADODB.CursorTypeEnum.adOpenForwardOnly)
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1049.asp'
		If IsDbNull(IIF(IsDBNull(vRS.Fields.Item(0).Value), Nothing, vRS.Fields.Item(0).Value)) Then
			vFileRevision = 1
		Else
			vFileRevision = CShort(IIF(IsDBNull(vRS.Fields.Item(0).Value), Nothing, vRS.Fields.Item(0).Value)) + 1
		End If
		vRS.Close()
		
		vPos = InStrRev(pFilename, ".")
		If vPos > 0 Then
			vFilename = Left(pFilename, vPos - 1) & "-" & vFileRevision & Mid(pFilename, vPos)
		Else
			vFilename = pFilename & "-" & vFileRevision
		End If
		vFilename = SafeFileName(vFilename)
		
		BeginTrans((vConn))
		vRS.Open("openwiki_attachments", vConn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdTable)
		vRS.AddNew()
		vRS.Fields("att_wrv_name").Value = gPage
		vRS.Fields("att_wrv_revision").Value = vPageRevision
		vRS.Fields("att_name").Value = pFilename
		vRS.Fields("att_revision").Value = vFileRevision
		vRS.Fields("att_hidden").Value = pHidden
		vRS.Fields("att_deprecated").Value = 0
		vRS.Fields("att_filename").Value = vFilename
		vRS.Fields("att_timestamp").Value = Now
		vRS.Fields("att_filesize").Value = pFilesize
		vRS.Fields("att_host").Value = vHost
		vRS.Fields("att_agent").Value = vUserAgent
		vRS.Fields("att_by").Value = vBy
		vRS.Fields("att_byalias").Value = vByAlias
		vRS.Fields("att_comment").Value = pComment
		vRS.Update()
		vRS.Close()
		
		Call SaveAttachmentLog(vConn, pFilename, vFileRevision, "uploaded")
		
		Call ClearDocumentCache(vConn)
		'Call ClearDocumentCache2(vConn, gPage)
		
		If pAddLink <> "" Then
			If OPENWIKI_DB_SYNTAX = DB_MYSQL Then
				vQuery = "UPDATE openwiki_revisions SET wrv_text = CONCAT(wrv_text, '" & vbCrLf & vbCrLf & "  * " & Replace(pFilename, "'", "''") & "') WHERE wrv_name = '" & Replace(gPage, "'", "''") & "' AND wrv_current = 1"
				vConn.Execute(vQuery)
			Else
				vQuery = "SELECT wrv_text FROM openwiki_revisions WHERE wrv_name = '" & Replace(gPage, "'", "''") & "' AND wrv_current = 1"
				vRS.Open(vQuery, vConn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
				If Not vRS.EOF Then
					vRS.Fields("wrv_text").Value = IIF(IsDBNull(vRS.Fields.Item("wrv_text").Value), Nothing, vRS.Fields.Item("wrv_text").Value) & vbCrLf & vbCrLf & "  * " & pFilename
					vRS.Update()
				End If
				vRS.Close()
			End If
		End If
		
		CommitTrans((vConn))
		
		SaveAttachmentMetaData = vFilename
	End Function
	
	
	Function HideAttachmentMetaData(ByRef pName As Object, ByRef pRevision As Object, ByRef pHide As Object) As Object
		Dim gPage As Object
		BeginTrans((vConn))
		vConn.Execute("UPDATE openwiki_attachments SET att_hidden = " & pHide & " WHERE att_wrv_name = '" & Replace(gPage, "'", "''") & "' AND att_name = '" & Replace(pName, "'", "''") & "' AND att_revision = " & pRevision)
		If pHide = 1 Then
			Call SaveAttachmentLog(vConn, pName, pRevision, "hidden")
		Else
			Call SaveAttachmentLog(vConn, pName, pRevision, "made visible")
		End If
		Call ClearDocumentCache(vConn)
		'Call ClearDocumentCache2(vConn, gPage)
		CommitTrans((vConn))
	End Function
	
	
	Function TrashAttachmentMetaData(ByRef pName As Object, ByRef pRevision As Object, ByRef pTrash As Object) As Object
		Dim gPage As Object
		BeginTrans((vConn))
		vConn.Execute("UPDATE openwiki_attachments SET att_deprecated = " & pTrash & " WHERE att_wrv_name = '" & Replace(gPage, "'", "''") & "' AND att_name = '" & Replace(pName, "'", "''") & "'")
		If pTrash = 1 Then
			Call SaveAttachmentLog(vConn, pName, pRevision, "deprecated")
		Else
			Call SaveAttachmentLog(vConn, pName, pRevision, "restored")
		End If
		Call ClearDocumentCache(vConn)
		'Call ClearDocumentCache2(vConn, gPage)
		CommitTrans((vConn))
	End Function
	
	
	Sub SaveAttachmentLog(ByRef pConn As Object, ByRef pName As Object, ByRef pFileRevision As Object, ByRef pAction As Object)
		Dim gPage As Object
		Dim GetRemoteHost As Object
		Dim vBy, vHost, vUserAgent, vByAlias As Object
		Dim pPagename, pPageRevision As Object
		
		vHost = GetRemoteHost()
		vUserAgent = HttpContext.Current.Request.ServerVariables("HTTP_USER_AGENT")
		'UPGRADE_NOTE: Global Sub/Function GetRemoteUser is not accessible
		vBy = GetRemoteUser()
		If vBy = "" Then
			vBy = vHost
		End If
		'UPGRADE_NOTE: Global Sub/Function GetRemoteAlias is not accessible
		vByAlias = GetRemoteAlias()
		
		vQuery = "SELECT att_wrv_name, att_wrv_revision FROM openwiki_attachments WHERE att_wrv_name = '" & Replace(gPage, "'", "''") & "' AND att_name = '" & Replace(pName, "'", "''") & "' AND att_revision = " & pFileRevision
		vRS.Open(vQuery, pConn, ADODB.CursorTypeEnum.adOpenForwardOnly)
		If vRS.EOF Then
			vRS.Close()
			Exit Sub
		End If
		pPagename = IIF(IsDBNull(vRS.Fields.Item("att_wrv_name").Value), Nothing, vRS.Fields.Item("att_wrv_name").Value)
		pPageRevision = IIF(IsDBNull(vRS.Fields.Item("att_wrv_revision").Value), Nothing, vRS.Fields.Item("att_wrv_revision").Value)
		vRS.Close()
		
		vQuery = "SELECT wrv_timestamp FROM openwiki_revisions WHERE wrv_name = '" & Replace(pPagename, "'", "''") & "' AND wrv_revision = " & pPageRevision
		vRS.Open(vQuery, pConn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
		If vRS.EOF Then
			vRS.Close()
			Exit Sub
		End If
		vRS.Fields("wrv_timestamp").Value = Now
		vRS.Update()
		vRS.Close()
		
		vRS.Open("openwiki_attachments_log", pConn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdTable)
		vRS.AddNew()
		vRS.Fields("ath_wrv_name").Value = pPagename
		vRS.Fields("ath_wrv_revision").Value = pPageRevision
		vRS.Fields("ath_name").Value = pName
		vRS.Fields("ath_revision").Value = pFileRevision
		vRS.Fields("ath_timestamp").Value = Now
		vRS.Fields("ath_agent").Value = vUserAgent
		vRS.Fields("ath_by").Value = vBy
		vRS.Fields("ath_byalias").Value = vByAlias
		vRS.Fields("ath_action").Value = pAction
		vRS.Update()
		vRS.Close()
	End Sub
	
	
	' Convert the filename to a filename with an extension that is safe
	' to be served by the webserver.
	Function SafeFileName(ByRef pFilename As Object) As Object
		Dim gNotAcceptedExtensions As Object
		Dim gDocExtensions As Object
		Dim vPos, vExtension As Object
		SafeFileName = pFilename
		vPos = InStrRev(pFilename, ".")
		If vPos > 0 Then
			vExtension = Mid(pFilename, vPos + 1)
			If gNotAcceptedExtensions = "" Then
				' accept nothing, except the ones enumerated in gDocExtensions
				If Not InStr("|" & gDocExtensions & "|", "|" & vExtension & "|") > 0 Then
					SafeFileName = SafeFileName & ".safe"
				End If
			Else
				' accept everything, except the ones enumerated in gNotAcceptedExtensions
				If InStr("|" & gNotAcceptedExtensions & "|", "|" & vExtension & "|") > 0 Then
					SafeFileName = SafeFileName & ".safe"
				End If
			End If
		End If
	End Function
	
	
	Sub GetAttachments(ByRef pPage As Object, ByRef pRevision As Object, ByRef pIncludeAllChangeRecords As Object)
		Dim vAttachment, vMaxRevision As Object
		If pIncludeAllChangeRecords Then
			' show all file revisions
			vQuery = "SELECT att_name, att_revision, att_hidden, att_deprecated, att_filename, att_timestamp, att_filesize, att_by, att_byalias, att_comment" & " FROM openwiki_attachments" & " WHERE att_wrv_name = '" & Replace(pPage.Name, "'", "''") & "'" & " AND   att_name = '" & Replace(CStr(HttpContext.Current.Request("file")), "'", "''") & "'" & " ORDER BY att_revision DESC"
		Else
			' show last file revision relative to page revision
			vQuery = "SELECT MAX(att_wrv_revision) FROM openwiki_attachments WHERE att_wrv_name = '" & Replace(pPage.Name, "'", "''") & "'"
			If pRevision > 0 Then
				vQuery = vQuery & " AND att_wrv_revision <= " & pRevision
			End If
			vRS.Open(vQuery, vConn, ADODB.CursorTypeEnum.adOpenForwardOnly)
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1049.asp'
			If IsDbNull(IIF(IsDBNull(vRS.Fields.Item(0).Value), Nothing, vRS.Fields.Item(0).Value)) Then
				vMaxRevision = 0
			Else
				vMaxRevision = CShort(IIF(IsDBNull(vRS.Fields.Item(0).Value), Nothing, vRS.Fields.Item(0).Value))
			End If
			vRS.Close()
			vQuery = "SELECT att_name, att_revision, att_hidden, att_deprecated, att_filename, att_timestamp, att_filesize, att_by, att_byalias, att_comment" & " FROM openwiki_attachments" & " WHERE att_wrv_name = '" & Replace(pPage.Name, "'", "''") & "'" & " AND   att_wrv_revision <= " & vMaxRevision & " ORDER BY att_name ASC, att_revision DESC"
		End If
		vRS.Open(vQuery, vConn, ADODB.CursorTypeEnum.adOpenForwardOnly)
		Do While Not vRS.EOF
			vAttachment = New Attachment
			vAttachment.Name = IIF(IsDBNull(vRS.Fields.Item("att_name").Value), Nothing, vRS.Fields.Item("att_name").Value)
			vAttachment.Revision = CShort(IIF(IsDBNull(vRS.Fields.Item("att_revision").Value), Nothing, vRS.Fields.Item("att_revision").Value))
			vAttachment.Hidden = CShort(IIF(IsDBNull(vRS.Fields.Item("att_hidden").Value), Nothing, vRS.Fields.Item("att_hidden").Value))
			vAttachment.Deprecated = CShort(IIF(IsDBNull(vRS.Fields.Item("att_deprecated").Value), Nothing, vRS.Fields.Item("att_deprecated").Value))
			vAttachment.Filename = IIF(IsDBNull(vRS.Fields.Item("att_filename").Value), Nothing, vRS.Fields.Item("att_filename").Value)
			vAttachment.Timestamp = IIF(IsDBNull(vRS.Fields.Item("att_timestamp").Value), Nothing, vRS.Fields.Item("att_timestamp").Value)
			vAttachment.Filesize = CInt(IIF(IsDBNull(vRS.Fields.Item("att_filesize").Value), Nothing, vRS.Fields.Item("att_filesize").Value))
			vAttachment.By = IIF(IsDBNull(vRS.Fields.Item("att_by").Value), Nothing, vRS.Fields.Item("att_by").Value)
			vAttachment.ByAlias = IIF(IsDBNull(vRS.Fields.Item("att_byalias").Value), Nothing, vRS.Fields.Item("att_byalias").Value)
			vAttachment.Comment = IIF(IsDBNull(vRS.Fields.Item("att_comment").Value), Nothing, vRS.Fields.Item("att_comment").Value)
			Call pPage.AddAttachment(vAttachment, Not pIncludeAllChangeRecords)
			vRS.MoveNext()
		Loop 
		vRS.Close()
		pPage.AttachmentsLoaded = True
	End Sub
	
	
	' pFilter --> 0=All, 1=NoMinorEdit, 2=OnlyMinorEdit
	Function TitleSearch(ByRef pPattern As Object, ByRef pDays As Object, ByRef pFilter As Object, ByRef pOrderBy As Object, ByRef pIncludeAttachmentChanges As Object) As Object
		Dim cAllowAttachments As Object
		Dim DB_ORACLE As Object
		Dim OPENWIKI_DB_SYNTAX As Object
		Dim vCurPage, vPage, vRegEx, vTitle, vList, vChange, vAttachmentChange As Object
		vList = New Vector
		vRegEx = New RegExp
		vRegEx.IgnoreCase = True
		vRegEx.Global = True
		'UPGRADE_NOTE: Global Sub/Function EscapePattern is not accessible
		vRegEx.Pattern = EscapePattern(pPattern)
		vQuery = "SELECT wpg_name, wpg_changes, wrv_revision, wrv_status, wrv_timestamp, wrv_minoredit, wrv_by, wrv_byalias, wrv_comment "
		If cAllowAttachments And pIncludeAttachmentChanges Then
			vQuery = vQuery & ", ath_name, ath_revision, ath_timestamp, ath_by, ath_byalias, ath_action "
			If OPENWIKI_DB_SYNTAX = DB_ORACLE Then
				vQuery = vQuery & " FROM   openwiki_pages, openwiki_revisions, openwiki_attachments_log " & " WHERE  wpg_name = wrv_name " & " AND    wrv_name = ath_wrv_name (+) " & " AND    wrv_revision = ath_wrv_revision (+)"
			Else
				vQuery = vQuery & " FROM (openwiki_pages LEFT JOIN openwiki_revisions ON openwiki_pages.wpg_name = openwiki_revisions.wrv_name) LEFT JOIN openwiki_attachments_log ON (openwiki_revisions.wrv_name = openwiki_attachments_log.ath_wrv_name) AND (openwiki_revisions.wrv_revision = openwiki_attachments_log.ath_wrv_revision) WHERE 1 = 1 "
			End If
		Else
			vQuery = vQuery & "FROM openwiki_pages, openwiki_revisions " & "WHERE wrv_name = wpg_name "
		End If
		If pDays > 0 Then
			' is there a database independent way to test the current date?
			'vQuery = vQuery & " AND wpg_timestamp >
		End If
		If pFilter = 0 Then
			vQuery = vQuery & " AND wpg_lastminor = wrv_revision"
		ElseIf pFilter = 1 Then 
			vQuery = vQuery & " AND wpg_lastmajor = wrv_revision"
		ElseIf pFilter = 2 Then 
			vQuery = vQuery & " AND wpg_lastminor = wrv_revision AND wrv_minoredit = 1"
		End If
		If pOrderBy = 1 Then
			vQuery = vQuery & " ORDER BY wrv_timestamp DESC"
		ElseIf pOrderBy = 2 Then 
			vQuery = vQuery & " ORDER BY wrv_timestamp"
		Else
			vQuery = vQuery & " ORDER BY wpg_name"
		End If
		If cAllowAttachments And pIncludeAttachmentChanges Then
			vQuery = vQuery & ", ath_timestamp DESC"
		End If
		vRS.Open(vQuery, vConn, ADODB.CursorTypeEnum.adOpenForwardOnly)
		Do While Not vRS.EOF
			If vRegEx.Test(vRS.Fields("wpg_name")) Then
				If vCurPage <> IIF(IsDBNull(vRS.Fields.Item("wpg_name").Value), Nothing, vRS.Fields.Item("wpg_name").Value) Then
					vCurPage = IIF(IsDBNull(vRS.Fields.Item("wpg_name").Value), Nothing, vRS.Fields.Item("wpg_name").Value)
					vPage = New WikiPage
					vPage.Name = IIF(IsDBNull(vRS.Fields.Item("wpg_name").Value), Nothing, vRS.Fields.Item("wpg_name").Value)
					vPage.Changes = CShort(IIF(IsDBNull(vRS.Fields.Item("wpg_changes").Value), Nothing, vRS.Fields.Item("wpg_changes").Value))
					vChange = vPage.AddChange
					vChange.Revision = CShort(IIF(IsDBNull(vRS.Fields.Item("wrv_revision").Value), Nothing, vRS.Fields.Item("wrv_revision").Value))
					vChange.Status = CShort(IIF(IsDBNull(vRS.Fields.Item("wrv_status").Value), Nothing, vRS.Fields.Item("wrv_status").Value))
					vChange.MinorEdit = CShort(IIF(IsDBNull(vRS.Fields.Item("wrv_minoredit").Value), Nothing, vRS.Fields.Item("wrv_minoredit").Value))
					vChange.Timestamp = IIF(IsDBNull(vRS.Fields.Item("wrv_timestamp").Value), Nothing, vRS.Fields.Item("wrv_timestamp").Value)
					vChange.By = IIF(IsDBNull(vRS.Fields.Item("wrv_by").Value), Nothing, vRS.Fields.Item("wrv_by").Value)
					vChange.ByAlias = IIF(IsDBNull(vRS.Fields.Item("wrv_byalias").Value), Nothing, vRS.Fields.Item("wrv_byalias").Value)
					vChange.Comment = IIF(IsDBNull(vRS.Fields.Item("wrv_comment").Value), Nothing, vRS.Fields.Item("wrv_comment").Value)
					vList.Push(vPage)
				End If
				If cAllowAttachments And pIncludeAttachmentChanges Then
					If (IIF(IsDBNull(vRS.Fields.Item("ath_name").Value), Nothing, vRS.Fields.Item("ath_name").Value) <> "") And (IIF(IsDBNull(vRS.Fields.Item("ath_timestamp").Value), Nothing, vRS.Fields.Item("ath_timestamp").Value) > DateAdd(Microsoft.VisualBasic.DateInterval.Hour, -24, Now())) Then
						vAttachmentChange = New AttachmentChange
						vAttachmentChange.Name = IIF(IsDBNull(vRS.Fields.Item("ath_name").Value), Nothing, vRS.Fields.Item("ath_name").Value)
						vAttachmentChange.Revision = CShort(IIF(IsDBNull(vRS.Fields.Item("ath_revision").Value), Nothing, vRS.Fields.Item("ath_revision").Value))
						vAttachmentChange.Timestamp = IIF(IsDBNull(vRS.Fields.Item("ath_timestamp").Value), Nothing, vRS.Fields.Item("ath_timestamp").Value)
						vAttachmentChange.By = IIF(IsDBNull(vRS.Fields.Item("ath_by").Value), Nothing, vRS.Fields.Item("ath_by").Value)
						vAttachmentChange.ByAlias = IIF(IsDBNull(vRS.Fields.Item("ath_byalias").Value), Nothing, vRS.Fields.Item("ath_byalias").Value)
						vAttachmentChange.Action = IIF(IsDBNull(vRS.Fields.Item("ath_action").Value), Nothing, vRS.Fields.Item("ath_action").Value)
						vChange.AddAttachmentChange(vAttachmentChange)
					End If
				End If
			End If
			vRS.MoveNext()
		Loop 
		vRS.Close()
		'UPGRADE_NOTE: Object vRegEx may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
		vRegEx = Nothing
		TitleSearch = vList
	End Function
	
	
	Function FullSearch(ByRef pPattern As Object, ByRef pIncludeTitles As Object) As Object
		Dim vChange, vList, vRegEx, vTitle, vRegEx2, vPage, vFound As Object
		'UPGRADE_NOTE: Global Sub/Function EscapePattern is not accessible
		pPattern = EscapePattern(pPattern)
		vList = New Vector
		vRegEx = New RegExp
		vRegEx.IgnoreCase = True
		vRegEx.Global = True
		If CStr(HttpContext.Current.Request("fromtitle")) = "true" Then
			vRegEx.Pattern = Replace(pPattern, "_", " ")
		Else
			vRegEx.Pattern = pPattern
		End If
		If pIncludeTitles Then
			vRegEx2 = New RegExp
			vRegEx2.IgnoreCase = True
			vRegEx2.Global = True
			vRegEx2.Pattern = pPattern
		End If
		vQuery = "SELECT * FROM openwiki_pages, openwiki_revisions WHERE wrv_name = wpg_name AND wrv_current = 1 AND wrv_text IS NOT NULL ORDER BY wpg_name"
		vRS.Open(vQuery, vConn, ADODB.CursorTypeEnum.adOpenForwardOnly)
		Do While Not vRS.EOF
			vFound = False
			If pIncludeTitles Then
				If vRegEx2.Test(vRS.Fields("wpg_name")) Then
					vFound = True
				End If
			End If
			If Not vFound Then
				If vRegEx.Test(vRS.Fields("wrv_text")) Then
					vFound = True
				End If
			End If
			If vFound Then
				vPage = New WikiPage
				vPage.Name = IIF(IsDBNull(vRS.Fields.Item("wpg_name").Value), Nothing, vRS.Fields.Item("wpg_name").Value)
				vPage.Changes = CShort(IIF(IsDBNull(vRS.Fields.Item("wpg_changes").Value), Nothing, vRS.Fields.Item("wpg_changes").Value))
				vChange = vPage.AddChange
				vChange.Revision = CShort(IIF(IsDBNull(vRS.Fields.Item("wrv_revision").Value), Nothing, vRS.Fields.Item("wrv_revision").Value))
				vChange.Status = CShort(IIF(IsDBNull(vRS.Fields.Item("wrv_status").Value), Nothing, vRS.Fields.Item("wrv_status").Value))
				vChange.MinorEdit = CShort(IIF(IsDBNull(vRS.Fields.Item("wrv_minoredit").Value), Nothing, vRS.Fields.Item("wrv_minoredit").Value))
				vChange.Timestamp = IIF(IsDBNull(vRS.Fields.Item("wrv_timestamp").Value), Nothing, vRS.Fields.Item("wrv_timestamp").Value)
				vChange.By = IIF(IsDBNull(vRS.Fields.Item("wrv_by").Value), Nothing, vRS.Fields.Item("wrv_by").Value)
				vChange.ByAlias = IIF(IsDBNull(vRS.Fields.Item("wrv_byalias").Value), Nothing, vRS.Fields.Item("wrv_byalias").Value)
				vChange.Comment = IIF(IsDBNull(vRS.Fields.Item("wrv_comment").Value), Nothing, vRS.Fields.Item("wrv_comment").Value)
				vList.Push(vPage)
			End If
			vRS.MoveNext()
		Loop 
		vRS.Close()
		'UPGRADE_NOTE: Object vRegEx may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
		vRegEx = Nothing
		'UPGRADE_NOTE: Object vRegEx2 may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
		vRegEx2 = Nothing
		FullSearch = vList
	End Function
	
	Function EquationSearch(ByRef pPattern As Object, ByRef pIncludeTitles As Object, ByRef pOrderBy As Object) As Object
		Dim gEquation As Object
		Dim gSpecialPagesPrefix As Object
		Dim cUseSpecialPagesPrefix As Object
		Dim vFound, vPage, vRegEx2, vTitle, vRegEx, vList, vChange, vText As Object
		'UPGRADE_NOTE: Global Sub/Function EscapePattern is not accessible
		pPattern = EscapePattern(pPattern)
		vList = New Vector
		vRegEx = New RegExp
		vRegEx.IgnoreCase = True
		vRegEx.Global = True
		If CStr(HttpContext.Current.Request("fromtitle")) = "true" Then
			vRegEx.Pattern = Replace(pPattern, "_", " ")
		Else
			vRegEx.Pattern = pPattern
		End If
		If pIncludeTitles Then
			vRegEx2 = New RegExp
			vRegEx2.IgnoreCase = True
			vRegEx2.Global = True
			vRegEx2.Pattern = pPattern
		End If
		vQuery = "SELECT * FROM openwiki_pages, openwiki_revisions WHERE wrv_name = wpg_name AND wrv_current = 1 AND wrv_text IS NOT NULL"
		If pOrderBy = 1 Then
			vQuery = vQuery & " ORDER BY wrv_timestamp DESC"
		ElseIf pOrderBy = 2 Then 
			vQuery = vQuery & " ORDER BY wrv_timestamp"
		Else
			vQuery = vQuery & " ORDER BY wpg_name"
		End If
		vRS.Open(vQuery, vConn, ADODB.CursorTypeEnum.adOpenForwardOnly)
		Do While Not vRS.EOF
			vText = IIF(IsDBNull(vRS.Fields.Item("wrv_text").Value), Nothing, vRS.Fields.Item("wrv_text").Value)
			vFound = False
			If pIncludeTitles Then
				If vRegEx2.Test(vRS.Fields("wpg_name")) Then
					vFound = True
				End If
			End If
			If Not vFound Then
				If vRegEx.Test(vText) Then
					vFound = True
				End If
			End If
			'UPGRADE_NOTE: Global Sub/Function m is not accessible
			If vFound And pIncludeTitles And cUseSpecialPagesPrefix And m(vRS.Fields("wpg_name"), "^" & gSpecialPagesPrefix, False, False) Then
				vFound = False
			End If
			If vFound Then
				vPage = New WikiPage
				vPage.Name = IIF(IsDBNull(vRS.Fields.Item("wpg_name").Value), Nothing, vRS.Fields.Item("wpg_name").Value)
				'UPGRADE_NOTE: Global Sub/Function s is not accessible
				Call s(vText, "<math>([\s\S]*?)<\/math>", "&CutEquation($1)", False, False)
				vPage.text = gEquation
				vPage.Changes = CShort(IIF(IsDBNull(vRS.Fields.Item("wpg_changes").Value), Nothing, vRS.Fields.Item("wpg_changes").Value))
				vChange = vPage.AddChange
				vChange.Revision = CShort(IIF(IsDBNull(vRS.Fields.Item("wrv_revision").Value), Nothing, vRS.Fields.Item("wrv_revision").Value))
				vChange.Status = CShort(IIF(IsDBNull(vRS.Fields.Item("wrv_status").Value), Nothing, vRS.Fields.Item("wrv_status").Value))
				vChange.MinorEdit = CShort(IIF(IsDBNull(vRS.Fields.Item("wrv_minoredit").Value), Nothing, vRS.Fields.Item("wrv_minoredit").Value))
				vChange.Timestamp = IIF(IsDBNull(vRS.Fields.Item("wrv_timestamp").Value), Nothing, vRS.Fields.Item("wrv_timestamp").Value)
				vChange.By = IIF(IsDBNull(vRS.Fields.Item("wrv_by").Value), Nothing, vRS.Fields.Item("wrv_by").Value)
				vChange.ByAlias = IIF(IsDBNull(vRS.Fields.Item("wrv_byalias").Value), Nothing, vRS.Fields.Item("wrv_byalias").Value)
				vChange.Comment = IIF(IsDBNull(vRS.Fields.Item("wrv_comment").Value), Nothing, vRS.Fields.Item("wrv_comment").Value)
				vList.Push(vPage)
			End If
			vRS.MoveNext()
		Loop 
		vRS.Close()
		'UPGRADE_NOTE: Object vRegEx may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
		vRegEx = Nothing
		'UPGRADE_NOTE: Object vRegEx2 may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
		vRegEx2 = Nothing
		EquationSearch = vList
	End Function
	
	Function GetPreviousRevision(ByRef pDiffType As Object, ByRef pDiffTo As Object) As Object
		Dim gPage As Object
		Dim vHost, vBy, vAgent As Object
		GetPreviousRevision = 0
		If pDiffTo <= 0 Then
			pDiffTo = 99999999
		End If
		vQuery = "SELECT wrv_revision, wrv_minoredit, wrv_by, wrv_host, wrv_agent FROM openwiki_revisions WHERE wrv_name = '" & Replace(gPage, "'", "''") & "' AND wrv_revision <= " & pDiffTo
		vQuery = vQuery & " ORDER BY wrv_revision DESC"
		vRS.Open(vQuery, vConn, ADODB.CursorTypeEnum.adOpenForwardOnly)
		If Not vRS.EOF Then
			vBy = IIF(IsDBNull(vRS.Fields.Item("wrv_by").Value), Nothing, vRS.Fields.Item("wrv_by").Value)
			vHost = IIF(IsDBNull(vRS.Fields.Item("wrv_host").Value), Nothing, vRS.Fields.Item("wrv_host").Value)
			vAgent = IIF(IsDBNull(vRS.Fields.Item("wrv_agent").Value), Nothing, vRS.Fields.Item("wrv_agent").Value)
		End If
		Do While Not vRS.EOF
			GetPreviousRevision = CShort(IIF(IsDBNull(vRS.Fields.Item("wrv_revision").Value), Nothing, vRS.Fields.Item("wrv_revision").Value))
			If pDiffType = 0 Then
				' previous major
				If CShort(IIF(IsDBNull(vRS.Fields.Item("wrv_minoredit").Value), Nothing, vRS.Fields.Item("wrv_minoredit").Value)) = 0 Then
					vRS.MoveNext()
					If Not vRS.EOF Then
						GetPreviousRevision = CShort(IIF(IsDBNull(vRS.Fields.Item("wrv_revision").Value), Nothing, vRS.Fields.Item("wrv_revision").Value))
					End If
					Exit Do
				End If
			ElseIf pDiffType = 1 Then 
				' previous minor
				vRS.MoveNext()
				If Not vRS.EOF Then
					GetPreviousRevision = CShort(IIF(IsDBNull(vRS.Fields.Item("wrv_revision").Value), Nothing, vRS.Fields.Item("wrv_revision").Value))
				End If
				Exit Do
			Else
				' previous author
				If IIF(IsDBNull(vRS.Fields.Item("wrv_by").Value), Nothing, vRS.Fields.Item("wrv_by").Value) <> vBy Or IIF(IsDBNull(vRS.Fields.Item("wrv_host").Value), Nothing, vRS.Fields.Item("wrv_host").Value) <> vHost Or IIF(IsDBNull(vRS.Fields.Item("wrv_agent").Value), Nothing, vRS.Fields.Item("wrv_agent").Value) <> vAgent Then
					Exit Do
				End If
			End If
			vRS.MoveNext()
		Loop 
		vRS.Close()
	End Function
	
	
	Function InterWiki() As Object
		Dim vTemp As Object
		vQuery = "SELECT wik_name, wik_url FROM openwiki_interwikis ORDER BY wik_name"
		vRS.Open(vQuery, vConn, ADODB.CursorTypeEnum.adOpenForwardOnly)
		Do While Not vRS.EOF
			'UPGRADE_NOTE: Global Sub/Function CDATAEncode is not accessible
			'UPGRADE_NOTE: Global Sub/Function PCDATAEncode is not accessible
			vTemp = vTemp & "<ow:interlink>" & "<ow:name>" & PCDATAEncode(vRS.Fields("wik_name")) & "</ow:name>" & "<ow:href>" & CDATAEncode(vRS.Fields("wik_url")) & "</ow:href>" & "<ow:class>" & CDATAEncode(LCase(Trim(IIF(IsDBNull(vRS.Fields.Item("wik_name").Value), Nothing, vRS.Fields.Item("wik_name").Value)))) & "</ow:class>" & "</ow:interlink>"
			vRS.MoveNext()
		Loop 
		vRS.Close()
		InterWiki = "<ow:interlinks>" & vTemp & "</ow:interlinks>"
	End Function
	
	Function ListRedirects() As Object
		Dim DB_ACCESS As Object
		Dim OPENWIKI_DB_SYNTAX As Object
		Dim vPageTo, vText, vTemp, vTempPage, vPageFrom, vPos As Object
		Dim vBuffTo, vBuffFrom, i As Object
		
		vBuffFrom = New Vector
		vBuffTo = New Vector
		
		If OPENWIKI_DB_SYNTAX = DB_ACCESS Then
			vQuery = "SELECT * FROM openwiki_pages, openwiki_revisions " & "WHERE wrv_name = wpg_name AND wrv_current = 1 AND wrv_text LIKE '[#]REDIRECT %' " & "ORDER BY wpg_name"
		Else
			vQuery = "SELECT * FROM openwiki_pages, openwiki_revisions " & "WHERE wrv_name = wpg_name AND wrv_current = 1 AND wrv_text LIKE '\#REDIRECT %' ESCAPE '\' " & "ORDER BY wpg_name"
		End If
		vRS.Open(vQuery, vConn, ADODB.CursorTypeEnum.adOpenForwardOnly)
		Do While Not vRS.EOF
			vText = IIF(IsDBNull(vRS.Fields.Item("wrv_text").Value), Nothing, vRS.Fields.Item("wrv_text").Value)
			vPageFrom = IIF(IsDBNull(vRS.Fields.Item("wpg_name").Value), Nothing, vRS.Fields.Item("wpg_name").Value)
			'UPGRADE_NOTE: Global Sub/Function m is not accessible
			If m(vText, "^#REDIRECT\s+", False, False) Then
				vPos = InStr(Len("#REDIRECT "), vText, vbCr)
				If vPos > 0 Then
					vPageTo = Trim(Mid(vText, Len("#REDIRECT "), vPos - Len("#REDIRECT ")))
				Else
					vPageTo = Trim(Mid(vText, Len("#REDIRECT ")))
				End If
				vBuffFrom.Push(vPageFrom)
				vBuffTo.Push(vPageTo)
			End If
			vRS.MoveNext()
		Loop 
		vRS.Close()
		
		If Not vBuffFrom.IsEmpty() Then
			For i = 0 To vBuffFrom.Count - 1
				vPageFrom = vBuffFrom.ElementAt(i)
				vPageTo = vBuffTo.ElementAt(i)
				'UPGRADE_NOTE: Global Sub/Function GetWikiLink is not accessible
				vTemp = vTemp & "<ow:redirect>" & "<ow:from>" & GetWikiLink("", vPageFrom, "") & "</ow:from>" & "<ow:to>" & GetWikiLink("", vPageTo, "") & "</ow:to>" & "</ow:redirect>"
			Next 
			ListRedirects = "<ow:redirectlinks>" & vTemp & "</ow:redirectlinks>"
		End If
		
		'UPGRADE_NOTE: Object vBuffFrom may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
		vBuffFrom = Nothing
		'UPGRADE_NOTE: Object vBuffTo may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
		vBuffTo = Nothing
	End Function
	
	Function GetInterWiki(ByRef pName As Object) As Object
		Dim OPENWIKI_DB As Object
		Dim gScriptName As Object
		If OPENWIKI_DB <> "" Then
			If pName = "This" Then
				GetInterWiki = gScriptName & "?p="
			Else
				vQuery = "SELECT wik_url FROM openwiki_interwikis WHERE wik_name = '" & Replace(pName, "'", "''") & "'"
				vRS.Open(vQuery, vConn, ADODB.CursorTypeEnum.adOpenForwardOnly)
				If Not vRS.EOF Then
					GetInterWiki = IIF(IsDBNull(vRS.Fields.Item("wik_url").Value), Nothing, vRS.Fields.Item("wik_url").Value)
				End If
				vRS.Close()
			End If
		End If
	End Function
	
	
	Function GetRSSFromCache(ByRef pURL As Object, ByRef pRefreshRate As Object, ByRef pFreshlyFromRemoteSite As Object, ByRef pRetryLater As Object) As Object
		Dim FormatDateISO8601 As Object
		Dim OPENWIKI_DB As Object
		Dim conn As ADODB.Connection
		Dim vRS As ADODB.Recordset
		Dim vNext, vLast, vRefreshRate As Object
		conn = New ADODB.Connection
		conn.Open(OPENWIKI_DB)
		vQuery = "SELECT rss_last, rss_next, rss_refreshrate, rss_cache FROM openwiki_rss WHERE rss_url = '" & Replace(pURL, "'", "''") & "'"
		vRS = New ADODB.Recordset
		vRS.Open(vQuery, conn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
		If vRS.EOF Then
			GetRSSFromCache = "notexists"
		Else
			vLast = IIF(IsDBNull(vRS.Fields.Item("rss_last").Value), Nothing, vRS.Fields.Item("rss_last").Value)
			vNext = IIF(IsDBNull(vRS.Fields.Item("rss_next").Value), Nothing, vRS.Fields.Item("rss_next").Value)
			vRefreshRate = CShort(IIF(IsDBNull(vRS.Fields.Item("rss_refreshrate").Value), Nothing, vRS.Fields.Item("rss_refreshrate").Value))
			If vRefreshRate <> pRefreshRate Then
				vNext = DateAdd(Microsoft.VisualBasic.DateInterval.Minute, pRefreshRate, vLast)
				vRS.Fields("rss_next").Value = vNext
				vRS.Fields("rss_refreshrate").Value = pRefreshRate
				vRS.Update()
			ElseIf pRetryLater Then 
				' retry a minute from now
				vNext = DateAdd(Microsoft.VisualBasic.DateInterval.Minute, 1, Now)
				vRS.Fields("rss_next").Value = vNext
				vRS.Update()
			End If
			
			If pFreshlyFromRemoteSite Or (DateDiff(Microsoft.VisualBasic.DateInterval.Minute, vNext, Now) < 0) Then
				GetRSSFromCache = "<ow:feed href='" & Replace(pURL, "&", "&amp;") & "' "
				If pFreshlyFromRemoteSite Then
					GetRSSFromCache = GetRSSFromCache & "fresh='true' "
				Else
					GetRSSFromCache = GetRSSFromCache & "fresh='false' "
				End If
				GetRSSFromCache = GetRSSFromCache & "last='" & FormatDateISO8601(vLast) & "' "
				GetRSSFromCache = GetRSSFromCache & "next='" & FormatDateISO8601(vNext) & "' "
				GetRSSFromCache = GetRSSFromCache & "refreshrate='" & pRefreshRate & "'>"
				GetRSSFromCache = GetRSSFromCache & IIF(IsDBNull(vRS.Fields.Item("rss_cache").Value), Nothing, vRS.Fields.Item("rss_cache").Value)
				GetRSSFromCache = GetRSSFromCache & "</ow:feed>"
			End If
			
		End If
		vRS.Close()
		conn.Close()
		'UPGRADE_NOTE: Object vRS may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
		vRS = Nothing
		'UPGRADE_NOTE: Object conn may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
		conn = Nothing
	End Function
	
	Sub SaveRSSToCache(ByRef pURL As Object, ByRef pRefreshRate As Object, ByRef pCache As Object)
		Dim OPENWIKI_DB As Object
		Dim conn As ADODB.Connection
		Dim vRS As ADODB.Recordset
		conn = New ADODB.Connection
		conn.Open(OPENWIKI_DB)
		vQuery = "SELECT * FROM openwiki_rss WHERE rss_url = '" & Replace(pURL, "'", "''") & "'"
		vRS = New ADODB.Recordset
		vRS.Open(vQuery, conn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
		If vRS.EOF Then
			vRS.Close()
			vRS.Open("openwiki_rss", conn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdTable)
			vRS.AddNew()
			vRS.Fields("rss_url").Value = pURL
		End If
		vRS.Fields("rss_last").Value = Now
		If pCache = "" Then
			vRS.Fields("rss_next").Value = DateAdd(Microsoft.VisualBasic.DateInterval.Minute, 30, Now) ' 30 minutes from now
		Else
			vRS.Fields("rss_next").Value = DateAdd(Microsoft.VisualBasic.DateInterval.Minute, pRefreshRate, Now)
		End If
		vRS.Fields("rss_refreshrate").Value = pRefreshRate
		vRS.Fields("rss_cache").Value = pCache
		vRS.Update()
		vRS.Close()
		conn.Close()
		'UPGRADE_NOTE: Object vRS may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
		vRS = Nothing
		'UPGRADE_NOTE: Object conn may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
		conn = Nothing
	End Sub
	
	Sub Aggregate(ByRef pURL As Object, ByRef pXmlDoc As Object)
		Dim FormatDateISO8601 As Object
		Dim sReturn As Object
		Dim gTimestampPattern As Object
		Dim OPENWIKI_DB As Object
		Dim conn As ADODB.Connection
		Dim vRS As ADODB.Recordset
		Dim vItems, vRoot, vItem As Object
		Dim vNow, vXmlIsland, vAgXmlIsland, i As Object
		Dim vRdfTimestamp, vRssLink, vRdfResource, vDcDate As Object
		
		On Error Resume Next
		'Response.Write("<p />Aggregating " & pURL & "<br />")
		
		vRoot = pXmlDoc.documentElement
		
		If vRoot.NodeName = "rss" Then
			vItems = vRoot.selectNodes("channel/item")
		ElseIf vRoot.getAttribute("xmlns") = "http://my.netscape.com/rdf/simple/0.9/" Then 
			vItems = vRoot.selectNodes("item")
		ElseIf vRoot.getAttribute("xmlns") = "http://purl.org/rss/1.0/" Then 
			vItems = vRoot.selectNodes("item")
		Else
			Exit Sub
		End If
		
		vNow = Now
		i = 0
		
		' TODO: find workaround for bug in MSXML v4
		If Not vRoot.selectSingleNode("channel/wiki:interwiki") Is Nothing Then
			'UPGRADE_NOTE: Global Sub/Function PCDATAEncode is not accessible
			vAgXmlIsland = "<ag:source><rdf:Description wiki:interwiki=""" & vRoot.selectSingleNode("channel/wiki:interwiki").text & """><rdf:value>" & PCDATAEncode(vRoot.selectSingleNode("channel/title").text) & "</rdf:value></rdf:Description></ag:source>"
		Else
			'UPGRADE_NOTE: Global Sub/Function PCDATAEncode is not accessible
			vAgXmlIsland = "<ag:source>" & PCDATAEncode(vRoot.selectSingleNode("channel/title").text) & "</ag:source>"
		End If
		'UPGRADE_NOTE: Global Sub/Function PCDATAEncode is not accessible
		vAgXmlIsland = vAgXmlIsland & "<ag:sourceURL>" & PCDATAEncode(vRoot.selectSingleNode("channel/link").text) & "</ag:sourceURL>"
		
		conn = New ADODB.Connection
		conn.Open(OPENWIKI_DB)
		vRS = New ADODB.Recordset
		
		' walk trough all item elements and store them in the openwiki_rss_aggregations table
		For	Each vItem In vItems
			vRssLink = vItem.selectSingleNode("link").text
			
			vRdfResource = vItem.getAttribute("rdf:about")
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1049.asp'
			If IsDbNull(vRdfResource) Then
				vRdfResource = vRssLink
			End If
			
			If vItem.selectSingleNode("ag:timestamp") Is Nothing Then
				vRdfTimestamp = DateAdd(Microsoft.VisualBasic.DateInterval.Second, i, vNow)
			Else
				vRdfTimestamp = vItem.selectSingleNode("ag:timestamp").text
				'UPGRADE_NOTE: Global Sub/Function s is not accessible
				Call s(vRdfTimestamp, gTimestampPattern, "&ToDateTime($1,$2,$3,$4,$5,$6,$7,$8,$9)", False, False)
				If DateDiff(Microsoft.VisualBasic.DateInterval.Day, vNow, sReturn) > 1 Then
					' we cannot take this date serious, it's too far in the future
					vRdfTimestamp = DateAdd(Microsoft.VisualBasic.DateInterval.Second, i, vNow)
				Else
					vRdfTimestamp = sReturn
					vAgXmlIsland = vItem.selectSingleNode("ag:source").xml & vItem.selectSingleNode("ag:sourceURL").xml
				End If
			End If
			i = i - 1
			
			'UPGRADE_NOTE: Global Sub/Function PCDATAEncode is not accessible
			vXmlIsland = "<title>" & PCDATAEncode(vItem.selectSingleNode("title").text) & "</title><link>" & PCDATAEncode(vItem.selectSingleNode("link").text) & "</link>"
			If Not vItem.selectSingleNode("description") Is Nothing Then
				'UPGRADE_NOTE: Global Sub/Function PCDATAEncode is not accessible
				vXmlIsland = vXmlIsland & "<description>" & PCDATAEncode(vItem.selectSingleNode("description").text) & "</description>"
			End If
			If Not vItem.selectSingleNode("dc:creator") Is Nothing Then
				vXmlIsland = vXmlIsland & vItem.selectSingleNode("dc:creator").xml
			End If
			If Not vItem.selectSingleNode("dc:contributor") Is Nothing Then
				vXmlIsland = vXmlIsland & vItem.selectSingleNode("dc:contributor").xml
			End If
			If vItem.selectSingleNode("dc:date") Is Nothing Then
				vDcDate = ""
			Else
				vDcDate = vItem.selectSingleNode("dc:date").text
				vXmlIsland = vXmlIsland & "<dc:date>" & vItem.selectSingleNode("dc:date").text & "</dc:date>"
			End If
			If Not vItem.selectSingleNode("wiki:version") Is Nothing Then
				vXmlIsland = vXmlIsland & "<wiki:version>" & vItem.selectSingleNode("wiki:version").text & "</wiki:version>"
			End If
			If Not vItem.selectSingleNode("wiki:status") Is Nothing Then
				vXmlIsland = vXmlIsland & "<wiki:status>" & vItem.selectSingleNode("wiki:status").text & "</wiki:status>"
			End If
			If Not vItem.selectSingleNode("wiki:importance") Is Nothing Then
				vXmlIsland = vXmlIsland & "<wiki:importance>" & vItem.selectSingleNode("wiki:importance").text & "</wiki:importance>"
			End If
			If Not vItem.selectSingleNode("wiki:diff") Is Nothing Then
				vXmlIsland = vXmlIsland & vItem.selectSingleNode("wiki:diff").xml
			End If
			If Not vItem.selectSingleNode("wiki:history") Is Nothing Then
				vXmlIsland = vXmlIsland & vItem.selectSingleNode("wiki:history").xml
			End If
			vXmlIsland = vXmlIsland & vAgXmlIsland & "<ag:timestamp>" & FormatDateISO8601(vRdfTimestamp) & "</ag:timestamp>"
			
			'UPGRADE_NOTE: Global Sub/Function PCDATAEncode is not accessible
			vXmlIsland = "<item rdf:about='" & PCDATAEncode(vRdfResource) & "'>" & vXmlIsland & "</item>"
			
			' TODO: erm... this is actually inefficient.. use better ADO techniques
			vQuery = "SELECT * FROM openwiki_rss_aggregations WHERE agr_feed='" & Replace(pURL, "'", "''") & "' AND agr_rsslink = '" & Replace(vRssLink, "'", "''") & "'"
			vRS.Open(vQuery, conn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
			If vRS.EOF Then
				vRS.Close()
				vRS.Open("openwiki_rss_aggregations", conn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdTable)
				vRS.AddNew()
				vRS.Fields("agr_feed").Value = pURL
				vRS.Fields("agr_resource").Value = vRdfResource
				vRS.Fields("agr_rsslink").Value = vRssLink
				vRS.Fields("agr_timestamp").Value = vRdfTimestamp
				vRS.Fields("agr_dcdate").Value = vDcDate
				vRS.Fields("agr_xmlisland").Value = vXmlIsland
				vRS.Update()
			ElseIf IIF(IsDBNull(vRS.Fields.Item("agr_dcdate").Value), Nothing, vRS.Fields.Item("agr_dcdate").Value) <> vDcDate Then 
				vRS.Fields("agr_resource").Value = vRdfResource
				vRS.Fields("agr_timestamp").Value = vRdfTimestamp
				vRS.Fields("agr_dcdate").Value = vDcDate
				vRS.Fields("agr_xmlisland").Value = vXmlIsland
				vRS.Update()
			End If
			vRS.Close()
		Next vItem
		
		conn.Close()
		'UPGRADE_NOTE: Object vRS may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
		vRS = Nothing
		'UPGRADE_NOTE: Object conn may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
		conn = Nothing
		
		'Response.Write("<p />Done aggregating " & pURL & "<br />")
	End Sub
	
	Function GetAggregation(ByRef pURLs As Object) As Object
		Dim PrettyWikiLink As Object
		Dim OPENWIKI_TITLE As Object
		Dim gPage As Object
		Dim gScriptName As Object
		Dim gServerRoot As Object
		Dim OPENWIKI_MAXNROFAGGR As Object
		Dim vTemp, vRdfSeq, vItems, i As Object
		vQuery = ""
		Do While Not pURLs.IsEmpty
			vQuery = vQuery & "'" & Replace(pURLs.Pop(), "'", "''") & "'"
			If pURLs.Count > 0 Then
				vQuery = vQuery & ","
			End If
		Loop 
		vQuery = "SELECT * FROM openwiki_rss_aggregations WHERE agr_feed IN (" & vQuery & ") ORDER BY agr_timestamp DESC"
		vRS.Open(vQuery, vConn, ADODB.CursorTypeEnum.adOpenForwardOnly)
		i = 0
		If OPENWIKI_MAXNROFAGGR <= 0 Then
			OPENWIKI_MAXNROFAGGR = 100
		End If
		Do While Not vRS.EOF
			i = i + 1
			If i > OPENWIKI_MAXNROFAGGR Then
				Exit Do
			End If
			'UPGRADE_NOTE: Global Sub/Function CDATAEncode is not accessible
			vTemp = CDATAEncode(vRS.Fields("agr_resource"))
			vRdfSeq = vRdfSeq & "<rdf:li rdf:resource='" & vTemp & "'/>"
			vItems = vItems & IIF(IsDBNull(vRS.Fields.Item("agr_xmlisland").Value), Nothing, vRS.Fields.Item("agr_xmlisland").Value)
			vRS.MoveNext()
		Loop 
		vRS.Close()
		'UPGRADE_NOTE: Global Sub/Function PCDATAEncode is not accessible
		'UPGRADE_NOTE: Global Sub/Function CDATAEncode is not accessible
		GetAggregation = "<?xml version='1.0' encoding='ISO-8859-1'?>" & vbCrLf & "<!-- All Your Wiki Are Belong To Us -->" & vbCrLf & "<rdf:RDF xmlns='http://purl.org/rss/1.0/' xmlns:rdf='http://www.w3.org/1999/02/22-rdf-syntax-ns#' xmlns:dc='http://purl.org/dc/elements/1.1/' xmlns:wiki='http://purl.org/rss/1.0/modules/wiki/' xmlns:ag='http://purl.org/rss/1.0/modules/aggregation/'>" & "<channel rdf:about='" & CDATAEncode(gServerRoot & gScriptName & "?p=" & gPage & "&a=rss") & "'>" & "<title>" & PCDATAEncode(OPENWIKI_TITLE & " -- " & PrettyWikiLink(gPage)) & "</title>" & "<link>" & PCDATAEncode(gServerRoot & gScriptName & "?" & gPage) & "</link>" & "<description>" & PCDATAEncode(OPENWIKI_TITLE & " -- " & PrettyWikiLink(gPage)) & "</description>" & "<image rdf:about='" & CDATAEncode(gServerRoot & "ow/images/aggregator.gif") & "'/>" & "<items><rdf:Seq>" & vRdfSeq & "</rdf:Seq></items>" & "</channel>" & "<image rdf:about='" & CDATAEncode(gServerRoot & "ow/images/aggregator.gif") & "'>" & "<title>" & PCDATAEncode(OPENWIKI_TITLE) & "</title>" & "<link>" & CDATAEncode(gServerRoot & gScriptName & "?p=" & gPage) & "</link>" & "<url>" & PCDATAEncode(gServerRoot & "ow/images/logo_aggregator.gif") & "</url>" & "</image>" & vItems & "</rdf:RDF>"
	End Function
	
	
	Private Function CreateDocKey(ByRef pSubKey As Object) As Object
		Dim Hash As Object
		Dim gRevision As Object
		Dim gCookieHash As Object
		Dim gFS As Object
		CreateDocKey = pSubKey & gFS & gCookieHash & gFS & gRevision & gFS & HttpContext.Current.Request.Cookies(gCookieHash & "?up")("pwl") & gFS & HttpContext.Current.Request.Cookies(gCookieHash & "?up")("new") & gFS & HttpContext.Current.Request.Cookies(gCookieHash & "?up")("emo")
		CreateDocKey = Hash(CreateDocKey)
	End Function
	
	Function GetDocumentCache(ByRef pSubKey As Object) As Object
		Dim gPage As Object
		vQuery = "SELECT chc_xmlisland FROM openwiki_cache WHERE chc_name = '" & Replace(gPage, "'", "''") & "' AND chc_hash = " & CreateDocKey(pSubKey)
		vRS.Open(vQuery, vConn, ADODB.CursorTypeEnum.adOpenForwardOnly)
		If vRS.EOF Then
			GetDocumentCache = ""
		Else
			GetDocumentCache = IIF(IsDBNull(vRS.Fields.Item("chc_xmlisland").Value), Nothing, vRS.Fields.Item("chc_xmlisland").Value)
		End If
		vRS.Close()
	End Function
	
	Sub SetDocumentCache(ByRef pSubKey As Object, ByRef pXmlStr As Object)
		Dim gPage As Object
		Dim vKey As Object
		vKey = CreateDocKey(pSubKey)
		vQuery = "SELECT chc_xmlisland FROM openwiki_cache WHERE chc_name = '" & Replace(gPage, "'", "''") & "' AND chc_hash = " & vKey
		vRS.Open(vQuery, vConn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
		If vRS.EOF Then
			vRS.Close()
			vRS.Open("openwiki_cache", vConn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdTable)
			vRS.AddNew()
			vRS.Fields("chc_name").Value = gPage
			vRS.Fields("chc_hash").Value = vKey
		End If
		vRS.Fields("chc_xmlisland").Value = pXmlStr
		vRS.Update()
		vRS.Close()
	End Sub
	
	Sub ClearDocumentCache(ByRef pConn As Object)
		pConn.Execute("DELETE FROM openwiki_cache")
	End Sub
	
	Sub ClearDocumentCache2(ByRef pConn As Object, ByRef pPagename As Object)
		If pConn = "" Then
			pConn = vConn
		End If
		pConn.Execute("DELETE FROM openwiki_cache WHERE chc_name = '" & Replace(pPagename, "'", "''") & "'")
	End Sub
End Class

'End Sub


Sub ToDateTime(ByRef pYear As Object, ByRef pMonth As Object, ByRef pDay As Object, ByRef pHour As Object, ByRef pMinutes As Object, ByRef pSeconds As Object, ByRef pPlusMinTZ As Object, ByRef pHourTZ As Object, ByRef pMinutesTZ As Object)
	sReturn = DateSerial(pYear, pMonth, pDay)
	If pPlusMinTZ = "-" Then
		sReturn = DateAdd(Microsoft.VisualBasic.DateInterval.Hour, pHour + pHourTZ, sReturn)
		sReturn = DateAdd(Microsoft.VisualBasic.DateInterval.Minute, pMinutes + pMinutesTZ, sReturn)
	ElseIf pPlusMinTZ = "+" Then 
		sReturn = DateAdd(Microsoft.VisualBasic.DateInterval.Hour, pHour - pHourTZ, sReturn)
		sReturn = DateAdd(Microsoft.VisualBasic.DateInterval.Minute, pMinutes - pMinutesTZ, sReturn)
	End If
	If pPlusMinTZ = "-" Or pPlusMinTZ = "+" Then
		' it's in GMT, now move it to OPENWIKI_TIMEZONE
		If Left(OPENWIKI_TIMEZONE, 1) = "-" Then
			sReturn = DateAdd(Microsoft.VisualBasic.DateInterval.Hour, -1 * CDbl(Mid(OPENWIKI_TIMEZONE, 2, 2)), sReturn)
			sReturn = DateAdd(Microsoft.VisualBasic.DateInterval.Minute, -1 * CDbl(Mid(OPENWIKI_TIMEZONE, 5, 2)), sReturn)
		Else
			sReturn = DateAdd(Microsoft.VisualBasic.DateInterval.Hour, CDbl(Mid(OPENWIKI_TIMEZONE, 2, 2)), sReturn)
			sReturn = DateAdd(Microsoft.VisualBasic.DateInterval.Minute, CDbl(Mid(OPENWIKI_TIMEZONE, 5, 2)), sReturn)
		End If
	End If
End Sub


Function EscapePattern(ByRef pPattern As Object) As Object
	Dim vRegEx As Object
	pPattern = Replace(pPattern, "''''''", "")
	vRegEx = New RegExp
	vRegEx.IgnoreCase = True
	vRegEx.Global = True
	vRegEx.Pattern = pPattern
	On Error Resume Next
	Err.Clear()
	vRegEx.Test("x")
	If Err.Number <> 0 Then
		pPattern = Replace(pPattern, "\", "\\")
		pPattern = Replace(pPattern, "(", "\(")
		pPattern = Replace(pPattern, ")", "\)")
		pPattern = Replace(pPattern, "[", "\[")
		pPattern = Replace(pPattern, "+", "\+")
		pPattern = Replace(pPattern, "*", "\*")
		pPattern = Replace(pPattern, "?", "\?")
	End If
	
	'Response.Write("Pattern : " & pPattern & "<br />")
	EscapePattern = pPattern
End Function

Sub CutEquation(ByRef pText As Object)
	sReturn = "<ow:math><![CDATA[" & Replace(pText, "]]>", "]]&gt;") & "]]></ow:math>"
	gEquation = "<ow:math><![CDATA[" & Replace(pText, "]]>", "]]&gt;") & "]]></ow:math>"
End Sub

<%FormatDateISO8601 = Year(pTimestamp) & "-"
vTemp = Month(pTimestamp)
If vTemp < 10 Then
	FormatDateISO8601 = FormatDateISO8601 & "0"
End If
FormatDateISO8601 = FormatDateISO8601 & vTemp & "-"
vTemp = Microsoft.VisualBasic.Day(pTimestamp)
If vTemp < 10 Then
	FormatDateISO8601 = FormatDateISO8601 & "0"
End If
FormatDateISO8601 = FormatDateISO8601 & vTemp & "T"
vTemp = Hour(pTimestamp)
If vTemp < 10 Then
	FormatDateISO8601 = FormatDateISO8601 & "0"
End If
FormatDateISO8601 = FormatDateISO8601 & vTemp & ":"
vTemp = Minute(pTimestamp)
If vTemp < 10 Then
	FormatDateISO8601 = FormatDateISO8601 & "0"
End If
FormatDateISO8601 = FormatDateISO8601 & vTemp & ":"
vTemp = Second(pTimestamp)
If vTemp < 10 Then
	FormatDateISO8601 = FormatDateISO8601 & "0"
End If
FormatDateISO8601 = FormatDateISO8601 & vTemp
FormatDateISO8601 = FormatDateISO8601 & OPENWIKI_TIMEZONE

%>
