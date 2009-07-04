<script language="VB" runat="Server">Dim vPos, vPos2 As Object


Dim gCookieTrail As Object

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
'      $Source: /usr/local/cvsroot/openwiki/dist/owbase/ow/owprocessor.asp,v $
'    $Revision: 1.3 $
'      $Author: pit $
' ---------------------------------------------------------------------------
'

Sub OwProcessRequest()
	Dim Execute As Object
	Dim ParseQueryString As Object
	Dim SERVER_PORT, SCRIPT_NAME, SERVER_NAME, SERVER_PORT_SECURE As Object
	SCRIPT_NAME = Request.ServerVariables("SCRIPT_NAME")
	SERVER_NAME = Request.ServerVariables("SERVER_NAME")
	SERVER_PORT = Request.ServerVariables("SERVER_PORT")
	SERVER_PORT_SECURE = Request.ServerVariables("SERVER_PORT_SECURE")
	
	If SERVER_PORT_SECURE = 0 Then
		gServerRoot = "http://" & SERVER_NAME
	Else
		gServerRoot = "https://" & SERVER_NAME
	End If
	If SERVER_PORT <> 80 Then
		gServerRoot = gServerRoot & ":" & SERVER_PORT
	End If
	gServerRoot = gServerRoot & Left(SCRIPT_NAME, InStrRev(SCRIPT_NAME, "/"))
	
	If OPENWIKI_SCRIPTNAME <> "" Then
		gScriptName = OPENWIKI_SCRIPTNAME
	Else
		gTemp = InStrRev(SCRIPT_NAME, "/")
		If gTemp > 0 Then
			gScriptName = Mid(SCRIPT_NAME, gTemp + 1)
		Else
			gScriptName = SCRIPT_NAME
		End If
	End If
	
	gCookieHash = "C" & Hash(gServerRoot & SCRIPT_NAME)
	
	If Request.Cookies(gCookieHash & "?up") <> "" Then
		If Request.Cookies(gCookieHash & "?up")("pwl") = "1" Then
			cPrettyLinks = 1
		Else
			cPrettyLinks = 0
		End If
		If Request.Cookies(gCookieHash & "?up")("new") = "1" Then
			cExternalOut = 1
		Else
			cExternalOut = 0
		End If
		If Request.Cookies(gCookieHash & "?up")("emo") = "1" Then
			cEmoticons = 1
		Else
			cEmoticons = 0
		End If
	End If
	
	gTransformer = New Transformer
	gNamespace = New OpenWikiNamespace
	
	InitLinkPatterns()
	ParseQueryString()
	
	If gReadPassword <> "" Then
		If gEditPassword = "" Then
			gEditPassword = gReadPassword
		End If
		gTemp = Request.Cookies(gCookieHash & "?pr")
		If gTemp <> gReadPassword Then
			gAction = "login"
		End If
	End If
	
	If Not m(OPENWIKI_TIMEZONE, "^[+|-](0\d|1[0-2]):[0-5]\d$", False, False) Then
		OPENWIKI_TIMEZONE = "+00:00"
	End If
	
	gActionReturn = False
	Execute("Call Action" & gAction)
	If Not gActionReturn Then
		Response.ContentType = "text/xml; charset:" & OPENWIKI_ENCODING & ";"
		Response.Write("<?xml version='1.0'?><error>Illegal action</error>")
		Response.End()
	End If
	
	'UPGRADE_NOTE: Object gTransformer may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
	gTransformer = Nothing
	'UPGRADE_NOTE: Object gNamespace may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
	gNamespace = Nothing
End Sub


Function TransformEmbedded(ByRef pText As Object) As Object
	Dim vPage As Object
	
	gScriptName = OPENWIKI_SCRIPTNAME
	gTransformer = New Transformer
	gNamespace = New OpenWikiNamespace
	gAction = "embedded"
	
	InitLinkPatterns()
	
	vPage = New WikiPage
	vPage.AddChange()
	vPage.Text = pText
	TransformEmbedded = gTransformer.Transform(vPage.ToXML(1))
End Function

'End Sub



Function GetIntParameter(ByRef pParam As Object) As Object
	GetIntParameter = Request(pParam)
	If IsNumeric(GetIntParameter) Then
		GetIntParameter = Int(GetIntParameter)
	Else
		GetIntParameter = 0
	End If
End Function



Function getUserPreferences() As Object
	Dim vValue, vMatches, vRegEx, vMatch, vUsername As Object
	Dim vRows, vCols, vBookmarks As Object
	vCols = Request.Cookies(gCookieHash & "?up")("cols")
	If vCols <= 0 Then
		vCols = 55
	End If
	vRows = Request.Cookies(gCookieHash & "?up")("rows")
	If vRows <= 0 Then
		vRows = 25
	End If
	vBookmarks = Request.Cookies(gCookieHash & "?up")("bm")
	If vBookmarks = "" Then
		vBookmarks = gDefaultBookmarks
	End If
	vRegEx = New RegExp
	vRegEx.IgnoreCase = False
	vRegEx.Global = True
	vRegEx.Pattern = "\s+([^ ]*)"
	vMatches = vRegEx.Execute(" " & Trim(vBookmarks))
	vBookmarks = ""
	For	Each vMatch In vMatches
		vValue = Mid(vMatch.Value, 2)
		vBookmarks = vBookmarks & ToLinkXML(vValue)
	Next vMatch
	vBookmarks = "<ow:bookmarks>" & vBookmarks & "</ow:bookmarks>"
	'UPGRADE_NOTE: Object vRegEx may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
	vRegEx = Nothing
	'UPGRADE_NOTE: Object vMatches may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
	vMatches = Nothing
	'UPGRADE_NOTE: Object vMatch may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
	vMatch = Nothing
	
	If Request.Cookies(gCookieHash & "?up") = "" Then
		If cPrettyLinks Then
			getUserPreferences = getUserPreferences & "<ow:prettywikilinks/>"
		End If
		If cExternalOut Then
			getUserPreferences = getUserPreferences & "<ow:opennew/>"
		End If
		If cEmoticons Then
			getUserPreferences = getUserPreferences & "<ow:emoticons/>"
		End If
		getUserPreferences = getUserPreferences & "<ow:bookmarksontop/><ow:editlinkontop/><ow:trailontop/>"
	Else
		If Request.Cookies(gCookieHash & "?up")("pwl") = "1" Then
			getUserPreferences = getUserPreferences & "<ow:prettywikilinks/>"
		End If
		If Request.Cookies(gCookieHash & "?up")("bmt") = "1" Then
			getUserPreferences = getUserPreferences & "<ow:bookmarksontop/>"
		End If
		If Request.Cookies(gCookieHash & "?up")("elt") = "1" Then
			getUserPreferences = getUserPreferences & "<ow:editlinkontop/>"
		End If
		If Request.Cookies(gCookieHash & "?up")("trt") = "1" Then
			getUserPreferences = getUserPreferences & "<ow:trailontop/>"
		End If
		If Request.Cookies(gCookieHash & "?up")("new") = "1" Then
			getUserPreferences = getUserPreferences & "<ow:opennew/>"
		End If
		If Request.Cookies(gCookieHash & "?up")("emo") = "1" Then
			getUserPreferences = getUserPreferences & "<ow:emoticons/>"
		End If
	End If
	
	vUsername = Request.Cookies(gCookieHash & "?up")("un")
	If cNTAuthentication = 1 And vUsername = "" Then
		vUsername = GetRemoteUser()
	End If
	
	getUserPreferences = "<ow:userpreferences>" & "<ow:cols>" & vCols & "</ow:cols>" & "<ow:rows>" & vRows & "</ow:rows>" & "<ow:username>" & vUsername & "</ow:username>" & vBookmarks & getUserPreferences & "</ow:userpreferences>"
End Function

Sub AddCookieTrail(ByRef pPage As Object)
	gCookieTrail.Push(pPage)
End Sub


Function GetCookieTrail() As Object
	Dim vElem, vCount, vTrailStr, vLast, vExists, i As Object
	
	vTrailStr = Request.Cookies(gCookieHash & "?tr")("trail")
	
	gCookieTrail = New Vector
	Call s(vTrailStr, "#(.*?)#", "&AddCookieTrail($1)", False, True)
	
	vTrailStr = ""
	vExists = False
	vCount = gCookieTrail.Count
	For i = 1 To vCount - 1
		vElem = gCookieTrail.ElementAt(i)
		If vElem = gPage Then
			vExists = True
		Else
			GetCookieTrail = GetCookieTrail & ToLinkXML(vElem)
			vTrailStr = vTrailStr & "#" & vElem & "#"
		End If
	Next 
	If vExists Or (vCount < OPENWIKI_MAXTRAIL) Then
		If vCount > 0 Then
			vElem = gCookieTrail.ElementAt(0)
			If vElem <> gPage Then
				GetCookieTrail = ToLinkXML(vElem) & GetCookieTrail
				vTrailStr = "#" & vElem & "#" & vTrailStr
			End If
		End If
		If gPage <> "" Then
			vElem = gPage
			GetCookieTrail = GetCookieTrail & ToLinkXML(vElem)
			vTrailStr = vTrailStr & "#" & vElem & "#"
		End If
	ElseIf vCount > 0 Then 
		vElem = gPage
		GetCookieTrail = GetCookieTrail & ToLinkXML(vElem)
		vTrailStr = vTrailStr & "#" & vElem & "#"
	End If
	
	
	Response.Cookies(gCookieHash & "?tr")("trail") = vTrailStr
	Response.Cookies(gCookieHash & "?tr")("last") = gPage
	
	'UPGRADE_NOTE: Object gCookieTrail may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
	gCookieTrail = Nothing
	GetCookieTrail = "<ow:trail>" & GetCookieTrail & "</ow:trail>"
End Function


Function ToLinkXML(ByRef pID As Object) As Object
	Dim PrettyWikiLink As Object
	Dim vTemp As Object
	If gAction = "print" Then
		vTemp = gScriptName & "?p=" & Server.URLEncode(pID) & "&amp;a=print"
	Else
		vTemp = gScriptName & "?" & Server.URLEncode(pID)
	End If
	ToLinkXML = "<ow:link name='" & CDATAEncode(pID) & "' href='" & vTemp & "'>" & PCDATAEncode(PrettyWikiLink(pID)) & "</ow:link>"
End Function



Function GetCookieTrail_Alternative() As Object
	Dim PrettyWikiLink As Object
	Dim vStart, i, vExists, vLast, vTrailStr, vCount, vElem, vPosLast, vEnd As Object
	
	vTrailStr = Request.Cookies(gCookieHash & "?tr")("trail")
	vLast = Request.Cookies(gCookieHash & "?tr")("last")
	
	gCookieTrail = New Vector
	Call s(vTrailStr, "#(.*?)#", "&AddCookieTrail($1)", False, True)
	
	vExists = False
	vCount = gCookieTrail.Count
	vPosLast = OPENWIKI_MAXTRAIL
	For i = 0 To vCount - 1
		vElem = gCookieTrail.ElementAt(i)
		If vElem = gPage Then
			vExists = True
		End If
		If vElem = vLast Then
			vPosLast = i
		End If
	Next 
	
	If vExists Then
		vStart = 0
		vEnd = vCount - 1
	ElseIf vPosLast < (OPENWIKI_MAXTRAIL - 1) Then 
		vStart = 0
		vEnd = vPosLast
	Else
		vStart = 1
		vEnd = vCount - 1
	End If
	
	vTrailStr = ""
	For i = vStart To vEnd
		vElem = gCookieTrail.ElementAt(i)
		GetCookieTrail = GetCookieTrail & "<ow:trailmark name='" & CDATAEncode(vElem) & "'>" & PCDATAEncode(PrettyWikiLink(vElem)) & "</ow:trailmark>"
		vTrailStr = vTrailStr & "#" & vElem & "#"
	Next 
	
	If (Not vExists) And ((vEnd - vStart + 1) < OPENWIKI_MAXTRAIL) Then
		vElem = gPage
		GetCookieTrail = GetCookieTrail & "<ow:trailmark name='" & CDATAEncode(vElem) & "'>" & PCDATAEncode(PrettyWikiLink(vElem)) & "</ow:trailmark>"
		vTrailStr = vTrailStr & "#" & vElem & "#"
		Response.Cookies(gCookieHash & "?tr")("trail") = vTrailStr
	End If
	
	Response.Cookies(gCookieHash & "?tr")("last") = gPage
	
	'UPGRADE_NOTE: Object gCookieTrail may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
	gCookieTrail = Nothing
	GetCookieTrail = "<ow:trail>" & GetCookieTrail & "</ow:trail>"
End Function



Private Function Hash(ByRef pText As Object) As Object
	Dim vCount, i, vMax As Object
	vMax = 2 ^ 30
	Hash = 0
	vCount = Len(CStr(pText))
	For i = 1 To vCount
		If Hash > vMax Then
			Hash = Hash - vMax
			Hash = Hash * 2
			Hash = Hash Or 1
		Else
			Hash = Hash * 2
		End If
		Hash = Hash Xor AscW(Mid(pText, i, 1))
	Next 
	If Hash = 0 Then
		Hash = 1
	End If
End Function

<%


' As you may notice I never use Request.Query and/or Request.Form, only Request(vSomeParam).
' The reason is that I plan to support platforms where submitted data cannot always go
' through a form, but only through the use of a URL. Another reason is that the presentation
' layer is separated from the logic, therefor no assumption should be made about whether the
' parameters are passed through the URL or posted as a form.
'_____________________________________________________________________________________________________________Sub ParseQueryString()
gPage = Request("p")
If gPage = "" Then
	gPage = Request.ServerVariables("QUERY_STRING")
	vPos = InStr(gPage, "&")
	vPos2 = InStr(gPage, "=")
	If vPos2 <= 0 Or vPos2 > vPos Then
		If vPos > 0 Then
			'Dim vArgs
			'vArgs = Mid(gPage, vPos)
			'Call s(vArgs, "\&(.*?)[^\&]", "&AddParameter($1,$2)", True, True)
			gPage = Left(gPage, vPos - 1)
		ElseIf gPage = "" Then 
			' ow.asp?, no parameters passed at all
			gPage = OPENWIKI_FRONTPAGE
		ElseIf vPos2 > 0 Then 
			' ow.asp?a=login, no page parameter
			gPage = "" ' gNamespace.Frontpage
		End If
	Else
		' ow.asp?foo=bar, no page posted, rescue to the frontpage
		gPage = OPENWIKI_FRONTPAGE
	End If
End If

gPage = URLDecode(gPage)

' determine MainPage/SubPage
vPos = InStr(gPage, "/")
If vPos = 1 Then
	gPage = OPENWIKI_FRONTPAGE & gPage
End If

gRevision = GetIntParameter("revision")

gAction = Request("a")
If gAction = "" Then
	gAction = "view"
End If

If CStr(Request("refresh")) <> "" Then
	cCacheXML = False
End If

gTxt = Request("txt")


%>
