
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
'      $Source: /usr/local/cvsroot/openwiki/dist/owbase/ow/owactions.asp,v $
'    $Revision: 1.4 $
'      $Author: pit $
' ---------------------------------------------------------------------------
'

Sub ActionXml()
	ActionView()
End Sub


Sub ActionRss()
	Dim MultiLineMarkup As Object
	Dim vPage, vXmlStr As Object
	If cAllowRSSExport Then
		If CStr(Request("p")) <> "" And cAllowAggregations Then
			vPage = gNamespace.GetPage(gPage, gRevision, True, False)
			gAggregateURLs = New Vector
			gRaw = New Vector
			MultiLineMarkup(vPage.Text) ' refreshes RSS feed(s) and fills the gAggregateURLs vector
			If gAggregateURLs.Count = 0 Then
				Response.ContentType = "text/xml; charset:" & OPENWIKI_ENCODING & ";"
				Response.Write("<?xml version='1.0'?><error>Nothing to aggregate</error>")
				Response.End()
			Else
				Response.ContentType = "text/xml; charset:" & OPENWIKI_ENCODING & ";"
				Response.Write(gNamespace.GetAggregation(gAggregateURLs))
				Response.End()
			End If
		Else
			If cCacheXML Then
				vXmlStr = gNamespace.GetDocumentCache("rss")
			End If
			If vXmlStr = "" Then
				gPage = OPENWIKI_RCNAME
				vPage = gNamespace.GetPage(gPage, gRevision, False, False)
				' make sure we execute only the RecentChanges macro
				vPage.Text = "<RecentChangesLong>"
				vXmlStr = gTransformer.TransformXmlStr(vPage.ToXML(1), "owrss10export.xsl")
				If cCacheXML Then
					Call gNamespace.SetDocumentCache("rss", vXmlStr)
				End If
			End If
			gActionReturn = True
		End If
	Else
		Response.ContentType = "text/xml; charset:" & OPENWIKI_ENCODING & ";"
		Response.Write("<?xml version='1.0'?><error>RSS feed disabled</error>")
		Response.End()
	End If
End Sub


Sub ActionRefresh()
	Dim MultiLineMarkup As Object
	Dim vPage As Object
	If OPENWIKI_SCRIPTTIMEOUT > 0 Then
		Server.ScriptTimeOut = OPENWIKI_SCRIPTTIMEOUT
	End If
	cCacheXML = False
	vPage = gNamespace.GetPage(gPage, gRevision, True, False)
	gAggregateURLs = New Vector
	gRaw = New Vector
	Call MultiLineMarkup(vPage.Text) ' refreshes RSS feed(s)
	Call gNamespace.ClearDocumentCache2("", gPage)
	If CStr(Request("redirect")) = "" Then
		Response.Redirect(gScriptName & "?" & Server.URLEncode(gPage))
	Else
		Response.Redirect(gScriptName & "?" & Server.URLEncode(Request("redirect")))
	End If
End Sub


Sub ActionNaked()
	gAction = "view"
	ActionView()
End Sub


Sub ActionPrint()
	Dim vXmlStr As Object
	cReadOnly = 1
	If cCacheXML Then
		vXmlStr = gNamespace.GetDocumentCache("print")
	End If
	If vXmlStr = "" Then
		vXmlStr = gNamespace.GetPageAndAttachments(gPage, gRevision, True, False).ToXML(1)
		If cCacheXML Then
			Call gNamespace.SetDocumentCache("print", vXmlStr)
		End If
	End If
	Call gTransformer.Transform(vXmlStr)
	gActionReturn = True
End Sub


Sub ActionView()
	Dim vXmlStr As Object
	If cNakedView Then
		gAction = "naked"
	End If
	If cAllowRSSExport And CStr(Request("v")) = "rss" Then
		Call gTransformer.TransformXmlStr(gNamespace.GetPage(gPage, gRevision, True, False).ToXML(1), "owrss10export.xsl")
	Else
		If cCacheXML Then
			vXmlStr = gNamespace.GetDocumentCache("view")
		End If
		If vXmlStr = "" Then
			vXmlStr = gNamespace.GetPageAndAttachments(gPage, gRevision, True, False).ToXML(1)
			If cCacheXML Then
				Call gNamespace.SetDocumentCache("view", vXmlStr)
			End If
		End If
		Call gTransformer.Transform(vXmlStr)
	End If
	gActionReturn = True
End Sub


Sub ActionPreview()
	Dim vPage As Object
	vPage = gNamespace.GetPage(gPage, 0, False, False)
	vPage.Text = Request("text")
	gAction = "naked"
	Call gTransformer.Transform(vPage.ToXML(1))
	gActionReturn = True
End Sub


Sub ActionDiff()
	Dim vPageTo, vDiffType, vDiffFrom, vXmlStr, vDiff, vDiffTo, vPageFrom, vMatcher As Object
	vDiff = GetIntParameter("diff")
	vDiffFrom = GetIntParameter("difffrom")
	vDiffTo = GetIntParameter("diffto")
	
	If vDiffFrom <> 0 Or vDiffTo <> 0 Then
		cCacheXML = False
	End If
	
	If cCacheXML Then
		vXmlStr = gNamespace.GetDocumentCache("diff" & vDiff)
	End If
	
	If vXmlStr = "" Then
		
		If vDiff = 0 Then
			If vDiffFrom = 0 Then
				' difference of prior major revision relative to vDiffTo
				vDiffType = "major"
				vDiffFrom = gNamespace.GetPreviousRevision(0, vDiffTo)
			Else
				' difference of selected revision relative to vDiffTo
				vDiffType = "selected"
			End If
		ElseIf vDiff = 1 Then 
			' difference of previous minor edit relative to vDiffTo
			vDiffType = "minor"
			vDiffFrom = gNamespace.GetPreviousRevision(1, vDiffTo)
		Else
			' difference of previous author edit relative to vDiffTo
			vDiffType = "author"
			vDiffFrom = gNamespace.GetPreviousRevision(2, vDiffTo)
		End If
		
		' difference of vDiffFrom to vDiffTo
		vPageFrom = gNamespace.GetPage(gPage, vDiffFrom, True, False)
		vPageTo = gNamespace.GetPageAndAttachments(gPage, vDiffTo, True, False)
		vDiffFrom = vPageFrom.GetLastChange().Revision
		vDiffTo = vPageTo.GetLastChange().Revision
		vXmlStr = "<ow:diff type='" & vDiffType & "' from='" & vDiffFrom & "' to='" & vDiffTo & "'>"
		If vDiffTo > vDiffFrom Then
			vMatcher = New Matcher
			vXmlStr = vXmlStr & vMatcher.Compare(Server.HTMLEncode(vPageFrom.Text), Server.HTMLEncode(vPageTo.Text))
		End If
		vXmlStr = vXmlStr & "</ow:diff>"
		vXmlStr = vXmlStr & vPageTo.ToXML(1)
		
		If cCacheXML Then
			Call gNamespace.SetDocumentCache("diff" & vDiff, vXmlStr)
		End If
	End If
	
	Call gTransformer.Transform(vXmlStr)
	'UPGRADE_NOTE: Object vMatcher may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
	vMatcher = Nothing
	'UPGRADE_NOTE: Object vPageTo may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
	vPageTo = Nothing
	'UPGRADE_NOTE: Object vPageFrom may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
	vPageFrom = Nothing
	
	gActionReturn = True
End Sub


Sub ActionEdit()
	Dim ASPTypeLibrary As Object
	Dim vChange, vPage, vXmlStr As Object
	Dim vComment, vNewRev, vMinorEdit, vText As Object
	Dim CaptchaCheck As Object
	
	If cReadOnly Then
		' TODO: generate <ow:error> tag into the XML output
		gAction = "view"
		ActionView()
		gActionReturn = True
		Exit Sub
	End If
	
	If gEditPassword <> "" Then
		If gEditPassword <> gReadPassword Then
			If Request.Cookies(gCookieHash & "?pe") <> gEditPassword Then
				If (cUseRecaptcha <> "1") Or (m(gPage, OPENWIKI_PROTECTEDPAGES, False, False)) Then
					Call ActionLogin()
					Exit Sub
				End If
			End If
		End If
	End If
	
	Dim vBacklink As Object
	If Request("save").Item <> "" Then
		vNewRev = Int(CDbl(Request("newrev")))
		vMinorEdit = Int(CDbl(Request("rc").Item)) Xor 1
		vComment = Trim(Request("comment").Item & "")
		vText = Request("text").Item
		
		If Len(CStr(vComment)) > 1000 Then
			vXmlStr = vXmlStr & "<ow:error code='1'>Maximum length for the comment is 1000 characters.</ow:error>"
		End If
		If Len(CStr(vText)) > OPENWIKI_MAXTEXT Then
			vXmlStr = vXmlStr & "<ow:error code='2'>Maximum length for the text is " & OPENWIKI_MAXTEXT & " characters.</ow:error>"
		End If
		
		If (Not m(gPage, OPENWIKI_PROTECTEDPAGES, False, False)) And cUseRecaptcha Then
			CaptchaCheck = RecaptchaConfirm(Request("recaptcha_challenge_field"), Request("recaptcha_response_field"))
			If CaptchaCheck <> "" Then
				vXmlStr = vXmlStr & "<ow:captcha_error>" & CaptchaCheck & "</ow:captcha_error>"
				vXmlStr = vXmlStr & "<ow:error code='5'>reCAPTCHA error. See details in reCAPTCHA form.</ow:error>"
			End If
		End If
		
		If vXmlStr <> "" Then
			vPage = gNamespace.GetPage(gPage, 0, False, False)
			vPage.Revision = gRevision
			vPage.Text = vText
			
			vChange = vPage.GetLastChange()
			vChange.Revision = vNewRev
			vChange.MinorEdit = vMinorEdit
			vChange.Comment = vComment
			vChange.Timestamp = Now
			vChange.UpdateBy()
			
			vXmlStr = vXmlStr & vPage.ToXML(2)
		ElseIf gNamespace.SavePage(vNewRev, vMinorEdit, vComment, vText) Then 
			Response.Redirect(gScriptName & "?" & Server.URLEncode(gPage))
		Else
			vPage = gNamespace.GetPage(gPage, 0, True, False)
			vChange = vPage.GetLastChange()
			vChange.Revision = vChange.Revision + 1
			vChange.MinorEdit = Int(CDbl(Request("rc").Item)) Xor 1
			vChange.Comment = Trim(Request("comment").Item & "")
			vChange.Timestamp = Now
			vChange.UpdateBy()
			vXmlStr = vXmlStr & "<ow:error code='4'>Somebody else just edited this page.</ow:error>"
			vXmlStr = vXmlStr & "<ow:textedits>" & PCDATAEncode(Request("text")) & "</ow:textedits>"
			vXmlStr = vXmlStr & vPage.ToXML(2)
		End If
		' now v0.78, let's see if someone's going to complain..
		'Elseif Request("preview") <> "" Then
		' pre 0.74 version code; now ActionPreview (i.e. ?a=preview is prefered method)
		
		'    vNewRev    = Int(Request("newrev"))
		'    vMinorEdit = Int(Request("rc")) Xor 1
		'    vComment   = Trim(Request("comment") & "")
		'    vText      = Request("text")
		'
		'    Set vPage = gNamespace.GetPage(gPage, 0, False, False)
		'    vPage.Revision = gRevision
		'    vPage.Text     = vText
		'
		'    Set vChange = vPage.GetLastChange()
		'    vChange.Revision  = vNewRev
		'    vChange.MinorEdit = vMinorEdit
		'    vChange.Comment   = vComment
		'    vChange.Timestamp = Now()
		'    vChange.UpdateBy()
		'
		'    vXmlStr = vPage.ToXML(3)
	ElseIf Request("cancel").Item <> "" Then 
		If gRevision = 0 Then
			vBacklink = gScriptName & "?" & Server.URLEncode(gPage)
		Else
			vBacklink = gScriptName & "?p=" & Server.URLEncode(gPage) & "&revision=" & gRevision
		End If
		Response.Redirect(vBacklink)
	Else
		' first time opening edit form
		vPage = gNamespace.GetPage(gPage, 0, True, False)
		If gRevision > 0 Then
			gTemp = gNamespace.GetPage(gPage, gRevision, True, False)
			vPage.Revision = gTemp.Revision
			vPage.Text = gTemp.Text
		End If
		
		If vPage.Revision = 0 And Request("template").Item <> "" Then
			gTemp = gNamespace.GetPage(URLDecode(Request("template")), 0, True, False)
			vPage.Text = gTemp.Text
		End If
		
		vChange = vPage.GetLastChange()
		vChange.Revision = vChange.Revision + 1
		vChange.MinorEdit = 0
		vChange.Comment = ""
		vChange.Timestamp = Now
		vChange.UpdateBy()
		
		vXmlStr = vPage.ToXML(2)
	End If
	
	Call gTransformer.Transform(vXmlStr)
	gActionReturn = True
End Sub


Sub ActionTitleSearch()
	Dim vXmlStr As Object
	vXmlStr = gNamespace.GetIndexSchemes.GetTitleSearch(gTxt)
	If cAllowRSSExport And Request("v").Item = "rss" Then
		Call gTransformer.TransformXmlStr(vXmlStr, "owsearchrss10export.xsl")
	Else
		Call gTransformer.Transform(vXmlStr)
	End If
	gActionReturn = True
End Sub


Sub ActionFullSearch()
	Dim vXmlStr As Object
	vXmlStr = gNamespace.GetIndexSchemes.GetFullSearch(gTxt, True)
	If cAllowRSSExport And Request("v").Item = "rss" Then
		Call gTransformer.TransformXmlStr(vXmlStr, "owsearchrss10export.xsl")
	Else
		Call gTransformer.Transform(vXmlStr)
	End If
	gActionReturn = True
End Sub


Sub ActionTextSearch()
	Dim vXmlStr As Object
	vXmlStr = gNamespace.GetIndexSchemes.GetFullSearch(gTxt, False)
	If cAllowRSSExport And Request("v").Item = "rss" Then
		Call gTransformer.TransformXmlStr(vXmlStr, "owsearchrss10export.xsl")
	Else
		Call gTransformer.Transform(vXmlStr)
	End If
	gActionReturn = True
End Sub


Sub ActionRandomPage()
	Randomize()
	If cUseSpecialPagesPrefix Then
		gTemp = gNamespace.TitleSearch("^(?!" & gSpecialPagesPrefix & ")" & ".*", 0, 0, 0, 0)
	Else
		gTemp = gNamespace.TitleSearch(".*", 0, 0, 0, 0)
	End If
	
	'    Response.Redirect(gScriptName & "?a=" & gAction & "&p=" & Server.URLEncode(gTemp.ElementAt(Int((gTemp.Count - 1) * Rnd)).Name) & "&redirect=" & Server.URLEncode(gPage))
	Response.Redirect(gScriptName & "?a=" & gAction & "&p=" & Server.URLEncode(gTemp.ElementAt(Int((gTemp.Count - 1) * Rnd())).Name))
End Sub


Sub ActionChanges()
	Dim vXmlStr As Object
	If cCacheXML Then
		vXmlStr = gNamespace.GetDocumentCache("changes")
	End If
	If vXmlStr = "" Then
		vXmlStr = gNamespace.GetPage(gPage, 0, False, True).ToXML(0)
		If cCacheXML Then
			Call gNamespace.SetDocumentCache("changes", vXmlStr)
		End If
	End If
	Call gTransformer.Transform(vXmlStr)
	gActionReturn = True
End Sub


Sub ActionUserPreferences()
	If Request("save").Item <> "" Then
		'UPGRADE_WARNING: Date was upgraded to Today and has a new behavior. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1041.asp'
		'UPGRADE_NOTE: Date operands have a different behavior in arithmetical operations
		Response.Cookies(gCookieHash & "?up").Expires = System.Date.FromOADate(Today.ToOADate + 60)
		Response.Cookies(gCookieHash & "?up")("un") = FreeToNormal_X(Request("username"), False)
		Response.Cookies(gCookieHash & "?up")("bm") = Request("bookmarks").Item
		Response.Cookies(gCookieHash & "?up")("cols") = Request("cols").Item
		Response.Cookies(gCookieHash & "?up")("rows") = Request("rows").Item
		Response.Cookies(gCookieHash & "?up")("pwl") = Request("prettywikilinks").Item
		Response.Cookies(gCookieHash & "?up")("bmt") = Request("bookmarksontop").Item
		Response.Cookies(gCookieHash & "?up")("elt") = Request("editlinkontop").Item
		Response.Cookies(gCookieHash & "?up")("trt") = Request("trailontop").Item
		Response.Cookies(gCookieHash & "?up")("new") = Request("opennew").Item
		Response.Cookies(gCookieHash & "?up")("emo") = Request("emoticons").Item
		Response.Redirect(gScriptName & "?p=" & Server.URLEncode(gPage) & "&up=1")
	ElseIf Request("clear").Item <> "" Then 
		Response.Cookies(gCookieHash & "?up").expires  = #01.01.1990#
		'UPGRADE_ISSUE: The preceding line couldn't be parsed. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1010.asp'
		Response.Cookies(gCookieHash & "?up") = ""
		Response.Redirect(gScriptName & "?p=" & Server.URLEncode(gPage) & "&up=2")
	End If
	gActionReturn = False
End Sub


Sub ActionLogout()
	Response.Cookies(gCookieHash & "?pr").Expires = #01.01.1990#
	'UPGRADE_ISSUE: The preceding line couldn't be parsed. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1010.asp'
	Response.Cookies(gCookieHash & "?pr") = ""
	Response.Cookies(gCookieHash & "?pe").Expires = #01.01.1990#
	'UPGRADE_ISSUE: The preceding line couldn't be parsed. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1010.asp'
	Response.Cookies(gCookieHash & "?pe") = ""
	Response.Redirect(gScriptName & "?" & Server.URLEncode(gPage))
End Sub


Sub ActionLogin()
	Dim vPwd, vMode, vXmlStr As Object
	If gAction = "edit" Then
		vMode = "edit"
		gAction = "login"
	Else
		vMode = Request("mode").Item
	End If
	vPwd = Request("pwd").Item
	If vMode = "edit" Then
		If vPwd = gEditPassword Then
			If Request("r").Item = "1" Then
				'UPGRADE_WARNING: Date was upgraded to Today and has a new behavior. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1041.asp'
				'UPGRADE_NOTE: Date operands have a different behavior in arithmetical operations
				Response.Cookies(gCookieHash & "?pe").Expires = System.Date.FromOADate(Today.ToOADate + 60)
			End If
			Response.Cookies(gCookieHash & "?pe") = vPwd
			Response.Redirect(gScriptName & "?" & Request("backlink").Item)
		End If
	Else
		If vPwd = gReadPassword Then
			If Request("r").Item = "1" Then
				'UPGRADE_WARNING: Date was upgraded to Today and has a new behavior. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1041.asp'
				'UPGRADE_NOTE: Date operands have a different behavior in arithmetical operations
				Response.Cookies(gCookieHash & "?pr").Expires = System.Date.FromOADate(Today.ToOADate + 60)
			End If
			Response.Cookies(gCookieHash & "?pr") = vPwd
			Response.Redirect(gScriptName & "?" & Request("backlink").Item)
		End If
	End If
	If vPwd <> "" Then
		vXmlStr = "<ow:error code='3'>Incorrect password</ow:error>"
	End If
	If Request("backlink").Item <> "" Then
		gTemp = Request("backlink").Item
	Else
		gTemp = Request.ServerVariables("QUERY_STRING")
		If gTemp = "" Then
			gTemp = OPENWIKI_FRONTPAGE
		End If
	End If
	vXmlStr = vXmlStr & "<ow:login"
	If vMode = "edit" Then
		vXmlStr = vXmlStr & " mode='edit'>"
	Else
		vXmlStr = vXmlStr & " mode='view'>"
	End If
	vXmlStr = vXmlStr & "<ow:backlink>" & PCDATAEncode(gTemp) & "</ow:backlink>"
	If Request("r").Item <> "" Then
		vXmlStr = vXmlStr & "<ow:rememberme>true</ow:rememberme>"
	End If
	vXmlStr = vXmlStr & "</ow:login>"
	Call gTransformer.Transform(vXmlStr)
	gActionReturn = True
End Sub


