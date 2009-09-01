Namespace Openwiki
    Module Actions
        Sub ActionXml()
            ActionView()
        End Sub

        Sub ActionRss()
            Dim vPage As WikiPage
            Dim vXmlStr As String = ""

            If cAllowRSSExport = 1 Then
                If HttpContext.Current.Request("p") <> "" And cAllowAggregations = 1 Then
                    vPage = gNamespace.GetPage(gPage, gRevision, 1, False)
                    gAggregateURLs = New Vector
                    gRaw = New Vector
                    MultiLineMarkup(vPage.Text)   ' refreshes RSS feed(s) and fills the gAggregateURLs vector
                    If gAggregateURLs.Count = 0 Then
                        HttpContext.Current.Response.ContentType = "text/xml; charset:" & OPENWIKI_ENCODING & ";"
                        HttpContext.Current.Response.Write("<?xml version='1.0'?><error>Nothing to aggregate</error>")
                        HttpContext.Current.Response.End()
                    Else
                        HttpContext.Current.Response.ContentType = "text/xml; charset:" & OPENWIKI_ENCODING & ";"
                        HttpContext.Current.Response.Write(gNamespace.GetAggregation(gAggregateURLs))
                        HttpContext.Current.Response.End()
                    End If
                Else
                    If cCacheXML = 1 Then
                        vXmlStr = gNamespace.GetDocumentCache("rss")
                    End If
                    If vXmlStr = "" Then
                        gPage = OPENWIKI_RCNAME
                        vPage = gNamespace.GetPage(gPage, gRevision, 0, False)
                        ' make sure we execute only the RecentChanges macro
                        vPage.Text = "<RecentChangesLong>"
                        vXmlStr = gTransformer.TransformXmlStr(vPage.ToXML(1), "owrss10export.xsl")
                        If cCacheXML = 1 Then
                            Call gNamespace.SetDocumentCache("rss", vXmlStr)
                        End If
                    End If
                    gActionReturn = True
                End If
            Else
                HttpContext.Current.Response.ContentType = "text/xml; charset:" & OPENWIKI_ENCODING & ";"
                HttpContext.Current.Response.Write("<?xml version='1.0'?><error>RSS feed disabled</error>")
                HttpContext.Current.Response.End()
            End If
        End Sub

        Sub ActionRefresh()
            Dim vPage As WikiPage

            If OPENWIKI_SCRIPTTIMEOUT > 0 Then
                HttpContext.Current.Server.ScriptTimeout = OPENWIKI_SCRIPTTIMEOUT
            End If
            cCacheXML = 0
            vPage = gNamespace.GetPage(gPage, gRevision, 1, False)
            gAggregateURLs = New Vector
            gRaw = New Vector
            MultiLineMarkup(vPage.Text)   ' refreshes RSS feed(s)
            gNamespace.ClearDocumentCache2(Nothing, gPage)
            If HttpContext.Current.Request("redirect") = "" Then
                HttpContext.Current.Response.Redirect(gScriptName & "?" & HttpContext.Current.Server.UrlEncode(gPage))
            Else
                HttpContext.Current.Response.Redirect(gScriptName & "?" & HttpContext.Current.Server.UrlEncode(HttpContext.Current.Request("redirect")))
            End If
        End Sub

        Sub ActionNaked()
            gAction = "view"
            ActionView()
        End Sub

        Sub ActionPrint()
            Dim vXmlStr As String = ""

            cReadOnly = 1
            'If cCacheXML Then
            '    vXmlStr = gNamespace.GetDocumentCache("print")
            'End If
            If vXmlStr = "" Then
                vXmlStr = gNamespace.GetPageAndAttachments(gPage, gRevision, 1, False).ToXML(1)
                'If cCacheXML Then
                '    Call gNamespace.SetDocumentCache("print", vXmlStr)
                'End If
            End If
            gTransformer.Transform(vXmlStr)
            gActionReturn = True
        End Sub

        Sub ActionView()
            Dim vXmlStr As String = ""

            If cNakedView = 1 Then
                gAction = "naked"
            End If
            If cAllowRSSExport = 1 And HttpContext.Current.Request("v") = "rss" Then
                Call gTransformer.TransformXmlStr(gNamespace.GetPage(gPage, gRevision, 1, False).ToXML(1), "owrss10export.xsl")
            Else
                If cCacheXML = 1 Then
                    vXmlStr = gNamespace.GetDocumentCache("view")
                End If
                If vXmlStr = "" Then
                    vXmlStr = gNamespace.GetPageAndAttachments(gPage, gRevision, 1, False).ToXML(1)
                    If cCacheXML = 1 Then
                        Call gNamespace.SetDocumentCache("view", vXmlStr)
                    End If
                End If
                Call gTransformer.Transform(vXmlStr)
            End If
            gActionReturn = True
        End Sub

        Sub ActionPreview()
            Dim vPage As WikiPage
            vPage = gNamespace.GetPage(gPage, 0, 0, False)
            vPage.Text = HttpContext.Current.Request("text")
            gAction = "naked"
            Call gTransformer.Transform(vPage.ToXML(1))
            gActionReturn = True
        End Sub

        Sub ActionDiff()
            Dim vXmlStr As String = ""
            Dim vDiff As Integer
            Dim vDiffFrom As Integer
            Dim vDiffTo As Integer
            Dim vDiffType As String
            Dim vPageFrom As WikiPage
            Dim vPageTo As WikiPage
            Dim vMatcher As Matcher

            vDiff = GetIntParameter("diff")
            vDiffFrom = GetIntParameter("difffrom")
            vDiffTo = GetIntParameter("diffto")

            If vDiffFrom <> 0 Or vDiffTo <> 0 Then
                cCacheXML = 0
            End If

            If cCacheXML = 1 Then
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
                vPageFrom = gNamespace.GetPage(gPage, vDiffFrom, 1, False)
                vPageTo = gNamespace.GetPageAndAttachments(gPage, vDiffTo, 1, False)
                vDiffFrom = vPageFrom.GetLastChange().Revision
                vDiffTo = vPageTo.GetLastChange().Revision
                vXmlStr = "<ow:diff type='" & vDiffType & "' from='" & vDiffFrom & "' to='" & vDiffTo & "'>"
                If vDiffTo > vDiffFrom Then
                    vMatcher = New Matcher
                    vXmlStr = vXmlStr & vMatcher.Compare(HttpContext.Current.Server.HtmlEncode(vPageFrom.Text), HttpContext.Current.Server.HtmlEncode(vPageTo.Text))
                End If
                vXmlStr = vXmlStr & "</ow:diff>"
                vXmlStr = vXmlStr & vPageTo.ToXML(1)

                If cCacheXML = 1 Then
                    Call gNamespace.SetDocumentCache("diff" & vDiff, vXmlStr)
                End If
            End If

            Call gTransformer.Transform(vXmlStr)
            vMatcher = Nothing
            vPageTo = Nothing
            vPageFrom = Nothing

            gActionReturn = True
        End Sub

        Sub ActionEdit()
            Dim vPage As WikiPage
            Dim vChange As Change
            Dim vXmlStr As String = ""
            Dim vNewRev As Integer
            Dim vMinorEdit As Integer
            Dim vComment As String
            Dim vText As String
            Dim CaptchaCheck As String
            Dim vTemp As WikiPage

            If cReadOnly = 1 Then
                ' TODO: generate <ow:error> tag into the XML output
                gAction = "view"
                ActionView()
                gActionReturn = True
                Exit Sub
            End If

            If gEditPassword <> "" Then
                If gEditPassword <> gReadPassword Then
                    If HttpContext.Current.Request.Cookies(gCookieHash & "?pe").Value <> gEditPassword Then
                        If (cUseRecaptcha <> 1) Or (m(gPage, OPENWIKI_PROTECTEDPAGES, False, False)) Then
                            ActionLogin()
                            Exit Sub
                        End If
                    End If
                End If
            End If

            If HttpContext.Current.Request("save") <> "" Then
                vNewRev = CInt(HttpContext.Current.Request("newrev"))
                vMinorEdit = CInt(CInt(HttpContext.Current.Request("rc")) Xor 1)
                vComment = Trim(HttpContext.Current.Request("comment") & "")
                vText = HttpContext.Current.Request("text")

                If Len(vComment) > 1000 Then
                    vXmlStr = vXmlStr & "<ow:error code='1'>Maximum length for the comment is 1000 characters.</ow:error>"
                End If
                If Len(vText) > OPENWIKI_MAXTEXT Then
                    vXmlStr = vXmlStr & "<ow:error code='2'>Maximum length for the text is " & OPENWIKI_MAXTEXT & " characters.</ow:error>"
                End If

                If (Not m(gPage, OPENWIKI_PROTECTEDPAGES, False, False)) And (cUseRecaptcha = 1) Then
                    CaptchaCheck = RecaptchaConfirm(HttpContext.Current.Request("recaptcha_challenge_field"), HttpContext.Current.Request("recaptcha_response_field"))
                    If CaptchaCheck <> "" Then
                        vXmlStr = vXmlStr & "<ow:captcha_error>" & CaptchaCheck & "</ow:captcha_error>"
                        vXmlStr = vXmlStr & "<ow:error code='5'>reCAPTCHA error. See details in reCAPTCHA form.</ow:error>"
                    End If
                End If

                If vXmlStr <> "" Then
                    vPage = gNamespace.GetPage(gPage, 0, 0, False)
                    vPage.Revision = gRevision
                    vPage.Text = vText

                    vChange = vPage.GetLastChange()
                    vChange.Revision = vNewRev
                    vChange.MinorEdit = vMinorEdit
                    vChange.Comment = vComment
                    vChange.Timestamp = Now()
                    vChange.UpdateBy()

                    vXmlStr = vXmlStr & vPage.ToXML(2)
                ElseIf gNamespace.SavePage(vNewRev, vMinorEdit, vComment, vText) Then
                    HttpContext.Current.Response.Redirect(gScriptName & "?" & HttpContext.Current.Server.UrlEncode(gPage))
                Else
                    vPage = gNamespace.GetPage(gPage, 0, 1, False)
                    vChange = vPage.GetLastChange()
                    vChange.Revision = vChange.Revision + 1
                    vChange.MinorEdit = CInt(HttpContext.Current.Request("rc")) Xor 1
                    vChange.Comment = Trim(HttpContext.Current.Request("comment") & "")
                    vChange.Timestamp = Now()
                    vChange.UpdateBy()
                    vXmlStr = vXmlStr & "<ow:error code='4'>Somebody else just edited this page.</ow:error>"
                    vXmlStr = vXmlStr & "<ow:textedits>" & PCDATAEncode(HttpContext.Current.Request("text")) & "</ow:textedits>"
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
            ElseIf HttpContext.Current.Request("cancel") <> "" Then
                Dim vBacklink As String

                If gRevision = 0 Then
                    vBacklink = gScriptName & "?" & HttpContext.Current.Server.UrlEncode(gPage)
                Else
                    vBacklink = gScriptName & "?p=" & HttpContext.Current.Server.UrlEncode(gPage) & "&revision=" & gRevision
                End If
                HttpContext.Current.Response.Redirect(vBacklink)
            Else
                ' first time opening edit form
                vPage = gNamespace.GetPage(gPage, 0, 1, False)
                If gRevision > 0 Then
                    vTemp = gNamespace.GetPage(gPage, gRevision, 1, False)
                    vPage.Revision = vTemp.Revision
                    vPage.Text = vTemp.Text
                End If

                If vPage.Revision = 0 And HttpContext.Current.Request("template") <> "" Then
                    vTemp = gNamespace.GetPage(URLDecode(HttpContext.Current.Request("template")), 0, 1, False)
                    vPage.Text = vTemp.Text
                End If

                vChange = vPage.getLastChange()
                vChange.Revision = vChange.Revision + 1
                vChange.MinorEdit = 0
                vChange.Comment = ""
                vChange.Timestamp = Now()
                vChange.UpdateBy()

                vXmlStr = vPage.ToXML(2)
            End If

            Call gTransformer.Transform(vXmlStr)
            gActionReturn = True
        End Sub

        Sub ActionTitleSearch()
            Dim vXmlStr As String
            vXmlStr = gNamespace.GetIndexSchemes.GetTitleSearch(gTxt)
            If cAllowRSSExport = 1 And HttpContext.Current.Request("v") = "rss" Then
                Call gTransformer.TransformXmlStr(vXmlStr, "owsearchrss10export.xsl")
            Else
                Call gTransformer.Transform(vXmlStr)
            End If
            gActionReturn = True
        End Sub

        Sub ActionFullSearch()
            Dim vXmlStr As String
            vXmlStr = gNamespace.GetIndexSchemes.GetFullSearch(gTxt, 1)
            If cAllowRSSExport = 1 And HttpContext.Current.Request("v") = "rss" Then
                Call gTransformer.TransformXmlStr(vXmlStr, "owsearchrss10export.xsl")
            Else
                Call gTransformer.Transform(vXmlStr)
            End If
            gActionReturn = True
        End Sub

        Sub ActionTextSearch()
            Dim vXmlStr As String
            vXmlStr = gNamespace.GetIndexSchemes.GetFullSearch(gTxt, 0)
            If cAllowRSSExport = 1 And HttpContext.Current.Request("v") = "rss" Then
                Call gTransformer.TransformXmlStr(vXmlStr, "owsearchrss10export.xsl")
            Else
                Call gTransformer.Transform(vXmlStr)
            End If
            gActionReturn = True
        End Sub

        Sub ActionRandomPage()
            Dim vTemp As Vector

            Randomize()
            If cUseSpecialPagesPrefix = 1 Then
                vTemp = gNamespace.TitleSearch("^(?!" & gSpecialPagesPrefix & ")" & ".*", 0, 0, 0, 0)
            Else
                vTemp = gNamespace.TitleSearch(".*", 0, 0, 0, 0)
            End If

            '    HttpContext.Current.Response.Redirect(gScriptName & "?a=" & gAction & "&p=" & HttpContext.Current.Server.URLEncode(vTemp.ElementAt(Int((vTemp.Count - 1) * Rnd)).Name) & "&redirect=" & HttpContext.Current.Server.URLEncode(gPage))
            HttpContext.Current.Response.Redirect(gScriptName _
                & "?a=" & gAction & "&p=" _
                & HttpContext.Current.Server.UrlEncode(CType(vTemp.ElementAt(CInt((vTemp.Count - 1) * Rnd())), WikiPage).Name))
        End Sub

        Sub ActionChanges()
            Dim vXmlStr As String = ""

            If cCacheXML = 1 Then
                vXmlStr = gNamespace.GetDocumentCache("changes")
            End If
            If vXmlStr = "" Then
                vXmlStr = gNamespace.GetPage(gPage, 0, 0, True).ToXML(0)
                If cCacheXML = 1 Then
                    Call gNamespace.SetDocumentCache("changes", vXmlStr)
                End If
            End If
            Call gTransformer.Transform(vXmlStr)
            gActionReturn = True
        End Sub

        Sub ActionUserPreferences()
            If HttpContext.Current.Request("save") <> "" Then
                HttpContext.Current.Response.Cookies(gCookieHash & "?up").Expires = Now().AddDays(60)
                HttpContext.Current.Response.Cookies(gCookieHash & "?up")("un") = FreeToNormal_X(HttpContext.Current.Request("username"), False)
                HttpContext.Current.Response.Cookies(gCookieHash & "?up")("bm") = HttpContext.Current.Request("bookmarks")
                HttpContext.Current.Response.Cookies(gCookieHash & "?up")("cols") = HttpContext.Current.Request("cols")
                HttpContext.Current.Response.Cookies(gCookieHash & "?up")("rows") = HttpContext.Current.Request("rows")
                HttpContext.Current.Response.Cookies(gCookieHash & "?up")("pwl") = HttpContext.Current.Request("prettywikilinks")
                HttpContext.Current.Response.Cookies(gCookieHash & "?up")("bmt") = HttpContext.Current.Request("bookmarksontop")
                HttpContext.Current.Response.Cookies(gCookieHash & "?up")("elt") = HttpContext.Current.Request("editlinkontop")
                HttpContext.Current.Response.Cookies(gCookieHash & "?up")("trt") = HttpContext.Current.Request("trailontop")
                HttpContext.Current.Response.Cookies(gCookieHash & "?up")("new") = HttpContext.Current.Request("opennew")
                HttpContext.Current.Response.Cookies(gCookieHash & "?up")("emo") = HttpContext.Current.Request("emoticons")
                HttpContext.Current.Response.Redirect(gScriptName & "?p=" & HttpContext.Current.Server.UrlEncode(gPage) & "&up=1")
            ElseIf HttpContext.Current.Request("clear") <> "" Then
                HttpContext.Current.Response.Cookies(gCookieHash & "?up").expires = #1/1/1990#
                HttpContext.Current.Response.Cookies(gCookieHash & "?up").Value = ""
                HttpContext.Current.Response.Redirect(gScriptName & "?p=" & HttpContext.Current.Server.UrlEncode(gPage) & "&up=2")
            End If
            gActionReturn = False
        End Sub

        Sub ActionLogout()
            HttpContext.Current.Response.Cookies(gCookieHash & "?pr").Expires = #1/1/1990#
            HttpContext.Current.Response.Cookies(gCookieHash & "?pr").Value = ""
            HttpContext.Current.Response.Cookies(gCookieHash & "?pe").Expires = #1/1/1990#
            HttpContext.Current.Response.Cookies(gCookieHash & "?pe").Value = ""
            HttpContext.Current.Response.Redirect(gScriptName & "?" & HttpContext.Current.Server.UrlEncode(gPage))
        End Sub

        Sub ActionLogin()
            Dim vMode As String
            Dim vPwd As String
            Dim vXmlStr As String = ""
            Dim vTemp As String

            If gAction = "edit" Then
                vMode = "edit"
                gAction = "login"
            Else
                vMode = HttpContext.Current.Request("mode")
            End If
            vPwd = HttpContext.Current.Request("pwd")
            If vMode = "edit" Then
                If vPwd = gEditPassword Then
                    If HttpContext.Current.Request("r") = "1" Then
                        HttpContext.Current.Response.Cookies(gCookieHash & "?pe").Expires = Now.AddDays(60)
                    End If
                    HttpContext.Current.Response.Cookies(gCookieHash & "?pe").Value = vPwd
                    HttpContext.Current.Response.Redirect(gScriptName & "?" & HttpContext.Current.Request("backlink"))
                End If
            Else
                If vPwd = gReadPassword Then
                    If HttpContext.Current.Request("r") = "1" Then
                        HttpContext.Current.Response.Cookies(gCookieHash & "?pr").Expires = Now.AddDays(60)
                    End If
                    HttpContext.Current.Response.Cookies(gCookieHash & "?pr").Value = vPwd
                    HttpContext.Current.Response.Redirect(gScriptName & "?" & HttpContext.Current.Request("backlink"))
                End If
            End If
            If vPwd <> "" Then
                vXmlStr = "<ow:error code='3'>Incorrect password</ow:error>"
            End If
            If HttpContext.Current.Request("backlink") <> "" Then
                vTemp = HttpContext.Current.Request("backlink")
            Else
                vTemp = HttpContext.Current.Request.ServerVariables("QUERY_STRING")
                If vTemp = "" Then
                    vTemp = OPENWIKI_FRONTPAGE
                End If
            End If
            vXmlStr = vXmlStr & "<ow:login"
            If vMode = "edit" Then
                vXmlStr = vXmlStr & " mode='edit'>"
            Else
                vXmlStr = vXmlStr & " mode='view'>"
            End If
            vXmlStr = vXmlStr & "<ow:backlink>" & PCDATAEncode(vTemp) & "</ow:backlink>"
            If HttpContext.Current.Request("r") <> "" Then
                vXmlStr = vXmlStr & "<ow:rememberme>true</ow:rememberme>"
            End If
            vXmlStr = vXmlStr & "</ow:login>"
            Call gTransformer.Transform(vXmlStr)
            gActionReturn = True
        End Sub
    End Module
End Namespace