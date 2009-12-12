Imports System.Text.RegularExpressions

Namespace Openwiki
    Module Wikification
        Function Wikify(ByVal pText As String) As String
            Dim vText As String
            Dim vTempPos As Integer
            Dim vTemp As String

            vText = pText

            gIncludingAsTemplate = False
            If gIncludeLevel = 0 Then
                gRaw = New Vector
                gBracketIndices = New Vector
                gTOC = New TableOfContents
                gCategories = New Vector

                If gAction <> "edit" And Not cEmbeddedMode = 1 Then
                    If Left(vText, 1) = "#" Then
                        If m(vText, "^#RANDOMPAGE", False, False) Then
                            ActionRandomPage()
                        ElseIf m(vText, "^#REDIRECT\s+", False, False) And HttpContext.Current.Request("redirect") = "" Then
                            vTempPos = InStr(10, vText, vbCr)
                            If vTempPos > 0 Then
                                vTemp = Trim(Mid(vText, 10, vTempPos - 10))
                            Else
                                vTemp = Trim(Mid(vText, 10))
                            End If
                            HttpContext.Current.Response.Redirect(gScriptName & "?a=" & gAction & "&p=" & HttpContext.Current.Server.UrlEncode(vTemp) & "&redirect=" & HttpContext.Current.Server.UrlEncode(FreeToNormal(gPage)))
                        ElseIf m(vText, "^#INCLUDE_AS_TEMPLATE", False, False) Then
                            vText = Mid(vText, Len("#INCLUDE_AS_TEMPLATE") + 1)
                        ElseIf m(vText, "^#MINOREDIT", False, False) Then
                            vText = Mid(vText, Len("#MINOREDIT") + 1)
                        ElseIf m(vText, "^#DEPRECATED", False, False) Then
                            'StoreRaw("#DEPRECATED")
                            StoreRaw("<ow:deprecated />")
                            vText = sReturn & Mid(vText, Len("#DEPRECATED") + 1)
                        End If
                        '                        vText = MyWikifyProcessingInstructions(vText)
                    End If
                End If
            Else
                If gAction <> "edit" And Not cEmbeddedMode = 1 Then
                    If Left(vText, 1) = "#" Then
                        If m(vText, "^#INCLUDE_AS_TEMPLATE", False, False) Then
                            vText = Mid(vText, 21)
                            gIncludingAsTemplate = True
                        End If
                    End If
                End If
            End If

            vText = MultiLineMarkup(vText)  ' Multi-line markup
            vText = WikiLinesToHtml(vText)  ' Line-oriented markup

            vText = s(vText, gFS & "(\d+)" & gFS, AddressOf GetRaw, False, True, New [String]() {"$1"})  ' Restore saved text
            vText = s(vText, gFS & "(\d+)" & gFS, AddressOf GetRaw, False, True, New [String]() {"$1"})  ' Restore nested saved text

            If gIncludeLevel = 0 Then
                If cUseHeadings = 1 Then
                    vText = s(vText, gFS & "(\=+)[ \t]+(.*?)[ \t]+\=+ " & gFS, AddressOf GetWikiHeading, False, True, New [String]() {"$1", "$2"})
                    '            vText = Replace(vText, gFS & "TOC" & gFS, gTOC.GetTOC)
                    vText = Replace(vText, gFS & "TOC" & gFS, "<ow:toc_root>" & gTOC.GetTOC & "</ow:toc_root>")
                    vText = Replace(vText, gFS & "TOCRight" & gFS, "<ow:toc_root align=""right"">" & gTOC.GetTOC & "</ow:toc_root>")
                End If

                Dim i As Integer
                If gCategories.Count > 0 Then
                    vText = vText & "<ow:categories>"
                    For i = 0 To gCategories.Count - 1
                        vText = vText & CStr(gCategories.ElementAt(i))
                    Next
                    vText = vText & "</ow:categories>"
                End If

                If InStr(gMacros, "Footnote") > 0 Then
                    vText = InsertFootnotes(vText)
                End If

                '                vText = MyLastMinuteChanges(vText)
                gRaw = Nothing
                gBracketIndices = Nothing
                gTOC = Nothing
                gCategories = Nothing
            End If

            Wikify = vText
        End Function

        Function MultiLineMarkup(ByVal pText As String) As String
            '            Dim vAttachmentPattern As String

            pText = Replace(pText & "", Chr(9), Space(8))
            'pText = Replace(pText, gFS, "")    ' remove separators

            If cRawHtml = 1 Then
                pText = s(pText, "<html>([\s\S]*?)<\/html>", AddressOf StoreHtml, True, True, New [String]() {"$1"})
            End If
            If cMathML = 1 Then
                pText = s(pText, "<math(\s[^<>/]+?)?>([\s\S]*?)<\/math>", AddressOf StoreMathML, True, True, New [String]() {"$1", "$2"})
            End If

            '            pText = MyMultiLineMarkupStart(pText)

            '       //*** THE @this DIRECTIVES START ***//
            '       // Inline tokens: @this,@username,@serverroot,@date,@time,@parent
            '       // Inline directives: @editthis,@printthis,@historythis,@attachmentthis,@xmlthis,@printthis
            '       // Gordon Bamber 20041007
            '       // precede a token with ~ to avoid autolinking: ~@this, ~@parent, ~@parent/~@parent for example
            '       // Also {{{also @this is not autolinked}}} <code>also @this is not autolinked</code>
            'pText = s(pText, "~(@\S+)", "&StoreRaw(""<tt>"" & $1 & ""</tt>"")", True, True)
            pText = s(pText, "~(@\S+)", AddressOf StoreRaw, True, True, New [String]() {"<tt>$1</tt>"})

            'pText = s(pText, "\{\{\{(.*?@\S+.*?)\}\}\}", "&StoreRaw(""<tt>"" & $1 & ""</tt>"")", True, True)
            pText = s(pText, "\{\{\{(.*?@\S+.*?)\}\}\}", AddressOf StoreRaw, True, True, New [String]() {"<tt>$1</tt>"})

            'pText = s(pText, "\<code\>(.*?@\S+.*?)\<\/code\>", "&StoreRaw(""<tt>"" & $1 & ""</tt>"")", True, True)
            pText = s(pText, "\<code\>(.*?@\S+.*?)\<\/code\>", AddressOf StoreRaw, True, True, New [String]() {"<tt>$1</tt>"})

            '       // First do the formatted versions //
            'pText = Replace(pText, "@editlink", "[" & gServerRoot & OPENWIKI_SCRIPTNAME & "?p=" & gPage & "&a=edit edit]", 1, -1, 1)
            'pText = Replace(pText, "@historylink", "[" & gServerRoot & OPENWIKI_SCRIPTNAME & "?p=" & gPage & "&a=changes history]", 1, -1, 1)
            'pText = Replace(pText, "@attachmentlink", "[" & gServerRoot & OPENWIKI_SCRIPTNAME & "?p=" & gPage & "&a=attach attachment]", 1, -1, 1)
            'pText = Replace(pText, "@xmllink", "[" & gServerRoot & OPENWIKI_SCRIPTNAME & "?p=" & gPage & "&a=xml&revision=" & gRevision & " xml]", 1, -1, 1)
            'pText = Replace(pText, "@printlink", "[" & gServerRoot & OPENWIKI_SCRIPTNAME & "?p=" & gPage & "&a=print&revision=" & gRevision & " print]", 1, -1, 1)
            ''       // (the global gServerRoot is initialised in owprocessor.asp)
            ''       // Then the unformatted versions //
            'pText = Replace(pText, "@editthis", gServerRoot & OPENWIKI_SCRIPTNAME & "?p=" & gPage & "&a=edit", 1, -1, 1)
            'pText = Replace(pText, "@historythis", gServerRoot & OPENWIKI_SCRIPTNAME & "?p=" & gPage & "&a=changes", 1, -1, 1)
            'pText = Replace(pText, "@attachmentthis", gServerRoot & OPENWIKI_SCRIPTNAME & "?p=" & gPage & "&a=attach", 1, -1, 1)
            'pText = Replace(pText, "@xmlthis", gServerRoot & OPENWIKI_SCRIPTNAME & "?p=" & gPage & "&a=xml&revision=" & gRevision & "", 1, -1, 1)
            'pText = Replace(pText, "@printthis", gServerRoot & OPENWIKI_SCRIPTNAME & "?p=" & gPage & "&a=print&revision=" & gRevision & "", 1, -1, 1)

            '       // Gordon Bamber 20041007
            '       // all the rest of the @tokens are done here
            pText = ReplacePageTokens(pText, gPage) '        // This function is also used by macro code
            '        // *** THE @this DIRECTIVES END ***//


            pText = QuoteXml(pText)
            If cRawHtml = 1 Then
                ' transform our field separator back
                pText = Replace(pText, "&#179;", gFS)
            End If
            pText = s(pText, " \\ *\r?\n", "", False, True)  ' Join lines with backslash at end



            ' The <nowiki> tag stores text with no markup (except quoting HTML)
            pText = s(pText, "\&lt;nowiki\&gt;([\s\S]*?)\&lt;\/nowiki\&gt;", AddressOf StoreRaw, True, True, New [String]() {"$1"})

            ' <!-- and --> mark commented block
            pText = s(pText, "\&lt;!--([\s\S]*?)--\&gt;", "", True, True)

            ' <code></code> and {{{ }}} do the same thing.
            pText = s(pText, "\{\{\{(.*?)\}\}\}", AddressOf StoreRaw, True, True, New [String]() {"<tt>$1</tt>"})
            pText = s(pText, "\&lt;code\&gt;(.*?)\&lt;\/code\&gt;", AddressOf StoreRaw, True, True, New [String]() {"<tt>$1</tt>"})
            pText = s(pText, "\{\{\{([\s\S]*?)\}\}\}", AddressOf StoreCode, True, True, New [String]() {"$1"})
            pText = s(pText, "\&lt;code\&gt;([\s\S]*?)\&lt;\/code\&gt;", AddressOf StoreCode, True, True, New [String]() {"$1"})
            pText = s(pText, "\&lt;pre\&gt;([\s\S]*?)\&lt;\/pre\&gt;", "<pre>$1</pre>", True, True)

            If cHtmlTags = 1 Then
                ' Scripting is currently possible with these tags, so they are *not* particularly "safe".
                Dim vTag As String
                For Each vTag In Split("b,i,u,font,big,small,sub,sup,h1,h2,h3,h4,h5,h6,cite,code,em,s,strike,strong,tt,var,div,span,center,blockquote,ol,ul,dl,table,caption,br,p,hr,li,dt,dd,tr,td,th", ",")
                    pText = s(pText, "\&lt;" & vTag & "(\s[^<>]+?)?\&gt;([\s\S]*?)\&lt;\/" & vTag & "\&gt;", "<" & vTag & "$1>$2</" & vTag & ">", True, True)
                Next
                For Each vTag In Split("br,p,hr,li,dt,dd,tr,td,th", ",")
                    pText = s(pText, "\&lt;" & vTag & "(\s[^<>/]+?)?\&gt;", "<" & vTag & "$1 />", True, True)
                Next
            End If

            If cHtmlLinks = 1 Then
                pText = s(pText, "\&lt;a\s([^<>]+?)\&gt;([\s\S]*?)\&lt;\/a\&gt;", AddressOf StoreHref, True, True, New [String]() {"$1", "$2"})
            End If

            If Not IsNothing(gAggregateURLs) Then
                ' we are in the process of refreshing RSS feeds
                If m(gMacros, "Include", True, True) Then
                    pText = s(pText, "\&lt;(Include)(\(.*?\))?(?:\s*\/)?\&gt;", AddressOf ExecMacro, True, True, New [String]() {"$1", "$2"})
                End If
                pText = s(pText, "\&lt;(Syndicate)(\(.*?\))?(?:\s*\/)?\&gt;", AddressOf ExecMacro, True, True, New [String]() {"$1", "$2"})
                MultiLineMarkup = pText
                Exit Function
            End If

            ' process macro's
            pText = s(pText, "\&lt;(" & gMacros & ")(\(.*?\))?(?:\s*\/)?\&gt;", AddressOf ExecMacro, True, True, New [String]() {"$1", "$2"})

            ' Category marks on wikipage
            pText = s(pText, gCategoryMarkPattern, AddressOf StoreCategoryMark, False, True, New [String]() {"$1"})

            If cFreeLinks = 1 Then
                pText = s(pText, "\[\[" & gFreeLinkPattern & "(?:\|([^\]]+))*\]\]", AddressOf StoreFreeLink, False, True, New [String]() {"$1", "$2"})
            End If

            ' Links like [URL] and [URL text of link]
            pText = s(pText, "\[" & gUrlPattern & "(\s+[^\]]+)*\]", AddressOf StoreBracketUrl, False, True, New [String]() {"$1", "$2"})
            pText = s(pText, "\[" & gInterLinkPattern & "(\s+[^\]]+)*\]", AddressOf StoreInterPage, False, True, New [String]() {"$1", "$2", "True"})
            pText = s(pText, "\[" & gISBNPattern & "([^\]]+)*\]", AddressOf StoreISBN, False, True, New [String]() {"$1", "$2", "True"})

            If cAllowAttachments = 1 Then
                ''Dim vAttachmentPattern
                'If Not IsNothing(gCurrentWorkingPages) Then
                '    ' we're including a page
                '    gTemp = gNamespace.GetPageAndAttachments(gCurrentWorkingPages.Top(), 0, True, False)
                'Else
                '    gTemp = gNamespace.GetPageAndAttachments(gPage, gRevision, True, False)
                'End If
                'vAttachmentPattern = gTemp.GetAttachmentPattern()
                'If vAttachmentPattern <> "" Then
                '    pText = s(pText, "\[(" & gTemp.GetAttachmentPattern & ")(\s+[^\]]+)*\]", "&StoreBracketAttachmentLink($1, $2)", False, True)
                'End If
            End If

            If cWikiLinks = 1 And cBracketText = 1 And cBracketWiki = 1 Then
                ' Local bracket-links
                pText = s(pText, "\[" & "(#?)" & gLinkPattern & "(\s+[^\]]+?)\]", AddressOf StoreBracketWikiLink, False, True, New [String]() {"$1", "$2", "$3"})
            End If

            pText = s(pText, gUrlPattern, AddressOf StoreUrl, False, True, New [String]() {"$1"})
            pText = s(pText, gInterLinkPattern, AddressOf StoreInterPage, False, True, New [String]() {"$1", "", "False"})
            pText = s(pText, gMailPattern, AddressOf StoreMail, False, True, New [String]() {"$1"})
            pText = s(pText, gISBNPattern, AddressOf StoreISBN, False, True, New [String]() {"$1", "", "False"})

            If cAllowAttachments = 1 Then
                'If Not IsNothing(gCurrentWorkingPages) Then
                '    ' we're including a page
                '    gTemp = gNamespace.GetPageAndAttachments(gCurrentWorkingPages.Top(), 0, True, False)
                'Else
                '    gTemp = gNamespace.GetPageAndAttachments(gPage, gRevision, True, False)
                'End If
                'vAttachmentPattern = gTemp.GetAttachmentPattern()
                'If vAttachmentPattern <> "" Then
                '    pText = s(pText, "(" & gTemp.GetAttachmentPattern & ")", "&StoreAttachmentLink($1)", False, True)
                'End If
            End If

            pText = s(pText, "-{4,}", "<hr />", False, True)
            pText = s(pText, "\&gt;\&gt;([\s\S]*?)\&lt;\&lt;", "<center>$1</center>", False, True)

            If cNewSkool = 1 Then
                pText = s(pText, "\*\*([^\s\*].*?)\*\*", "<b>$1</b>", False, True)
                pText = s(pText, "\/\/([^\s\/].*?)\/\/", "<i>$1</i>", False, True)
                pText = s(pText, "__([^\s_].*?)__", "<span style=""text-decoration: underline"">$1</span>", False, True)
                pText = s(pText, "--([^\s-].*?)--", "<span style=""text-decoration: line-through"">$1</span>", False, True)
                pText = s(pText, "!!([^\s!].*?)!!", "<big>$1</big>", False, True)
                pText = s(pText, "\^\^([^\s\^].*?)\^\^", "<sup>$1</sup>", False, True)
                pText = s(pText, "vv([^\sv].*?)vv", "<sub>$1</sub>", False, True)
                'pText = s(pText, " --", " &#173;", False, True)
            End If

            If cUseHeadings = 1 And cWikifyHeaders = 0 Then
                pText = s(pText, gHeaderPattern, AddressOf StoreWikiHeading, False, True, New [String]() {"$1", "$2", "$3"})
            End If

            If cWikiLinks = 1 Then
                If OPENWIKI_STOPWORDS <> "" Then
                    gStopWords = gNamespace.GetPage(OPENWIKI_STOPWORDS, 0, 1, False).Text
                    gStopWords = Replace(gStopWords & "", Chr(9), " ")
                    gStopWords = Replace(gStopWords, gFS, "")    ' remove separators
                    gStopWords = Replace(gStopWords, vbCr, " ")
                    gStopWords = Replace(gStopWords, vbLf, " ")
                    gStopWords = Trim(gStopWords)
                    gStopWords = s(gStopWords, "\s+", "|", False, True)
                End If

                If gStopWords <> "" Then
                    pText = s(pText, "\b(" & gStopWords & ")\b", AddressOf StoreRaw, True, True, New [String]() {"$1"})
                End If

                If cNewSkool = 1 Then
                    pText = s(pText, "(~?)" & gLinkPattern, AddressOf GetWikiLink, False, True, New [String]() {"$1", "$2", ""})
                Else
                    pText = s(pText, gLinkPattern, AddressOf GetWikiLink, False, True, New [String]() {"", "$1", ""})
                End If
            End If

            If cOldSkool = 1 Then
                ' The quote markup patterns avoid overlapping tags (with 5 quotes)
                ' by matching the inner quotes for the strong pattern.
                pText = Replace(pText, "''''''", "")
                pText = s(pText, "('*)'''(.*?)'''", "$1<strong>$2</strong>", False, True)
                pText = s(pText, "''(.*?)''", "<em>$1</em>", False, True)
            End If

            If Not cHtmlTags = 1 Then
                ' I disabled this because I don't like this way of quoting
                ' Enabling this forces editors to use "correct" HTML, i.e. XHTML.
                ' E.g. <b><i>bla</b></i> will fail, because it's not valid XHTML. -- LaurensPit
                'pText = s(pText, "\&lt;b\&gt;(.*?)\&lt;\/b\&gt;", "<b>$1</b>", True, True)
                'pText = s(pText, "\&lt;i\&gt;(.*?)\&lt;\/i\&gt;", "<i>$1</i>", True, True)
                'pText = s(pText, "\&lt;u\&gt;(.*?)\&lt;\/u\&gt;", "<u>$1</u>", True, True)
                'pText = s(pText, "\&lt;strong\&gt;(.*?)\&lt;\/strong\&gt;", "<strong>$1</strong>", True, True)
                'pText = s(pText, "\&lt;em\&gt;(.*?)\&lt;\/em\&gt;", "<em>$1</em>", True, True)
            End If

            If cEmoticons = 1 Then
                pText = s(pText, "\s\:\-?\)($|\s)", " <span><img src=""" & OPENWIKI_ICONPATH & "/emoticon-smile.gif"" width=""14"" height=""12"" alt=""""/></span>$1", True, True)
                pText = s(pText, "\s\;\-?\)($|\s)", " <span><img src=""" & OPENWIKI_ICONPATH & "/emoticon-wink.gif"" width=""14"" height=""12"" alt=""""/></span>$1", True, True)
                pText = s(pText, "\s\:\-?\(($|\s)", " <span><img src=""" & OPENWIKI_ICONPATH & "/emoticon-sad.gif"" width=""14"" height=""12"" alt=""""/></span>$1", True, True)
                pText = s(pText, "\s\:\-?\|($|\s)", " <span><img src=""" & OPENWIKI_ICONPATH & "/emoticon-ambivalent.gif"" width=""14"" height=""12"" alt=""""/></span>$1", True, True)
                pText = s(pText, "\s\:\-?D($|\s)", " <span><img src=""" & OPENWIKI_ICONPATH & "/emoticon-laugh.gif"" width=""14"" height=""12"" alt=""""/></span>$1", True, True)
                pText = s(pText, "\s\:\-?O($|\s)", " <span><img src=""" & OPENWIKI_ICONPATH & "/emoticon-surprised.gif"" width=""14"" height=""12"" alt=""""/></span>$1", True, True)
                pText = s(pText, "\s\:\-?P($|\s)", " <span><img src=""" & OPENWIKI_ICONPATH & "/emoticon-tongue-in-cheek.gif"" width=""14"" height=""12"" alt=""""/></span>$1", True, True)
                pText = s(pText, "\s\:\-?S($|\s)", " <span><img src=""" & OPENWIKI_ICONPATH & "/emoticon-unsure.gif"" width=""14"" height=""12"" alt=""""/></span>$1", True, True)
                pText = s(pText, "(^|\s)\(([Y|N|L|U|K|G|F|P|B|D|T|C|I|H|S|8|E|M])\)($|\s)", "$1<span><img src=""" & OPENWIKI_ICONPATH & "/emoticon-$2.gif"" width=""14"" height=""12"" alt=""""/></span>$3", True, True)
                pText = s(pText, "(^|\s)\(\*\)($|\s)", "$1<span><img src=""" & OPENWIKI_ICONPATH & "/emoticon-star.gif"" width=""14"" height=""12"" alt=""""/></span>$2", True, True)
                pText = s(pText, "(^|\s)\(\@\)($|\s)", "$1<span><img src=""" & OPENWIKI_ICONPATH & "/emoticon-cat.gif"" width=""14"" height=""12"" alt=""""/></span>$2", True, True)
                pText = s(pText, "(^|\s)\/i\\($|\s)", "$1<span><img src=""" & OPENWIKI_ICONPATH & "/icon-info.gif"" width=""16"" height=""16"" alt=""""/></span>$2", True, True)
                pText = s(pText, "(^|\s)\/w\\($|\s)", "$1<span><img src=""" & OPENWIKI_ICONPATH & "/icon-warning.gif"" width=""16"" height=""16"" alt=""""/></span>$2", True, True)
                pText = s(pText, "(^|\s)\/s\\($|\s)", "$1<span><img src=""" & OPENWIKI_ICONPATH & "/icon-error.gif"" width=""16"" height=""16"" alt=""""/></span>$2", True, True)
            End If

            If cUseHeadings = 1 And cWikifyHeaders = 1 Then
                pText = s(pText, gHeaderPattern, AddressOf StoreWikiHeading, False, True, New [String]() {"$1", "$2", "$3"})
            End If

            '            pText = MyMultiLineMarkupEnd(pText)

            MultiLineMarkup = pText
        End Function

        Function WikiLinesToHtml(ByVal pText As String) As String
            Dim vTagStack As TagStack
            'vRegEx()
            Dim vMatch As Match
            Dim vMatches As MatchCollection
            Dim vLine As String
            Dim vFirstChar As String
            Dim vCode As String = ""
            Dim vDepth As Integer
            Dim vPos As Integer
            '            Dim vStart As Integer
            Dim vAttrs As String = ""
            Dim vCodeOpen As String = ""
            Dim vCodeClose As String = ""
            Dim vCodeList As String, vCodeItem As String
            Dim vInTable As Integer
            Dim vInInfobox As Integer
            Dim vText As String
            Dim vResult As String

            If IsNothing(pText) Then
                Return Nothing
            End If

            vText = ""
            vDepth = 0
            vInTable = 0
            vInInfobox = 0

            vTagStack = New TagStack

            'vRegEx = New Regexp
            'vRegEx.IgnoreCase = False
            'vRegEx.Global = True
            'vRegEx.Pattern = ".+"
            vMatches = Regex.Matches(pText, ".+")
            For Each vMatch In vMatches
                'vLine = vMatch.Value
                vLine = RTrim(Replace(vMatch.Value, vbCr, ""))
                vLine = s(vLine, "^\s*$", "<p></p>", False, True)  ' Blank lines

                ' The following piece of code is not as bad as you could hope for      
                vFirstChar = Left(vLine, 1)
                If (vFirstChar = " ") Or (vFirstChar = Chr(8)) Then
                    If (vDepth = 0) And (vInTable > 0) Then
                        vText = vText & vbCrLf & "</table>" & vbCrLf
                        vInTable = 0
                    End If

                    vAttrs = ""
                    gListSet = False    ' Dictionary Lists processing block when True
                    vLine = s(vLine, "^(\s+)\;(.*?) \:", AddressOf SetListValues, False, True, New [String]() {"True", "$1", "<dt>$2</dt><dd>"})
                    If gListSet Then
                        vCode = "dl"
                        vCodeList = "dl"
                        vCodeItem = "dd"
                        vCodeOpen = vCodeList
                        vDepth = Len(gDepth) \ 2

                        vLine = vTagStack.ProcessLine(vDepth, vCodeItem) & vLine
                        vCodeClose = vTagStack.ProcessCodeClose(vDepth, vCodeItem, vCodeList)
                        Call vTagStack.NestList(vDepth, vCodeItem, vCodeList)
                    Else    ' Indented lists processing block when True
                        vLine = s(vLine, "^(\s+)\:\s(.*?)$", AddressOf SetListValues, False, True, New [String]() {"True", "$1", "<dt /><dd>$2"})
                        If gListSet Then
                            vCode = "dl"
                            vCodeList = "dl"
                            vCodeItem = "dd"
                            vCodeOpen = vCodeList
                            vDepth = Len(gDepth) \ 2

                            vLine = vTagStack.ProcessLine(vDepth, vCodeItem) & vLine
                            vCodeClose = vTagStack.ProcessCodeClose(vDepth, vCodeItem, vCodeList)
                            Call vTagStack.NestList(vDepth, vCodeItem, vCodeList)
                        Else ' Unordered lists processing block when True
                            vLine = s(vLine, "^(\s+)\*\s(.*?)$", AddressOf SetListValues, False, True, New [String]() {"True", "$1", "<li>$2"})
                            If gListSet Then
                                vCode = "ul"
                                vCodeList = "ul"
                                vCodeItem = "li"
                                vCodeOpen = vCodeList
                                vDepth = Len(gDepth) \ 2

                                vLine = vTagStack.ProcessLine(vDepth, vCodeItem) & vLine
                                vCodeClose = vTagStack.ProcessCodeClose(vDepth, vCodeItem, vCodeList)
                                Call vTagStack.NestList(vDepth, vCodeItem, vCodeList)
                            Else
                                vLine = s(vLine, "^(\s+)([0-9aAiI]\.(?:#\d+)? )", AddressOf SetListValues, False, True, New [String]() {"True", "$1", "$2"})
                                If gListSet Then
                                    vPos = InStr(vLine, " ")
                                    '                            vCode  = Left(vLine, vPos - 1)
                                    '			                vCodeOpen = vCode
                                    '            			    vCodeClose = vCode                            
                                    vLine = "<li>" & Mid(vLine, vPos + 1) ' & "</li>"

                                    '                            vPos   = InStr(vCode, "#")
                                    '                            vStart = ""
                                    '                            If vPos > 0 Then
                                    '                                vStart = "start=""" & Mid(vCode, vPos + 1) & """"
                                    '                            End If
                                    '                            vCode = Left(vCode, 1)
                                    '			                vCodeOpen = vCode
                                    '            			    vCodeClose = vCode                            
                                    '                            If IsNumeric(vCode) Then
                                    '                                vAttrs = " type=""1"""
                                    '                            Else
                                    '                                vAttrs = " type=""" & vCode & """"
                                    '                            End If
                                    '                            If vStart <> "" Then
                                    '                                vAttrs = vAttrs & " " & vStart
                                    '                            End If
                                    vCode = "ol"
                                    vCodeList = "ol"
                                    vCodeItem = "li"
                                    vCodeOpen = vCodeList
                                    vDepth = Len(gDepth) \ 2

                                    vLine = vTagStack.ProcessLine(vDepth, vCodeItem) & vLine
                                    vCodeClose = vTagStack.ProcessCodeClose(vDepth, vCodeItem, vCodeList)
                                    Call vTagStack.NestList(vDepth, vCodeItem, vCodeList)
                                ElseIf vDepth > 0 And vCode <> "pre" Then
                                    Dim vTemp As String
                                    vTemp = Trim(vLine)
                                    If (Left(vTemp, 2) = "||") And (Right(vTemp, 2) = "||") Then
                                        vLine = vTemp
                                    ElseIf vInTable = 0 Then
                                        vText = vText & "<br />"
                                    End If
                                Else
                                    vCode = "pre"
                                    vCodeOpen = vCode
                                    vCodeClose = vCode
                                    vDepth = 1
                                    vTagStack.Depth = 1
                                End If  ' If gListSet Then .. Else
                            End If  ' If gListSet Then .. Else
                        End If  ' If gListSet Then .. Else
                    End If  ' If gListSet Then .. Else
                Else    ' If (vFirstChar = " ") Or (vFirstChar = Chr(8)) Then
                    If (vDepth > 0) And (vInTable > 0) Then
                        vText = vText & vbCrLf & "</table>" & vbCrLf
                        vInTable = 0
                    End If

                    vDepth = 0
                    vTagStack.Depth = 0
                End If

                Do While (vTagStack.Count > vDepth)  ' vDepth has decreased
                    vText = vText & "</" & vTagStack.Pop() & ">" & vbCrLf
                Loop

                If (vDepth > 0) Then
                    If vDepth > gIndentLimit Then
                        vDepth = gIndentLimit
                    End If
                    Do While (vTagStack.Count < vDepth) ' vDepth has increased
                        vTagStack.Push(vCodeClose)
                        vText = vText & "<" & vCodeOpen & vAttrs & ">" & vbCrLf
                    Loop
                    '            If Not vTagStack.IsEmpty Then
                    '                If vTagStack.Top <> vCodeClose Then
                    '                    vText = vText & "</" & vTagStack.Pop() & ">" & vbCRLF & "<" & vCodeOpen & vAttrs & ">"
                    '                    vTagStack.Push(vCodeClose)
                    '                End If
                    '            End If
                End If

                If Left(vLine, 2) = "||" And Right(vLine, 2) = "||" Then
                    ' tables
                    Dim vTR As String
                    Dim vTD As String
                    Dim vColSpan As String = ""
                    Dim vColSpanPos As Integer
                    Dim vNrOfTDs As Integer
                    Dim vSaveReturn As String

                    vTR = vLine
                    vNrOfTDs = 0
                    vResult = ""

                    Do While vTR <> ""
                        gListSet = False
                        vTD = s(vTR, "^(\|{2,})(.*?)\|\|", AddressOf SetListValues, False, True, New [String]() {"True", "$1", "$2"})
                        If gListSet Then
                            vColSpanPos = Len(gDepth) \ 2
                            vNrOfTDs = vNrOfTDs + vColSpanPos
                            If vColSpanPos = 1 Then
                                vColSpan = "<td class=""wiki"">"
                            Else
                                vColSpan = "<td class=""wiki"" align=""center"" colspan=""" & vColSpan & """>"
                            End If
                            vSaveReturn = sReturn
                            If Trim(sReturn) = "" Then
                                sReturn = "&#160;"
                            End If
                            vResult = vResult & vColSpan & sReturn & "</td>"
                            'Response.Write("GOT: " & HttpContext.Current.Server.HTMLEncode(vResult) & "<br>")
                            vTR = Mid(vTR, Len(gDepth) + Len(vSaveReturn) + 1)
                        Else
                            vTR = ""
                        End If
                    Loop    ' Do While vTR <> ""

                    If (vInTable > 0) And (vInTable <> vNrOfTDs) Then
                        vText = vText & vbCrLf & "</table>" & vbCrLf
                        vInTable = 0
                    End If

                    If vInTable = 0 Then
                        vText = vText & "<table cellspacing=""0"" cellpadding=""2"" border=""1"" class=""wiki"">"
                        vInTable = vNrOfTDs
                    End If
                    vText = vText & vbCrLf & "<tr class=""wiki"">" & vResult & "</tr>"
                ElseIf vInTable > 0 Then
                    vText = vText & vbCrLf & "</table>" & vbCrLf
                    vInTable = 0
                End If  ' If Left(vLine, 2) = "||" And Right(vLine, 2) = "||" Then

                If Left(vLine, 9) = "{{Infobox" Then
                    ' infoboxes: first line
                    vText = vText & "<ow:infobox>"
                    vInInfobox = 1
                End If

                If Left(vLine, 1) = "|" And vInInfobox > 0 Then
                    ' infoboxes: content
                    Dim vInfoboxRow As String

                    vResult = ""

                    gListSet = False
                    '            HttpContext.Current.Response.Write("vLine..." & "<br>")
                    '            HttpContext.Current.Response.Write(vLine & "<br>")
                    vInfoboxRow = s(vLine, "^\|(.*?)=(.*)$", AddressOf WikifyInfoboxContent, False, True, New [String]() {"$1", "$2"})
                    If Trim(sReturn) = "" Then
                        sReturn = "&#160;"
                    End If
                    vResult = sReturn

                    vText = vText & vResult
                End If

                If vInTable = 0 And vInInfobox = 0 Then
                    ' do not put wiki lines of tables to output
                    vText = vText & vLine & vbCrLf
                End If

                If Left(vLine, 2) = "}}" And vInInfobox > 0 Then
                    ' infoboxes: last line
                    vText = vText & vbCrLf & "</ow:infobox>" & vbCrLf
                    vInInfobox = 0
                End If

            Next    ' For Each vMatch In vMatches

            If vInTable > 0 Then
                vText = vText & vbCrLf & "</table>" & vbCrLf
            End If

            Do While Not vTagStack.IsEmpty
                vText = vText & "</" & vTagStack.Pop() & ">" & vbCrLf
            Loop

            '            vRegEx = Nothing
            vTagStack = Nothing

            WikiLinesToHtml = vText
        End Function    ' WikiLinesToHtml(pText)

        Dim gListSet As Boolean
        Dim gDepth As String

        Sub SetListValues(ByVal pListSet As Boolean, ByVal pDepth As String, ByVal pText As String)
            gListSet = pListSet
            gDepth = pDepth
            sReturn = pText
        End Sub

        Sub WikifyInfoboxContent(ByVal pParameterName As String, ByVal pParameterValue As String)
            Dim vParameterName As String, vParameterValue As String

            vParameterName = Trim(pParameterName)
            vParameterValue = Trim(pParameterValue)

            If vParameterName = "name" Then
                sReturn = "<ow:infobox_name>" & Trim(vParameterValue) & "</ow:infobox_name>" & vbCrLf
            Else
                sReturn = "<ow:param_name>" & Trim(vParameterName) & "</ow:param_name>" & vbCrLf
                sReturn = sReturn & "<ow:param_val>" & Trim(vParameterValue) & "</ow:param_val>" & vbCrLf
                sReturn = vbCrLf & "<ow:infobox_row>" & vbCrLf & sReturn & "</ow:infobox_row>"
            End If
        End Sub

        Function QuoteXml(ByRef pText As String) As String
            QuoteXml = Replace(pText, "&", "&amp;")
            QuoteXml = Replace(QuoteXml, "<", "&lt;")
            QuoteXml = Replace(QuoteXml, ">", "&gt;")

            ' In XML data HTML character references are invalid (unless these are
            ' defined in the DTD). Special characters can be entered in XML without
            ' the use of character references. Make sure you've set the constant
            ' OPENWIKI_ENCODING correct though in owconfig.asp and also the encoding
            ' attribute at the first line of the stylesheets.
            If cAllowCharRefs = 1 Then
                QuoteXml = s(QuoteXml, "\&amp;([#a-zA-Z0-9]+);", AddressOf StoreCharRef, False, True, New [String]() {"$1"})
            End If
        End Function

        Function CDATAEncode(ByVal pText As String) As String
            If pText <> "" Then
                CDATAEncode = Replace(pText, "&", "&amp;")
                CDATAEncode = Replace(CDATAEncode, "<", "&lt;")
                CDATAEncode = Replace(CDATAEncode, "'", "&apos;")
            End If
        End Function

        Function PCDATAEncode(ByVal pText As String) As String
            If pText <> "" Then
                PCDATAEncode = Replace(pText, "&", "&amp;")
                PCDATAEncode = Replace(PCDATAEncode, "<", "&lt;")
                PCDATAEncode = Replace(PCDATAEncode, "]]>", "]]&gt;")
            End If
        End Function

        Function URLDecode(ByVal pURL As String) As String
            Dim vPos As Integer
            If pURL <> "" Then
                pURL = Replace(pURL, "+", " ")
                vPos = InStr(pURL, "%")
                Do While vPos > 0
                    pURL = Left(pURL, vPos - 1) _
                         & Chr(CInt("&H" & Mid(pURL, vPos + 1, 2))) _
                         & Mid(pURL, vPos + 3)
                    vPos = InStr(vPos + 1, pURL, "%")
                Loop
            End If
            URLDecode = pURL
        End Function

        Sub StoreRaw(ByVal pText As String)
            gRaw.Push(pText)
            sReturn = gFS & (gRaw.Count - 1) & gFS
        End Sub

        Sub GetRaw(ByVal pIndex As Integer)
            sReturn = CStr(gRaw.ElementAt(pIndex))
        End Sub

        Sub StoreCharRef(ByVal pText As String)
            StoreHtml("&" & pText & ";")
        End Sub

        Sub StoreHtml(ByVal pText As String)
            StoreRaw("<ow:html><![CDATA[" & Replace(pText, "]]>", "]]&gt;") & "]]></ow:html>")
        End Sub

        Sub StoreMathML(ByVal pDisplay As String, ByVal pText As String)
            If Trim(pDisplay) = "display=""inline""" Then
                StoreRaw("<ow:math" & pDisplay & "><ow:display>inline</ow:display><![CDATA[" & Replace(pText, "]]>", "]]&gt;") & "]]></ow:math>")
            Else
                StoreRaw("<ow:math" & pDisplay & "><![CDATA[" & Replace(pText, "]]>", "]]&gt;") & "]]></ow:math>")
            End If
        End Sub

        Sub StoreCode(ByVal pText As String)
            Call WriteDebug("StoreCode entered with", "", 100)
            Call WriteDebug("pText", pText, 100)

            StoreRaw("<pre class=""code"">" & s(pText, "'''(.*?)'''", "<b>$1</b>", False, True) & "</pre>")
            Call WriteDebug("StoreCode finished", "", 100)
        End Sub

        Sub StoreMail(ByVal pText As String)
            StoreRaw("<a href=""mailto:" & pText & """ class=""external"">" & pText & "</a>")
        End Sub

        Sub StoreUrl(ByVal pURL As String)
            Call UrlLink(pURL)
            StoreRaw(gTempLink)
            sReturn = sReturn & gTempJunk
        End Sub

        Sub StoreBracketUrl(ByVal pURL As String, ByVal pText As String)
            If pText = "" Then
                If cUseLinkIcons = 1 Then
                    pText = pURL
                End If
            Else
                If cBracketText = 0 Then
                    sReturn = "[" & pURL & " " & pText & "]"
                    Exit Sub
                End If
            End If
            StoreRaw(GetExternalLink(pURL, pText, "", True))
        End Sub

        Sub StoreHref(ByVal pAnchor As String, ByVal pText As String)
            Dim vLink As String
            vLink = "<a " & pAnchor
            If cExternalOut = 1 Then
                If Not m(pAnchor, " target=\""", True, True) Then
                    vLink = vLink & " onclick=""return !window.open(this.href)"""
                End If
            End If
            If Not m(pAnchor, " class=\""", True, True) Then
                vLink = vLink & " class=""external"""
            End If
            vLink = vLink & ">"
            vLink = vLink & pText & "</a>"
            StoreRaw(vLink)
        End Sub

        Sub StoreFreeLink(ByVal pID As String, ByVal pText As String)
            Dim vTemp As String

            ' trim spaces before/after subpages
            pID = s(pID, "\s*\/\s*", "/", False, True)
            vTemp = GetWikiLink("", Trim(pID), Trim(pText))
            If Left(vTemp, 1) <> "<" Then
                sReturn = "[[" & pID & pText & "]]"
            Else
                StoreRaw(vTemp)
            End If
        End Sub

        Sub StoreBracketWikiLink(ByVal pPrefix As String, ByVal pID As String, ByVal pText As String)
            Dim vTemp As String

            If pID = gPage Then
                ' don't link to oneself
                sReturn = pText
            Else
                vTemp = GetWikiLink(pPrefix, pID, LTrim(pText))
                If Left(vTemp, 1) <> "<" Then
                    sReturn = "[" & pPrefix & pID & pText & "]"
                Else
                    StoreRaw(vTemp)
                End If
            End If
        End Sub

        Sub StoreInterPage(ByVal pID As String, ByVal pText As String, ByVal pUseBrackets As Boolean)
            Dim vPos As Integer
            Dim vSite As String = ""
            Dim vRemotePage As String = ""
            Dim vURL As String = ""
            Dim vTemp As String = ""
            Dim vClass As String = ""

            If pUseBrackets Then
                gTempLink = pID
                gTempJunk = ""
            Else
                SplitUrlPunct(pID)
            End If
            vPos = InStr(gTempLink, ":")
            If vPos > 0 Then
                vSite = Left(gTempLink, vPos - 1)
                vRemotePage = Mid(gTempLink, vPos + 1)
                vURL = gNamespace.GetInterWiki(vSite)
                vClass = LCase(Trim(vSite))
            End If
            If vURL = "" Then
                sReturn = pID & pText
                If pUseBrackets Then
                    sReturn = "[" & sReturn & "]"
                End If
            Else
                If pText = "" Then
                    If pUseBrackets And cBracketIndex = 1 And (cUseLinkIcons = 0) Then
                        pText = ""
                    Else
                        ' pText = Mid(pID, Len(vSite) + 2)
                        ' pText = pID
                        pText = gTempLink
                    End If
                ElseIf cBracketText = 0 Then
                    If pUseBrackets Then
                        sReturn = "[" & pID & pText & "]"
                        Exit Sub
                    End If
                End If
                If vPos > 0 Then
                    If InStr(vURL, "$1") > 0 Then
                        vURL = Replace(vURL, "$1", vRemotePage)
                    Else
                        vURL = vURL & vRemotePage
                    End If
                Else
                    vURL = vURL & vRemotePage
                End If
                vURL = Replace(vURL, "&", "&amp;")
                vURL = Replace(vURL, "&amp;amp;", "&amp;")  ' correction back
                If vSite = "This" Then
                    StoreRaw("<ow:link name=""" & pText & """ href=""" & vURL & """ date=""" & FormatDateISO8601(Now()) & """>" & pText & "</ow:link>" & gTempJunk)
                Else
                    StoreRaw(GetExternalLink_x(vURL, pText, vSite, pUseBrackets, vClass) & gTempJunk)
                End If
            End If
        End Sub

        Sub StoreISBN(ByVal pNumber As String, ByVal pText As String, ByVal pUseBrackets As Boolean)
            If pText <> "" And cBracketText = 0 And pUseBrackets Then
                sReturn = "[ISBN" & pNumber & pText & "]"
            Else
                Dim vRawPrint As String
                Dim vNumber As String
                Dim vText As String

                vRawPrint = Replace(pNumber, " ", "")
                vNumber = Replace(vRawPrint, "-", "")

                If Len(vNumber) = 11 Then
                    If UCase(Right(vNumber, 1)) = "X" Then
                        pText = Right(vNumber, 1) & pText
                        vNumber = Left(vNumber, 10)
                    End If
                End If

                If Len(vNumber) <> 10 Then
                    If pText = "" Then
                        sReturn = "ISBN " & pNumber
                    Else
                        sReturn = "[ISBN " & pNumber & pText & "]"
                    End If
                Else
                    If pText = "" Then
                        If pUseBrackets And cBracketIndex = 1 And (cUseLinkIcons = 0) Then
                            vText = ""
                        Else
                            vText = "ISBN " & vRawPrint
                        End If
                    Else
                        vText = pText
                    End If
                    sReturn = GetExternalLink("http://www.amazon.com/exec/obidos/ISBN=" & vNumber, vText, "Amazon", pUseBrackets) _
                            & " (" & GetExternalLink("http://shop.barnesandnoble.com/bookSearch/isbnInquiry.asp?isbn=" & vNumber, "alternate", "Barnes & Noble", False) _
                            & ", " & GetExternalLink("http://www1.fatbrain.com/asp/bookinfo/bookinfo.asp?theisbn=" & vNumber, "alternate", "FatBrain", False) & ")"

                    If (pText = "") And (Right(pNumber, 1) = " ") Then
                        sReturn = sReturn & " "
                    End If
                    StoreRaw(sReturn)
                End If
            End If
        End Sub

        Sub StoreWikiHeading(ByVal pSymbols As String, ByVal pText As String, ByVal pTrailer As String)
            StoreRaw(gFS & pSymbols & " " & pText & " " & pSymbols & " " & gFS)
            sReturn = sReturn & pTrailer
        End Sub

        Sub GetWikiHeading(ByVal pSymbols As String, ByVal pText As String)
            Dim vLevel As Integer, vTemp As String
            vLevel = Len(pSymbols)
            If vLevel > 6 Then
                vLevel = 6
            End If
            vTemp = s(pText, "<ow:link name=""(.*?)"" href=.*?</ow:link>", "$1", False, False)
            '    Call gTOC.AddTOC(vLevel, "<li><a href=""#h" & gTOC.Count & """>" & vTemp & "</a></li>")
            '    Call gTOC.AddTOC(vLevel, "<ow:toctext>" _
            '    	& "<number>" & gTOC.Count & "</number>" _
            '    	& "<level>" & vLevel & "</level>" _
            '    	& "<number_trail>" & gTOC.CurNum & "</number_trail>" _
            '    	& "<text>" & vTemp & "</text>" _
            '    	& "</ow:toctext>")
            gTOC.AddTOC(vLevel, vTemp)
            sReturn = "<a id=""h" & (gTOC.Count - 1) & """/><h" & vLevel & ">" & pText & "</h" & vLevel & ">"
        End Sub


        Sub StoreBracketAttachmentLink(ByVal pName As String, ByVal pText As String)
            Dim vTemp As String

            vTemp = AttachmentLink(pName, pText)
            If vTemp = "" Then
                sReturn = "[" & pName & " " & pText & "]"
            Else
                StoreRaw(vTemp)
            End If
        End Sub

        Sub StoreAttachmentLink(ByVal pName As String)
            Dim vTemp As String

            vTemp = AttachmentLink(pName, "")
            If vTemp = "" Then
                sReturn = pName
            Else
                StoreRaw(vTemp)
            End If
        End Sub

        Function PrettyWikiLink(ByVal pID As String) As String
            If cPrettyLinks = 1 Then
                PrettyWikiLink = s(pID, "([a-z\xdf-\xff0-9])([A-Z\xc0-\xde]+)", "$1 $2", False, True)
            Else
                PrettyWikiLink = pID
            End If
            If cFreeLinks = 1 Then
                PrettyWikiLink = Replace(PrettyWikiLink, "_", " ")
            End If
        End Function


        Function GetWikiLink(ByVal pPrefix As String, ByVal pID As String, ByVal pText As String) As String
            '	Response.Write("GetWikiLink entered pPrefix=" & pPrefix & " pID=" & pID & " pText=" & pText & "<br>")
            Dim vID As String
            Dim vPage As WikiPage
            Dim vAnchor As String = ""
            Dim vTemplate As String = ""
            Dim vTemp As Integer

            If pPrefix = "~" Then
                GetWikiLink = pID
                sReturn = GetWikiLink
                Exit Function
            End If

            If pPrefix = "#" Then
                vAnchor = "#" & pID
                pID = gPage
            ElseIf pID = gPage Then
                ' don't link to oneself
                GetWikiLink = PrettyWikiLink(pID)
                sReturn = GetWikiLink
                Exit Function
            End If

            ' detect anchor
            vTemp = InStr(pID, "#")
            If vTemp > 0 Then
                vAnchor = Mid(pID, vTemp)
                pID = Left(pID, vTemp - 1)
            End If

            ' detect template
            vTemp = InStr(pID, "-&gt;")
            If vTemp > 0 Then
                vTemplate = Left(pID, vTemp - 1)
                pID = Mid(pID, vTemp + 5)
            End If

            vID = AbsoluteName(pID)

            vPage = gNamespace.GetPage(vID, 0, 0, False)
            vPage.Anchor = vAnchor
            If vPage.Exists Then
                If pText = "" Then
                    GetWikiLink = vPage.ToLinkXML(PrettyWikiLink(pID), vTemplate, True)
                Else
                    GetWikiLink = vPage.ToLinkXML(pText, vTemplate, False)
                End If
            Else
                If cReadOnly = 1 Or gAction = "print" Then
                    GetWikiLink = pID & vAnchor
                Else
                    If pText = "" Then
                        pText = pID
                    End If

                    If cFreeLinks = 1 Then
                        If InStr(pText, " ") > 0 Then
                            pText = "[" & pText & "]" ' Add brackets so boundaries are obvious
                        End If
                    End If

                    ' non existent link
                    GetWikiLink = vPage.ToLinkXML(pText, vTemplate, True)
                End If
            End If
            sReturn = GetWikiLink
        End Function


        Function AbsoluteName(ByVal pID As String) As String
            Dim vPos As Integer
            '           Dim vTemp As String
            Dim vCurrentPage As String
            Dim vMainpage As String

            If Not gIncludingAsTemplate And Not IsNothing(gCurrentWorkingPages) Then
                vCurrentPage = CStr(gCurrentWorkingPages.Top())
            Else
                vCurrentPage = gPage
            End If

            ' asbolute subpage
            vPos = InStr(vCurrentPage, "/")
            If vPos > 0 Then
                vMainpage = Left(vCurrentPage, vPos - 1)
            Else
                vMainpage = vCurrentPage
            End If
            AbsoluteName = s(pID, "^/", vMainpage & "/", False, True)

            ' relative subpage
            AbsoluteName = s(AbsoluteName, "^\./", vCurrentPage & "/", False, True)

            If cFreeLinks = 1 Then
                AbsoluteName = FreeToNormal(AbsoluteName)
            End If
        End Function

        Function FreeToNormal(ByVal pID As String) As String
            Dim vID As String
            vID = Replace(pID, " ", "_")
            vID = UCase(Left(vID, 1)) & Mid(vID, 2)
            If InStr(vID, "_") > 0 Then
                vID = s(vID, "__+", "_", False, True)
                vID = s(vID, "^_", "", False, True)
                vID = s(vID, "_$", "", False, True)
                If cUseSubpage = 1 Then
                    vID = s(vID, "_\/", "/", False, True)
                    vID = s(vID, "\/_", "/", False, True)
                End If
            End If
            If cFreeUpper = 1 Then
                vID = s(vID, "([-_\.,\(\)\/])([a-z])", AddressOf Capitalize, False, True, New [String]() {"$1", "$2"})
            End If
            FreeToNormal = vID
        End Function

        Function FreeToNormal_X(ByVal pID As String, ByVal pUseUCase As Boolean) As String
            Dim vID As String
            vID = Replace(pID, " ", "_")
            If pUseUCase Then
                vID = UCase(Left(vID, 1)) & Mid(vID, 2)
            End If
            If InStr(vID, "_") > 0 Then
                vID = s(vID, "__+", "_", False, True)
                vID = s(vID, "^_", "", False, True)
                vID = s(vID, "_$", "", False, True)
                If cUseSubpage = 1 Then
                    vID = s(vID, "_\/", "/", False, True)
                    vID = s(vID, "\/_", "/", False, True)
                End If
            End If
            If cFreeUpper = 1 Then
                vID = s(vID, "([-_\.,\(\)\/])([a-z])", AddressOf Capitalize, False, True, New [String]() {"$1", "$2"})
            End If
            FreeToNormal_X = vID
        End Function

        Sub Capitalize(ByVal pChars As String, ByVal pWord As String)
            sReturn = pChars & UCase(Left(pWord, 1)) & Mid(pWord, 2)
        End Sub

        Function GetExternalLink(ByVal pURL As String, ByVal pText As String, ByVal pTitle As String, ByVal pUseBrackets As Boolean) As String
            Dim vLink As String
            Dim vLinkedImage As Boolean
            '            Dim vTemp

            If pUseBrackets And pText = "" Then
                If cBracketIndex = 1 Then
                    pText = "[" & GetBracketUrlIndex(pURL) & "]"
                Else
                    pText = pURL
                End If
            Else
                pText = Trim(pText)
            End If

            If cAllowAttachments = 1 And (Left(pURL, 13) = "attachment://") Then
                If pUseBrackets And cShowBrackets = 1 Then
                    pText = "[" & pText & "]"
                End If
                GetExternalLink = AttachmentLink(Mid(pURL, 14), pText)
                If GetExternalLink = "" Then
                    GetExternalLink = "[" & pURL & " " & pText & "]"
                End If
                Exit Function
            End If

            vLink = "<a href='" & pURL & "' class='external'"
            If cExternalOut = 1 Then
                vLink = vLink & " onclick=""return !window.open(this.href)"""
            End If
            If pTitle <> "" Then
                vLink = vLink & " title='" & CDATAEncode(pTitle) & "'"
            End If
            vLink = vLink & ">"

            vLinkedImage = False
            If pText <> "" Then
                If m(pText, gImagePattern, False, True) Then
                    pText = "<span><img src=""" & pText & """ alt=""""/></span>"
                    vLinkedImage = True
                End If
            End If

            If pUseBrackets And cUseLinkIcons = 1 And Not vLinkedImage Then
                Dim vScheme As String
                '                Dim vImg As String
                Dim vPos As Integer

                vPos = InStr(pURL, ":")
                vScheme = Left(pURL, vPos - 1)
                '        vImg = "/wiki-" & vScheme & ".gif"" width=""12"" height=""12"""
                '        vLink = vLink & "<img src=""" & OPENWIKI_ICONPATH & vImg & " border=""0"" hspace=""4"" alt=""""/>" & pText
                vLink = vLink & pText
            Else
                If vLinkedImage Then
                    vLink = vLink & pText
                Else
                    If pUseBrackets And cShowBrackets = 1 Then
                        vLink = vLink & "["
                    End If
                    vLink = vLink & pText
                    If pUseBrackets And cShowBrackets = 1 Then
                        vLink = vLink & "]"
                    End If
                End If
            End If
            vLink = vLink & "</a>"
            GetExternalLink = vLink
        End Function

        Function GetExternalLink_x(ByVal pURL As String, ByVal pText As String, ByVal pTitle As String, ByVal pUseBrackets As Boolean, ByVal pClass As String) As String
            Dim vLink As String
            Dim vLinkedImage As Boolean
            '            Dim vTemp As String

            If pUseBrackets And pText = "" Then
                If cBracketIndex = 1 Then
                    pText = "[" & GetBracketUrlIndex(pURL) & "]"
                Else
                    pText = pURL
                End If
            Else
                pText = Trim(pText)
            End If

            If cAllowAttachments = 1 And (Left(pURL, 13) = "attachment://") Then
                If pUseBrackets And cShowBrackets = 1 Then
                    pText = "[" & pText & "]"
                End If
                GetExternalLink_x = AttachmentLink(Mid(pURL, 14), pText)
                If GetExternalLink_x = "" Then
                    GetExternalLink_x = "[" & pURL & " " & pText & "]"
                End If
                Exit Function
            End If

            vLink = "<a href='" & pURL & "' class='external " & pClass & "'"
            If cExternalOut = 1 Then
                vLink = vLink & " onclick=""return !window.open(this.href)"""
            End If
            If pTitle <> "" Then
                vLink = vLink & " title='" & CDATAEncode(pTitle) & "'"
            End If
            vLink = vLink & ">"

            vLinkedImage = False
            If pText <> "" Then
                If m(pText, gImagePattern, False, True) Then
                    pText = "<span><img src=""" & pText & """ alt=""""/></span>"
                    vLinkedImage = True
                End If
            End If

            If pUseBrackets And cUseLinkIcons = 1 And Not vLinkedImage Then
                Dim vScheme As String
                '                Dim vImg As String
                Dim vPos As Integer

                vPos = InStr(pURL, ":")
                vScheme = Left(pURL, vPos - 1)
                '        vImg = "/wiki-" & vScheme & ".gif"" width=""12"" height=""12"""
                '        vLink = vLink & "<img src=""" & OPENWIKI_ICONPATH & vImg & " border=""0"" hspace=""4"" alt=""""/>" & pText
                vLink = vLink & pText
            Else
                If vLinkedImage Then
                    vLink = vLink & pText
                Else
                    If pUseBrackets And cShowBrackets = 1 Then
                        vLink = vLink & "["
                    End If
                    vLink = vLink & pText
                    If pUseBrackets And cShowBrackets = 1 Then
                        vLink = vLink & "]"
                    End If
                End If
            End If
            vLink = vLink & "</a>"
            GetExternalLink_x = vLink
        End Function

        Function GetBracketUrlIndex(ByVal pID As String) As Integer
            Dim i As Integer
            Dim vCount As Integer

            vCount = gBracketIndices.Count
            For i = 0 To vCount
                If CStr(gBracketIndices.ElementAt(i)) = pID Then
                    GetBracketUrlIndex = i + 1
                    Exit Function
                End If
            Next
            gBracketIndices.Push(pID)
            GetBracketUrlIndex = gBracketIndices.Count
        End Function

        Sub UrlLink(ByVal pURL As String)
            Dim vLink As String = ""

            SplitUrlPunct(pURL)
            If cNetworkFile = 1 And (Left(pURL, 5) = "file:") Then
                ' only do remote file:// links. No file:///c|/windows.
                If (Left(pURL, 8) <> "file:///") Then
                    gTempLink = "<a href=""" & gTempLink & """>" & gTempLink & "</a>"
                End If
                Exit Sub
            ElseIf cAllowAttachments = 1 And (Left(pURL, 13) = "attachment://") Then
                gTempLink = AttachmentLink(Mid(gTempLink, 14), "")
                If gTempLink = "" Then
                    gTempLink = pURL
                End If
                Exit Sub
            End If
            ' restricted image URLs so that mailto:foo@bar.gif is not an image
            If cLinkImages = 1 Then
                If m(gTempLink, gImagePattern, False, True) Then
                    vLink = "<span><img src=""" & gTempLink & """ alt=""""/></span>"
                End If
            End If
            If vLink = "" Then
                vLink = "<a href=""" & gTempLink & """ class=""external"""
                If cExternalOut = 1 Then
                    vLink = vLink & " onclick=""return !window.open(this.href)"""
                End If
                vLink = vLink & ">" & gTempLink & "</a>"
            End If
            gTempLink = vLink
        End Sub

        Dim gTempLink As String, gTempJunk As String
        Sub SplitUrlPunct(ByVal pURL As String)
            Dim vTemp As Integer

            If Len(pURL) > 2 Then
                If Right(pURL, 2) = """""" Then
                    gTempLink = Mid(pURL, 1, Len(pURL) - 2)
                    gTempJunk = ""
                    Exit Sub
                End If
            End If

            gTempLink = s(pURL, "([^a-zA-Z0-9\/\xc0-\xff]+)$", "", False, True)
            gTempJunk = Mid(pURL, Len(gTempLink) + 1)

            'Response.Write("GOT: " & HttpContext.Current.Server.HTMLEncode(gTempLink) & "  :  " & HttpContext.Current.Server.HTMLEncode(gTempJunk)& "<br>")

            ' check the rare case where a semicolon was actually part of the link
            ' e.g. http://x.com?x=<y> is, at this point, translated to <a ...>http://x.com?x=&lt;y&gt</a>;
            ' which is invalid XML
            If Left(gTempJunk, 1) = ";" Then
                vTemp = InStrRev(gTempLink, "&")
                If vTemp > 0 Then
                    Dim vPosSemiColon As Integer

                    vPosSemiColon = InStrRev(gTempLink, ";")
                    If vPosSemiColon < vTemp Then
                        ' invalid XML, restore
                        gTempLink = gTempLink & ";"
                        gTempJunk = Mid(gTempJunk, 2)
                    End If
                End If
            End If
        End Sub

        Function AttachmentLink(ByVal pName As String, ByVal pText As String) As String
            Dim vPos As Integer
            Dim vPagename As String
            Dim vPage As WikiPage
            Dim vAttachment As Attachment, vText As String

            If pText = "" Then
                vText = pName
            Else
                vText = Trim(pText)
            End If
            vPos = InStrRev(pName, "/")
            If vPos > 1 Then
                vPagename = Left(pName, vPos - 1)
                pName = Mid(pName, vPos + 1)
            ElseIf Not IsNothing(gCurrentWorkingPages) Then
                ' we're including a page
                vPagename = CStr(gCurrentWorkingPages.Top())
            Else
                vPagename = gPage
            End If

            vPage = gNamespace.GetPageAndAttachments(vPagename, gRevision, 1, False)
            vAttachment = vPage.GetAttachment(pName)
            If vAttachment Is Nothing Then
                AttachmentLink = ""
                'AttachmentLink = "<ow:link name='" & CDATAEncode(pName) & "'" _
                '     & " href='" & gScriptName & "?p=" & HttpContext.Current.Server.URLEncode(gPage) & "&amp;a=attach'" _
                '     & " attachment='true'>" _
                '     & PCDATAEncode(vText) & "</ow:link>"
            ElseIf vAttachment.Deprecated = 1 Then
                AttachmentLink = ""
            Else
                AttachmentLink = vAttachment.ToXML(vPagename, vText)
            End If
        End Function

        Function InsertFootnotes(ByVal pText As String) As String
            pText = s(pText, gFS & gFS & "(.*?)" & gFS & gFS, AddressOf AddFootnote, False, True, New [String]() {"$1"})
            If Not gFootnotes Is Nothing Then
                Dim i As Integer ', vCount
                pText = pText & "<ow:footnotes>"
                For i = 0 To gFootnotes.Count - 1
                    pText = pText & "<ow:footnote index='" & (i + 1) & "'>" & CStr(gFootnotes.ElementAt(i)) & "</ow:footnote>"
                Next
                pText = pText & "</ow:footnotes>"
                gFootnotes = Nothing
            End If
            InsertFootnotes = pText
        End Function

        Dim gFootnotes As Vector
        Sub AddFootnote(ByVal pParam As String)
            If Not Not IsNothing(gFootnotes) Then
                gFootnotes = New Vector
            End If
            gFootnotes.Push(pParam)
            sReturn = "<sup><a href='#footnote" & gFootnotes.Count & "' class='footnote'>" & gFootnotes.Count & "</a></sup>"
        End Sub

        Sub StoreCategoryMark(ByVal pParam As String)
            Dim vID As String

            vID = "Category" & pParam
            gCategories.Push("<ow:category>" & "<name>" & pParam & "</name>" & GetWikiLink("", vID, "") & "</ow:category>")
            sReturn = ""
        End Sub

        Function ReplacePageTokens(ByVal pPagename As String, ByVal pRootpage As String) As String
            '        // Called in MultiLineMarkup and in many macros
            '        // When called in MultilineMarkup, the token cannot be escaped
            '        // inside a <code>block </code>
            '        // If a Page Name contains the tokens
            '        // @this,@parent,@grandparent or @greatgrandparent
            '        // Then the tokens are replaced by relations of pRootpage
            '
            '        // Example1: pPagename=@this/ChildPage   pRootPage=AnyPage/SubPage
            '        // Result would be: AnyPage/SubPage/ChildPage
            '
            '        // Example2: pPagename=@parent/ChildPage   pRootPage=AnyPage/SubPage
            '        // Result would be: AnyPage/ChildPage
            Dim aResult As String
            'Dim tPage As String
            'Dim p As Integer
            If (pRootpage = "") Then pRootpage = gPage '        // Default if blank 2nd parameter
            '        If (Instr(pPagename,"@")= -1) then
            '                // Getoutofjail
            '                ReplacePageTokens=pPageName
            '                Exit Function
            '        End If
            '        // Replace the fixed var tokens
            aResult = pPagename
            aResult = Replace(aResult, "@this", pRootpage, 1, -1, CompareMethod.Text)
            'aResult = Replace(aResult, "@username", gNamespace.FetchUserName(), 1, -1, 1)
            'aResult = Replace(aResult, "@serverroot", gServerRoot, 1, -1, 1)
            'aResult = Replace(aResult, "@date", FormatDateTime(Now(), 1), 1, -1, 1)
            'aResult = Replace(aResult, "@time", FormatDateTime(Now(), 3), 1, -1, 1)
            'If cUseMultipleParents Then
            '    aResult = Replace(aResult, "@greatgreatgreatgreatgrandparent", "@parent/@parent/@parent/@parent/@parent/@parent")
            '    aResult = Replace(aResult, "@greatgreatgreatgrandparent", "@parent/@parent/@parent/@parent/@parent")
            '    aResult = Replace(aResult, "@greatgreatgrandparent", "@parent/@parent/@parent/@parent")
            '    aResult = Replace(aResult, "@greatgrandparent", "@parent/@parent/@parent")
            '    aResult = Replace(aResult, "@grandparent", "@parent/@parent")
            '    If InStr(aResult, "@parent") > 0 Then
            '        ReplacePageTokens = ReplaceParents(aResult, pRootpage)
            '    Else
            '        ReplacePageTokens = aResult
            '    End If
            '    Exit Function
            'End If

            'aResult = s(aResult, "@parent/", "This syntax is not allowed!", True, True)
            'If InStrRev(pRootpage, "/") Then
            '    tPage = gPage
            '    p = InStrRev(tPage, "/")
            '    If (p > 1) Then
            '        tPage = Left(pRootpage, p - 1)
            '        aResult = Replace(aResult, "@parent", tPage)
            '        p = InStrRev(tPage, "/")
            '        If (p > 1) Then
            '            tPage = Left(pRootpage, p - 1)
            '            aResult = Replace(aResult, "@grandparent", tPage)
            '            p = InStrRev(tPage, "/")
            '            If (p > 1) Then
            '                tPage = Left(pRootpage, p - 1)
            '                aResult = Replace(aResult, "@greatgrandparent", tPage)
            '                p = InStrRev(tPage, "/")
            '                If (p > 1) Then
            '                    tPage = Left(pRootpage, p - 1)
            '                    aResult = Replace(aResult, "@greatgreatgrandparent", tPage)
            '                    p = InStrRev(tPage, "/")
            '                End If
            '            End If
            '        End If
            '    End If
            'End If
            ReplacePageTokens = aResult
        End Function

    End Module
End Namespace