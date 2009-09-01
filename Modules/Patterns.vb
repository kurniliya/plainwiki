Namespace Openwiki
    Module Patterns
        Sub InitLinkPatterns()
            Dim vUpperLetter As String, vLowerLetter As String, vAnyLetter As String, vQDelim As String
            vUpperLetter = "A-Z"
            vLowerLetter = "a-z"
            If cNonEnglish = 1 Then
                vUpperLetter = vUpperLetter & "\xc0-\xde"
                vLowerLetter = vLowerLetter & "\xdf-\xff"
            End If
            vAnyLetter = vUpperLetter & vLowerLetter
            If Not (cSimpleLinks = 1) Then
                'vLowerLetter = vLowerLetter & "_0-9"
                vAnyLetter = vAnyLetter & "_0-9"
            End If
            vUpperLetter = "[" & vUpperLetter & "]"
            vLowerLetter = "[" & vLowerLetter & "]"
            vAnyLetter = "[" & vAnyLetter & "]"

            vQDelim = "(?:"""")?"     ' Optional quote delimiter (not in output)

            ' Main link pattern: lowercase between uppercase, then anything:
            ' i.e. basic CamelHumpedWordPattern.
            gLinkPattern = vUpperLetter & "+" & vLowerLetter & "+" & vUpperLetter & vAnyLetter & "*"

            If (cAcronymLinks = 1) Then
                ' acronyms: three or more upper case letters
                gLinkPattern = gLinkPattern & "|" & vUpperLetter & "{3,}\b"
            End If

            If (cUseSpecialPagesPrefix = 1) Then
                gLinkPattern = gLinkPattern & "|" & gSpecialPagesPrefix & vAnyLetter & "*"
            End If

            ' Optional subpage link pattern: uppercase, lowercase, then anything
            If (cUseSubpage = 1) Then
                gSubpagePattern = "\/" & vUpperLetter & "+" & vLowerLetter & "+" & vAnyLetter & "*"
                ' Loose pattern: If subpage is used, subpage may be simple name
                gLinkPattern = "(?:(?:(?:(?:" & gLinkPattern & ")|(?:\.))?(?:" & gSubpagePattern & ")+)|" & gLinkPattern & ")"
                ' Strict pattern: both sides must be the main LinkPattern
                ' gLinkPattern = "((?:(?:" & gLinkPattern & ")?\/)?" & gLinkPattern & ")"
            End If

            If (cTemplateLinking = 1) Then
                'main link pattern (gLinkPattern) looks like TemplateName->PageName or just PageName.
                gLinkPattern = "(?:(?:" & gLinkPattern & "\-&gt;" & gLinkPattern & ")|(?:" & gLinkPattern & "))"
            End If

            ' add anchor pattern
            gLinkPattern = "(" & gLinkPattern & "(?:#" & vAnyLetter & "+)?" & ")"

            ' add optional quote delimiter
            gLinkPattern = gLinkPattern & vQDelim

            ' Inter-site convention: sites must start with uppercase letter
            ' (Uppercase letter avoids confusion with URLs)
            gInterSitePattern = vUpperLetter & vAnyLetter & "+"
            gInterLinkPattern = "((?:" & gInterSitePattern & ":[^\]\s\""<>" & gFS & "]+)" & vQDelim & ")"

            If (cFreeLinks = 1) Then
                ' Note: the - character must be first in vAnyLetter definition
                If (cNonEnglish = 1) Then
                    vAnyLetter = "[-,.()'# _0-9A-Za-z\xc0-\xff]"
                Else
                    vAnyLetter = "[-,.()'# _0-9A-Za-z]"
                End If

                If (cUseSubpage = 1) Then
                    gFreeLinkPattern = "((?:" & vAnyLetter & "+)(?:\/" & vAnyLetter & "+)*)"
                    If (cUseSpecialPagesPrefix = 1) Then
                        gFreeLinkPattern = "(" _
                         & "(?:(?:" & vAnyLetter & "+)(?:\/" & vAnyLetter & "+)*)" _
                         & "|" _
                         & "(?:" & gSpecialPagesPrefix & "(?:" & vAnyLetter & "+)(?:\/" & vAnyLetter & "+)*)" _
                         & ")"
                    End If
                Else
                    gFreeLinkPattern = "(" & vAnyLetter & "+)"
                    If (cUseSpecialPagesPrefix = 1) Then
                        gFreeLinkPattern = "(" _
                         & "(?:" & vAnyLetter & "+)" _
                         & "|" _
                         & "(?:" & gSpecialPagesPrefix & vAnyLetter & "+)" _
                         & ")"
                    End If
                End If
                'gFreeLinkPattern = gFreeLinkPattern & vQDelim
            End If


            ' Url-style links are delimited by one of:
            '   1.  Whitespace                           (kept in output)
            '   2.  Left or right angle-bracket (< or >) (kept in output)
            '   3.  Right square-bracket (])             (kept in output)
            '   4.  A single double-quote (")            (kept in output)
            '   5.  A gFS (field separator) character    (kept in output)
            '   6.  A double double-quote ("")           (removed from output)

            gUrlProtocols = "http|https|ftp|afs|news|nntp|mid|cid|mailto|wais|prospero|telnet|gopher"
            If (cNetworkFile = 1) Then
                gUrlProtocols = gUrlProtocols & "|outlook|file"
            End If
            If (cAllowAttachments = 1) Then
                gUrlProtocols = gUrlProtocols & "|attachment"
            End If
            gUrlPattern = "((?:(?:" & gUrlProtocols & "):[^\]\s\""<>" & gFS & "]+)" & vQDelim & ")"
            gMailPattern = "([-\w._+]+\@[\w.-]+\.[\w.-]+\w)"
            gImageExtensions = "gif|jpg|png|bmp|jpeg"
            gImagePattern = "^(http:|https:|ftp:).+\.(" & gImageExtensions & ")$"
            gDocExtensions = gImageExtensions & "|doc|htm|html|xsl|xml|ps|txt|zip|gz|mov|avi|mpeg|mpg|mp3|pdf|ppt|chm"
            gISBNPattern = "ISBN:?([0-9- xX]{10,})"
            gHeaderPattern = "(\=+)[ \t]+(.*?)[ \t]+\=+([ \t]*\r?\n)?"

            ' see comments in owattach.asp
            'gNotAcceptedExtensions = "asp|cdx|asa|htr|idc|shtm|shtml|stm|printer|php|pl|plx|py"  ' default app mappings using W2000
            'gNotAcceptedExtensions = gNotAcceptedExtensions & "|asax|ascx|ashx|asmx|aspx|axd|vsdisco|rem|soap|config|cs|csproj|vb|vbproj|webinfo|licx|resx|resources  ' dotnet extensions

            gTimestampPattern = "(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2}):(\d{2})(?:([+|-])(0\d|1[0-2]):([0-5]\d))?)?"

            If (cEmbeddedMode = 1) Then
                gMacros = "BR|TableOfContents|Icon|Anchor|Date|Time|DateTime|Footnote"
            Else
                ' override this in mymacros.asp or mywikify.asp if you want
                gMacros = "BR|RecentChanges|RecentChangesLong|TitleSearch|FullSearch|TextSearch|TableOfContents|WordIndex|TitleIndex|GoTo|RandomPage|" _
                        & "InterWiki|SystemInfo|Include|" _
                        & "PageCount|UserPreferences|Icon|Anchor|" _
                        & "Date|Time|DateTime|Syndicate|Aggregate|Footnote|" _
                        & "TableOfContentsRight|EquationSearch|ListRedirects|RecentNewPages|RecentEquations"
            End If
        End Sub
    End Module
End Namespace