Namespace Openwiki
    Module Macros
        Sub ExecMacro(ByVal pMacro As String, ByVal pParams As String)
            ' On error resume next should be on, because in the event someone does e.g. <bogusmacroname>
            ' then it should nicely return.
            ' The side effect of having this option on is that if a programming error occurs in the
            ' processing of a macro, the programmer won't notice it.
            '    On Error Resume Next
            Dim vMacro As String = Nothing
            Dim vParams As String = Nothing
            Dim vPos As Integer
            'Dim vTemp1 As String
            'Dim vTemp2 As String
            Dim vParamArray() As String

            vMacro = pMacro
            vParams = pParams.Trim("()".ToCharArray)
            If vParams <> "" Then
                If IsNumeric(vParams) Then
                    If InStr(vParams, ",") > 0 Then
                        vMacro = vMacro & "P"
                    End If
                Else
                    If Mid(vParams, 2, 1) = """" Then
                        vPos = InStr(3, vParams, """")
                        If InStr(vPos, vParams, ",") > 0 Then
                            vMacro = vMacro & "P"
                        End If
                    Else
                        vPos = InStr(vParams, ",")
                        If vPos > 0 Then
                            'vTemp1 = Mid(vParams, 2, vPos - 2)
                            'If Not IsNumeric(vTemp1) Then
                            '    vTemp1 = """" & vTemp1 & """"
                            'End If
                            'vTemp2 = Mid(vParams, vPos + 1, Len(vParams) - vPos - 1)
                            'If Not IsNumeric(vTemp2) Then
                            '    vTemp2 = """" & vTemp2 & """"
                            'End If
                            ''vParams = "(" & vTemp1 & "," & vTemp2 & ")"
                            'vParams = vTemp1 & "," & vTemp2
                            vMacro = vMacro & "P"
                            'Else
                            'vParams = "(""" & Mid(vParams, 2, Len(vParams) - 2) & """)"
                            'vParams = Mid(vParams, 2, Len(vParams) - 2)
                        End If
                    End If
                End If
                vMacro = vMacro & "P"
            End If

            gMacroReturn = ""
            'vCmd = "Macro" & vMacro & vParams
            'vCmd = Replace(vCmd, vbCrLf, """ & vbCRLF & """)
            ''Response.Write("<br />MACRO CMD: " & HttpContext.Current.Server.HTMLEncode(vCmd))
            ''Execute("Call " & vCmd)

            Select Case vMacro
                ' Macros without parameters
                Case "TableOfContents"
                    MacroTableOfContents()
                Case "TableOfContentsRight"
                    MacroTableOfContentsRight()
                Case "BR"
                    MacroBR()
                Case "TitleSearch"
                    MacroTitleSearch()
                Case "FullSearch"
                    MacroFullSearch()
                Case "TextSearch"
                    MacroTextSearch()
                Case "GoTo"
                    MacroGoTo()
                Case "SystemInfo"
                    MacroSystemInfo()
                Case "Date"
                    MacroDate()
                Case "Time"
                    MacroTime()
                Case "DateTime"
                    MacroDateTime()
                Case "PageCount"
                    MacroPageCount()
                Case "RecentChanges"
                    MacroRecentChanges()
                Case "RecentChangesLong"
                    MacroRecentChangesLong()
                Case "TitleIndex"
                    MacroTitleIndex()
                Case "WordIndex"
                    MacroWordIndex()
                Case "ListRedirects"
                    MacroListRedirects()
                Case "RandomPage"
                    MacroRandomPage()
                Case "InterWiki"
                    MacroInterWiki()
                Case "UserPreferences"
                    MacroUserPreferences()
                Case "CollapseClose"
                    MacroCollapseClose()
                    ' Macros with one parameter
                Case "TitleSearchP"
                    MacroTitleSearchP(vParams)
                Case "AnchorP"
                    MacroAnchorP(vParams)
                Case "FullSearchP"
                    MacroFullSearchP(vParams)
                Case "EquationSearchP"
                    MacroEquationSearchP(vParams)
                Case "TextSearchP"
                    MacroTextSearchP(vParams)
                Case "RecentChangesP"
                    MacroRecentChangesP(CInt(vParams))
                Case "RandomPageP"
                    MacroRandomPageP(CInt(vParams))
                Case "IconP"
                    MacroIconP(vParams)
                Case "IncludeP"
                    MacroIncludeP(vParams)
                Case "ImageP"
                    MacroImageP(vParams)
                Case "CollapseOpenP"
                    MacroCollapseOpenP(vParams)
                    ' Macros with several parameters
                Case "RecentEquationsPP"
                    vParamArray = vParams.Split(CChar(","))
                    MacroRecentEquationsPP(vParamArray(0), CInt(vParamArray(1)), CInt(vParamArray(2)))
                Case "ImagePP"
                    vParamArray = vParams.Split(CChar(","))
                    MacroImagePP(vParamArray(0), vParamArray(1))
            End Select

            If gMacroReturn = "" Then
                sReturn = "&lt;" & pMacro & pParams & "&gt;"
            Else
                StoreRaw(gMacroReturn)
            End If
        End Sub

        Sub MacroTableOfContents()
            If cUseHeadings = 1 Then
                gMacroReturn = gFS & "TOC" & gFS
                ' at the end of the Wikify function this pattern will be
                ' replaced by the actual table of contents
            End If
        End Sub

        Sub MacroTableOfContentsRight()
            If cUseHeadings = 1 Then
                gMacroReturn = gFS & "TOCRight" & gFS
                ' at the end of the Wikify function this pattern will be
                ' replaced by the actual table of contents
            End If
        End Sub

        Sub MacroBR()
            gMacroReturn = "<br />"
        End Sub

        Sub MacroTitleSearch()
            gMacroReturn = "<form id=""TitleSearch"" action=""" & CDATAEncode(gScriptName) & """ method=""get""><div><input type=""hidden"" name=""a"" value=""titlesearch""/><input type=""text"" name=""txt"" value=""" & CDATAEncode(gTxt) & """ ondblclick='event.cancelBubble=true;' /><input id=""mts"" type=""submit"" value=""Go""/></div></form>"
        End Sub

        Sub MacroTitleSearchP(ByVal pParam As String)
            gMacroReturn = gNamespace.GetIndexSchemes.GetTitleSearch(pParam)
        End Sub

        Sub MacroFullSearch()
            gMacroReturn = "<form id=""FullSearch"" action=""" & CDATAEncode(gScriptName) & """ method=""get""><div><input type=""hidden"" name=""a"" value=""fullsearch""/><input type=""text"" name=""txt"" value=""" & CDATAEncode(gTxt) & """ ondblclick='event.cancelBubble=true;' /><input id=""mfs"" type=""submit"" value=""Go""/></div></form>"
        End Sub

        Sub MacroFullSearchP(ByVal pParam As String)
            gMacroReturn = gNamespace.GetIndexSchemes.GetFullSearch(pParam, 1)
        End Sub

        Sub MacroEquationSearchP(ByVal pParam As String)
            gMacroReturn = gNamespace.GetIndexSchemes.GetEquationSearch(pParam, 1)
        End Sub

        Sub MacroTextSearch()
            gMacroReturn = "<form id=""TextSearch"" action=""" & CDATAEncode(gScriptName) & """ method=""get""><div><input type=""hidden"" name=""a"" value=""textsearch""/><input type=""text"" name=""txt"" value=""" & CDATAEncode(gTxt) & """ ondblclick='event.cancelBubble=true;' /><input id=""mfs"" type=""submit"" value=""Go""/></div></form>"
        End Sub

        Sub MacroTextSearchP(ByVal pParam As String)
            gMacroReturn = gNamespace.GetIndexSchemes.GetFullSearch(pParam, 0)
        End Sub

        Sub MacroGoTo()
            gMacroReturn = "<form id=""GoTo"" action=""" & CDATAEncode(gScriptName) & """ method=""get""><div><input type=""text"" name=""p"" value="""" ondblclick='event.cancelBubble=true;' /><input id=""goto"" type=""submit"" value=""Go""/></div></form>"
        End Sub

        Sub MacroSystemInfo()
            '            On Error Resume Next
            Dim vFSO As Scripting.FileSystemObject
            Dim vFile As Scripting.File
            Dim vRev As String
            Dim vConn As ADODB.Connection

            vFSO = New Scripting.FileSystemObject
            vFile = vFSO.GetFile(HttpContext.Current.Server.MapPath(HttpContext.Current.Request.ServerVariables("SCRIPT_NAME")))

            vRev = Mid(OPENWIKI_REVISION, 12, Len(OPENWIKI_REVISION) - 13)
            gMacroReturn = "<table class=""systeminfo"">" _
                    & "<tr><td>OpenWiki Version:</td><td>" & OPENWIKI_VERSION & " rev." & vRev & "</td></tr>" _
                    & "<tr><td>XML Schema Version:</td><td>" & OPENWIKI_XMLVERSION & "</td></tr>" _
                    & "<tr><td>Namespace:</td><td>" & OPENWIKI_NAMESPACE & "</td></tr>" _
                    & "<tr><td>" & ScriptEngine & " Version:</td><td>" & ScriptEngineMajorVersion & "." & ScriptEngineMinorVersion & "." & ScriptEngineBuildVersion & "</td></tr>" _
                    & "<tr><td>MSXML Version:</td><td>" & MSXML_VERSION & "</td></tr>"

            vConn = New ADODB.Connection
            gMacroReturn = gMacroReturn & "<tr><td>ADO Version:</td><td>" & vConn.Version & "</td></tr>"
            vFile = Nothing
            vFSO = Nothing
            gMacroReturn = gMacroReturn & "<tr><td>Nr Of Pages:</td><td>" & gNamespace.GetPageCount() & "</td></tr>"
            gMacroReturn = gMacroReturn & "<tr><td>Nr Of Revisions:</td><td>" & gNamespace.GetRevisionsCount() & "</td></tr>"
            'gMacroReturn = gMacroReturn & "<tr><td>Now:</td><td>" & FormatDate(Now()) & " " & FormatTime(Now()) & "</td></tr>"
            gMacroReturn = gMacroReturn & "</table>"
        End Sub

        Sub MacroDate()
            cCacheXML = 0
            gMacroReturn = FormatDate(Now())
        End Sub

        Sub MacroTime()
            cCacheXML = 0
            gMacroReturn = FormatTime(Now())
        End Sub

        Sub MacroDateTime()
            cCacheXML = 0
            gMacroReturn = FormatDate(Now()) & " " & FormatTime(Now())
        End Sub

        Sub MacroPageCount()
            gMacroReturn = CStr(gNamespace.GetPageCount())
        End Sub

        Sub MacroRecentChanges()
            Call MacroRecentChangesPP(OPENWIKI_RCDAYS, 9999)
        End Sub

        Sub MacroRecentChangesP(ByVal pParams As Integer)
            Call MacroRecentChangesPP(pParams, 9999)
        End Sub

        Sub MacroRecentChangesPP(ByVal pDays As Integer, ByVal pNrOfChanges As Integer)
            If Not IsNumeric(pDays) Or Not IsNumeric(pNrOfChanges) Then
                Exit Sub
            End If
            If pDays <= 0 Then
                pDays = OPENWIKI_RCDAYS
            End If
            If pNrOfChanges <= 0 Then
                pNrOfChanges = 0
            End If
            gMacroReturn = gNamespace.GetIndexSchemes.GetRecentChanges(pDays, pNrOfChanges, 1, 1)
        End Sub

        Sub MacroRecentChangesLong()
            Dim vDays As Integer
            Dim vMaxNrOfChanges As Integer
            Dim vFilter As Integer

            vDays = GetIntParameter("days")
            vMaxNrOfChanges = GetIntParameter("max")
            vFilter = GetIntParameter("filter")
            If vDays <= 0 Then
                vDays = OPENWIKI_RCDAYS
            End If
            If vMaxNrOfChanges <= 0 Then
                If gAction = "rss" Then
                    vMaxNrOfChanges = 15
                Else
                    vMaxNrOfChanges = 9999
                End If
            End If
            If vFilter = 0 Then
                vFilter = 1  ' major edits only
            ElseIf vFilter = 3 Then
                vFilter = 0  ' major and minor edits
            End If
            ' vFilter = 2  ' minor edits only
            gMacroReturn = gNamespace.GetIndexSchemes.GetRecentChanges(vDays, vMaxNrOfChanges, vFilter, 0)
        End Sub

        Sub MacroTitleIndex()
            gMacroReturn = gNamespace.GetIndexSchemes.GetTitleIndex
        End Sub

        Sub MacroWordIndex()
            gMacroReturn = gNamespace.GetIndexSchemes.GetWordIndex
        End Sub

        Sub MacroListRedirects()
            gMacroReturn = gNamespace.ListRedirects()
        End Sub

        Sub MacroRandomPage()
            gMacroReturn = gNamespace.GetIndexSchemes.GetRandomPage(1)
        End Sub

        Sub MacroRandomPageP(ByVal pParam As Integer)
            If IsNumeric(pParam) Then
                gMacroReturn = gNamespace.GetIndexSchemes.GetRandomPage(pParam)
            End If
        End Sub

        Sub MacroCollapseOpenP(ByVal pCaption As String)
            gMacroReturn = "<div class=""NavFrame collapsed"">" _
                & "<div class=""NavHead"">" _
                & pCaption _
                & "</div>" _
                & "<div class=""NavContent"">"
        End Sub

        Sub MacroCollapseClose()
            gMacroReturn = "</div>" _
                & "</div>"
        End Sub

        Sub MacroImageP(ByVal pParam As String)
            gMacroReturn = "<div class=""thumb tright"">" _
                & "<div class=""thumbinner"">" _
                & "<img class=""thumbimage"" src=""" & pParam & """/>" _
                & "</div>" _
                & "</div>"
        End Sub

        Sub MacroImagePP(ByVal pImage As String, ByVal pCaption As String)
            gMacroReturn = "<div class=""thumb tnone"">" _
                & "<div class=""thumbinner"">" _
                & "<img class=""thumbimage"" src=""" & pImage & """/>" _
                & "<div class=""thumbcaption"">" _
                & pCaption _
                & "</div>" _
                & "</div>" _
                & "</div>"
        End Sub

        Sub MacroIconP(ByVal pParam As String)
            gMacroReturn = "<img src=""" & OPENWIKI_ICONPATH & "/" & pParam & ".gif"" border=""0"" alt=""" & pParam & """/>"
        End Sub

        Sub MacroAnchorP(ByVal pParam As String)
            gMacroReturn = "<a id='" & CDATAEncode(pParam) & "'></a>"
        End Sub

        Sub MacroIncludeP(ByVal pParam As String)
            Dim i As Integer
            '            Dim vCount As Integer
            Dim vID As String

            If IsNothing(gCurrentWorkingPages) Then
                gCurrentWorkingPages = New Vector
                gCurrentWorkingPages.Push(gPage)
            End If
            For i = 0 To gCurrentWorkingPages.Count - 1
                If UCase(CStr(gCurrentWorkingPages.ElementAt(i))) = UCase(pParam) Then
                    Exit Sub
                End If
            Next

            vID = AbsoluteName(pParam)

            gIncludeLevel = gIncludeLevel + 1
            If (gIncludeLevel <= OPENWIKI_MAXINCLUDELEVEL) Then
                Dim vPage As WikiPage
                vPage = gNamespace.GetPageAndAttachments(vID, 0, 1, False)
                If vPage.Exists Then
                    gCurrentWorkingPages.Push(vPage.Name)
                    gMacroReturn = vPage.ToXML(1)
                    gCurrentWorkingPages.Pop()
                End If
            End If
            gIncludeLevel = gIncludeLevel - 1
        End Sub

        Sub MacroInterWiki()
            gMacroReturn = gNamespace.InterWiki()
        End Sub

        Sub MacroUserPreferences()
            gMacroReturn = ""
            If HttpContext.Current.Request.QueryString("up") = "1" Then
                gMacroReturn = gMacroReturn & "<ow:message code=""userpreferences_saved""/>"
            ElseIf HttpContext.Current.Request.QueryString("up") = "2" Then
                gMacroReturn = gMacroReturn & "<ow:message code=""userpreferences_cleared""/>"
            End If
            gMacroReturn = gMacroReturn & "<ow:userpreferences/>"
        End Sub

        Function FormatDate(ByVal pTimestamp As Date) As String
            ' TODO: apply user preferences
            FormatDate = (Month(pTimestamp)) & "/" & Day(pTimestamp) & "/" & Year(pTimestamp)
        End Function

        Function FormatTime(ByVal pTimestamp As Date) As String
            ' TODO: apply user preferences
            FormatTime = FormatDateTime(pTimestamp, DateFormat.ShortTime)  ' 4 = vbShortTime
        End Function


        Sub MacroFootnoteP(ByVal pText As String)
            ' processed at the end of wikify
            gMacroReturn = gFS & gFS & pText & gFS & gFS
        End Sub


        Sub MacroAggregateP(ByVal pPage As String)
            If cAllowAggregations <> 1 Then
                Exit Sub
            End If

            If HttpContext.Current.Request("preview") <> "" Then
                Exit Sub
            End If

            pPage = AbsoluteName(pPage)

            Dim vPage As WikiPage
            vPage = gNamespace.GetPage(pPage, gRevision, 1, False)
            gAggregateURLs = New Vector
            MultiLineMarkup(vPage.Text)   ' refreshes RSS feed(s) and fills the gAggregateURLs vector
            gMacroReturn = GetAggregation(pPage)
            gAggregateURLs = Nothing
        End Sub

        Sub MacroSyndicateP(ByVal pURL As String)
            Call MacroSyndicatePP(pURL, 240)  ' default = 4 * 60 minutes
        End Sub

        Sub MacroSyndicatePP(ByVal pURL As String, ByVal pRefreshRate As Integer)
            Dim vURL As String
            Dim vCache As String = ""
            Dim vRefreshURL As String

            If HttpContext.Current.Request("preview") <> "" Then
                Exit Sub
            End If

            vURL = Replace(pURL, "&amp;", "&")
            If Not m(vURL, "^https?://", False, False) Or Not IsNumeric(pRefreshRate) Then
                Exit Sub
            End If
            If pRefreshRate < 0 Then
                pRefreshRate = 0
            End If

            If Not IsNothing(gAggregateURLs) And cAllowAggregations = 1 Then
                gAggregateURLs.Push(vURL)
            End If

            vRefreshURL = URLDecode(HttpContext.Current.Request("refreshurl"))

            If (gAction <> "refresh") Or ((vRefreshURL <> "") And (vRefreshURL <> vURL)) Then
                vCache = gNamespace.GetRSSFromCache(vURL, pRefreshRate, False, False)
                If vCache = "notexists" Then
                    If cAllowNewSyndications = 0 Then
                        Exit Sub
                    End If
                ElseIf vCache <> "" Then
                    gMacroReturn = vCache
                    Exit Sub
                End If
            End If
            If gAction = "refresh" Or vRefreshURL = vURL Or vCache = "notexists" Or gNrOfRSSRetrievals < OPENWIKI_MAXWEBGETS Then
                gMacroReturn = RetrieveRSSFeed(vURL)
                gNrOfRSSRetrievals = gNrOfRSSRetrievals + 1
            End If
            If gMacroReturn = "" Then
                ' failure to retrieve RSS feed from remote source
                If vCache = "notexists" Then
                    Call gNamespace.SaveRSSToCache(vURL, pRefreshRate, "")
                    gMacroReturn = gNamespace.GetRSSFromCache(vURL, pRefreshRate, True, False)
                Else
                    ' retry later, and get the cached version
                    gMacroReturn = gNamespace.GetRSSFromCache(vURL, pRefreshRate, True, True)
                    If gMacroReturn = "notexists" Then
                        gMacroReturn = ""
                    End If
                End If
            Else
                Call gNamespace.SaveRSSToCache(vURL, pRefreshRate, gMacroReturn)
                gMacroReturn = gNamespace.GetRSSFromCache(vURL, pRefreshRate, True, False)
            End If
        End Sub

        Sub MacroRecentNewPagesPP(ByVal pDays As Integer _
            , ByVal pNrOfChanges As Integer)
            If Not IsNumeric(pDays) Or Not IsNumeric(pNrOfChanges) Then
                Exit Sub
            End If
            If pDays <= 0 Then
                pDays = OPENWIKI_RCDAYS
            End If
            If pNrOfChanges <= 0 Then
                pNrOfChanges = 0
            End If
            gMacroReturn = gNamespace.GetIndexSchemes.GetRecentNewPages(pDays, pNrOfChanges, 1, 1)
        End Sub


        Sub MacroRecentEquationsPP(ByVal pPattern As String _
            , ByVal pDays As Integer _
            , ByVal pNrOfChanges As Integer)
            If Not IsNumeric(pDays) Or Not IsNumeric(pNrOfChanges) Then
                Exit Sub
            End If
            If pDays <= 0 Then
                pDays = OPENWIKI_RCDAYS
            End If
            If pNrOfChanges <= 0 Then
                pNrOfChanges = 0
            End If
            gMacroReturn = gNamespace.GetIndexSchemes.GetRecentEquations(pPattern, 1, pDays, pNrOfChanges)
        End Sub
    End Module
End Namespace