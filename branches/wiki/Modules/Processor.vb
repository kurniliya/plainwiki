Imports System.Text.RegularExpressions

Namespace Openwiki
    Module Processor
        Sub OwProcessRequest()
            Dim SCRIPT_NAME As String
            Dim SERVER_NAME As String
            Dim SERVER_PORT As Integer
            Dim SERVER_PORT_SECURE As Integer
            Dim SlashPos As Integer
            Dim CookieReadPassword As String

            SCRIPT_NAME = HttpContext.Current.Request.ServerVariables("SCRIPT_NAME")
            SERVER_NAME = HttpContext.Current.Request.ServerVariables("SERVER_NAME")
            SERVER_PORT = CInt(HttpContext.Current.Request.ServerVariables("SERVER_PORT"))
            SERVER_PORT_SECURE = CInt(HttpContext.Current.Request.ServerVariables("SERVER_PORT_SECURE"))

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
                SlashPos = InStrRev(SCRIPT_NAME, "/")
                If SlashPos > 0 Then
                    gScriptName = Mid(SCRIPT_NAME, SlashPos + 1)
                Else
                    gScriptName = SCRIPT_NAME
                End If
            End If

            gCookieHash = "C" & Hash(gServerRoot & SCRIPT_NAME)

            If Not (HttpContext.Current.Request.Cookies(gCookieHash & "?up") Is Nothing) Then
                If HttpContext.Current.Request.Cookies(gCookieHash & "?up")("pwl") = "1" Then
                    cPrettyLinks = 1
                Else
                    cPrettyLinks = 0
                End If
                If HttpContext.Current.Request.Cookies(gCookieHash & "?up")("new") = "1" Then
                    cExternalOut = 1
                Else
                    cExternalOut = 0
                End If
                If HttpContext.Current.Request.Cookies(gCookieHash & "?up")("emo") = "1" Then
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
                CookieReadPassword = HttpContext.Current.Request.Cookies(gCookieHash & "?pr").Value
                If CookieReadPassword <> gReadPassword Then
                    gAction = "login"
                End If
            End If

            If Not m(OPENWIKI_TIMEZONE, "^[+|-](0\d|1[0-2]):[0-5]\d$", False, False) Then
                OPENWIKI_TIMEZONE = "+00:00"
            End If

            gActionReturn = False

            'Execute("Call Action" & gAction)
            Select Case gAction
                Case "attach"
                    ActionAttach()
                Case "attachchanges"
                    ActionAttachchanges()
                Case "changes"
                    ActionChanges()
                Case "diff"
                    ActionDiff()
                Case "edit"
                    ActionEdit()
                Case "fullsearch"
                    ActionFullSearch()
                Case "hidefile"
                    ActionHidefile()
                Case "login"
                    ActionLogin()
                Case "logout"
                    ActionLogout()
                Case "naked"
                    ActionNaked()
                Case "preview"
                    ActionPreview()
                Case "print"
                    ActionPrint()
                Case "randompage"
                    ActionRandomPage()
                Case "refresh"
                    ActionRefresh()
                Case "rss"
                    ActionRss()
                Case "textsearch"
                    ActionTextSearch()
                Case "titlesearch"
                    ActionTitleSearch()
                Case "trashfile"
                    ActionTrashfile()
                Case "undohidefile"
                    ActionUndohidefile()
                Case "undotrashfile"
                    ActionUndotrashfile()
                Case "upload"
                    ActionUpload()
                Case "userpreferences"
                    ActionUserPreferences()
                Case "view"
                    ActionView()
                Case "xml"
                    ActionXml()
            End Select

            If Not gActionReturn Then
                HttpContext.Current.Response.ContentType = "text/xml; charset:" & OPENWIKI_ENCODING & ";"
                HttpContext.Current.Response.Write("<?xml version='1.0'?><error>Illegal action</error>")
                HttpContext.Current.Response.End()
            End If

            gTransformer = Nothing
            gNamespace = Nothing
        End Sub

        Function TransformEmbedded(ByVal pText As String) As String
            Dim vPage As WikiPage

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



        ' As you may notice I never use HttpContext.Current.Request.Query and/or HttpContext.Current.Request.Form, only Request(vSomeParam).
        ' The reason is that I plan to support platforms where submitted data cannot always go
        ' through a form, but only through the use of a URL. Another reason is that the presentation
        ' layer is separated from the logic, therefor no assumption should be made about whether the
        ' parameters are passed through the URL or posted as a form.
        '______________________________________________________________________________________________________________
        Sub ParseQueryString()
            gPage = HttpContext.Current.Request("p")
            Dim vPos As Integer, vPos2 As Integer
            If gPage = "" Then
                gPage = HttpContext.Current.Request.ServerVariables("QUERY_STRING")
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
                        gPage = ""   ' gNamespace.Frontpage
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

            gAction = HttpContext.Current.Request("a")
            If gAction = "" Then
                gAction = "view"
            End If

            If HttpContext.Current.Request("refresh") <> "" Then
                cCacheXML = 0
            End If

            gTxt = HttpContext.Current.Request("txt")
        End Sub


        Function GetIntParameter(ByVal pParam As String) As Integer
            Dim Temp As String
            Temp = HttpContext.Current.Request(pParam)
            If IsNumeric(Temp) Then
                GetIntParameter = CInt(Temp)
            Else
                GetIntParameter = 0
            End If
        End Function


        Function getUserPreferences() As String
            Dim vMatches As MatchCollection
            Dim vMatch As Match
            Dim vValue As String
            Dim vUsername As String = Nothing
            Dim vCols As Integer
            Dim vRows As Integer
            Dim vBookmarks As String = Nothing

            getUserPreferences = ""

            If Not HttpContext.Current.Request.Cookies(gCookieHash & "?up") Is Nothing Then

                vCols = CInt(HttpContext.Current.Request.Cookies(gCookieHash & "?up")("cols"))
                If vCols <= 0 Then
                    vCols = 55
                End If
                vRows = CInt(HttpContext.Current.Request.Cookies(gCookieHash & "?up")("rows"))
                If vRows <= 0 Then
                    vRows = 25
                End If
                vBookmarks = HttpContext.Current.Request.Cookies(gCookieHash & "?up")("bm")
                If vBookmarks = "" Then
                    vBookmarks = gDefaultBookmarks
                End If
                'vRegEx = New RegEx
                'vRegEx.IgnoreCase = False
                'vRegEx.Global = True
                'vRegEx.Pattern = "\s+([^ ]*)"
                vMatches = Regex.Matches(" " & Trim(vBookmarks), "\s+([^ ]*)")
                vBookmarks = ""
                For Each vMatch In vMatches
                    vValue = Mid(vMatch.Value, 2)
                    vBookmarks = vBookmarks & ToLinkXML(vValue)
                Next
                vBookmarks = "<ow:bookmarks>" & vBookmarks & "</ow:bookmarks>"
                '            vRegEx = Nothing
                vMatches = Nothing
                vMatch = Nothing

                If Not (HttpContext.Current.Request.Cookies(gCookieHash & "?up") Is Nothing) Then
                    If (cPrettyLinks = 1) Then
                        getUserPreferences = getUserPreferences & "<ow:prettywikilinks/>"
                    End If
                    If (cExternalOut = 1) Then
                        getUserPreferences = getUserPreferences & "<ow:opennew/>"
                    End If
                    If (cEmoticons = 1) Then
                        getUserPreferences = getUserPreferences & "<ow:emoticons/>"
                    End If
                    getUserPreferences = getUserPreferences & "<ow:bookmarksontop/><ow:editlinkontop/><ow:trailontop/>"
                Else
                    If HttpContext.Current.Request.Cookies(gCookieHash & "?up")("pwl") = "1" Then
                        getUserPreferences = getUserPreferences & "<ow:prettywikilinks/>"
                    End If
                    If HttpContext.Current.Request.Cookies(gCookieHash & "?up")("bmt") = "1" Then
                        getUserPreferences = getUserPreferences & "<ow:bookmarksontop/>"
                    End If
                    If HttpContext.Current.Request.Cookies(gCookieHash & "?up")("elt") = "1" Then
                        getUserPreferences = getUserPreferences & "<ow:editlinkontop/>"
                    End If
                    If HttpContext.Current.Request.Cookies(gCookieHash & "?up")("trt") = "1" Then
                        getUserPreferences = getUserPreferences & "<ow:trailontop/>"
                    End If
                    If HttpContext.Current.Request.Cookies(gCookieHash & "?up")("new") = "1" Then
                        getUserPreferences = getUserPreferences & "<ow:opennew/>"
                    End If
                    If HttpContext.Current.Request.Cookies(gCookieHash & "?up")("emo") = "1" Then
                        getUserPreferences = getUserPreferences & "<ow:emoticons/>"
                    End If
                End If

                vUsername = HttpContext.Current.Request.Cookies(gCookieHash & "?up")("un")
                If cNTAuthentication = 1 And vUsername = "" Then
                    vUsername = GetRemoteUser()
                End If

            End If

            getUserPreferences = "<ow:userpreferences>" _
                    & "<ow:cols>" & vCols & "</ow:cols>" _
                    & "<ow:rows>" & vRows & "</ow:rows>" _
                    & "<ow:username>" & vUsername & "</ow:username>" _
                    & vBookmarks _
                    & getUserPreferences _
                    & "</ow:userpreferences>"
        End Function

        Dim gCookieTrail As Vector
        Sub AddCookieTrail(ByVal pPage As String)
            gCookieTrail.Push(pPage)
        End Sub

        Function GetCookieTrail() As String
            Dim vTrailStr As String
            '            Dim vLast As Integer
            Dim vCount As Integer
            Dim vExists As Boolean
            Dim vElem As String
            Dim i As Integer

            GetCookieTrail = ""

            If Not HttpContext.Current.Request.Cookies(gCookieHash & "?tr") Is Nothing Then
                vTrailStr = HttpContext.Current.Request.Cookies(gCookieHash & "?tr")("trail")

                gCookieTrail = New Vector
                s(vTrailStr, "#(.*?)#", "&AddCookieTrail($1)", False, True)

                vTrailStr = ""
                vExists = False
                vCount = gCookieTrail.Count
                For i = 1 To vCount - 1
                    vElem = CStr(gCookieTrail.ElementAt(i))
                    If vElem = gPage Then
                        vExists = True
                    Else
                        GetCookieTrail = GetCookieTrail & ToLinkXML(vElem)
                        vTrailStr = vTrailStr & "#" & vElem & "#"
                    End If
                Next
                If vExists Or (vCount < OPENWIKI_MAXTRAIL) Then
                    If vCount > 0 Then
                        vElem = CStr(gCookieTrail.ElementAt(0))
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

                HttpContext.Current.Response.Cookies(gCookieHash & "?tr")("trail") = vTrailStr
                HttpContext.Current.Response.Cookies(gCookieHash & "?tr")("last") = gPage
            End If

            gCookieTrail = Nothing
            GetCookieTrail = "<ow:trail>" & GetCookieTrail & "</ow:trail>"
        End Function

        Function ToLinkXML(ByVal pID As String) As String
            Dim vTemp As String
            If gAction = "print" Then
                vTemp = gScriptName & "?p=" & HttpContext.Current.Server.UrlEncode(pID) & "&amp;a=print"
            Else
                vTemp = gScriptName & "?" & HttpContext.Current.Server.UrlEncode(pID)
            End If
            ToLinkXML = "<ow:link name=""" & CDATAEncode(pID) & """ href=""" & vTemp & """>" & PCDATAEncode(PrettyWikiLink(pID)) & "</ow:link>"
        End Function


        Function GetCookieTrail_Alternative() As String
            Dim vTrailStr As String, vLast As String, vCount As Integer, vExists As Boolean, vElem As String, i As Integer, vPosLast As Integer, vStart As Integer, vEnd As Integer

            vTrailStr = HttpContext.Current.Request.Cookies(gCookieHash & "?tr")("trail")
            vLast = HttpContext.Current.Request.Cookies(gCookieHash & "?tr")("last")

            gCookieTrail = New Vector
            Call s(vTrailStr, "#(.*?)#", "&AddCookieTrail($1)", False, True)

            vExists = False
            vCount = gCookieTrail.Count
            vPosLast = OPENWIKI_MAXTRAIL
            For i = 0 To vCount - 1
                vElem = CStr(gCookieTrail.ElementAt(i))
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
                vElem = CStr(gCookieTrail.ElementAt(i))
                GetCookieTrail_Alternative = GetCookieTrail() & "<ow:trailmark name='" & CDATAEncode(vElem) & "'>" & PCDATAEncode(PrettyWikiLink(vElem)) & "</ow:trailmark>"
                vTrailStr = vTrailStr & "#" & vElem & "#"
            Next

            If (Not vExists) And ((vEnd - vStart + 1) < OPENWIKI_MAXTRAIL) Then
                vElem = gPage
                GetCookieTrail_Alternative = GetCookieTrail() & "<ow:trailmark name='" & CDATAEncode(vElem) & "'>" & PCDATAEncode(PrettyWikiLink(vElem)) & "</ow:trailmark>"
                vTrailStr = vTrailStr & "#" & vElem & "#"
                HttpContext.Current.Response.Cookies(gCookieHash & "?tr")("trail") = vTrailStr
            End If

            HttpContext.Current.Response.Cookies(gCookieHash & "?tr")("last") = gPage

            gCookieTrail = Nothing
            GetCookieTrail_Alternative = "<ow:trail>" & GetCookieTrail() & "</ow:trail>"
        End Function


        Public Function Hash(ByVal pText As String) As Integer
            Dim i As Integer, vCount As Integer, vMax As Integer
            vMax = CInt(2 ^ 30)
            Hash = 0
            vCount = Len(pText)
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

    End Module
End Namespace