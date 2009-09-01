Namespace Openwiki
    Module Rss
        Function RetrieveRSSFeed(ByVal pURL As String) As String
            Dim vXmlDoc As MSXML2.FreeThreadedDOMDocument60
            Dim vRoot As MSXML2.IXMLDOMElement
            Dim vXslFilename As String

            '            On Error Resume Next
            'Response.Write("Retrieving " & pURL & "<br />")

            vXmlDoc = RetrieveXML(pURL)

            vRoot = vXmlDoc.documentElement

            ' determine the type of the feed
            If vRoot.NodeName = "rss" Then
                vXslFilename = "owrss091.xsl"
            ElseIf vRoot.NodeName = "scriptingNews" Then
                vXslFilename = "owscriptingnews.xsl"
            ElseIf vRoot.getAttribute("xmlns").ToString = "http://my.netscape.com/rdf/simple/0.9/" Then
                vXslFilename = "owrss09.xsl"
            ElseIf vRoot.getAttribute("xmlns").ToString = "http://purl.org/rss/1.0/" Then
                ' TODO: find workaround for bug in MSXML v4
                If Not vRoot.selectSingleNode("item/ag:source") Is Nothing Then
                    vXslFilename = "owrss10aggr.xsl"
                Else
                    vXslFilename = "owrss10.xsl"
                End If
            Else
                Exit Function
            End If

            If cAllowAggregations = 1 Then
                Call gNamespace.Aggregate(pURL, vXmlDoc)
            End If

            RetrieveRSSFeed = gTransformer.TransformXmlDoc(vXmlDoc, vXslFilename)

            ' strip away any <script> elements, rigorously
            ' avoid running security risk of malicious javascript code
            RetrieveRSSFeed = s(RetrieveRSSFeed, "<script(.*?)script>", "", True, True)
        End Function


        ' retrieve the XML data from the given URL
        Function RetrieveXML(ByVal pURL As String) As MSXML2.FreeThreadedDOMDocument60
            Dim vXmlDoc As MSXML2.FreeThreadedDOMDocument60
            Dim vXmlHttp As MSXML2.ServerXMLHTTP60
            Dim vXmlStr As String
            Dim vPos As Integer
            Dim vPosEnd As Integer

            'If MSXML_VERSION <> 3 Then
            '    vXmlHttp = HttpContext.Current.Server.CreateObject("Msxml2.ServerXMLHTTP." & MSXML_VERSION & ".0")
            'Else
            '    vXmlHttp = HttpContext.Current.Server.CreateObject("Msxml2.ServerXMLHTTP")
            'End If
            vXmlHttp = New MSXML2.ServerXMLHTTP60

            vXmlHttp.Open("GET", pURL, False)
            vXmlHttp.send("")

            vXmlDoc = CType(vXmlHttp.responseXML, MSXML2.FreeThreadedDOMDocument60)
            If vXmlDoc.xml = "" Then
                ' sometimes (quite often actually) an RSS feed can't be
                ' loaded into the DOM directly. This is usually because the
                ' feed is send with content-type text/plain instead of text/xml.
                ' For example, the RSS feeds from kuro5hin and salon.com won't
                ' load properly, resulting in an empty XML document object.
                '
                ' therefore, alternative method: first get the document as a string.
                vXmlStr = vXmlHttp.ResponseText

                ' unbelievable, but true, valid ISO-8859-1 characters in the vXmlStr
                ' variable won't load in a DOM document, here's an (imperfect) trick:
                vXmlStr = HttpContext.Current.Server.HtmlEncode(vXmlStr)
                vXmlStr = Replace(vXmlStr, "&gt;", ">")
                vXmlStr = Replace(vXmlStr, "&lt;", "<")
                vXmlStr = Replace(vXmlStr, "&amp;", "&")
                vXmlStr = Replace(vXmlStr, "&quot;", """")
                vXmlStr = Replace(vXmlStr, "&#65535;", "?")

                ' the next stumbling block is that some contain the
                ' <!DOCTYPE ...> string which, although it's perfectly valid
                ' in XML world, for some really maddening reason won't load
                ' into an XML document object as well.
                '
                ' therefore, first strip it away
                vPos = InStr(vXmlStr, "<!DOCTYPE ")
                If vPos > 0 Then
                    vPosEnd = InStr(vPos, vXmlStr, ">")
                    If vPosEnd > 0 Then
                        ' note: conveniently assume UTF-8 encoding
                        vXmlStr = "<?xml version='1.0'?>" & Mid(vXmlStr, vPosEnd + 1)
                    End If
                End If
                'Response.Write("<b><a href='" & pURL & "' target='_blank'>" & pURL & "</a></b><br />" & HttpContext.Current.Server.HTMLEncode(vXmlStr) & "<br /><br />")

                ' and finally we can, hopefully, get it loaded as an xml document object
                'If MSXML_VERSION <> 3 Then
                '    vXmlDoc = HttpContext.Current.Server.CreateObject("Msxml2.FreeThreadedDOMDocument." & MSXML_VERSION & ".0")
                '    'vXslDoc.ResolveExternals = True
                '    'vXslDoc.setProperty("AllowXsltScript", True)
                'Else
                '    vXmlDoc = HttpContext.Current.Server.CreateObject("Msxml2.FreeThreadedDOMDocument")
                'End If
                vXmlDoc = New MSXML2.FreeThreadedDOMDocument60

                vXmlDoc.async = False
                If Not vXmlDoc.loadXML(vXmlStr) Then
                    ' sometimes this fails because of character endoding issues.
                    ' if anyone knows a solid way to load XML feeds from other
                    ' servers, plz let us know! -- LP
                    'Response.Write("<p><b>Error</b> " & vXmlDoc.parseError.reason & " line: " & vXmlDoc.parseError.Line & " col: " & vXmlDoc.parseError.linepos & "</p>")
                    Exit Function
                End If
            End If
            RetrieveXML = vXmlDoc
        End Function


        Function GetAggregation(ByVal pPage As String) As String
            Dim vXmlStr As String
            Dim vXmlDoc As MSXML2.FreeThreadedDOMDocument60

            '            On Error Resume Next

            If Not IsReference(gAggregateURLs) Then
                Exit Function
            End If
            If gAggregateURLs.Count = 0 Then
                Exit Function
            End If

            vXmlStr = gNamespace.GetAggregation(gAggregateURLs)

            'If MSXML_VERSION <> 3 Then
            '    vXmlDoc = HttpContext.Current.Server.CreateObject("Msxml2.FreeThreadedDOMDocument." & MSXML_VERSION & ".0")
            '    'vXslDoc.ResolveExternals = True
            '    'vXslDoc.setProperty("AllowXsltScript", True)
            'Else
            '    vXmlDoc = HttpContext.Current.Server.CreateObject("Msxml2.FreeThreadedDOMDocument")
            'End If
            vXmlDoc = New MSXML2.FreeThreadedDOMDocument60

            vXmlDoc.async = False
            If Not vXmlDoc.loadXML(vXmlStr) Then
                'Response.Write("<p><b>Error</b> " & vXmlDoc.parseError.reason & " line: " & vXmlDoc.parseError.Line & " col: " & vXmlDoc.parseError.linepos & "</p>")
                Exit Function
            End If

            vXmlStr = gTransformer.TransformXmlDoc(vXmlDoc, "owrss10aggr.xsl")

            ' strip away any <script> elements, rigorously
            ' avoid running security risk of malicious javascript code
            vXmlStr = s(vXmlStr, "<script(.*?)script>", "", True, True)

            GetAggregation = "<ow:aggregation href='" & CDATAEncode(gScriptName & "?p=" & pPage & "&a=rss") & "' " _
                           & "refreshURL='" & CDATAEncode(gScriptName & "?p=" & pPage & "&a=refresh&redirect=" & gPage) & "' "
            If Not vXmlDoc.documentElement.selectSingleNode("item/ag:timestamp") Is Nothing Then
                GetAggregation = GetAggregation & "last='" & vXmlDoc.documentElement.selectSingleNode("item/ag:timestamp").text & "' "
            End If
            If HttpContext.Current.Request("refresh") = "" Then
                GetAggregation = GetAggregation & "fresh='false'"
            Else
                GetAggregation = GetAggregation & "fresh='true'"
            End If
            GetAggregation = GetAggregation & ">" & vXmlStr & "</ow:aggregation>"
        End Function
    End Module
End Namespace