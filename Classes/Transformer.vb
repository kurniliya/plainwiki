Imports System.Text.RegularExpressions

Namespace Openwiki
    Public Class Transformer
        Private vXmlDoc As MSXML2.FreeThreadedDOMDocument60
        Private vXslDoc As MSXML2.FreeThreadedDOMDocument60
        Private vXslTemplate As MSXML2.XSLTemplate60
        Private vXslProc As MSXML2.IXSLProcessor
        Private vIsIE As Boolean
        Private vIsGecko As Boolean
        Private vIsMathPlayer As Boolean

        '    Public Property Let MSXML_VERSION(pMSXML_VERSION)
        '        vMSXML_VERSION = pMSXML_VERSION
        '    End Property
        '
        '    Public Property Let Cache(pCacheXSL)
        '        vCacheXSL = pCacheXSL
        '    End Property
        '
        '    Public Property Let StylesheetsDir(pDir)
        '        vStylesheetsDir = pDir
        '    End Property
        '
        '    Public Property Let Encoding(pEncoding)
        '        vEncoding = pEncoding
        '    End Property
        '
        '    Public Property Let WriteToOutput(pWriteToOutput)
        '        vWriteToOutput = pWriteToOutput
        '    End Property

        Public Sub New()
            'If MSXML_VERSION <> 3 Then
            '    vXmlDoc = HttpContext.Current.Server.CreateObject("Msxml2.FreeThreadedDOMDocument." & MSXML_VERSION & ".0")
            vXmlDoc = New MSXML2.FreeThreadedDOMDocument60

            'vXslDoc.resolveExternals = True
            'vXslDoc.setProperty("AllowXsltScript", True)


            'Else
            'vXmlDoc = HttpContext.Current.Server.CreateObject("Msxml2.FreeThreadedDOMDocument")
            'End If

            vXmlDoc.async = False
            vXmlDoc.preserveWhiteSpace = True

            'If Not IsReference(vXmlDoc) Then
            '    ' As this is the first time we try to instantiate the XML Doc object
            '    ' let's assume the user hasn't configured his/her owconfig file
            '    ' correctly yet. Switch MS XML Version to try again.
            '    '            Dim MSXML_VERSION_OLD

            '    MSXML_VERSION_OLD = MSXML_VERSION

            '    If MSXML_VERSION <> 3 Then
            '        MSXML_VERSION = 3
            '    Else
            '        MSXML_VERSION = 6
            '    End If
            '    If MSXML_VERSION <> 3 Then
            '        vXmlDoc = HttpContext.Current.Server.CreateObject("Msxml2.FreeThreadedDOMDocument." & MSXML_VERSION & ".0")
            '        vXslDoc.resolveExternals = True
            '        vXslDoc.setProperty("AllowXsltScript", True)
            '    Else
            '        vXmlDoc = HttpContext.Current.Server.CreateObject("Msxml2.FreeThreadedDOMDocument")
            '    End If
            '    vXmlDoc.async = False
            '    vXmlDoc.preserveWhiteSpace = True
            '    If Not IsReference(vXmlDoc) Then
            '        EndWithErrorMessage()
            '    ElseIf MSXML_VERSION = 3 Then
            '        HttpContext.Current.Response.Write("<b>WARNING:</b>You have configured your OpenWiki to use the MSXML v" & MSXML_VERSION_OLD & " component, but you don't appear to have this installed. The application now falls back to use the MSXML v3 component. Please update your config file (usually file owconfig_default.asp) or install MSXML v" & MSXML_VERSION_OLD & ".<br />")
            '        HttpContext.Current.Response.End()
            '    Else
            '        HttpContext.Current.Response.Write("<b>WARNING:</b>You've configured your OpenWiki to use the MSXML v3 component, but you don't appear to have this installed. The application now falls back to use the MSXML v6 component. Please update your config file (usually file owconfig_default.asp) or install MSXML v6.<br />")
            '        HttpContext.Current.Response.End()
            '    End If
            'End If

            Dim vTemp As String
            vTemp = HttpContext.Current.Request.ServerVariables("HTTP_USER_AGENT")
            If (InStr(vTemp, "MSIE ") > 0) Then
                vIsIE = True
            Else
                vIsIE = False
            End If

            If (InStr(vTemp, "MathPlayer ") > 0) Then
                vIsMathPlayer = True
            Else
                vIsMathPlayer = False
            End If

            If (InStr(vTemp, " Gecko") > 0) Then
                vIsGecko = True
            Else
                vIsGecko = False
            End If
        End Sub

        Protected Overrides Sub Finalize()
            vXslProc = Nothing
            vXslTemplate = Nothing
            vXslDoc = Nothing
            vXmlDoc = Nothing
        End Sub

        Private Sub EndWithErrorMessage()
            HttpContext.Current.Response.Write("<h2>Error: Missing MSXML Parser 3.0 Release</h2>")
            HttpContext.Current.Response.Write("In order for this script to work correctly the component " _
                         & "MSXML Parser 3.0 Release " _
                         & "or a higher version needs to be installed on the server. " _
                         & "You can download this component from " _
                         & "<a href=""http://msdn.microsoft.com/xml"">http://msdn.microsoft.com/xml</a>.")
            HttpContext.Current.Response.End()
        End Sub

        Private Sub ProcessIEWithoutMathPlayer()
            HttpContext.Current.Response.Redirect("static/processie/processie.html")
        End Sub

        Public Sub LoadXSL(ByVal pFilename As String)
            '            On Error Resume Next
            vXslTemplate = Nothing
            'vXslTemplate = ""
            If cCacheXSL = 1 Then
                If Not HttpContext.Current.Application("ow__" & pFilename) Is Nothing Then
                    vXslTemplate = CType(HttpContext.Current.Application("ow__" & pFilename), MSXML2.XSLTemplate60)
                End If
            End If
            If vXslTemplate Is Nothing Then
                'If MSXML_VERSION <> 3 Then
                '    vXslDoc = HttpContext.Current.Server.CreateObject("Msxml2.FreeThreadedDOMDocument." & MSXML_VERSION & ".0")
                vXslDoc = New MSXML2.FreeThreadedDOMDocument60
                vXslDoc.resolveExternals = True
                vXslDoc.setProperty("AllowXsltScript", True)
                'Else
                '    vXslDoc = HttpContext.Current.Server.CreateObject("Msxml2.FreeThreadedDOMDocument")
                'End If
                vXslDoc.async = False
                If Not vXslDoc.load(HttpContext.Current.Server.MapPath(OPENWIKI_STYLESHEETS & pFilename)) Then
                    HttpContext.Current.Response.Write("<p><b>Error in " & pFilename & ":</b> " & vXslDoc.parseError.reason & " line: " & vXslDoc.parseError.line & " col: " & vXslDoc.parseError.linepos & "</p>")
                    HttpContext.Current.Response.End()
                End If
                '                If MSXML_VERSION <> 3 Then
                'vXslTemplate = HttpContext.Current.Server.CreateObject("Msxml2.XSLTemplate." & MSXML_VERSION & ".0")
                'Else
                '    vXslTemplate = HttpContext.Current.Server.CreateObject("Msxml2.XSLTemplate")
                'End If
                vXslTemplate = New MSXML2.XSLTemplate60

                If vXslTemplate Is Nothing Then
                    EndWithErrorMessage()
                End If

                vXslTemplate.stylesheet = vXslDoc
                If Err.Number <> 0 Then
                    HttpContext.Current.Response.Write("<p><b>Error in an included stylesheet</p>")
                    HttpContext.Current.Response.End()
                End If
                'If cCacheXSL Then
                '    HttpContext.Current.Application("ow__" & pFilename) = vXslTemplate
                'End If
            End If

            vXslProc = vXslTemplate.createProcessor()
            If vXslProc Is Nothing Then
                EndWithErrorMessage()
            End If
        End Sub

        Public Function TransformXmlDoc(ByVal pXmlDoc As MSXML2.FreeThreadedDOMDocument60 _
            , ByVal pXslFilename As String) _
        As String
            LoadXSL(pXslFilename)
            vXslProc.input = pXmlDoc
            vXslProc.transform()
            TransformXmlDoc = CStr(vXslProc.output)
        End Function

        Public Function Transform(ByVal pXmlStr As String) As String
            Transform = TransformXmlStr(pXmlStr, "ow.xsl")
        End Function

        Public Function TransformXmlStr(ByVal pXmlStr As String _
            , ByVal pXslFilename As String) _
        As String
            Dim vXmlStr As String

            vXmlStr = "<?xml version='1.0' encoding='" & OPENWIKI_ENCODING & "'?>" & vbCrLf _
                    & gNamespace.ToXML(pXmlStr)

            'Response.ContentType = "text/html"
            'Response.Write(vXmlStr)
            'Response.Write(Server.HTMLEncode(vXmlStr) & "<br /><br />" & vbCRLF & vbCRLF)
            'Response.End

            If (gAction = "xml") Or (HttpContext.Current.Request.QueryString("xml") = "1") Then
                '            If vIsIE Or gAction = "xml" Then
                HttpContext.Current.Response.ContentType = "text/xml; charset=" & OPENWIKI_ENCODING & ";"
                HttpContext.Current.Response.Write(vXmlStr)
                HttpContext.Current.Response.End()
                '            Else
                '                pXslFilename = "xmldisplay.xsl"
                '            End If
            End If

            ' XML error handling: shows source text of a page with lines numbered
            If Not vXmlDoc.loadXML(vXmlStr) Then
                Dim vText As String
                Dim vLineNum As Integer
                Dim vMatches As MatchCollection
                Dim vMatch As Match
                Dim vErrorLine As Integer
                vText = HttpContext.Current.Server.HtmlEncode(vXmlStr)
                vLineNum = 1
                vErrorLine = vXmlDoc.parseError.line

                'vRegEx = New Regexp
                'vRegEx.IgnoreCase = False
                'vRegEx.Global = True
                'vRegEx.Pattern = ".+"
                vMatches = Regex.Matches(vText, ".+")

                HttpContext.Current.Response.ContentType = "text/html;" ' charset=" & OPENWIKI_ENCODING & ";"
                HttpContext.Current.Response.Write("<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">")
                HttpContext.Current.Response.Write("<html xmlns=""http://www.w3.org/1999/xhtml"">")
                HttpContext.Current.Response.Write("<head><title>Invalid XML document</title></head>")
                HttpContext.Current.Response.Write("<body><b>Invalid XML document</b>:<br /><br />")
                HttpContext.Current.Response.Write(vXmlDoc.parseError.reason & " line: " & _
                  "<a href=#ErrorLine>" & _
                  vErrorLine & _
                  "</a>" & _
                  " col: " & vXmlDoc.parseError.linepos)
                HttpContext.Current.Response.Write("<br /><br /><hr />")
                HttpContext.Current.Response.Write("<pre>")

                For Each vMatch In vMatches
                    HttpContext.Current.Response.Write(vLineNum & ": ")
                    If vLineNum = vErrorLine Then
                        HttpContext.Current.Response.Write("<a name=""ErrorLine"" />")
                        HttpContext.Current.Response.Write("<font style=""BACKGROUND-COLOR: red"">")
                    End If
                    HttpContext.Current.Response.Write(vMatch.Value & "<br />")
                    If vLineNum = vErrorLine Then
                        HttpContext.Current.Response.Write("</font>")
                    End If
                    vLineNum = vLineNum + 1
                Next

                HttpContext.Current.Response.Write("</pre>")
                HttpContext.Current.Response.Write("</body></html>")
            ElseIf vIsIE And cUseXhtmlHttpHeaders = 1 And Not vIsMathPlayer Then
                ProcessIEWithoutMathPlayer()
            Else
                LoadXSL(pXslFilename)
                vXslProc.input = vXmlDoc
                vXslProc.transform()

                TransformXmlStr = CStr(vXslProc.output)

                If cEmbeddedMode = 0 Then
                    If gAction = "edit" Then
                        If (cUseXhtmlHttpHeaders = 1) Then
                            ' 						IE+MathPlayer workaround: in ContentType must be specified just "content type"
                            '						Response.ContentType = "application/xhtml+xml; charset=" & OPENWIKI_ENCODING & ";"
                            HttpContext.Current.Response.ContentType = "application/xhtml+xml"
                        Else
                            HttpContext.Current.Response.ContentType = "text/html; charset=" & OPENWIKI_ENCODING & ";"
                        End If
                        HttpContext.Current.Response.Expires = 0   ' expires in a minute
                    ElseIf gAction = "rss" Then
                        HttpContext.Current.Response.ContentType = "text/xml; charset=" & OPENWIKI_ENCODING & ";"
                    Else
                        If cUseXhtmlHttpHeaders = 1 Then
                            ' 						IE+MathPlayer workaround: in ContentType must be specified just "content type"
                            '						Response.ContentType = "application/xhtml+xml; charset=" & OPENWIKI_ENCODING & ";"
                            HttpContext.Current.Response.ContentType = "application/xhtml+xml"
                        Else
                            HttpContext.Current.Response.ContentType = "text/html; charset=" & OPENWIKI_ENCODING & ";"
                        End If
                        HttpContext.Current.Response.Expires = -1  ' expires now
                        '                    HttpContext.Current.Response.AddHeader "Last-modified", DateToHTTPDate(gLastModified)
                        HttpContext.Current.Response.AddHeader("Last-modified", DateToHTTPDate(Now))
                        'Response.ExpiresAbsolute = Now() - 1
                        'Response.AddHeader "Cache-Control", "must-revalidate"
                        HttpContext.Current.Response.AddHeader("Cache-Control", "no-cache")
                    End If
                    HttpContext.Current.Response.Write(TransformXmlStr)
                End If

            End If
        End Function

    End Class
End Namespace