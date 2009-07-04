
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
'      $Source: /usr/local/cvsroot/openwiki/dist/owbase/ow/owtransformer.asp,v $
'    $Revision: 1.6 $
'      $Author: pit $
' ---------------------------------------------------------------------------
'

Class Transformer
Private vXmlDoc As MSXML2.FreeThreadedDOMDocument
	Private vXslDoc As MSXML2.FreeThreadedDOMDocument
	Private vXslTemplate As MSXML2.XSLTemplate
	Private vIsGecko, vXslProc, vIsIE, vIsMathPlayer As Object
	
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
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1061.asp'
	Private Sub Class_Initialize_Renamed()
		Dim MSXML_VERSION As Object
		On Error Resume Next
		If MSXML_VERSION <> 3 Then
'UPGRADE_NOTE: The 'Msxml2.FreeThreadedDOMDocument." & MSXML_VERSION & ".0' object is not registered in the migration machine. Copy this link in your browser for more: ms-its:C:\Soft\Dev\ASP to ASP.NET Migration Assistant\AspToAspNet.chm::/1016.htm
			vXmlDoc = HttpContext.Current.Server.CreateObject("Msxml2.FreeThreadedDOMDocument." & MSXML_VERSION & ".0")
			vXslDoc.ResolveExternals = True
			vXslDoc.setProperty("AllowXsltScript", True)
		Else
			vXmlDoc = New MSXML2.FreeThreadedDOMDocument
		End If
		vXmlDoc.async = False
		vXmlDoc.preserveWhiteSpace = True
		Dim MSXML_VERSION_OLD As Object
		If IsNothing(vXmlDoc) Then
			' As this is the first time we try to instantiate the XML Doc object
			' let's assume the user hasn't configured his/her owconfig file
			' correctly yet. Switch MS XML Version to try again.
			
			MSXML_VERSION_OLD = MSXML_VERSION
			
			If MSXML_VERSION <> 3 Then
				MSXML_VERSION = 3
			Else
				MSXML_VERSION = 6
			End If
			If MSXML_VERSION <> 3 Then
'UPGRADE_NOTE: The 'Msxml2.FreeThreadedDOMDocument." & MSXML_VERSION & ".0' object is not registered in the migration machine. Copy this link in your browser for more: ms-its:C:\Soft\Dev\ASP to ASP.NET Migration Assistant\AspToAspNet.chm::/1016.htm
				vXmlDoc = HttpContext.Current.Server.CreateObject("Msxml2.FreeThreadedDOMDocument." & MSXML_VERSION & ".0")
				vXslDoc.ResolveExternals = True
				vXslDoc.setProperty("AllowXsltScript", True)
			Else
				vXmlDoc = New MSXML2.FreeThreadedDOMDocument
			End If
			vXmlDoc.async = False
			vXmlDoc.preserveWhiteSpace = True
			If IsNothing(vXmlDoc) Then
				EndWithErrorMessage()
			ElseIf MSXML_VERSION = 3 Then 
				HttpContext.Current.Response.Write("<b>WARNING:</b>You have configured your OpenWiki to use the MSXML v" & MSXML_VERSION_OLD & " component, but you don't appear to have this installed. The application now falls back to use the MSXML v3 component. Please update your config file (usually file owconfig_default.asp) or install MSXML v" & MSXML_VERSION_OLD & ".<br />")
				HttpContext.Current.Response.End()
			Else
				HttpContext.Current.Response.Write("<b>WARNING:</b>You've configured your OpenWiki to use the MSXML v3 component, but you don't appear to have this installed. The application now falls back to use the MSXML v6 component. Please update your config file (usually file owconfig_default.asp) or install MSXML v6.<br />")
				HttpContext.Current.Response.End()
			End If
		End If
		
		Dim vTemp As Object
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
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1061.asp'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object vXslProc may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
		vXslProc = Nothing
		'UPGRADE_NOTE: Object vXslTemplate may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
		vXslTemplate = Nothing
		'UPGRADE_NOTE: Object vXslDoc may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
		vXslDoc = Nothing
		'UPGRADE_NOTE: Object vXmlDoc may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
		vXmlDoc = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	Private Sub EndWithErrorMessage()
		HttpContext.Current.Response.Write("<h2>Error: Missing MSXML Parser 3.0 Release</h2>")
		HttpContext.Current.Response.Write("In order for this script to work correctly the component " & "MSXML Parser 3.0 Release " & "or a higher version needs to be installed on the server. " & "You can download this component from " & "<a href=""http://msdn.microsoft.com/xml"">http://msdn.microsoft.com/xml</a>.")
		HttpContext.Current.Response.End()
	End Sub
	
	Private Sub ProcessIEWithoutMathPlayer()
		HttpContext.Current.Response.Redirect("static/processie/processie.html")
	End Sub
	
	Public Sub LoadXSL(ByRef pFilename As Object)
		Dim OPENWIKI_STYLESHEETS As Object
		Dim MSXML_VERSION As Object
		Dim cCacheXSL As Object
		On Error Resume Next
		'UPGRADE_NOTE: Object vXslTemplate may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
		vXslTemplate = Nothing
		vXslTemplate = ""
		If cCacheXSL = 1 Then
			If Not IsNothing(HttpContext.Current.Application("ow__" & pFilename)) Then
				vXslTemplate = HttpContext.Current.Application("ow__" & pFilename)
			End If
		End If
		If IsNothing(vXslTemplate) Then
			If MSXML_VERSION <> 3 Then
'UPGRADE_NOTE: The 'Msxml2.FreeThreadedDOMDocument." & MSXML_VERSION & ".0' object is not registered in the migration machine. Copy this link in your browser for more: ms-its:C:\Soft\Dev\ASP to ASP.NET Migration Assistant\AspToAspNet.chm::/1016.htm
				vXslDoc = HttpContext.Current.Server.CreateObject("Msxml2.FreeThreadedDOMDocument." & MSXML_VERSION & ".0")
				vXslDoc.ResolveExternals = True
				vXslDoc.setProperty("AllowXsltScript", True)
			Else
				vXslDoc = New MSXML2.FreeThreadedDOMDocument
			End If
			vXslDoc.async = False
			If Not vXslDoc.load(HttpContext.Current.Server.MapPath(OPENWIKI_STYLESHEETS & pFilename)) Then
				HttpContext.Current.Response.Write("<p><b>Error in " & pFilename & ":</b> " & vXslDoc.parseError.reason & " line: " & vXslDoc.parseError.Line & " col: " & vXslDoc.parseError.linepos & "</p>")
				HttpContext.Current.Response.End()
			End If
			If MSXML_VERSION <> 3 Then
'UPGRADE_NOTE: The 'Msxml2.XSLTemplate." & MSXML_VERSION & ".0' object is not registered in the migration machine. Copy this link in your browser for more: ms-its:C:\Soft\Dev\ASP to ASP.NET Migration Assistant\AspToAspNet.chm::/1016.htm
				vXslTemplate = HttpContext.Current.Server.CreateObject("Msxml2.XSLTemplate." & MSXML_VERSION & ".0")
			Else
				vXslTemplate = New MSXML2.XSLTemplate
			End If
			If IsNothing(vXslTemplate) Then
				EndWithErrorMessage()
			End If
			vXslTemplate.stylesheet = vXslDoc
			If Err.Number <> 0 Then
				HttpContext.Current.Response.Write("<p><b>Error in an included stylesheet</p>")
				HttpContext.Current.Response.End()
			End If
			If cCacheXSL Then
				HttpContext.Current.Application("ow__" & pFilename) = vXslTemplate
			End If
		End If
		vXslProc = vXslTemplate.createProcessor()
		If IsNothing(vXslProc) Then
			EndWithErrorMessage()
		End If
		On Error GoTo 0
	End Sub
	
	Public Function TransformXmlDoc(ByRef pXmlDoc As Object, ByRef pXslFilename As Object) As Object
		LoadXSL((pXslFilename))
		vXslProc.input = pXmlDoc
		vXslProc.Transform()
		TransformXmlDoc = vXslProc.output
	End Function
	
	Public Function Transform(ByRef pXmlStr As Object) As Object
		Transform = TransformXmlStr(pXmlStr, "ow.xsl")
	End Function
	
	Public Function TransformXmlStr(ByRef pXmlStr As Object, ByRef pXslFilename As Object) As Object
		Dim cEmbeddedMode As Object
		Dim cUseXhtmlHttpHeaders As Object
		Dim gAction As Object
		Dim gNamespace As Object
		Dim OPENWIKI_ENCODING As Object
		Dim vXmlStr As Object
		vXmlStr = "<?xml version='1.0' encoding='" & OPENWIKI_ENCODING & "'?>" & vbCrLf & gNamespace.ToXML(pXmlStr)
		
		'Response.ContentType = "text/html"
		'Response.Write(vXmlStr)
		'Response.Write(Server.HTMLEncode(vXmlStr) & "<br /><br />" & vbCRLF & vbCRLF)
		'Response.End
		
		If gAction = "xml" Or InStrRev(HttpContext.Current.Request.QueryString, "&xml=1") > 0 Then
			'            If vIsIE Or gAction = "xml" Then
			HttpContext.Current.Response.ContentType = "text/xml; charset=" & OPENWIKI_ENCODING & ";"
			HttpContext.Current.Response.Write(vXmlStr)
			HttpContext.Current.Response.End()
			'            Else
			'                pXslFilename = "xmldisplay.xsl"
			'            End If
		End If
		
		' XML error handling: shows source text of a page with lines numbered
		Dim vMatch, vLineNum, vRegEx, vText, vMatches, vErrorLine As Object
		If Not vXmlDoc.loadXML(vXmlStr) Then
			vText = HttpContext.Current.Server.HTMLEncode(vXmlStr)
			vLineNum = 1
			vErrorLine = vXmlDoc.parseError.Line
			
			vRegEx = New RegExp
			vRegEx.IgnoreCase = False
			vRegEx.Global = True
			vRegEx.Pattern = ".+"
			vMatches = vRegEx.Execute(vText)
			
			HttpContext.Current.Response.ContentType = "text/html;" ' charset=" & OPENWIKI_ENCODING & ";"
			HttpContext.Current.Response.Write("<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">")
			HttpContext.Current.Response.Write("<html xmlns=""http://www.w3.org/1999/xhtml"">")
			HttpContext.Current.Response.Write("<head><title>Invalid XML document</title></head>")
			HttpContext.Current.Response.Write("<body><b>Invalid XML document</b>:<br /><br />")
			HttpContext.Current.Response.Write(vXmlDoc.parseError.reason & " line: " & "<a href=#ErrorLine>" & vErrorLine & "</a>" & " col: " & vXmlDoc.parseError.linepos)
			HttpContext.Current.Response.Write("<br /><br /><hr />")
			HttpContext.Current.Response.Write("<pre>")
			
			For	Each vMatch In vMatches
				HttpContext.Current.Response.Write(vLineNum & ": ")
				If vLineNum = vErrorLine Then
					HttpContext.Current.Response.Write("<a name=""ErrorLine"" />")
					HttpContext.Current.Response.Write("<font style=""BACKGROUND-COLOR: red"">")
				End If
				HttpContext.Current.Response.Write(vMatch & "<br />")
				If vLineNum = vErrorLine Then
					HttpContext.Current.Response.Write("</font>")
				End If
				vLineNum = vLineNum + 1
			Next vMatch
			
			HttpContext.Current.Response.Write("</pre>")
			HttpContext.Current.Response.Write("</body></html>")
		ElseIf vIsIE And cUseXhtmlHttpHeaders And Not vIsMathPlayer Then 
			ProcessIEWithoutMathPlayer()
		Else
			LoadXSL((pXslFilename))
			vXslProc.input = vXmlDoc
			vXslProc.Transform()
			
			TransformXmlStr = vXslProc.output
			
			If cEmbeddedMode = 0 Then
				If gAction = "edit" Then
					If cUseXhtmlHttpHeaders Then
						' 						IE+MathPlayer workaround: in ContentType must be specified just "content type"
						'						Response.ContentType = "application/xhtml+xml; charset=" & OPENWIKI_ENCODING & ";"
						HttpContext.Current.Response.ContentType = "application/xhtml+xml"
					Else
						HttpContext.Current.Response.ContentType = "text/html; charset=" & OPENWIKI_ENCODING & ";"
					End If
					HttpContext.Current.Response.Expires = 0 ' expires in a minute
				ElseIf gAction = "rss" Then 
					HttpContext.Current.Response.ContentType = "text/xml; charset=" & OPENWIKI_ENCODING & ";"
				Else
					If cUseXhtmlHttpHeaders Then
						' 						IE+MathPlayer workaround: in ContentType must be specified just "content type"
						'						Response.ContentType = "application/xhtml+xml; charset=" & OPENWIKI_ENCODING & ";"
						HttpContext.Current.Response.ContentType = "application/xhtml+xml"
					Else
						HttpContext.Current.Response.ContentType = "text/html; charset=" & OPENWIKI_ENCODING & ";"
					End If
					HttpContext.Current.Response.Expires = -1 ' expires now
					'                    Response.AddHeader "Last-modified", DateToHTTPDate(gLastModified)
					'UPGRADE_NOTE: Global Sub/Function DateToHTTPDate is not accessible
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

