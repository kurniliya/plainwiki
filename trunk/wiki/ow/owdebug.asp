<%

Sub WriteDebug(pText, pParam, pDebugLevel)
	If OPENWIKI_DEBUGLEVEL >= pDebugLevel  Then
		Response.Write("<ow:debug>")
		Response.Write("<ow:date>" & Now() & "</ow:date>")
		Response.Write("<ow:text>" & pText & "</ow:text>")
		Response.Write("<ow:value>" & Server.HTMLEncode(pParam) & "</ow:value>")		
		Response.Write("</ow:debug>")	    
	End If
End Sub

%>