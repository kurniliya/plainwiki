<%

Sub WriteDebug(pText, pParam, pDebugLevel)
	If OPENWIKI_DEBUGLEVEL >= pDebugLevel  Then
	    Response.Write(Now() & ": " & pText & " " & Server.HTMLEncode(pParam) & "<br>")
	End If
End Sub

%>