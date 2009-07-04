
' server side support for reCAPTCHA
' initially taken from http://wiki.recaptcha.net/index.php/Overview#Classic_ASP
' returns "" if correct, otherwise it returns the error response
Function RecaptchaConfirm(ByRef rechallenge As Object, ByRef reresponse As Object) As Object
	
	Dim VarString As Object
	VarString = "privatekey=" & OPENWIKI_RECAPTCHAPRIVATEKEY & "&remoteip=" & Request.ServerVariables("REMOTE_ADDR") & "&challenge=" & rechallenge & "&response=" & reresponse
	
	Dim objXmlHttp As MSXML2.ServerXMLHTTP
	objXmlHttp = New MSXML2.ServerXMLHTTP
	objXmlHttp.open("POST", "http://api-verify.recaptcha.net/verify", False)
	objXmlHttp.setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
	objXmlHttp.send(VarString)
	
	Dim ResponseString As Object
	ResponseString = Split(objXmlHttp.responseText, vbLf)
	'UPGRADE_NOTE: Object objXmlHttp may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
	objXmlHttp = Nothing
	
	If ResponseString(0) = "true" Then
		'They answered correctly
		RecaptchaConfirm = ""
	Else
		'They answered incorrectly
		RecaptchaConfirm = ResponseString(1)
	End If
	
End Function

