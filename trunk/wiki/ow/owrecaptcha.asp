<%
 ' server side support for reCAPTCHA
 ' initially taken from http://wiki.recaptcha.net/index.php/Overview#Classic_ASP
 ' returns "" if correct, otherwise it returns the error response
Function RecaptchaConfirm(rechallenge, reresponse)
	
	Dim VarString
	VarString = _
	     "privatekey=" & OPENWIKI_RECAPTCHAPRIVATEKEY & _
	     "&remoteip=" & Request.ServerVariables("REMOTE_ADDR") & _
	     "&challenge=" & rechallenge & _
	     "&response=" & reresponse
	
	Dim objXmlHttp
	Set objXmlHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
	objXmlHttp.open "POST", "http://api-verify.recaptcha.net/verify", False
	objXmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXmlHttp.send VarString
	
	Dim ResponseString
	ResponseString = split(objXmlHttp.responseText, vbLF)
	Set objXmlHttp = Nothing
	
	If ResponseString(0) = "true" Then
	'They answered correctly
		RecaptchaConfirm = ""
	Else
	'They answered incorrectly
		RecaptchaConfirm = ResponseString(1)
	End If

End Function
%>