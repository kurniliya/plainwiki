Namespace Openwiki
    Module Recaptcha
        ' server side support for reCAPTCHA
        ' initially taken from http://wiki.recaptcha.net/index.php/Overview#Classic_ASP
        ' returns "" if correct, otherwise it returns the error response

        Private OPENWIKI_RECAPTCHAPRIVATEKEY As String

        Function RecaptchaConfirm(ByVal rechallenge As String, ByVal reresponse As String) As String

            Dim VarString As String
            VarString = _
                 "privatekey=" & OPENWIKI_RECAPTCHAPRIVATEKEY & _
                 "&remoteip=" & HttpContext.Current.Request.ServerVariables("REMOTE_ADDR") & _
                 "&challenge=" & rechallenge & _
                 "&response=" & reresponse

            Dim objXmlHttp As MSXML2.ServerXMLHTTP
            objXmlHttp = New MSXML2.ServerXMLHTTP
            objXmlHttp.open("POST", "http://api-verify.recaptcha.net/verify", False)
            objXmlHttp.setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
            objXmlHttp.send(VarString)

            Dim ResponseString() As String
            ResponseString = Split(objXmlHttp.responseText, vbLf)
            objXmlHttp = Nothing

            If ResponseString(0) = "true" Then
                'They answered correctly
                RecaptchaConfirm = ""
            Else
                'They answered incorrectly
                RecaptchaConfirm = ResponseString(1)
            End If

        End Function
    End Module
End Namespace