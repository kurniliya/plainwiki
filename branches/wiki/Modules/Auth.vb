Namespace Openwiki
    Module Auth
        Function GetRemoteUser() As String
            Dim vPos As Integer
            GetRemoteUser = HttpContext.Current.Request.ServerVariables("REMOTE_USER")
            If cStripNTDomain = 1 Then
                vPos = InStr(GetRemoteUser, "\")
                If vPos > 0 Then
                    GetRemoteUser = Mid(GetRemoteUser, vPos + 1)
                End If
            End If
        End Function

        Function GetRemoteAlias() As String
            GetRemoteAlias = HttpContext.Current.Request.Cookies(gCookieHash & "?up")("un")
        End Function

        ' http://support.microsoft.com/support/kb/articles/Q245/5/74.ASP
        '______________________________________________________________________________________________________________
        Function GetRemoteHost() As String
            Dim vHost As String = ""

            If cUseLookup = 1 Then
                vHost = HttpContext.Current.Request.ServerVariables("REMOTE_HOST")
            End If
            If Not cUseLookup = 1 Or vHost = "" Then
                vHost = HttpContext.Current.Request.ServerVariables("REMOTE_ADDR")
            End If
            GetRemoteHost = vHost
        End Function

        ' you need administrator rights to do this
        Sub EnableRemoteHostLookup(ByVal pCurrentWebOnly As Boolean)
            'Dim oIIS
            'Dim vWebsite
            'Dim vEnableRevDNS
            'Dim vDisableRevDNS

            'vEnableRevDNS = 1
            'vDisableRevDNS = 0

            'If pCurrentWebOnly Then
            '    Dim vPos
            '    vWebsite = HttpContext.Current.Request.ServerVariables("INSTANCE_META_PATH")
            '    vPos = InStrRev(vWebsite, "/")
            '    If vPos > 0 Then
            '        vWebsite = "/" & Mid(vWebsite, vPos + 1) & "/ROOT"
            '    Else
            '        Exit Sub
            '    End If
            'End If

            'oIIS = GetObject("IIS://localhost/w3svc" & vWebsite)
            'oIIS.Put("EnableReverseDNS", vEnableRevDNS)
            'oIIS.SetInfo()
            'oIIS = Nothing
        End Sub

    End Module
End Namespace