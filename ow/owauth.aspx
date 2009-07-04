<script language="VB" runat="Server">
' http://support.microsoft.com/support/kb/articles/Q245/5/74.ASP
'_____________________________________________________________________________________________________________Function GetRemoteHost()
Dim vHost As Object

</script>

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
'      $Source: /usr/local/cvsroot/openwiki/dist/owbase/ow/owauth.asp,v $
'    $Revision: 1.2 $
'      $Author: pit $
' ---------------------------------------------------------------------------
'


Function GetRemoteUser() As Object
	Dim vPos As Object
	GetRemoteUser = Request.ServerVariables("REMOTE_USER")
	If cStripNTDomain Then
		vPos = InStr(GetRemoteUser, "\")
		If vPos > 0 Then
			GetRemoteUser = Mid(GetRemoteUser, vPos + 1)
		End If
	End If
End Function


Function GetRemoteAlias() As Object
	GetRemoteAlias = Request.Cookies(gCookieHash & "?up")("un")
End Function

'End Sub


' you need administrator rights to do this
Sub EnableRemoteHostLookup(ByRef pCurrentWebOnly As Object)
	Dim oIIS As Object
	Dim vWebsite As Object
	Dim vEnableRevDNS As Object
	Dim vDisableRevDNS As Object
	
	vEnableRevDNS = 1
	vDisableRevDNS = 0
	
	Dim vPos As Object
	If pCurrentWebOnly Then
		vWebsite = Request.ServerVariables("INSTANCE_META_PATH")
		vPos = InStrRev(vWebsite, "/")
		If vPos > 0 Then
			vWebsite = "/" & Mid(vWebsite, vPos + 1) & "/ROOT"
		Else
			Exit Sub
		End If
	End If
	
	oIIS = GetObject("IIS://localhost/w3svc" & vWebsite)
	oIIS.Put("EnableReverseDNS", vEnableRevDNS)
	oIIS.SetInfo()
	'UPGRADE_NOTE: Object oIIS may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
	oIIS = Nothing
End Sub

<%If cUseLookup Then
	vHost = Request.ServerVariables("REMOTE_HOST")
End If
If Not cUseLookup Or vHost = "" Then
	vHost = Request.ServerVariables("REMOTE_ADDR")
End If
GetRemoteHost = vHost

%>
