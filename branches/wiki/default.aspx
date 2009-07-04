<%@ Page Language="VB" EnableSessionState=False explicit="true" %>
<%@ Import namespace="ADODB" %>
<%@ Import namespace="MSXML2" %>
<script language="VB" runat="Server">
Dim GetRemoteHost As Object
Dim FormatDateISO8601 As Object
Dim pTimestamp As Object
Dim PrettyWikiLink As Object
Dim pID As Object
Dim QuoteXml As Object
Dim Wikify As Object
Dim WikiLinesToHtml As Object
Dim MultiLineMarkup As Object
Dim StoreRaw As Object
Dim pText As Object
Dim ScriptEngineMinorVersion_Renamed As Object
Dim ScriptEngineMajorVersion_Renamed As Object
</script>

<%--
	'
	' see more examples in owaction.asp
	'
	
	%><%	'#$>$#C:\Sources\plainwiki\trunk\wiki\ow\my\myactions.asp|%>

<%
--%>

<%--
<!-- #INCLUDE FILE="ow/my/mymacros.aspx" -->
--%>
<!--   // -->

<%--
<!-- #INCLUDE FILE="ow.aspx" -->
--%>
<!--   // -->

<%--
<%'#$<$#C:\Sources\plainwiki\trunk\wiki\ow.asp|%>
<%
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
'
' ---------------------------------------------------------------------------
'      $Source: /usr/local/cvsroot/openwiki/dist/owbase/ow.asp,v $
'    $Revision: 1.2 $
'      $Author: pit $
' ---------------------------------------------------------------------------
'
%>
--%>
<%--
<!-- #INCLUDE FILE="ow/owall.aspx" -->
<!--   // -->
--%>
<%--
<%'#$<$#C:\Sources\plainwiki\trunk\wiki\ow\owall.asp|%>
<%' scripts that need to be used all together %>
--%>
<!-- #INCLUDE FILE="ow/owpreamble.aspx" -->
<!--   // -->
<!-- #INCLUDE FILE="ow/owconfig_default.aspx" -->
<!--   // -->
<!-- #INCLUDE FILE="ow/owprocessor.aspx" -->
<!--   // -->
<!-- #INCLUDE FILE="ow/owpatterns.aspx" -->
<!--   // -->
<!-- #INCLUDE FILE="ow/owwikify.aspx" -->
<!--   // -->
<!-- #INCLUDE FILE="ow/owmacros.aspx" -->
<!--   // -->
<!-- #INCLUDE FILE="ow/owactions.aspx" -->
<!--   // -->
<!-- #INCLUDE FILE="ow/owtoc.aspx" -->
<!--   // -->

<%--
<!-- #INCLUDE FILE="ow/owattach.aspx" -->
--%>

<!--   // -->
<!-- #INCLUDE FILE="ow/owpage.aspx" -->
<!--   // -->
<!-- #INCLUDE FILE="ow/owindex.aspx" -->
<!--   // -->
<!-- #INCLUDE FILE="ow/owdb.aspx" -->
<!--   // -->
<!-- #INCLUDE FILE="ow/owrss.aspx" -->
<!--   // -->
<!-- #INCLUDE FILE="ow/owauth.aspx" -->
<!--   // -->
<!-- #INCLUDE FILE="ow/owtransformer.aspx" -->
<!--   // -->
<!-- #INCLUDE FILE="ow/owhttpdate.aspx" -->
<!--   // -->
<!-- #INCLUDE FILE="ow/owtagstack.aspx" -->
<!--   // -->
<!-- #INCLUDE FILE="ow/owrecaptcha.aspx" -->
<!--   // -->
<!-- #INCLUDE FILE="ow/owdebug.aspx" -->
<!--   // -->

<%--
<%' scripts that each can be used on their own %>
--%>

<!-- #INCLUDE FILE="ow/owregexp.aspx" -->
<!--   // -->
<!-- #INCLUDE FILE="ow/owvector.aspx" -->
<!--   // -->
<!-- #INCLUDE FILE="ow/owdiff.aspx" -->
<!--   // -->
<!-- #INCLUDE FILE="ow/owado.aspx" -->
<!--   // -->


<%--
<%' scripts containing your custom build code %>
--%>

<!-- #INCLUDE FILE="ow/my/mywikify.aspx" -->
<!--   // -->

<%--
<!-- #INCLUDE FILE="ow/my/myactions.aspx" -->
--%>

<!--   // -->

<%--
<%'#$<$#C:\Sources\plainwiki\trunk\wiki\ow\my\myactions.asp|%>
<%
'
' add your custom made actions, must start with the string Action, e.g.
'
'  Sub ActionDoSomething()
'      gActionReturn = True
--%>