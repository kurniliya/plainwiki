﻿Imports System.Text.RegularExpressions

Namespace Openwiki

    Module Regexp
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
        '      $Source: /usr/local/cvsroot/openwiki/dist/owbase/ow/owregexp.asp,v $
        '    $Revision: 1.3 $
        '      $Author: pit $
        ' ---------------------------------------------------------------------------
        '
        ' These functions simulate the m and s operations as available in the
        ' programming language perl. You can usually literally copy perl regular
        ' expressions and expect them to work with these functions.
        '
        ' In perl you can do something like:
        '
        '     s/A(.*?)B(.*?)C/&MyMethod($1, $2)/ge
        '
        ' When the match is made this will call the sub "MyMethod", pass it
        ' the two matched variables, and finally the match is substituted
        ' by whatever the sub returns.
        '
        ' The function s below can behave in a similar manner. The perl expression
        ' shown above would be written as:
        '
        '     myText = s(myText, "A(.*?)B(.*?)C", "&MyMethod($1, $2)", True, True)
        '
        ' In ASP the sub must return the value to be substituted via the global
        ' parameter sReturn. E.g.
        '
        '     Sub MyMethod(pParam1, pParam2)
        '         If pParam1 = pParam2 Then
        '             sReturn = "Same"
        '         Else
        '             sReturn = "Different"
        '         End If
        '     End Sub
        '

        ' Global register which subs (called by s) should use to return their value.
        Public sReturn As String

        ' Reuse regular expression object
        'On Error Resume Next

        'Dim gRegEx As Regex

        Sub TestRegex()
            'gRegEx = New Regex
            'If Not IsReference(gRegEx) Then
            '    HttpContext.Current.Response.Write("<h2>Error:</h2><p>Probable cause: Registry permission problem.</p>")
            '    HttpContext.Current.Response.Write("This is a known problem with Microsoft.<br />" _
            '                 & "You can find more information about this problem in the following  " _
            '                 & "<a href=""http://support.microsoft.com/support/kb/articles/Q274/0/38.ASP"">Microsoft knowledge base article</a>.")
            '    HttpContext.Current.Response.End()
            'End If
            'On Error GoTo 0
        End Sub

        Function m _
        (ByVal pText As String, ByVal pPattern As String, _
         ByVal pIgnoreCase As Boolean, ByVal pGlobal As Boolean) As Boolean
            Dim vRegexOptions As RegexOptions

            If IsNothing(pText) Then
                Return False
            End If

            If pIgnoreCase Then
                vRegexOptions = RegexOptions.IgnoreCase
            End If

            Return Regex.IsMatch(pText, pPattern, vRegexOptions)
        End Function


        Function s( _
            ByVal pText As String _
            , ByVal pSearchPattern As String _
            , ByVal pReplacePattern As String _
            , ByVal pIgnoreCase As Boolean _
            , ByVal pGlobal As Boolean _
        ) As String
            Dim vRegexOptions As RegexOptions

            If IsNothing(pText) Then
                Return Nothing
            End If

            If pIgnoreCase Then
                vRegexOptions = RegexOptions.IgnoreCase
            End If

            If (Left(pReplacePattern, 1) <> "&") Then
                '    '                Call WriteDebug("Replacement is not procedural", "", 100)
                s = Regex.Replace(pText, pSearchPattern, pReplacePattern, vRegexOptions)
            Else
                Return Nothing
                '    Dim vText As String
                '    Dim vPrevLastIndex As Integer
                '    Dim vPrevNewPos As Integer
                '    Dim vMatch As Match
                '    Dim vMatches As MatchCollection
                '    Dim vSubMatch
                '    Dim i As Integer
                '    Dim vCmd As String
                '    Dim vReplacement As String

                '    vText = pText
                '    vPrevLastIndex = 0
                '    vPrevNewPos = 0

                '    pReplacePattern = Mid(pReplacePattern, 2)
                '    '                Call WriteDebug("pText before Execute", pText, 100)
                '    vMatches = Regex.Matches(pText, pSearchPattern)
                '    For Each vMatch In vMatches
                '        '                    Call WriteDebug("vMatch", vMatch.Value, 100)
                '        vCmd = pReplacePattern
                '        '                   Call WriteDebug("REGEXP CMD before For Each", vCmd, 100)

                '        i = 0
                '        For Each vSubMatch In vMatch.SubMatches
                '            '                      Call WriteDebug("SubMatch", vSubMatch, 100)
                '            vCmd = Replace(vCmd, "$" & (i + 1), """" & Replace(vSubMatch, """", """""") & """")
                '            i = i + 1
                '        Next

                '        '                 Call WriteDebug("REGEXP CMD after For Each", vCmd, 100)

                '        sReturn = ""
                '        vCmd = Replace(vCmd, vbCrLf, """ & vbCRLF & """)
                '        '                   Call WriteDebug("REGEXP CMD after replace", vCmd, 100)
                '        '                    Execute("Call " & vCmd)
                '        vReplacement = sReturn

                '        ' replace vMatch.Value in vText by vReplacement
                '        vPrevNewPos = vPrevNewPos + (vMatch.FirstIndex - vPrevLastIndex)
                '        vText = Mid(vText, 1, vPrevNewPos) & vReplacement & Mid(vText, vPrevNewPos + vMatch.Length + 1)
                '        vPrevNewPos = vPrevNewPos + Len(vReplacement) + 1
                '        vPrevLastIndex = vMatch.FirstIndex + vMatch.Length + 1
                '    Next
                '    Return vText
            End If
        End Function
    End Module
End Namespace