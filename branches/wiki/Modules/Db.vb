Imports System.Text.RegularExpressions

Namespace Openwiki
    Module db
        Public gEquation As String

        Function FormatDateISO8601(ByVal pTimestamp As Date) As String
            Dim vTemp As Integer
            FormatDateISO8601 = Year(pTimestamp) & "-"
            vTemp = Month(pTimestamp)
            If vTemp < 10 Then
                FormatDateISO8601 = FormatDateISO8601 & "0"
            End If
            FormatDateISO8601 = FormatDateISO8601 & vTemp & "-"
            vTemp = Day(pTimestamp)
            If vTemp < 10 Then
                FormatDateISO8601 = FormatDateISO8601 & "0"
            End If
            FormatDateISO8601 = FormatDateISO8601 & vTemp & "T"
            vTemp = Hour(pTimestamp)
            If vTemp < 10 Then
                FormatDateISO8601 = FormatDateISO8601 & "0"
            End If
            FormatDateISO8601 = FormatDateISO8601 & vTemp & ":"
            vTemp = Minute(pTimestamp)
            If vTemp < 10 Then
                FormatDateISO8601 = FormatDateISO8601 & "0"
            End If
            FormatDateISO8601 = FormatDateISO8601 & vTemp & ":"
            vTemp = Second(pTimestamp)
            If vTemp < 10 Then
                FormatDateISO8601 = FormatDateISO8601 & "0"
            End If
            FormatDateISO8601 = FormatDateISO8601 & vTemp
            FormatDateISO8601 = FormatDateISO8601 & OPENWIKI_TIMEZONE
        End Function

        Sub ToDateTime(ByVal pYear As Integer _
            , ByVal pMonth As Integer _
            , ByVal pDay As Integer _
            , ByVal pHour As Integer _
            , ByVal pMinutes As Integer _
            , ByVal pSeconds As Integer _
            , ByVal pPlusMinTZ As String _
            , ByVal pHourTZ As Integer _
            , ByVal pMinutesTZ As Integer)

            sReturn = CStr(DateSerial(pYear, pMonth, pDay))
            If pPlusMinTZ = "-" Then
                sReturn = CStr(DateAdd("h", pHour + pHourTZ, sReturn))
                sReturn = CStr(DateAdd("n", pMinutes + pMinutesTZ, sReturn))
            ElseIf pPlusMinTZ = "+" Then
                sReturn = CStr(DateAdd("h", pHour - pHourTZ, sReturn))
                sReturn = CStr(DateAdd("n", pMinutes - pMinutesTZ, sReturn))
            End If
            If pPlusMinTZ = "-" Or pPlusMinTZ = "+" Then
                ' it's in GMT, now move it to OPENWIKI_TIMEZONE
                If Left(OPENWIKI_TIMEZONE, 1) = "-" Then
                    sReturn = CStr(DateAdd("h", -1 * CInt(Mid(OPENWIKI_TIMEZONE, 2, 2)), sReturn))
                    sReturn = CStr(DateAdd("n", -1 * CInt(Mid(OPENWIKI_TIMEZONE, 5, 2)), sReturn))
                Else
                    sReturn = CStr(DateAdd("h", CInt(Mid(OPENWIKI_TIMEZONE, 2, 2)), sReturn))
                    sReturn = CStr(DateAdd("n", CInt(Mid(OPENWIKI_TIMEZONE, 5, 2)), sReturn))
                End If
            End If
        End Sub

        Function EscapePattern(ByVal pPattern As String) As String
            pPattern = Replace(pPattern, "''''''", "")
            pPattern = Replace(pPattern, "\", "\\")
            pPattern = Replace(pPattern, "(", "\(")
            pPattern = Replace(pPattern, ")", "\)")
            pPattern = Replace(pPattern, "[", "\[")
            pPattern = Replace(pPattern, "+", "\+")
            pPattern = Replace(pPattern, "*", "\*")
            pPattern = Replace(pPattern, "?", "\?")
            EscapePattern = pPattern
        End Function

        Sub CutEquation(ByVal pText As String)
            sReturn = "<ow:math><![CDATA[" & Replace(pText, "]]>", "]]&gt;") & "]]></ow:math>"
            gEquation = "<ow:math><![CDATA[" & Replace(pText, "]]>", "]]&gt;") & "]]></ow:math>"
        End Sub

    End Module
End Namespace