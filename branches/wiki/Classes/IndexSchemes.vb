Imports System.Text.RegularExpressions

Namespace Openwiki
    Public Class IndexSchemes
        'Private Sub Class_Initialize()
        'End Sub

        'Private Sub Class_Terminate()
        'End Sub

        Public Function GetRecentChanges(ByVal pDays As Integer _
            , ByVal pMaxNrOfChanges As Integer _
            , ByVal pFilter As Integer _
            , ByVal pShortVersion As Integer) _
        As String
            Dim vList As Vector
            Dim i As Integer
            Dim j As Integer
            Dim vCount As Integer
            Dim vResult As String = ""
            Dim vElem As WikiPage
            Dim vChange As Change
            Dim vTimestamp As Date

            If pMaxNrOfChanges > 0 Then
                vTimestamp = Now().AddDays(-1 * pDays)
                vList = gNamespace.TitleSearch(".*", pDays, pFilter, 1, 1)
                vCount = vList.Count - 1
                For i = 0 To vCount
                    vElem = CType(vList.ElementAt(i), WikiPage)
                    vChange = vElem.GetLastChange()
                    If vChange.Timestamp > vTimestamp Then
                        vResult = vResult & vElem.ToXML(0)
                        j = j + 1
                        If j >= pMaxNrOfChanges Then
                            Exit For
                        End If
                    End If
                Next
            End If
            GetRecentChanges = "<ow:recentchanges"
            If pFilter = 0 Or pFilter = 1 Then
                GetRecentChanges = GetRecentChanges & " majoredits='true'"
            Else
                GetRecentChanges = GetRecentChanges & " majoredits='false'"
            End If
            If pFilter = 0 Or pFilter = 2 Then
                GetRecentChanges = GetRecentChanges & " minoredits='true'"
            Else
                GetRecentChanges = GetRecentChanges & " minoredits='false'"
            End If
            If pShortVersion = 1 Then
                GetRecentChanges = GetRecentChanges & " short='true'"
            Else
                GetRecentChanges = GetRecentChanges & " short='false'"
            End If
            GetRecentChanges = GetRecentChanges & ">" & vResult & "</ow:recentchanges>"
        End Function

        Public Function GetRecentNewPages(ByVal pDays As Integer _
            , ByVal pMaxNrOfChanges As Integer _
            , ByVal pFilter As Integer _
            , ByVal pShortVersion As Integer) _
        As String
            Dim vList As Vector
            Dim i As Integer
            Dim j As Integer
            Dim vCount As Integer
            Dim vResult As String = ""
            Dim vElem As WikiPage
            Dim vChange As Change
            Dim vTimestamp As Date

            If pMaxNrOfChanges > 0 Then
                vTimestamp = Now().AddDays(-1 * pDays)
                vList = gNamespace.TitleSearch(".*", pDays, pFilter, 1, 0)
                vCount = vList.Count - 1
                For i = 0 To vCount
                    vElem = CType(vList.ElementAt(i), WikiPage)
                    vChange = vElem.GetLastChange()
                    If vChange.Timestamp > vTimestamp And vChange.Status = "new" Then
                        vResult = vResult & vElem.ToXML(0)
                        j = j + 1
                        If j >= pMaxNrOfChanges Then
                            Exit For
                        End If
                    End If
                Next
            End If
            GetRecentNewPages = "<ow:recentchanges"
            If pFilter = 0 Or pFilter = 1 Then
                GetRecentNewPages = GetRecentNewPages & " majoredits='true'"
            Else
                GetRecentNewPages = GetRecentNewPages & " majoredits='false'"
            End If
            If pFilter = 0 Or pFilter = 2 Then
                GetRecentNewPages = GetRecentNewPages & " minoredits='true'"
            Else
                GetRecentNewPages = GetRecentNewPages & " minoredits='false'"
            End If
            If pShortVersion = 1 Then
                GetRecentNewPages = GetRecentNewPages & " short='true'"
            Else
                GetRecentNewPages = GetRecentNewPages & " short='false'"
            End If
            GetRecentNewPages = GetRecentNewPages & ">" & vResult & "</ow:recentchanges>"
        End Function

        Public Function GetTitleSearch(ByVal pPattern As String) As String
            Dim vList As Vector
            Dim i As Integer
            Dim vCount As Integer
            Dim vResult As String = ""

            vList = gNamespace.TitleSearch(pPattern, 0, 0, 0, 0)
            vCount = vList.Count - 1
            For i = 0 To vCount
                vResult = vResult & CType(vList.ElementAt(i), WikiPage).ToXML(0)
            Next
            GetTitleSearch = "<ow:titlesearch value='" & CDATAEncode(pPattern) & "' pagecount='" & gNamespace.GetPageCount() & "'>" & vResult & "</ow:titlesearch>"
        End Function

        Public Function GetFullSearch(ByVal pPattern As String _
            , ByVal pIncludeTitles As Integer) _
        As String
            Dim vList As Vector
            Dim i As Integer
            Dim vCount As Integer
            Dim vResult As String = ""

            vList = gNamespace.FullSearch(pPattern, pIncludeTitles)
            vCount = vList.Count - 1
            For i = 0 To vCount
                vResult = vResult & CType(vList.ElementAt(i), WikiPage).ToXML(0)
            Next
            GetFullSearch = "<ow:fullsearch value='" & CDATAEncode(pPattern) & "' pagecount='" & gNamespace.GetPageCount() & "'>" & vResult & "</ow:fullsearch>"
        End Function

        Public Function GetEquationSearch(ByVal pPattern As String _
            , ByVal pIncludeTitles As Integer) _
        As String
            Dim vList As Vector
            Dim i As Integer
            Dim vCount As Integer
            Dim vResult As String = ""

            vList = gNamespace.EquationSearch(pPattern, pIncludeTitles, 0)
            vCount = vList.Count - 1
            For i = 0 To vCount
                vResult = vResult & CType(vList.ElementAt(i), WikiPage).ToXML(4)
            Next
            GetEquationSearch = "<ow:equationsearch value='" & CDATAEncode(pPattern) & "' pagecount='" & gNamespace.GetPageCount() & "'>" & vResult & "</ow:equationsearch>"
        End Function

        Public Function GetRecentEquations(ByVal pPattern As String _
            , ByVal pIncludeTitles As Integer _
            , ByVal pDays As Integer _
            , ByVal pMaxNrOfChanges As Integer) _
        As String
            Dim vList As Vector
            Dim i As Integer
            Dim j As Integer
            Dim vCount As Integer
            Dim vResult As String = ""
            Dim vElem As WikiPage
            Dim vChange As Change
            Dim vTimestamp As Date

            If pMaxNrOfChanges > 0 Then
                vTimestamp = Now().AddDays(-1 * pDays)
                vList = gNamespace.EquationSearch(pPattern, pIncludeTitles, 1)
                vCount = vList.Count - 1

                For i = 0 To vCount
                    vElem = CType(vList.ElementAt(i), WikiPage)
                    vChange = vElem.GetLastChange()
                    If vChange.Timestamp > vTimestamp And vChange.Status <> "deleted" Then
                        vResult = vResult & vElem.ToXML(4)
                        j = j + 1
                        If j >= pMaxNrOfChanges Then
                            Exit For
                        End If
                    End If
                Next

            End If
            GetRecentEquations = "<ow:equationsearch value='" & CDATAEncode(pPattern) & "' pagecount='" & gNamespace.GetPageCount() & "'>" & vResult & "</ow:equationsearch>"
        End Function

        Public Function GetRandomPage(ByVal pNrOfPages As Integer) As String
            Dim vList As Vector
            Dim i As Integer
            Dim vCount As Integer
            Dim vIndex As Integer
            Dim vResult As String = ""

            vList = gNamespace.TitleSearch(".*", 0, 0, 0, 0)
            Randomize()
            vCount = vList.Count - 1
            For i = 1 To pNrOfPages
                vIndex = CInt(vCount * Rnd())
                vResult = vResult & CType(vList.ElementAt(vIndex), WikiPage).ToXML(0)
            Next
            GetRandomPage = "<ow:randompages>" & vResult & "</ow:randompages>"
        End Function

        Public Function GetTemplates(ByVal pPattern As String) As String
            Dim vList As Vector
            Dim i As Integer
            Dim vCount As Integer
            Dim vResult As String = ""

            vList = gNamespace.TitleSearch(pPattern, 0, 0, 0, 0)
            vCount = vList.Count - 1
            For i = 0 To vCount
                vResult = vResult & CType(vList.ElementAt(i), WikiPage).ToXML(0)
            Next
            GetTemplates = "<ow:templates>" & vResult & "</ow:templates>"
        End Function

        Public Function GetTitleIndex() As String
            Dim vList As Vector
            Dim i As Integer
            Dim vCount As Integer
            Dim vResult As String = ""

            If cUseSpecialPagesPrefix = 1 Then
                vList = gNamespace.TitleSearch("^(?!" & gSpecialPagesPrefix & ")" & ".*", 0, 0, 0, 0)
            Else
                vList = gNamespace.TitleSearch(".*", 0, 0, 0, 0)
            End If
            vCount = vList.Count - 1
            For i = 0 To vCount
                vResult = vResult & CType(vList.ElementAt(i), WikiPage).ToXML(0)
            Next
            GetTitleIndex = "<ow:titleindex>" & vResult & "</ow:titleindex>"
        End Function

        ' This function is pure crap! really really bad!
        ' needs a totally different implementation
        ' either needs an NT service or something similar that runs daily to
        ' generate the meta-data, or keep track of this meta-data when saving
        ' a new page.
        ' Also generate meta-data about concepts like TwinPages, MetaWiki, etc.
        Public Function GetWordIndex() As String
            Dim vList As Vector
            Dim i As Integer
            Dim j As Integer
            Dim vCount As Integer
            Dim vResult As String = ""
            Dim vWords As Vector
            Dim vValues As Vector
            Dim vElem As WikiPage
            Dim vTitle As String
            Dim vMatches As MatchCollection
            Dim vMatch As Match
            Dim vLast As String
            Dim vLastIndex As Integer

            vWords = New Vector
            vValues = New Vector
            'vRegEx = New RegExp
            'vRegEx.IgnoreCase = False
            'vRegEx.Global = True
            'vRegEx.Pattern = "[A-Z\xc0-\xde]+[a-z\xdf-\xff]+"
            vList = gNamespace.TitleSearch(".*", 0, 0, 0, 0)
            vCount = vList.Count
            For i = 0 To vCount - 1
                vElem = CType(vList.ElementAt(i), WikiPage)
                vTitle = PrettyWikiLink(vElem.Name)
                vMatches = Regex.Matches(vTitle, "[A-Z\xc0-\xde]+[a-z\xdf-\xff]+")
                For Each vMatch In vMatches
                    vWords.Push(vMatch.Value)
                    vValues.Push("<ow:word value='" & CDATAEncode(vMatch.Value) & "'>" & vElem.ToXML(0) & "</ow:word>")
                Next
            Next

            vCount = vWords.Count - 1
            For i = 0 To vCount
                vLast = "\xff\xff\xff\xff\xff"
                vLastIndex = 0
                For j = 0 To vCount
                    If CStr(vWords.ElementAt(j)) < vLast Then
                        vLast = CStr(vWords.ElementAt(j))
                        vLastIndex = j
                    End If
                Next
                vWords.SetElementAt(vLastIndex, "\xff\xff\xff\xff\xff")
                vResult = vResult & CStr(vValues.ElementAt(vLastIndex))
            Next

            vWords = Nothing
            vValues = Nothing
            GetWordIndex = "<ow:wordindex>" & vResult & "</ow:wordindex>"
        End Function

    End Class
End Namespace