
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
'      $Source: /usr/local/cvsroot/openwiki/dist/owbase/ow/owindex.asp,v $
'    $Revision: 1.3 $
'      $Author: pit $
' ---------------------------------------------------------------------------
'

Class IndexSchemes
'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1061.asp'
	Private Sub Class_Initialize_Renamed()
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1061.asp'
	Private Sub Class_Terminate_Renamed()
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	Public Function GetRecentChanges(ByRef pDays As Object, ByRef pMaxNrOfChanges As Object, ByRef pFilter As Object, ByRef pShortVersion As Object) As Object
		Dim gNamespace As Object
		Dim vChange, vResult, i, vList, vCount, j, vElem, vTimestamp As Object
		If pMaxNrOfChanges > 0 Then
			'UPGRADE_NOTE: Date operands have a different behavior in arithmetical operations
			vTimestamp = System.Date.FromOADate(Now.ToOADate - CDate(pDays).ToOADate)
			vList = gNamespace.TitleSearch(".*", pDays, pFilter, 1, 1)
			vCount = vList.Count - 1
			For i = 0 To vCount
				vElem = vList.ElementAt(i)
				vChange = vElem.GetLastChange()
				If vChange.Timestamp > vTimestamp Then
					vResult = vResult & vElem.ToXML(False)
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
		If pShortVersion Then
			GetRecentChanges = GetRecentChanges & " short='true'"
		Else
			GetRecentChanges = GetRecentChanges & " short='false'"
		End If
		GetRecentChanges = GetRecentChanges & ">" & vResult & "</ow:recentchanges>"
	End Function
	
	Public Function GetRecentNewPages(ByRef pDays As Object, ByRef pMaxNrOfChanges As Object, ByRef pFilter As Object, ByRef pShortVersion As Object) As Object
		Dim gNamespace As Object
		Dim vChange, vResult, i, vList, vCount, j, vElem, vTimestamp As Object
		If pMaxNrOfChanges > 0 Then
			'UPGRADE_NOTE: Date operands have a different behavior in arithmetical operations
			vTimestamp = System.Date.FromOADate(Now.ToOADate - CDate(pDays).ToOADate)
			vList = gNamespace.TitleSearch(".*", pDays, pFilter, 1, 0)
			vCount = vList.Count - 1
			For i = 0 To vCount
				vElem = vList.ElementAt(i)
				vChange = vElem.GetLastChange()
				If vChange.Timestamp > vTimestamp And vChange.Status = "new" Then
					vResult = vResult & vElem.ToXML(False)
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
		If pShortVersion Then
			GetRecentNewPages = GetRecentNewPages & " short='true'"
		Else
			GetRecentNewPages = GetRecentNewPages & " short='false'"
		End If
		GetRecentNewPages = GetRecentNewPages & ">" & vResult & "</ow:recentchanges>"
	End Function
	
	Public Function GetTitleSearch(ByRef pPattern As Object) As Object
		Dim gNamespace As Object
		Dim vCount, vList, i, vResult As Object
		vList = gNamespace.TitleSearch(pPattern, 0, 0, 0, 0)
		vCount = vList.Count - 1
		For i = 0 To vCount
			vResult = vResult & vList.ElementAt(i).ToXML(False)
		Next 
		'UPGRADE_NOTE: Global Sub/Function CDATAEncode is not accessible
		GetTitleSearch = "<ow:titlesearch value='" & CDATAEncode(pPattern) & "' pagecount='" & gNamespace.GetPageCount() & "'>" & vResult & "</ow:titlesearch>"
	End Function
	
	Public Function GetFullSearch(ByRef pPattern As Object, ByRef pIncludeTitles As Object) As Object
		Dim gNamespace As Object
		Dim vCount, vList, i, vResult As Object
		vList = gNamespace.FullSearch(pPattern, pIncludeTitles)
		vCount = vList.Count - 1
		For i = 0 To vCount
			vResult = vResult & vList.ElementAt(i).ToXML(False)
		Next 
		'UPGRADE_NOTE: Global Sub/Function CDATAEncode is not accessible
		GetFullSearch = "<ow:fullsearch value='" & CDATAEncode(pPattern) & "' pagecount='" & gNamespace.GetPageCount() & "'>" & vResult & "</ow:fullsearch>"
	End Function
	
	Public Function GetEquationSearch(ByRef pPattern As Object, ByRef pIncludeTitles As Object) As Object
		Dim gNamespace As Object
		Dim vCount, vList, i, vResult As Object
		vList = gNamespace.EquationSearch(pPattern, pIncludeTitles, 0)
		vCount = vList.Count - 1
		For i = 0 To vCount
			vResult = vResult & vList.ElementAt(i).ToXML(4)
		Next 
		'UPGRADE_NOTE: Global Sub/Function CDATAEncode is not accessible
		GetEquationSearch = "<ow:equationsearch value='" & CDATAEncode(pPattern) & "' pagecount='" & gNamespace.GetPageCount() & "'>" & vResult & "</ow:equationsearch>"
	End Function
	
	Public Function GetRecentEquations(ByRef pPattern As Object, ByRef pIncludeTitles As Object, ByRef pDays As Object, ByRef pMaxNrOfChanges As Object) As Object
		Dim gNamespace As Object
		Dim vChange, vResult, j, vList, i, vCount, vElem, vTimestamp As Object
		If pMaxNrOfChanges > 0 Then
			'UPGRADE_NOTE: Date operands have a different behavior in arithmetical operations
			vTimestamp = System.Date.FromOADate(Now.ToOADate - CDate(pDays).ToOADate)
			vList = gNamespace.EquationSearch(pPattern, pIncludeTitles, 1)
			vCount = vList.Count - 1
			
			For i = 0 To vCount
				vElem = vList.ElementAt(i)
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
		'UPGRADE_NOTE: Global Sub/Function CDATAEncode is not accessible
		GetRecentEquations = "<ow:equationsearch value='" & CDATAEncode(pPattern) & "' pagecount='" & gNamespace.GetPageCount() & "'>" & vResult & "</ow:equationsearch>"
	End Function
	
	Public Function GetRandomPage(ByRef pNrOfPages As Object) As Object
		Dim gNamespace As Object
		Dim vIndex, i, vList, vCount, vResult As Object
		vList = gNamespace.TitleSearch(".*", 0, 0, 0, 0)
		Randomize()
		vCount = vList.Count - 1
		For i = 1 To pNrOfPages
			vIndex = Int(vCount * Rnd())
			vResult = vResult & vList.ElementAt(vIndex).ToXML(False)
		Next 
		GetRandomPage = "<ow:randompages>" & vResult & "</ow:randompages>"
	End Function
	
	Public Function GetTemplates(ByRef pPattern As Object) As Object
		Dim gNamespace As Object
		Dim vCount, vList, i, vResult As Object
		vList = gNamespace.TitleSearch(pPattern, 0, 0, 0, 0)
		vCount = vList.Count - 1
		For i = 0 To vCount
			vResult = vResult & vList.ElementAt(i).ToXML(False)
		Next 
		GetTemplates = "<ow:templates>" & vResult & "</ow:templates>"
	End Function
	
	Public Function GetTitleIndex() As Object
		Dim cUseSpecialPagesPrefix As Object
		Dim gSpecialPagesPrefix As Object
		Dim gNamespace As Object
		Dim i, vList, vCount, vResult As Object
		If cUseSpecialPagesPrefix Then
			vList = gNamespace.TitleSearch("^(?!" & gSpecialPagesPrefix & ")" & ".*", 0, 0, 0, 0)
		Else
			vList = gNamespace.TitleSearch(".*", 0, 0, 0, 0)
		End If
		vCount = vList.Count - 1
		For i = 0 To vCount
			vResult = vResult & vList.ElementAt(i).ToXML(False)
		Next 
		GetTitleIndex = "<ow:titleindex>" & vResult & "</ow:titleindex>"
	End Function
	
	' This function is pure crap! really really bad!
	' needs a totally different implementation
	' either needs an NT service or something similar that runs daily to
	' generate the meta-data, or keep track of this meta-data when saving
	' a new page.
	' Also generate meta-data about concepts like TwinPages, MetaWiki, etc.
	Public Function GetWordIndex() As Object
		Dim PrettyWikiLink As Object
		Dim gNamespace As Object
		Dim vKeys, vMatches, vValues, vTitle, j, vCount, vList, i, vElem, vWords, vRegEx, vMatch, vResult As Object
		Dim vLast, vLastIndex As Object
		vWords = New Vector
		vValues = New Vector
		vRegEx = New RegExp
		vRegEx.IgnoreCase = False
		vRegEx.Global = True
		vRegEx.Pattern = "[A-Z\xc0-\xde]+[a-z\xdf-\xff]+"
		vList = gNamespace.TitleSearch(".*", 0, 0, 0, 0)
		vCount = vList.Count
		For i = 0 To vCount - 1
			vElem = vList.ElementAt(i)
			vTitle = PrettyWikiLink(vElem.Name)
			vMatches = vRegEx.Execute(vTitle)
			For	Each vMatch In vMatches
				vWords.Push(vMatch.Value)
				'UPGRADE_NOTE: Global Sub/Function CDATAEncode is not accessible
				vValues.Push("<ow:word value='" & CDATAEncode(vMatch.Value) & "'>" & vElem.ToXML(False) & "</ow:word>")
			Next vMatch
		Next 
		
		vCount = vWords.Count - 1
		For i = 0 To vCount
			vLast = "\xff\xff\xff\xff\xff"
			vLastIndex = 0
			For j = 0 To vCount
				If vWords.ElementAt(j) < vLast Then
					vLast = vWords.ElementAt(j)
					vLastIndex = j
				End If
			Next 
			vWords.SetElementAt(vLastIndex, "\xff\xff\xff\xff\xff")
			vResult = vResult & vValues.ElementAt(vLastIndex)
		Next 
		
		'UPGRADE_NOTE: Object vWords may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
		vWords = Nothing
		'UPGRADE_NOTE: Object vValues may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
		vValues = Nothing
		'UPGRADE_NOTE: Object vRegEx may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
		vRegEx = Nothing
		GetWordIndex = "<ow:wordindex>" & vResult & "</ow:wordindex>"
	End Function
End Class

