
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
'      $Source: /usr/local/cvsroot/openwiki/dist/owbase/ow/owdiff.asp,v $
'    $Revision: 1.2 $
'      $Author: pit $
' ---------------------------------------------------------------------------
'
' Computes the difference between two page revisions.
'

Class Matcher
Private vLineBreak As Object
	Private vLineOriented As Object
	Private vA As Object
	Private vB As Object
	Private vBhash As Scripting.Dictionary
	Private vOut As Object
	Private vOutlen As Object
	Private vDebug As Object
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1061.asp'
	Private Sub Class_Initialize_Renamed()
		vLineBreak = vbCrLf
		vLineOriented = True
		vOut = ""
		vOutlen = 0
		vDebug = False
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
	
	Public WriteOnly Property Preformatted() As Object
		Set(ByVal Value As Object)
			If Value Then
				vLineBreak = vbCrLf
			Else
				vLineBreak = "<br/>"
			End If
		End Set
	End Property
	
	Public WriteOnly Property LineOriented() As Object
		Set(ByVal Value As Object)
			vLineOriented = Value
		End Set
	End Property
	
	Public WriteOnly Property Debug() As Object
		Set(ByVal Value As Object)
			vDebug = Value
		End Set
	End Property
	
	Public WriteOnly Property Outlen() As Object
		Set(ByVal Value As Object)
			vOutlen = Value
		End Set
	End Property
	
	Private Function Tokenize(ByRef pText As Object) As Object
		Dim vMatch2, vRegEx2, vMatches, vRegEx, vMatch, vMatches2, vValue As Object
		Tokenize = New Vector
		vRegEx = New RegExp
		vRegEx.IgnoreCase = False
		vRegEx.Global = True
		vRegEx.Pattern = ".+"
		pText = Replace(pText, Chr(9), Space(8))
		If Not vLineOriented Then
			vRegEx2 = New RegExp
			vRegEx2.IgnoreCase = False
			vRegEx2.Global = True
			vRegEx2.Pattern = "\s*\S+"
		End If
		vMatches = vRegEx.Execute(pText)
		For	Each vMatch In vMatches
			vValue = Replace(vMatch.Value, vbCr, "")
			If vLineOriented Then
				Tokenize.Push(vValue)
			Else
				If Trim(vValue) = "" Then
					Tokenize.Push(vValue)
				Else
					vMatches2 = vRegEx2.Execute(vValue)
					For	Each vMatch2 In vMatches2
						Tokenize.Push(vMatch2.Value)
					Next vMatch2
				End If
				Tokenize.Push(vbCrLf)
			End If
		Next vMatch
		If vValue = "" Then
			Tokenize.Push("")
		ElseIf Not vLineOriented Then 
			Tokenize.Pop()
		End If
		'UPGRADE_NOTE: Object vRegEx may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
		vRegEx = Nothing
		If Not vLineOriented Then
			'UPGRADE_NOTE: Object vRegEx2 may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
			vRegEx2 = Nothing
		End If
	End Function
	
	Private Sub HashB()
		Dim vElem, i, vList As Object
		vBhash = New Scripting.Dictionary
		For i = 0 To vB.Count - 1
			vElem = vB.ElementAt(i)
			If Trim(vElem) <> "" And vElem <> vbCrLf Then
				If vBhash.Exists(vElem) Then
					vList = vBhash.Item(vElem)
					vList.Push(i)
				Else
					vList = New Vector
					vList.Push(i)
					vBhash.Add(vElem, vList)
				End If
			End If
		Next 
	End Sub
	
	Private bestStartA As Object
	Private bestStartB As Object
	Private bestSize As Object
	' find longest matching block in vA[pALow,pAHigh] and vB[pBLow,pBHigh]
	Private Sub FindLongestMatch(ByRef pALow As Object, ByRef pAHigh As Object, ByRef pBLow As Object, ByRef pBHigh As Object)
		Dim x, j, i, k, vList As Object
		bestStartA = pALow
		bestStartB = pBLow
		bestSize = 0
		
		Dim vNewLen, vLen, vElem As Object
		vLen = New Vector
		vLen.Dimension = vB.Count
		For i = pALow To pAHigh
			vNewLen = New Vector
			vNewLen.Dimension = vB.Count
			vElem = vA.ElementAt(i)
			If vBhash.Exists(vElem) Then
				vList = vBhash.Item(vElem)
				For x = 0 To vList.Count - 1
					j = vList.ElementAt(x)
					If j > pBHigh Then
						Exit For
					End If
					If j >= pBLow Then
						If j > 0 Then
							k = vLen.ElementAt(j - 1) + 1
						Else
							k = 1
						End If
						vNewLen.SetElementAt(j, k)
						If k > bestSize Then
							bestStartA = i - k + 1
							bestStartB = j - k + 1
							bestSize = k
						End If
					End If
				Next 
			End If
			vLen = vNewLen
		Next 
		
		' add junk on both sides
		Do While bestStartA > pALow And bestStartB > pBLow
			If (Trim(vA.ElementAt(bestStartA - 1)) = "" Or vA.ElementAt(bestStartA - 1) = vbCrLf) And (Trim(vB.ElementAt(bestStartB - 1)) = "" Or vB.ElementAt(bestStartB - 1) = vbCrLf) Then
				bestStartA = bestStartA - 1
				bestStartB = bestStartB - 1
				bestSize = bestSize + 1
			Else
				Exit Do
			End If
		Loop 
		Do While bestStartA + bestSize <= pAHigh And bestStartB + bestSize <= pBHigh
			If (Trim(vA.ElementAt(bestStartA + bestSize)) = "" Or vA.ElementAt(bestStartA + bestSize) = vbCrLf) And (Trim(vB.ElementAt(bestStartB + bestSize)) = "" Or vB.ElementAt(bestStartB + bestSize) = vbCrLf) Then
				bestSize = bestSize + 1
			Else
				Exit Do
			End If
		Loop 
	End Sub
	
	Private Sub SplitLine(ByRef pLine As Object)
		Dim i As Object
		Do 
			i = InStrRev(pLine, " ", 80)
			If i > 0 Then
				vOut = vOut & Left(pLine, i) & vLineBreak
				pLine = LTrim(Mid(pLine, i))
			Else
				vOut = vOut & pLine
			End If
		Loop While i > 0
	End Sub
	
	Private Sub Output(ByRef pTag As Object, ByRef pVector As Object, ByRef pFrom As Object, ByRef pTo As Object)
		Dim i, vElem As Object
		
		If pTag = "delete" Then
			vOut = vOut & "<strike class='diff'>"
		ElseIf pTag = "insert" Then 
			vOut = vOut & "<u class='diff'>"
		End If
		
		For i = pFrom To pTo
			vElem = pVector.ElementAt(i)
			If vElem = vbCrLf Then
				vElem = vLineBreak
				vOutlen = 0
			ElseIf vElem = "" Then 
				vElem = "  "
			End If
			
			vOutlen = vOutlen + Len(CStr(vElem))
			If vOutlen > 80 Then
				If Len(CStr(vElem)) > 80 Then
					SplitLine((vElem))
					vElem = ""
				Else
					vOut = vOut & vLineBreak
					vElem = LTrim(vElem)
					vOutlen = Len(CStr(vElem))
				End If
			End If
			
			vOut = vOut & vElem
			
			If vLineOriented Then
				vOut = vOut & vLineBreak
				vOutlen = 0
			End If
		Next 
		
		If pTag = "delete" Then
			vOut = vOut & "</strike>"
		ElseIf pTag = "insert" Then 
			vOut = vOut & "</u>"
		End If
	End Sub
	
	Private Sub InnerReplace(ByRef pAFrom As Object, ByRef pATo As Object, ByRef pBFrom As Object, ByRef pBTo As Object)
		Dim vText2, i, vText1, vMatcher As Object
		vText1 = ""
		vText2 = ""
		For i = pAFrom To pATo
			vText1 = vText1 & vA.ElementAt(i)
			If i < pATo Then
				vText1 = vText1 & vbCrLf
			End If
		Next 
		For i = pBFrom To pBTo
			vText2 = vText2 & vB.ElementAt(i)
			If i < pBTo Then
				vText2 = vText2 & vbCrLf
			End If
		Next 
		vMatcher = New Matcher
		vMatcher.Outlen = vOutlen
		vMatcher.LineOriented = False
		vMatcher.Debug = vDebug
		vOut = vOut & vMatcher.Compare(vText1, vText2) & vLineBreak
	End Sub
	
	Private Sub Out(ByRef vAFound As Object, ByRef vBFound As Object, ByRef vSize As Object)
		If matchedI < vAFound And matchedJ < vBFound Then
			If vLineOriented Then
				Call InnerReplace(matchedI, vAFound - 1, matchedJ, vBFound - 1)
			Else
				Call Output("delete", vA, matchedI, vAFound - 1)
				' TODO: maybe, add "<br/>" when the intraline deleted was part of the last line
				Call Output("insert", vB, matchedJ, vBFound - 1)
			End If
		ElseIf matchedI < vAFound Then 
			Call Output("delete", vA, matchedI, vAFound - 1)
		ElseIf matchedJ < vBFound Then 
			Call Output("insert", vB, matchedJ, vBFound - 1)
		End If
		If vSize > 0 Then
			Call Output("equal", vA, vAFound, vAFound + vSize - 1)
		End If
	End Sub
	
	Dim matchedI, matchedJ As Object
	'  match between [pALow,pAHigh] and [pBLow,pBHigh]
	Private Sub GetMatchingBlocks(ByRef pDepth As Object, ByRef pALow As Object, ByRef pAHigh As Object, ByRef pBLow As Object, ByRef pBHigh As Object)
		If pDepth = 1 Then
			matchedI = 0
			matchedJ = 0
		End If
		
		Call FindLongestMatch(pALow, pAHigh, pBLow, pBHigh)
		
		Dim vBFound, vAFound, vSize As Object
		If bestSize > 0 Then
			vAFound = bestStartA
			vBFound = bestStartB
			vSize = bestSize
			
			If pALow < vAFound And pBLow < vBFound Then
				Call GetMatchingBlocks(pDepth + 1, pALow, vAFound - 1, pBLow, vBFound - 1)
			End If
			
			Call Out(vAFound, vBFound, vSize)
			
			matchedI = vAFound + vSize
			matchedJ = vBFound + vSize
			
			If matchedI <= pAHigh And matchedJ <= pBHigh Then
				Call GetMatchingBlocks(pDepth + 1, matchedI, pAHigh, matchedJ, pBHigh)
			End If
		End If
		
		If pDepth = 1 Then
			Call Out(vA.Count, vB.Count, 0)
		End If
	End Sub
	
	Public Function Compare(ByRef pText1 As Object, ByRef pText2 As Object) As Object
		vOut = ""
		vA = Tokenize(pText1)
		vB = Tokenize(pText2)
		HashB()
		Call GetMatchingBlocks(1, 0, vA.Count - 1, 0, vB.Count - 1)
		Compare = vOut
	End Function
End Class

