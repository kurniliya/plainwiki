Imports System.Text.RegularExpressions

Namespace Openwiki

    Public Class Matcher
        Private vLineBreak As String
        Private vLineOriented As Boolean
        Private vA As Vector
        Private vB As Vector
        Private vBhash As Scripting.Dictionary
        Private vOut As String
        Private vOutlen As Integer
        Private vDebug As Boolean
        Private bestStartA As Integer
        Private bestStartB As Integer
        Private bestSize As Integer
        Private matchedI As Integer
        Private matchedJ As Integer

        Public Sub New()
            vLineBreak = vbCrLf
            vLineOriented = True
            vOut = ""
            vOutlen = 0
            vDebug = False
        End Sub

        Public WriteOnly Property Preformatted() As Boolean
            Set(ByVal pPreformatted As Boolean)
                If pPreformatted Then
                    vLineBreak = vbCrLf
                Else
                    vLineBreak = "<br/>"
                End If
            End Set
        End Property

        Public WriteOnly Property LineOriented() As Boolean
            Set(ByVal pLineOriented As Boolean)
                vLineOriented = pLineOriented
            End Set
        End Property

        Public WriteOnly Property Debug() As Boolean
            Set(ByVal pDebug As Boolean)
                vDebug = pDebug
            End Set
        End Property

        Public WriteOnly Property Outlen() As Integer
            Set(ByVal pOutlen As Integer)
                vOutlen = pOutlen
            End Set
        End Property

        Private Function Tokenize(ByVal pText As String) As Vector
            Dim vMatches As MatchCollection
            Dim vMatch As Match
            Dim vMatches2 As MatchCollection
            Dim vMatch2 As Match
            Dim vValue As String = ""

            Tokenize = New Vector
            'vRegEx = New Regex
            'vRegEx.IgnoreCase = False
            'vRegEx.Global = True
            'vRegEx.Pattern = ".+"
            pText = Replace(pText, Chr(9), Space(8))
            If Not vLineOriented Then
                'vRegEx2 = New Regex
                'vRegEx2.IgnoreCase = False
                'vRegEx2.Global = True
                'vRegEx2.Pattern = "\s*\S+"
            End If
            vMatches = Regex.Matches(pText, ".+")
            For Each vMatch In vMatches
                vValue = Replace(vMatch.Value, vbCr, "")
                If vLineOriented Then
                    Tokenize.Push(vValue)
                Else
                    If Trim(vValue) = "" Then
                        Tokenize.Push(vValue)
                    Else
                        vMatches2 = Regex.Matches(vValue, "\s*\S+")
                        For Each vMatch2 In vMatches2
                            Tokenize.Push(vMatch2.Value)
                        Next
                    End If
                    Tokenize.Push(vbCrLf)
                End If
            Next
            If vValue = "" Then
                Tokenize.Push("")
            ElseIf Not vLineOriented Then
                Tokenize.Pop()
            End If
        End Function

        Private Sub HashB()
            Dim i As Integer
            Dim vElem As Object
            Dim vList As Object

            'vBhash = CreateObject("Scripting.Dictionary")
            vBhash = New Scripting.Dictionary
            For i = 0 To vB.Count - 1
                vElem = CStr(vB.ElementAt(i))
                If Trim(CStr(vElem)) <> "" And CStr(vElem) <> vbCrLf Then
                    If vBhash.Exists(vElem) Then
                        vList = CType(vBhash.Item(vElem), Vector)
                        CType(vList, Vector).Push(i)
                    Else
                        vList = New Vector
                        CType(vList, Vector).Push(i)
                        vBhash.Add(vElem, vList)
                    End If
                End If
            Next
        End Sub


        ' find longest matching block in vA[pALow,pAHigh] and vB[pBLow,pBHigh]
        Private Sub FindLongestMatch(ByVal pALow As Integer _
            , ByVal pAHigh As Integer _
            , ByVal pBLow As Integer _
            , ByVal pBHigh As Integer)

            Dim i As Integer
            Dim j As Integer
            Dim k As Integer
            Dim x As Integer
            Dim vList As Vector
            Dim vLen As Vector
            Dim vNewLen As Vector
            Dim vElem As Object

            bestStartA = pALow
            bestStartB = pBLow
            bestSize = 0


            vLen = New Vector
            vLen.Dimension = vB.Count
            For i = pALow To pAHigh
                vNewLen = New Vector
                vNewLen.Dimension = vB.Count
                vElem = vA.ElementAt(i)
                If vBhash.Exists(vElem) Then
                    vList = CType(vBhash.Item(vElem), Vector)
                    For x = 0 To vList.Count - 1
                        j = CInt(vList.ElementAt(x))
                        If j > pBHigh Then
                            Exit For
                        End If
                        If j >= pBLow Then
                            If j > 0 Then
                                k = CInt(vLen.ElementAt(j - 1)) + 1
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
                If (Trim(CStr(vA.ElementAt(bestStartA - 1))) = "" _
                    Or CStr(vA.ElementAt(bestStartA - 1)) = vbCrLf) _
                    And (Trim(CStr(vB.ElementAt(bestStartB - 1))) = "" _
                    Or CStr(vB.ElementAt(bestStartB - 1)) = vbCrLf) Then

                    bestStartA = bestStartA - 1
                    bestStartB = bestStartB - 1
                    bestSize = bestSize + 1
                Else
                    Exit Do
                End If
            Loop
            Do While bestStartA + bestSize <= pAHigh And bestStartB + bestSize <= pBHigh
                If (Trim(CStr(vA.ElementAt(bestStartA + bestSize))) = "" _
                    Or CStr(vA.ElementAt(bestStartA + bestSize)) = vbCrLf) _
                    And (Trim(CStr(vB.ElementAt(bestStartB + bestSize))) = "" _
                    Or CStr(vB.ElementAt(bestStartB + bestSize)) = vbCrLf) Then

                    bestSize = bestSize + 1
                Else
                    Exit Do
                End If
            Loop
        End Sub

        Private Sub SplitLine(ByVal pLine As String)
            Dim i As Integer
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

        Private Sub Output(ByVal pTag As String _
            , ByVal pVector As Vector _
            , ByVal pFrom As Integer _
            , ByVal pTo As Integer)
            Dim i As Integer
            Dim vElem As String

            If pTag = "delete" Then
                vOut = vOut & "<strike class='diff'>"
            ElseIf pTag = "insert" Then
                vOut = vOut & "<u class='diff'>"
            End If

            For i = pFrom To pTo
                vElem = CStr(pVector.ElementAt(i))
                If vElem = vbCrLf Then
                    vElem = vLineBreak
                    vOutlen = 0
                ElseIf vElem = "" Then
                    vElem = "  "
                End If

                vOutlen = vOutlen + Len(vElem)
                If vOutlen > 80 Then
                    If Len(vElem) > 80 Then
                        SplitLine(vElem)
                        vElem = ""
                    Else
                        vOut = vOut & vLineBreak
                        vElem = LTrim(vElem)
                        vOutlen = Len(vElem)
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

        Private Sub InnerReplace(ByVal pAFrom As Integer _
            , ByVal pATo As Integer _
            , ByVal pBFrom As Integer _
            , ByVal pBTo As Integer)

            Dim i As Integer
            Dim vText1 As String
            Dim vText2 As String
            Dim vMatcher As Matcher

            vText1 = ""
            vText2 = ""
            For i = pAFrom To pATo
                vText1 = vText1 & CStr(vA.ElementAt(i))
                If i < pATo Then
                    vText1 = vText1 & vbCrLf
                End If
            Next
            For i = pBFrom To pBTo
                vText2 = vText2 & CStr(vB.ElementAt(i))
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

        Private Sub Out(ByVal vAFound As Integer _
            , ByVal vBFound As Integer _
            , ByVal vSize As Integer)
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

        '  match between [pALow,pAHigh] and [pBLow,pBHigh]
        Private Sub GetMatchingBlocks(ByVal pDepth As Integer _
            , ByVal pALow As Integer _
            , ByVal pAHigh As Integer _
            , ByVal pBLow As Integer _
            , ByVal pBHigh As Integer)

            Dim vAFound As Integer
            Dim vBFound As Integer
            Dim vSize As Integer

            If pDepth = 1 Then
                matchedI = 0
                matchedJ = 0
            End If

            Call FindLongestMatch(pALow, pAHigh, pBLow, pBHigh)

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

        Public Function Compare(ByVal pText1 As String, ByVal pText2 As String) As String
            vOut = ""
            vA = Tokenize(pText1)
            vB = Tokenize(pText2)
            HashB()
            GetMatchingBlocks(1, 0, vA.Count - 1, 0, vB.Count - 1)
            Compare = vOut
        End Function
    End Class
End Namespace