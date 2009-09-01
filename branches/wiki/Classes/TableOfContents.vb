Namespace Openwiki
    Public Class TableOfContents
        Private vTOC As String, vTOCStart As Integer, vTOCCurLevel As Integer, myCount As Integer
        Private vCurNum As Integer
        Private vNumberStack As Vector
        Private vLevelRaised As Boolean

        Public Sub New()
            vTOCStart = 0
            vTOCCurLevel = -1
            myCount = 0
            vCurNum = 0
            vNumberStack = New Vector
            '        vNumberStack.Push(1)        
            vLevelRaised = False
        End Sub

        Protected Overrides Sub Finalize()
            vNumberStack = Nothing
        End Sub

        Public ReadOnly Property Count() As Integer
            Get
                Count = myCount + 1
            End Get
        End Property

        Public ReadOnly Property NumberTrace() As String
            Get
                NumberTrace = vNumberStack.Trace(".")
            End Get
        End Property

        Public ReadOnly Property CurNum() As Integer
            Get
                CurNum = CInt(vNumberStack.Top)
            End Get
        End Property

        Public Sub AddTOC(ByVal pLevel As Integer, ByVal pText As String)
            '        Dim i
            Dim vStr As String

            If vTOCStart = 0 Then
                vTOCStart = pLevel
                vTOCCurLevel = pLevel - 1
            End If

            ' 17/05/2009 IK: may be this block is unnecassary and not straightforward
            ' but MediaWiki works exactly so
            If pLevel > vTOCCurLevel + 1 Then
                pLevel = vTOCCurLevel + 1
            End If

            vLevelRaised = False

            Do While (vTOCCurLevel < pLevel)
                vTOC = vTOC & "<" & GetTOCElement(vTOCCurLevel) & ">" & vbCrLf
                vTOCCurLevel = vTOCCurLevel + 1
                vNumberStack.Push(1)
                vLevelRaised = True
            Loop
            Do While (vTOCCurLevel > pLevel)
                vTOC = vTOC & "</" & GetTOCEndElement() & ">" & vbCrLf
                vTOCCurLevel = vTOCCurLevel - 1
                vNumberStack.Pop()
            Loop
            Do While (vTOCStart > pLevel)
                vTOCStart = vTOCStart - 1
                vTOC = "<" & GetTOCElement(vTOCStart) & ">" & vTOC & vbCrLf
            Loop

            ' 25/02/2002 LP: commented because a <br/> after an <li> shows
            ' up bad in Opera. Besides, imo it's unnecessary.
            'If cNumTOC = 0 then
            '    vTOC = vTOC & pStr & "<br />" & vbCRLF
            'Else
            '    vTOC = vTOC & pStr & vbCRLF
            'End if

            If Not vLevelRaised Then
                vCurNum = CInt(vNumberStack.Pop)
                vNumberStack.Push(vCurNum + 1)
            End If

            vStr = "<ow:toctext>" _
               & "<number>" & Count & "</number>" _
               & "<level>" & vTOCCurLevel & "</level>" _
               & "<number_trace>" & NumberTrace & "</number_trace>" _
               & "<text>" & pText & "</text>" _
               & "</ow:toctext>"

            vTOC = vTOC & vStr & vbCrLf

            myCount = myCount + 1
        End Sub

        Public Function GetTOC() As String
            Do While (vTOCCurLevel >= vTOCStart)
                vTOC = vTOC & "</" & GetTOCEndElement() & ">" & vbCrLf
                vTOCCurLevel = vTOCCurLevel - 1
            Loop
            GetTOC = vTOC
        End Function

        Private Function GetTOCElement(ByVal pLevel As Integer) As String
            If cNumTOC = 0 Then
                '    GetTOCElement = "dl"
                '	Changed because multi level "dl"'s don't indent in FF
                GetTOCElement = "ow:toc mode=""indented"""
                '        Elseif pLevel = 0 Then
                '            GetTOCElement = "ol"
                '        Elseif pLevel = 1 Then
                '            GetTOCElement = "ol type=""I"""
                '        Elseif pLevel = 2 Then
                '            GetTOCElement = "ol type=""a"""
                '        Elseif pLevel = 3 Then
                '            GetTOCElement = "ol type=""i"""
                '        Elseif pLevel = 4 Then
                '            GetTOCElement = "ol type=""1"""
            Else
                '            GetTOCElement = "ol"
                GetTOCElement = "ow:toc"
            End If
        End Function

        Private Function GetTOCEndElement() As String
            If cNumTOC = 0 Then
                '    GetTOCEndElement = "dl"
                '	Changed because multi level "dl"'s don't intend in FF        
                GetTOCEndElement = "ow:toc"
            Else
                GetTOCEndElement = "ow:toc"
            End If
        End Function

    End Class
End Namespace