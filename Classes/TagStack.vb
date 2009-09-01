Namespace Openwiki
    Public Class TagStack
        Private myStack As Vector
        Private myDepth As Integer
        Private myCount As Integer

        Public Sub New()
            myStack = New Vector
            myDepth = 0
        End Sub

        Protected Overrides Sub Finalize()
            myStack = Nothing
        End Sub

        Public ReadOnly Property Count() As Integer
            Get
                Count = myStack.Count
            End Get
        End Property

        Public Property Depth() As Integer
            Get
                Depth = myDepth
            End Get

            Set(ByVal pDepth As Integer)
                myDepth = pDepth
            End Set
        End Property

        Public Function IsEmpty() As Boolean
            IsEmpty = myStack.IsEmpty
        End Function

        Public Sub NestList(ByVal pDepth As Integer _
            , ByVal pCodeItem As String _
            , ByVal pCodeList As String)
            Dim i As Integer
            Dim vTempEl As String

            If myDepth < pDepth Then
                For i = myStack.Count() - 1 To 0 Step -1
                    vTempEl = CStr(myStack.ElementAt(i))
                    If Left(vTempEl, Len(pCodeItem & "></")) = pCodeItem & "></" Then
                        myStack.SetElementAt(i, Mid(vTempEl, Len(pCodeItem & "></") + 1, Len(vTempEl) - Len(pCodeItem & "></")))
                        Exit For
                    End If
                Next
            End If

            If myDepth > pDepth Then
                For i = 0 To myStack.Count() - 1
                    vTempEl = CStr(myStack.ElementAt(i))
                    If Left(vTempEl, Len(pCodeList)) = pCodeList Then
                        myStack.SetElementAt(i, pCodeItem & "></" & vTempEl)
                        Exit For
                    End If
                Next
            End If

            myDepth = pDepth
        End Sub

        Public Function Pop() As String
            Pop = CStr(myStack.Pop())
        End Function

        Public Function ProcessCodeClose(ByVal pDepth As Integer _
            , ByVal pCodeItem As String _
            , ByVal pCodeList As String) _
        As String
            If pDepth = 1 Then
                ProcessCodeClose = pCodeItem & "></" & pCodeList
            Else
                ProcessCodeClose = pCodeItem & "></" & pCodeList & "></" & pCodeItem
            End If
        End Function

        Public Function ProcessLine(ByVal pDepth As Integer _
            , ByVal pCodeItem As String) _
        As String
            If pDepth = myDepth Then
                ProcessLine = "</" & pCodeItem & ">"
            Else
                ProcessLine = ""
            End If
        End Function

        Public Sub Push(ByVal pElem As String)
            myStack.Push(pElem)
        End Sub

        Public Function Top() As String
            Top = CStr(myStack.Top())
        End Function

    End Class
End Namespace