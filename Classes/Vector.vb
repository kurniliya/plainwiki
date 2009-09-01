Namespace Openwiki
    Public Class Vector
        Private myStack As Object()
        Private myCount As Integer

        Public Sub New()
            ReDim myStack(8)
            myCount = -1
        End Sub

        'Private Sub Class_Terminate()
        'End Sub

        Public WriteOnly Property Dimension() As Integer
            Set(ByVal pDim As Integer)
                ReDim myStack(pDim)
            End Set
        End Property

        Public ReadOnly Property Count() As Integer
            Get
                Count = myCount + 1
            End Get
        End Property

        Public Sub Push(ByVal pElem As Object)
            myCount = myCount + 1
            If (UBound(myStack) < myCount) Then
                ReDim Preserve myStack(UBound(myStack) * 2)
            End If
            Call SetElementAt(myCount, pElem)
        End Sub

        Public Function Pop() As Object
            If myCount >= LBound(myStack) And myCount <= UBound(myStack) Then
                If IsReference(myStack(myCount)) Then
                    Pop = myStack(myCount)
                Else
                    Pop = myStack(myCount)
                End If
                myCount = myCount - 1
            End If
        End Function

        Public Function Top() As Object
            If myCount >= LBound(myStack) And myCount <= UBound(myStack) Then
                If IsReference(myStack(myCount)) Then
                    Top = myStack(myCount)
                Else
                    Top = myStack(myCount)
                End If
            End If
        End Function

        Public Function ElementAt(ByVal pIndex As Integer) As Object
            If IsReference(myStack(pIndex)) Then
                ElementAt = myStack(pIndex)
            Else
                ElementAt = myStack(pIndex)
            End If
        End Function

        Public Sub SetElementAt(ByVal pIndex As Integer, ByVal pValue As Object)
            If IsReference(pValue) Then
                myStack(pIndex) = pValue
            Else
                myStack(pIndex) = pValue
            End If
        End Sub

        Public Sub RemoveElementAt(ByVal pIndex As Integer)
            Do While pIndex < myCount
                Call SetElementAt(pIndex, ElementAt(pIndex + 1))
                pIndex = pIndex + 1
            Loop
            myCount = myCount - 1
        End Sub

        Public Function IsEmpty() As Boolean
            IsEmpty = (myCount < 0)
        End Function

        Public Function Trace(ByVal pSeparator As String) As String
            Dim i As Integer
            Trace = ""

            If Not IsEmpty() Then
                For i = LBound(myStack) To Count - 1
                    If Not IsReference(myStack(i)) Then
                        If Trace = "" Then
                            Trace = CStr(myStack(i))
                        Else
                            Trace = Trace & pSeparator & CStr(myStack(i))
                        End If
                    End If
                Next
            End If
        End Function
    End Class
End Namespace