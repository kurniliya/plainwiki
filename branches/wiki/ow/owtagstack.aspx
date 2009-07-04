
'
' Implements a stack of tags wich is used in owwikify.asp.
'

Class TagStack
Private myStack As Object
	Private myDepth As Object
	Private myCount As Object
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1061.asp'
	Private Sub Class_Initialize_Renamed()
		myStack = New Vector
		myDepth = 0
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1061.asp'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object myStack may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
		myStack = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	Public ReadOnly Property Count() As Object
		Get
			Count = myStack.Count
		End Get
	End Property
	
	
	Public Property Depth() As Object
		Get
			Depth = myDepth
		End Get
		Set(ByVal Value As Object)
			myDepth = Value
		End Set
	End Property
	
	Public Function IsEmpty() As Object
		IsEmpty = myStack.IsEmpty
	End Function
	
	Public Sub NestList(ByRef pDepth As Object, ByRef pCodeItem As Object, ByRef pCodeList As Object)
		Dim i As Object
		Dim vTempEl As Object
		
		If myDepth < pDepth Then
			For i = myStack.Count() - 1 To 0 Step -1
				vTempEl = myStack.ElementAt(i)
				If Left(vTempEl, Len(pCodeItem & "></")) = pCodeItem & "></" Then
					myStack.SetElementAt(i, Mid(vTempEl, Len(pCodeItem & "></") + 1, Len(CStr(vTempEl)) - Len(pCodeItem & "></")))
					Exit For
				End If
			Next 
		End If
		
		If myDepth > pDepth Then
			For i = 0 To myStack.Count() - 1
				vTempEl = myStack.ElementAt(i)
				If Left(vTempEl, Len(CStr(pCodeList))) = pCodeList Then
					myStack.SetElementAt(i, pCodeItem & "></" & vTempEl)
					Exit For
				End If
			Next 
		End If
		
		myDepth = pDepth
	End Sub
	
	Public Function Pop() As Object
		Pop = myStack.Pop()
	End Function
	
	Public Function ProcessCodeClose(ByRef pDepth As Object, ByRef pCodeItem As Object, ByRef pCodeList As Object) As Object
		If pDepth = 1 Then
			ProcessCodeClose = pCodeItem & "></" & pCodeList
		Else
			ProcessCodeClose = pCodeItem & "></" & pCodeList & "></" & pCodeItem
		End If
	End Function
	
	Public Function ProcessLine(ByRef pDepth As Object, ByRef pCodeItem As Object) As Object
		If pDepth = myDepth Then
			ProcessLine = "</" & pCodeItem & ">"
		Else
			ProcessLine = ""
		End If
	End Function
	
	Public Sub Push(ByRef pElem As Object)
		myStack.Push(pElem)
	End Sub
	
	Public Function Top() As Object
		Top = myStack.Top()
	End Function
End Class

