<%
'
' Implements a stack of tags wich is used in owwikify.asp.
'

Class TagStack
	Private myStack
	Private myDepth
	Private myCount
	
    Private Sub Class_Initialize()
	    Set myStack = new Vector
	    myDepth = 0
    End Sub	

    Private Sub Class_Terminate()
    	Set myStack = Nothing
    End Sub
    
    Public Property Get Count()
        Count = myStack.Count
    End Property

    Public Property Get Depth()
    	Depth = myDepth
    End Property
    
    Public Property Let Depth(pDepth)
    	myDepth = pDepth
    End Property

    Public Function IsEmpty()
        IsEmpty = myStack.IsEmpty
    End Function

	Public Sub NestList(pDepth, pCodeItem, pCodeList)
		Dim i
		Dim vTempEl
		
	    If myDepth < pDepth Then		        
	        For i = myStack.Count() - 1 to 0 step -1
	        	vTempEl = myStack.ElementAt(i)
	        	if Left(vTempEl, Len(pCodeItem & "></")) = pCodeItem & "></" Then
	        		myStack.SetElementAt i, Mid(vTempEl, Len(pCodeItem & "></") + 1, Len(vTempEl) - Len(pCodeItem & "></"))
	        		Exit For					
	        	End If
	        Next
	    End If

	    If myDepth > pDepth Then				        
	        For i = 0 to myStack.Count() - 1
	        	vTempEl = myStack.ElementAt(i)
	        	if Left(vTempEl, Len(pCodeList)) = pCodeList Then
	        		myStack.SetElementAt i , pCodeItem & "></" & vTempEl
	        		Exit For					
	        	End If
	        Next
	    End If
	    
	    myDepth = pDepth	
	End Sub
    
    Public Function Pop()
	    Pop = myStack.Pop()
    End Function 
    
	Public Function ProcessCodeClose(pDepth, pCodeItem, pCodeList)
	    If pDepth = 1 Then
	    	ProcessCodeClose = pCodeItem & "></" & pCodeList
	    Else
	        ProcessCodeClose = pCodeItem & "></" & pCodeList & "></" & pCodeItem
	    End If
	End Function       
    
	Public Function ProcessLine(pDepth, pCodeItem)
	    If pDepth = myDepth Then
	        ProcessLine =  "</" & pCodeItem & ">"
	    Else
	    	ProcessLine = ""
	    End If
	End Function    
    
    Public Sub Push(pElem)
    	myStack.Push(pElem)
    End Sub    

    Public Function Top()
	    Top = myStack.Top()
    End Function 	
		
End Class

%>