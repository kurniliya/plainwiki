<script language="VB" runat="Server">
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
'      $Source: /usr/local/cvsroot/openwiki/dist/owbase/ow/owvector.asp,v $
'    $Revision: 1.2 $
'      $Author: pit $
' ---------------------------------------------------------------------------
'
' Implements a resizable array.
'

Class Vector
Private myStack() As String
	Private myCount As Double
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1061.asp'
	Private Sub Class_Initialize_Renamed()
		ReDim myStack(8)
		myCount = -1
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
	
	Public WriteOnly Property Dimension() As Object
		Set(ByVal Value As Object)
			ReDim myStack(Value)
		End Set
	End Property
	
	Public ReadOnly Property Count() As Short
		Get
			Count = myCount + 1
		End Get
	End Property
	
	Public Sub Push(ByRef pElem As String)
		myCount = myCount + 1
		If (UBound(myStack) < myCount) Then
			ReDim Preserve myStack(UBound(myStack) * 2)
		End If
		Call SetElementAt(myCount, pElem)
	End Sub
	
	Public Function Pop() As String
		If myCount >= LBound(myStack) And myCount <= UBound(myStack) Then
			If Not IsNothing(myStack(myCount)) Then
				Pop = myStack(myCount)
			Else
				Pop = myStack(myCount)
			End If
			myCount = myCount - 1
		End If
	End Function
	
	Public Function Top() As String
		If myCount >= LBound(myStack) And myCount <= UBound(myStack) Then
			If Not IsNothing(myStack(myCount)) Then
				Top = myStack(myCount)
			Else
				Top = myStack(myCount)
			End If
		End If
	End Function
	
	Public Function ElementAt(ByRef pIndex As Double) As String
		If Not IsNothing(myStack(pIndex)) Then
			ElementAt = myStack(pIndex)
		Else
			ElementAt = myStack(pIndex)
		End If
	End Function
	
	Public Sub SetElementAt(ByRef pIndex As Double, ByRef pValue As String)
		If Not IsNothing(pValue) Then
			myStack(pIndex) = pValue
		Else
			myStack(pIndex) = pValue
		End If
	End Sub
	
	Public Sub RemoveElementAt(ByRef pIndex As Double)
		Do While pIndex < myCount
			Call SetElementAt(pIndex, ElementAt(pIndex + 1))
			pIndex = pIndex + 1
		Loop 
		myCount = myCount - 1
	End Sub
	
	Public Function IsEmpty() As Boolean
		IsEmpty = (myCount < 0)
	End Function
	
	Public Function Trace(ByRef pSeparator As String) As String
		Dim i As Integer
		If Not IsEmpty() Then
			For i = LBound(myStack) To Count - 1
				If IsNothing(myStack(i)) Then
					If Trace = "" Then
						Trace = myStack(i)
					Else
						Trace = Trace & pSeparator & myStack(i)
					End If
				End If
			Next 
		End If
	End Function
End Class

</script>
