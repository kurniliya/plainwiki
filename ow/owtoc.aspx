
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
'      $Source: /usr/local/cvsroot/openwiki/dist/owbase/ow/owtoc.asp,v $
'    $Revision: 1.2 $
'      $Author: pit $
' ---------------------------------------------------------------------------
'

Class TableOfContents
Private vTOCCurLevel, vTOC, vTOCStart, myCount As Object
	Private vCurNum As Object
	Private vNumberStack As Object
	Private vLevelRaised As Object
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1061.asp'
	Private Sub Class_Initialize_Renamed()
		vTOCStart = 0
		vTOCCurLevel = -1
		myCount = 0
		vCurNum = 0
		vNumberStack = New Vector
		'        vNumberStack.Push(1)        
		vLevelRaised = False
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1061.asp'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object vNumberStack may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
		vNumberStack = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	Public ReadOnly Property Count() As Object
		Get
			Count = myCount + 1
		End Get
	End Property
	
	Public ReadOnly Property NumberTrace() As Object
		Get
			NumberTrace = vNumberStack.Trace(".")
		End Get
	End Property
	
	Public ReadOnly Property CurNum() As Object
		Get
			CurNum = vNumberStack.Top
		End Get
	End Property
	
	Public Sub AddTOC(ByRef pLevel As Object, ByRef pText As Object)
		'        Dim i
		Dim vStr As Object
		
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
			vCurNum = vNumberStack.Pop
			vNumberStack.Push(vCurNum + 1)
		End If
		
		vStr = "<ow:toctext>" & "<number>" & Count & "</number>" & "<level>" & vTOCCurLevel & "</level>" & "<number_trace>" & NumberTrace & "</number_trace>" & "<text>" & pText & "</text>" & "</ow:toctext>"
		
		vTOC = vTOC & vStr & vbCrLf
		
		myCount = myCount + 1
	End Sub
	
	Public Function GetTOC() As Object
		Do While (vTOCCurLevel >= vTOCStart)
			vTOC = vTOC & "</" & GetTOCEndElement() & ">" & vbCrLf
			vTOCCurLevel = vTOCCurLevel - 1
		Loop 
		GetTOC = vTOC
	End Function
	
	Private Function GetTOCElement(ByRef pLevel As Object) As Object
		Dim cNumTOC As Object
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
	
	Private Function GetTOCEndElement() As Object
		Dim cNumTOC As Object
		If cNumTOC = 0 Then
			'    GetTOCEndElement = "dl"
			'	Changed because multi level "dl"'s don't intend in FF        
			GetTOCEndElement = "ow:toc"
		Else
			GetTOCEndElement = "ow:toc"
		End If
	End Function
End Class

