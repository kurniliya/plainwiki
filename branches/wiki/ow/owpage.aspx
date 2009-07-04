
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
'      $Source: /usr/local/cvsroot/openwiki/dist/owbase/ow/owpage.asp,v $
'    $Revision: 1.3 $
'      $Author: pit $
' ---------------------------------------------------------------------------
'

Class WikiPage
Private vLastMinor, vRevision, vAnchor, vName, vText, vChanges, vLastMajor As Object
	Private vChangesCollection As Object
	Private vAttachmentCollection As Object
	Private vAttachmentsLoaded As Object
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1061.asp'
	Private Sub Class_Initialize_Renamed()
		vText = ""
		vRevision = 0
		vChanges = 0
		vChangesCollection = New Vector
		vAttachmentsLoaded = False
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1061.asp'
	Private Sub Class_Terminate_Renamed()
		' You are the weakest link. Goodbye.
		'UPGRADE_NOTE: Object vChangesCollection may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
		vChangesCollection = Nothing
		'UPGRADE_NOTE: Object vAttachmentCollection may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
		vAttachmentCollection = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	
	Public Property Name() As Object
		Get
			Return vName
		End Get
		Set(ByVal Value As Object)
			vName = Value
		End Set
	End Property
	
	
	Public Property Anchor() As Object
		Get
			Return vAnchor
		End Get
		Set(ByVal Value As Object)
			vAnchor = Value
		End Set
	End Property
	
	
	Public Property Revision() As Object
		Get
			Revision = vRevision
		End Get
		Set(ByVal Value As Object)
			vRevision = Value
		End Set
	End Property
	
	
	Public Property Changes() As Object
		Get
			Changes = vChanges
		End Get
		Set(ByVal Value As Object)
			vChanges = Value
		End Set
	End Property
	
	
	Public Property LastMinor() As Object
		Get
			LastMinor = vLastMinor
		End Get
		Set(ByVal Value As Object)
			vLastMinor = Value
		End Set
	End Property
	
	
	Public Property LastMajor() As Object
		Get
			LastMajor = vLastMajor
		End Get
		Set(ByVal Value As Object)
			vLastMajor = Value
		End Set
	End Property
	
	
	Public Property Text() As Object
		Get
			Return vText
		End Get
		Set(ByVal Value As Object)
			vText = Value
		End Set
	End Property
	
	
	Public Property AttachmentsLoaded() As Object
		Get
			AttachmentsLoaded = vAttachmentsLoaded
		End Get
		Set(ByVal Value As Object)
			vAttachmentsLoaded = Value
		End Set
	End Property
	
	Public Function AddChange() As Object
		Dim vChange As Object
		vChange = New Change
		vChangesCollection.Push(vChange)
		AddChange = vChange
	End Function
	
	Public Function GetLastChange() As Object
		GetLastChange = vChangesCollection.ElementAt(0)
	End Function
	
	Public Function Exists() As Object
		If vChangesCollection.ElementAt(0).Timestamp = "" Then
			Exists = False
		Else
			Exists = True
		End If
	End Function
	
	Public Sub AddAttachment(ByRef pAttachment As Object, ByRef pStoreMaxRevOnly As Object)
		Dim vCount, i, vAttachment As Object
		If IsNothing(vAttachmentCollection) Then
			vAttachmentCollection = New Vector
		End If
		If pStoreMaxRevOnly Then
			For i = 0 To vAttachmentCollection.Count - 1
				vAttachment = vAttachmentCollection.ElementAt(i)
				If vAttachment.Name = pAttachment.Name Then
					If vAttachment.Revision < pAttachment.Revision Then
						vAttachmentCollection.RemoveElementAt(i)
						vAttachmentCollection.Push(pAttachment)
					End If
					Exit Sub
				End If
			Next 
		End If
		vAttachmentCollection.Push(pAttachment)
	End Sub
	
	Public Function GetAttachment(ByRef pName As Object) As Object
		Dim vCount, i, vAttachment As Object
		If Not IsNothing(vAttachmentCollection) Then
			For i = 0 To vAttachmentCollection.Count - 1
				vAttachment = vAttachmentCollection.ElementAt(i)
				If vAttachment.Name = pName Then
					GetAttachment = vAttachment
					Exit Function
				End If
			Next 
		End If
		'UPGRADE_NOTE: Object GetAttachment may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
		GetAttachment = Nothing
	End Function
	
	Public Function GetAttachmentPattern() As Object
		Dim vCount, i, vAttachment As Object
		GetAttachmentPattern = ""
		If Not IsNothing(vAttachmentCollection) Then
			For i = 0 To vAttachmentCollection.Count - 1
				vAttachment = vAttachmentCollection.ElementAt(i)
				If i > 0 Then
					GetAttachmentPattern = GetAttachmentPattern & "|"
				End If
				GetAttachmentPattern = GetAttachmentPattern & Replace(vAttachment.Name, ".", "\.")
			Next 
		End If
	End Function
	
	Public Function ToLinkXML(ByRef pText As Object, ByRef pTemplate As Object, ByRef pAddPath As Object) As Object
		Dim FormatDateISO8601 As Object
		Dim gAction As Object
		Dim cTemplateLinking As Object
		Dim cDirectEdit As Object
		Dim gScriptName As Object
		Dim vLastChange, vTemp As Object
		vLastChange = vChangesCollection.ElementAt(0)
		If vLastChange.Timestamp = "" Then
			'UPGRADE_NOTE: Global Sub/Function CDATAEncode is not accessible
			ToLinkXML = "<ow:link name='" & CDATAEncode(vName) & "' "
			If vAnchor <> "" Then
				'UPGRADE_NOTE: Global Sub/Function CDATAEncode is not accessible
				ToLinkXML = ToLinkXML & " anchor='" & CDATAEncode(vAnchor) & "'"
			End If
			ToLinkXML = ToLinkXML & " href='" & gScriptName & "?p=" & HttpContext.Current.Server.URLEncode(vName)
			If cDirectEdit = 1 Or (cTemplateLinking = 1 And pTemplate <> "") Then
				ToLinkXML = ToLinkXML & "&amp;a=edit"
			End If
			If cTemplateLinking = 1 And pTemplate <> "" Then
				ToLinkXML = ToLinkXML & "&amp;template=" & pTemplate
			End If
			ToLinkXML = ToLinkXML & "'>"
		Else
			If gAction = "print" Then
				vTemp = gScriptName & "?p=" & HttpContext.Current.Server.URLEncode(vName) & "&amp;a=print"
			Else
				vTemp = gScriptName & "?" & HttpContext.Current.Server.URLEncode(vName)
			End If
			'UPGRADE_NOTE: Global Sub/Function CDATAEncode is not accessible
			ToLinkXML = "<ow:link name='" & CDATAEncode(vName) & "'"
			If vAnchor <> "" Then
				'UPGRADE_NOTE: Global Sub/Function CDATAEncode is not accessible
				ToLinkXML = ToLinkXML & " anchor='" & CDATAEncode(vAnchor) & "'"
			End If
			ToLinkXML = ToLinkXML & " href='" & vTemp & "' date='" & FormatDateISO8601(vLastChange.Timestamp) & "'>"
		End If
		
		'UPGRADE_NOTE: Global Sub/Function PCDATAEncode is not accessible
		ToLinkXML = ToLinkXML & PCDATAEncode(pText) & "</ow:link>"
	End Function
	
	Public Function ToXML(ByRef pIncludeText As Object) As Object
		Dim cAllowAttachments As Object
		Dim cEmbeddedMode As Object
		Dim Wikify As Object
		Dim OPENWIKI_TEMPLATES As Object
		Dim gNamespace As Object
		Dim PrettyWikiLink As Object
		Dim i, vCount As Object
		'UPGRADE_NOTE: Global Sub/Function CDATAEncode is not accessible
		ToXML = "<ow:page name='" & CDATAEncode(vName) & "' changes='" & vChanges & "'"
		If vLastMinor > 0 Then
			ToXML = ToXML & " lastminor='" & vLastMinor & "'"
		End If
		If vLastMajor > 0 Then
			ToXML = ToXML & " lastmajor='" & vLastMajor & "'"
		End If
		If vRevision > 0 Then
			ToXML = ToXML & " revision='" & vRevision & "'"
		End If
		ToXML = ToXML & ">" & vbCrLf
		ToXML = ToXML & ToLinkXML(PrettyWikiLink(vName), "", True)
		If vChangesCollection.ElementAt(0).Revision > 0 Then
			vCount = vChangesCollection.Count - 1
			For i = 0 To vCount
				ToXML = ToXML & vChangesCollection.ElementAt(i).ToXML()
			Next 
		End If
		If pIncludeText = 1 Or pIncludeText = 3 Then
			If cEmbeddedMode = 0 And Trim(vText) = "" Then
				ToXML = gNamespace.GetIndexSchemes().GetTemplates(OPENWIKI_TEMPLATES) & ToXML & "<ow:body/>"
			Else
				ToXML = ToXML & "<ow:body>" & vbCrLf & Wikify(vText) & "</ow:body>" & vbCrLf
			End If
		End If
		If pIncludeText = 2 Or pIncludeText = 3 Then
			'UPGRADE_NOTE: Global Sub/Function PCDATAEncode is not accessible
			ToXML = ToXML & "<ow:raw>" & PCDATAEncode(vText) & "</ow:raw>"
		End If
		If pIncludeText = 4 Then
			ToXML = ToXML & "<ow:equation>" & vbCrLf & vText & "</ow:equation>" & vbCrLf
		End If
		
		If cAllowAttachments Then
			If Not IsNothing(vAttachmentCollection) Then
				vCount = vAttachmentCollection.Count - 1
				If vCount >= 0 Then
					ToXML = ToXML & "<ow:attachments>"
					For i = 0 To vCount
						ToXML = ToXML & vAttachmentCollection.ElementAt(i).ToXML(vName, "")
					Next 
					ToXML = ToXML & "</ow:attachments>"
				End If
			End If
		End If
		
		ToXML = ToXML & "</ow:page>" & vbCrLf
	End Function
End Class


Class Change
Private vByAlias, vMinorEdit, vRevision, vStatus, vTimestamp, vBy, vComment As Object
	Private vAttachmentChanges As Object
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1061.asp'
	Private Sub Class_Initialize_Renamed()
		vStatus = "new"
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1061.asp'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object vAttachmentChanges may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
		vAttachmentChanges = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	
	Public Property Status() As Object
		Get
			Status = vStatus
		End Get
		Set(ByVal Value As Object)
			Select Case Value
				Case 1
					vStatus = "new"
				Case 2
					vStatus = "updated"
				Case 3
					vStatus = "deleted"
				Case Else
					' must never happen
					HttpContext.Current.Response.Write("DING DONG !!!")
					vStatus = "unknown"
			End Select
		End Set
	End Property
	
	
	Public Property Revision() As Object
		Get
			Revision = vRevision
		End Get
		Set(ByVal Value As Object)
			vRevision = Value
		End Set
	End Property
	
	
	Public Property Timestamp() As Object
		Get
			Timestamp = vTimestamp
		End Get
		Set(ByVal Value As Object)
			vTimestamp = Value
		End Set
	End Property
	
	
	Public Property MinorEdit() As Object
		Get
			MinorEdit = vMinorEdit
		End Get
		Set(ByVal Value As Object)
			If Value = 1 Or Value = "1" Or Value = "true" Or Value = "on" Then
				vMinorEdit = 1
			Else
				vMinorEdit = 0
			End If
		End Set
	End Property
	
	
	Public Property By() As Object
		Get
			By = vBy
		End Get
		Set(ByVal Value As Object)
			Dim cMaskIPAddress As Object
			If cMaskIPAddress Then
				'UPGRADE_NOTE: Global Sub/Function s is not accessible
				vBy = s(Value, "\.\d+$", ".xxx", False, True)
			Else
				vBy = Value
			End If
		End Set
	End Property
	
	
	Public Property ByAlias() As Object
		Get
			ByAlias = vByAlias
		End Get
		Set(ByVal Value As Object)
			vByAlias = Value
		End Set
	End Property
	
	
	Public Property Comment() As Object
		Get
			Comment = vComment
		End Get
		Set(ByVal Value As Object)
			vComment = Value
		End Set
	End Property
	
	Public Sub UpdateBy()
		'UPGRADE_NOTE: Global Sub/Function GetRemoteUser is not accessible
		If (GetRemoteUser() <> vBy) Then
			vStatus = "updated"
		End If
	End Sub
	
	Public Sub AddAttachmentChange(ByRef pAttachmentChange As Object)
		If IsNothing(vAttachmentChanges) Then
			vAttachmentChanges = New Vector
		End If
		vAttachmentChanges.Push(pAttachmentChange)
	End Sub
	
	Public Function ToXML() As Object
		Dim gLastModified As Object
		Dim FormatDateISO8601 As Object
		Dim PrettyWikiLink As Object
		ToXML = ToXML & "<ow:change revision='" & vRevision & "' status='" & vStatus & "'"
		If vMinorEdit = 1 Then
			ToXML = ToXML & " minor='true'>" & vbCrLf
		Else
			ToXML = ToXML & " minor='false'>" & vbCrLf
		End If
		'UPGRADE_NOTE: Global Sub/Function CDATAEncode is not accessible
		ToXML = ToXML & "<ow:by name='" & CDATAEncode(vBy) & "'"
		If vByAlias <> "" Then
			'UPGRADE_NOTE: Global Sub/Function PCDATAEncode is not accessible
			'UPGRADE_NOTE: Global Sub/Function CDATAEncode is not accessible
			ToXML = ToXML & " alias='" & CDATAEncode(vByAlias) & "'>" & PCDATAEncode(PrettyWikiLink(vByAlias)) & "</ow:by>" & vbCrLf
		Else
			ToXML = ToXML & "/>"
		End If
		ToXML = ToXML & "<ow:date>" & FormatDateISO8601(vTimestamp) & "</ow:date>" & vbCrLf
		gLastModified = vTimestamp
		If vComment <> "" Then
			'UPGRADE_NOTE: Global Sub/Function PCDATAEncode is not accessible
			ToXML = ToXML & "<ow:comment>" & PCDATAEncode(vComment) & "</ow:comment>"
		End If
		Dim i As Object
		If Not IsNothing(vAttachmentChanges) Then
			For i = 0 To vAttachmentChanges.Count - 1
				ToXML = ToXML & vAttachmentChanges.ElementAt(i).ToXML()
			Next 
		End If
		ToXML = ToXML & "</ow:change>" & vbCrLf
	End Function
End Class


Class Attachment
Private vByAlias, vFilesize, vFilename, vHidden, vName, vRevision, vDeprecated, vTimestamp, vBy, vComment As Object
	
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
	
	
	Public Property Name() As Object
		Get
			Return vName
		End Get
		Set(ByVal Value As Object)
			vName = Value
		End Set
	End Property
	
	
	Public Property Revision() As Object
		Get
			Revision = vRevision
		End Get
		Set(ByVal Value As Object)
			vRevision = Value
		End Set
	End Property
	
	
	Public Property Hidden() As Object
		Get
			Hidden = vHidden
		End Get
		Set(ByVal Value As Object)
			vHidden = Value
		End Set
	End Property
	
	
	Public Property Deprecated() As Object
		Get
			Deprecated = vDeprecated
		End Get
		Set(ByVal Value As Object)
			vDeprecated = Value
		End Set
	End Property
	
	
	Public Property Filename() As Object
		Get
			Filename = vFilename
		End Get
		Set(ByVal Value As Object)
			vFilename = Value
		End Set
	End Property
	
	
	Public Property Timestamp() As Object
		Get
			Timestamp = vTimestamp
		End Get
		Set(ByVal Value As Object)
			vTimestamp = Value
		End Set
	End Property
	
	
	Public Property Filesize() As Object
		Get
			Filesize = vFilesize
		End Get
		Set(ByVal Value As Object)
			vFilesize = Value
		End Set
	End Property
	
	
	Public Property By() As Object
		Get
			By = vBy
		End Get
		Set(ByVal Value As Object)
			Dim cMaskIPAddress As Object
			If cMaskIPAddress Then
				'UPGRADE_NOTE: Global Sub/Function s is not accessible
				vBy = s(Value, "\.\d+$", ".xxx", False, True)
			Else
				vBy = Value
			End If
		End Set
	End Property
	
	
	Public Property ByAlias() As Object
		Get
			ByAlias = vByAlias
		End Get
		Set(ByVal Value As Object)
			vByAlias = Value
		End Set
	End Property
	
	
	Public Property Comment() As Object
		Get
			Comment = vComment
		End Get
		Set(ByVal Value As Object)
			vComment = Value
		End Set
	End Property
	
	Private Function GetIcon() As Object
		Dim gDocExtensions As Object
		Dim vPos, vExtension As Object
		vPos = InStrRev(vName, ".")
		If vPos > 0 Then
			vExtension = Mid(vName, vPos + 1)
			'UPGRADE_NOTE: Global Sub/Function m is not accessible
			If Not m(vExtension, "(" & gDocExtensions & ")", True, True) Then
				vExtension = "empty"
			End If
		Else
			vExtension = "empty"
		End If
		GetIcon = vExtension
	End Function
	
	Private Function FormatSize(ByRef pSize As Object) As Object
		FormatSize = Int((pSize / 1000) + 1)
		If FormatSize >= 1000000 Then
			'UPGRADE_WARNING: Mod has a new behavior. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1041.asp'
			FormatSize = Int((FormatSize / 1000000) + 1) & "," & (Int((FormatSize / 1000) + 1) - Int((FormatSize / 1000000) + 1)) & "," & (FormatSize Mod 1000)
		ElseIf FormatSize >= 1000 Then 
			'UPGRADE_WARNING: Mod has a new behavior. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1041.asp'
			FormatSize = Int((FormatSize / 1000) + 1) & "," & (FormatSize Mod 1000)
		End If
	End Function
	
	Public Function ToLinkXML(ByRef pHref As Object, ByRef pText As Object) As Object
		Dim FormatDateISO8601 As Object
		'UPGRADE_NOTE: Global Sub/Function PCDATAEncode is not accessible
		'UPGRADE_NOTE: Global Sub/Function CDATAEncode is not accessible
		ToLinkXML = "<ow:link name='" & CDATAEncode(vName) & "'" & " href='" & pHref & "' date='" & FormatDateISO8601(vTimestamp) & "'" & " attachment='true'>" & PCDATAEncode(pText) & "</ow:link>" & vbCrLf
	End Function
	
	Public Function ToXML(ByRef pPagename As Object, ByRef pText As Object) As Object
		Dim gLastModified As Object
		Dim FormatDateISO8601 As Object
		Dim PrettyWikiLink As Object
		Dim gImageExtensions As Object
		Dim vAttachmentLink, vIsImage As Object
		'UPGRADE_NOTE: Global Sub/Function CDATAEncode is not accessible
		ToXML = ToXML & "<ow:attachment name='" & CDATAEncode(vName) & "' revision='" & vRevision & "' hidden='"
		If vHidden = 1 Then
			ToXML = ToXML & "true"
		Else
			ToXML = ToXML & "false"
		End If
		ToXML = ToXML & "' deprecated='"
		If vDeprecated = 1 Then
			ToXML = ToXML & "true"
		Else
			ToXML = ToXML & "false"
		End If
		ToXML = ToXML & "'>"
		'UPGRADE_NOTE: Global Sub/Function GetAttachmentLink is not accessible
		vAttachmentLink = GetAttachmentLink(pPagename, vFilename)
		'UPGRADE_NOTE: Global Sub/Function m is not accessible
		If m(vAttachmentLink, "\.(" & gImageExtensions & ")$", True, True) Then
			vIsImage = "true"
		Else
			vIsImage = "false"
		End If
		'UPGRADE_NOTE: Global Sub/Function PCDATAEncode is not accessible
		'UPGRADE_NOTE: Global Sub/Function CDATAEncode is not accessible
		ToXML = ToXML & "<ow:file icon='" & GetIcon() & "' size='" & FormatSize(vFilesize) & "' href='" & CDATAEncode(vAttachmentLink) & "' image='" & vIsImage & "'>" & PCDATAEncode(vName) & "</ow:file>"
		'UPGRADE_NOTE: Global Sub/Function CDATAEncode is not accessible
		ToXML = ToXML & "<ow:by name='" & CDATAEncode(vBy) & "'"
		If vByAlias <> "" Then
			'UPGRADE_NOTE: Global Sub/Function PCDATAEncode is not accessible
			'UPGRADE_NOTE: Global Sub/Function CDATAEncode is not accessible
			ToXML = ToXML & " alias='" & CDATAEncode(vByAlias) & "'>" & PCDATAEncode(PrettyWikiLink(vByAlias)) & "</ow:by>" & vbCrLf
		Else
			ToXML = ToXML & "/>"
		End If
		ToXML = ToXML & "<ow:date>" & FormatDateISO8601(vTimestamp) & "</ow:date>" & vbCrLf
		gLastModified = vTimestamp
		If vComment <> "" Then
			'UPGRADE_NOTE: Global Sub/Function PCDATAEncode is not accessible
			ToXML = ToXML & "<ow:comment>" & PCDATAEncode(vComment) & "</ow:comment>"
		End If
		'UPGRADE_NOTE: Global Sub/Function PCDATAEncode is not accessible
		ToXML = ToXML & PCDATAEncode(pText)
		ToXML = ToXML & "</ow:attachment>"
	End Function
End Class


Class Attachment
Private vByAlias, vFilesize, vFilename, vHidden, vName, vRevision, vDeprecated, vTimestamp, vBy, vComment As Object
	
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
	
	
	Public Property Name() As Object
		Get
			Return vName
		End Get
		Set(ByVal Value As Object)
			vName = Value
		End Set
	End Property
	
	
	Public Property Revision() As Object
		Get
			Revision = vRevision
		End Get
		Set(ByVal Value As Object)
			vRevision = Value
		End Set
	End Property
	
	
	Public Property Hidden() As Object
		Get
			Hidden = vHidden
		End Get
		Set(ByVal Value As Object)
			vHidden = Value
		End Set
	End Property
	
	
	Public Property Deprecated() As Object
		Get
			Deprecated = vDeprecated
		End Get
		Set(ByVal Value As Object)
			vDeprecated = Value
		End Set
	End Property
	
	
	Public Property Filename() As Object
		Get
			Filename = vFilename
		End Get
		Set(ByVal Value As Object)
			vFilename = Value
		End Set
	End Property
	
	
	Public Property Timestamp() As Object
		Get
			Timestamp = vTimestamp
		End Get
		Set(ByVal Value As Object)
			vTimestamp = Value
		End Set
	End Property
	
	
	Public Property Filesize() As Object
		Get
			Filesize = vFilesize
		End Get
		Set(ByVal Value As Object)
			vFilesize = Value
		End Set
	End Property
	
	
	Public Property By() As Object
		Get
			By = vBy
		End Get
		Set(ByVal Value As Object)
			Dim cMaskIPAddress As Object
			If cMaskIPAddress Then
				'UPGRADE_NOTE: Global Sub/Function s is not accessible
				vBy = s(Value, "\.\d+$", ".xxx", False, True)
			Else
				vBy = Value
			End If
		End Set
	End Property
	
	
	Public Property ByAlias() As Object
		Get
			ByAlias = vByAlias
		End Get
		Set(ByVal Value As Object)
			vByAlias = Value
		End Set
	End Property
	
	
	Public Property Comment() As Object
		Get
			Comment = vComment
		End Get
		Set(ByVal Value As Object)
			vComment = Value
		End Set
	End Property
	
	Private Function GetIcon() As Object
		Dim gDocExtensions As Object
		Dim vPos, vExtension As Object
		vPos = InStrRev(vName, ".")
		If vPos > 0 Then
			vExtension = Mid(vName, vPos + 1)
			'UPGRADE_NOTE: Global Sub/Function m is not accessible
			If Not m(vExtension, "(" & gDocExtensions & ")", True, True) Then
				vExtension = "empty"
			End If
		Else
			vExtension = "empty"
		End If
		GetIcon = vExtension
	End Function
	
	Private Function FormatSize(ByRef pSize As Object) As Object
		FormatSize = Int((pSize / 1000) + 1)
		If FormatSize >= 1000000 Then
			'UPGRADE_WARNING: Mod has a new behavior. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1041.asp'
			FormatSize = Int((FormatSize / 1000000) + 1) & "," & (Int((FormatSize / 1000) + 1) - Int((FormatSize / 1000000) + 1)) & "," & (FormatSize Mod 1000)
		ElseIf FormatSize >= 1000 Then 
			'UPGRADE_WARNING: Mod has a new behavior. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1041.asp'
			FormatSize = Int((FormatSize / 1000) + 1) & "," & (FormatSize Mod 1000)
		End If
	End Function
	
	Public Function ToLinkXML(ByRef pHref As Object, ByRef pText As Object) As Object
		Dim FormatDateISO8601 As Object
		'UPGRADE_NOTE: Global Sub/Function PCDATAEncode is not accessible
		'UPGRADE_NOTE: Global Sub/Function CDATAEncode is not accessible
		ToLinkXML = "<ow:link name='" & CDATAEncode(vName) & "'" & " href='" & pHref & "' date='" & FormatDateISO8601(vTimestamp) & "'" & " attachment='true'>" & PCDATAEncode(pText) & "</ow:link>" & vbCrLf
	End Function
	
	Public Function ToXML(ByRef pPagename As Object, ByRef pText As Object) As Object
		Dim gLastModified As Object
		Dim FormatDateISO8601 As Object
		Dim PrettyWikiLink As Object
		Dim gImageExtensions As Object
		Dim vAttachmentLink, vIsImage As Object
		'UPGRADE_NOTE: Global Sub/Function CDATAEncode is not accessible
		ToXML = ToXML & "<ow:attachment name='" & CDATAEncode(vName) & "' revision='" & vRevision & "' hidden='"
		If vHidden = 1 Then
			ToXML = ToXML & "true"
		Else
			ToXML = ToXML & "false"
		End If
		ToXML = ToXML & "' deprecated='"
		If vDeprecated = 1 Then
			ToXML = ToXML & "true"
		Else
			ToXML = ToXML & "false"
		End If
		ToXML = ToXML & "'>"
		'UPGRADE_NOTE: Global Sub/Function GetAttachmentLink is not accessible
		vAttachmentLink = GetAttachmentLink(pPagename, vFilename)
		'UPGRADE_NOTE: Global Sub/Function m is not accessible
		If m(vAttachmentLink, "\.(" & gImageExtensions & ")$", True, True) Then
			vIsImage = "true"
		Else
			vIsImage = "false"
		End If
		'UPGRADE_NOTE: Global Sub/Function PCDATAEncode is not accessible
		'UPGRADE_NOTE: Global Sub/Function CDATAEncode is not accessible
		ToXML = ToXML & "<ow:file icon='" & GetIcon() & "' size='" & FormatSize(vFilesize) & "' href='" & CDATAEncode(vAttachmentLink) & "' image='" & vIsImage & "'>" & PCDATAEncode(vName) & "</ow:file>"
		'UPGRADE_NOTE: Global Sub/Function CDATAEncode is not accessible
		ToXML = ToXML & "<ow:by name='" & CDATAEncode(vBy) & "'"
		If vByAlias <> "" Then
			'UPGRADE_NOTE: Global Sub/Function PCDATAEncode is not accessible
			'UPGRADE_NOTE: Global Sub/Function CDATAEncode is not accessible
			ToXML = ToXML & " alias='" & CDATAEncode(vByAlias) & "'>" & PCDATAEncode(PrettyWikiLink(vByAlias)) & "</ow:by>" & vbCrLf
		Else
			ToXML = ToXML & "/>"
		End If
		ToXML = ToXML & "<ow:date>" & FormatDateISO8601(vTimestamp) & "</ow:date>" & vbCrLf
		gLastModified = vTimestamp
		If vComment <> "" Then
			'UPGRADE_NOTE: Global Sub/Function PCDATAEncode is not accessible
			ToXML = ToXML & "<ow:comment>" & PCDATAEncode(vComment) & "</ow:comment>"
		End If
		'UPGRADE_NOTE: Global Sub/Function PCDATAEncode is not accessible
		ToXML = ToXML & PCDATAEncode(pText)
		ToXML = ToXML & "</ow:attachment>"
	End Function
End ClassChange

