Namespace Openwiki
    Public Class WikiPage
        Private vName As String, vAnchor As String, vText As String, vRevision As Integer, vChanges As Integer, vLastMinor As Integer, vLastMajor As Integer
        Private vChangesCollection As Vector
        Private vAttachmentCollection As Vector
        Private vAttachmentsLoaded As Boolean

        Public Sub New()
            vText = ""
            vRevision = 0
            vChanges = 0
            vChangesCollection = New Vector
            vAttachmentsLoaded = False
        End Sub

        Protected Overrides Sub Finalize()
            ' You are the weakest link. Goodbye.
            vChangesCollection = Nothing
            vAttachmentCollection = Nothing
        End Sub

        Public Property Name() As String
            Get
                Return vName
            End Get
            Set(ByVal pName As String)
                vName = pName
            End Set
        End Property

        Public Property Anchor() As String
            Get
                Return vAnchor
            End Get
            Set(ByVal pAnchor As String)
                vAnchor = pAnchor
            End Set
        End Property

        Public Property Revision() As Integer
            Get
                Return vRevision
            End Get
            Set(ByVal pRevision As Integer)
                vRevision = pRevision
            End Set
        End Property

        Public Property Changes() As Integer
            Get
                Return vChanges
            End Get
            Set(ByVal pChanges As Integer)
                vChanges = pChanges
            End Set
        End Property

        Public Property LastMinor() As Integer
            Get
                Return vLastMinor
            End Get
            Set(ByVal pLastMinor As Integer)
                vLastMinor = pLastMinor
            End Set
        End Property

        Public Property LastMajor() As Integer
            Get
                Return vLastMajor
            End Get
            Set(ByVal pLastMajor As Integer)
                vLastMajor = pLastMajor
            End Set
        End Property

        Public Property Text() As String
            Get
                Return vText
            End Get
            Set(ByVal pText As String)
                vText = pText
            End Set
        End Property

        Public Function AddChange() As Change
            Dim vChange As Change
            vChange = New Change
            vChangesCollection.Push(vChange)
            AddChange = vChange
        End Function

        Public Function GetLastChange() As Change
            GetLastChange = CType(vChangesCollection.ElementAt(0), Change)
        End Function

        Public Function Exists() As Boolean
            If Not IsReference(CType(vChangesCollection.ElementAt(0), Change).Timestamp) Then
                Exists = False
            Else
                Exists = True
            End If
        End Function

        Public Property AttachmentsLoaded() As Boolean
            Get
                Return vAttachmentsLoaded
            End Get
            Set(ByVal pAttachmentsLoaded As Boolean)
                vAttachmentsLoaded = pAttachmentsLoaded
            End Set
        End Property

        Public Sub AddAttachment(ByVal pAttachment As Attachment, ByVal pStoreMaxRevOnly As Boolean)
            Dim i As Integer
            'Dim vCount As Integer
            Dim vAttachment As Attachment

            If Not IsReference(vAttachmentCollection) Then
                vAttachmentCollection = New Vector
            End If
            If pStoreMaxRevOnly Then
                For i = 0 To vAttachmentCollection.Count - 1
                    vAttachment = CType(vAttachmentCollection.ElementAt(i), Attachment)
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

        Public Function GetAttachment(ByVal pName As String) As Attachment
            Dim i As Integer
            '            Dim vCount As Integer
            Dim vAttachment As Attachment

            If IsReference(vAttachmentCollection) Then
                For i = 0 To vAttachmentCollection.Count - 1
                    vAttachment = CType(vAttachmentCollection.ElementAt(i), Attachment)
                    If vAttachment.Name = pName Then
                        GetAttachment = vAttachment
                        Exit Function
                    End If
                Next
            End If
            GetAttachment = Nothing
        End Function

        Public Function GetAttachmentPattern() As String
            Dim i As Integer
            '            Dim vCount As Integer
            Dim vAttachment As Attachment

            GetAttachmentPattern = ""
            If IsReference(vAttachmentCollection) Then
                For i = 0 To vAttachmentCollection.Count - 1
                    vAttachment = CType(vAttachmentCollection.ElementAt(i), Attachment)
                    If i > 0 Then
                        GetAttachmentPattern = GetAttachmentPattern & "|"
                    End If
                    GetAttachmentPattern = GetAttachmentPattern & Replace(vAttachment.Name, ".", "\.")
                Next
            End If
        End Function

        Public Function ToLinkXML(ByVal pText As String, ByVal pTemplate As String, ByVal pAddPath As Boolean) As String
            Dim vLastChange As Change, vTemp As String
            vLastChange = CType(vChangesCollection.ElementAt(0), Change)
            If Not IsReference(vLastChange.Timestamp) Then
                ToLinkXML = "<ow:link name='" & CDATAEncode(vName) & "' "
                If vAnchor <> "" Then
                    ToLinkXML = ToLinkXML & " anchor='" & CDATAEncode(vAnchor) & "'"
                End If
                ToLinkXML = ToLinkXML & " href='" & gScriptName & "?p=" & HttpContext.Current.Server.UrlEncode(vName)
                If cDirectEdit = 1 Or (cTemplateLinking = 1 And pTemplate <> "") Then
                    ToLinkXML = ToLinkXML & "&amp;a=edit"
                End If
                If cTemplateLinking = 1 And pTemplate <> "" Then
                    ToLinkXML = ToLinkXML & "&amp;template=" & pTemplate
                End If
                ToLinkXML = ToLinkXML & "'>"
            Else
                If gAction = "print" Then
                    vTemp = gScriptName & "?p=" & HttpContext.Current.Server.UrlEncode(vName) & "&amp;a=print"
                Else
                    vTemp = gScriptName & "?" & HttpContext.Current.Server.UrlEncode(vName)
                End If
                ToLinkXML = "<ow:link name='" & CDATAEncode(vName) & "'"
                If vAnchor <> "" Then
                    ToLinkXML = ToLinkXML & " anchor='" & CDATAEncode(vAnchor) & "'"
                End If
                ToLinkXML = ToLinkXML & " href='" & vTemp & "' date='" & FormatDateISO8601(vLastChange.Timestamp) & "'>"
            End If

            ToLinkXML = ToLinkXML & PCDATAEncode(pText) & "</ow:link>"
        End Function

        Public Function ToXML(ByVal pIncludeText As Integer) As String
            Dim i As Integer
            Dim vCount As Integer

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
            If CType(vChangesCollection.ElementAt(0), Change).Revision > 0 Then
                vCount = vChangesCollection.Count - 1
                For i = 0 To vCount
                    ToXML = ToXML & CType(vChangesCollection.ElementAt(i), Change).ToXML()
                Next
            End If
            If pIncludeText = 1 Or pIncludeText = 3 Then
                If cEmbeddedMode = 0 And Trim(vText) = "" Then
                    ToXML = gNamespace.GetIndexSchemes().GetTemplates(OPENWIKI_TEMPLATES) _
                          & ToXML & "<ow:body/>"
                Else
                    ToXML = ToXML & "<ow:body>" & vbCrLf & Wikify(vText) & "</ow:body>" & vbCrLf
                End If
            End If
            If pIncludeText = 2 Or pIncludeText = 3 Then
                ToXML = ToXML & "<ow:raw>" & PCDATAEncode(vText) & "</ow:raw>"
            End If
            If pIncludeText = 4 Then
                ToXML = ToXML & "<ow:equation>" & vbCrLf & vText & "</ow:equation>" & vbCrLf
            End If

            If (cAllowAttachments = 1) Then
                If IsReference(vAttachmentCollection) Then
                    vCount = vAttachmentCollection.Count - 1
                    If vCount >= 0 Then
                        ToXML = ToXML & "<ow:attachments>"
                        For i = 0 To vCount
                            ToXML = ToXML & CType(vAttachmentCollection.ElementAt(i), Attachment).ToXML(vName, "")
                        Next
                        ToXML = ToXML & "</ow:attachments>"
                    End If
                End If
            End If

            ToXML = ToXML & "</ow:page>" & vbCrLf
        End Function
    End Class
End Namespace