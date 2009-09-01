Namespace Openwiki
    Public Class Change
        Private vStatus As String, vRevision As Integer, vTimestamp As Date, vMinorEdit As Integer, vBy As String, vByAlias As String, vComment As String
        Private vAttachmentChanges As Vector

        Public Sub New()
            vStatus = "new"
        End Sub

        Protected Overrides Sub Finalize()
            vAttachmentChanges = Nothing
        End Sub

        Public Property Status() As String
            Get
                Return vStatus
            End Get
            Set(ByVal pStatus As String)
                Select Case pStatus
                    Case "1"
                        vStatus = "new"
                    Case "2"
                        vStatus = "updated"
                    Case "3"
                        vStatus = "deleted"
                    Case Else
                        ' must never happen
                        HttpContext.Current.Response.Write("DING DONG !!!")
                        vStatus = "unknown"
                End Select
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

        Public Property Timestamp() As Date
            Get
                Return vTimestamp
            End Get
            Set(ByVal pTimestamp As Date)
                vTimestamp = pTimestamp
            End Set
        End Property

        Public Property MinorEdit() As Integer
            Get
                Return vMinorEdit
            End Get
            Set(ByVal pMinorEdit As Integer)
                '                If CInt(pMinorEdit) = 1 Or pMinorEdit = "1" Or pMinorEdit = "true" Or pMinorEdit = "on" Then
                If pMinorEdit = 1 Then
                    vMinorEdit = 1
                Else
                    vMinorEdit = 0
                End If
            End Set
        End Property

        Public Property By() As String
            Get
                Return vBy
            End Get
            Set(ByVal pBy As String)
                If (cMaskIPAddress = 1) Then
                    vBy = s(pBy, "\.\d+$", ".xxx", False, True)
                Else
                    vBy = pBy
                End If
            End Set
        End Property

        Public Property ByAlias() As String
            Get
                Return vByAlias
            End Get
            Set(ByVal pByAlias As String)
                vByAlias = pByAlias
            End Set
        End Property

        Public Sub UpdateBy()
            If (GetRemoteUser() <> vBy) Then
                vStatus = "updated"
            End If
        End Sub

        Public Property Comment() As String
            Get
                Return vComment
            End Get
            Set(ByVal pComment As String)
                vComment = pComment
            End Set
        End Property

        Public Sub AddAttachmentChange(ByVal pAttachmentChange As AttachmentChange)
            If vAttachmentChanges Is Nothing Then
                vAttachmentChanges = New Vector
            End If
            vAttachmentChanges.Push(pAttachmentChange)
        End Sub

        Public Function ToXML() As String
            ToXML = "<ow:change revision='" & vRevision & "' status='" & vStatus & "'"
            If vMinorEdit = 1 Then
                ToXML = ToXML & " minor='true'>" & vbCrLf
            Else
                ToXML = ToXML & " minor='false'>" & vbCrLf
            End If
            ToXML = ToXML & "<ow:by name='" & CDATAEncode(vBy) & "'"
            If vByAlias <> "" Then
                ToXML = ToXML & " alias='" & CDATAEncode(vByAlias) & "'>" & PCDATAEncode(PrettyWikiLink(vByAlias)) & "</ow:by>" & vbCrLf
            Else
                ToXML = ToXML & "/>"
            End If
            ToXML = ToXML & "<ow:date>" & FormatDateISO8601(vTimestamp) & "</ow:date>" & vbCrLf
            gLastModified = vTimestamp
            If vComment <> "" Then
                ToXML = ToXML & "<ow:comment>" & PCDATAEncode(vComment) & "</ow:comment>"
            End If
            If Not vAttachmentChanges Is Nothing Then
                Dim i As Integer
                For i = 0 To vAttachmentChanges.Count - 1
                    ToXML = ToXML & CType(vAttachmentChanges.ElementAt(i), Change).ToXML()
                Next
            End If
            ToXML = ToXML & "</ow:change>" & vbCrLf
        End Function
    End Class
End Namespace