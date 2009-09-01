Namespace Openwiki
    Public Class AttachmentChange
        Private vName As String
        Private vRevision As Integer
        Private vTimestamp As Date
        Private vBy As String
        Private vByAlias As String
        Private vAction As String

        'Private Sub Class_Initialize()
        'End Sub

        'Private Sub Class_Terminate()
        'End Sub

        Public Property Name() As String
            Get
                Return vName
            End Get
            Set(ByVal pName As String)
                vName = pName
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

        Public Property By() As String
            Get
                Return vBy
            End Get
            Set(ByVal pBy As String)
                If cMaskIPAddress = 1 Then
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

        Public Property Action() As String
            Get
                Return vAction
            End Get
            Set(ByVal pAction As String)
                vAction = pAction
            End Set
        End Property

        Public Function ToXML() As String
            ToXML = "<ow:attachmentchange name='" & CDATAEncode(vName) & "' revision='" & vRevision & "'>"
            ToXML = ToXML & "<ow:by name='" & CDATAEncode(vBy) & "'"
            If vByAlias <> "" Then
                ToXML = ToXML & " alias='" & CDATAEncode(vByAlias) & "'>" & PCDATAEncode(PrettyWikiLink(vByAlias)) & "</ow:by>" & vbCrLf
            Else
                ToXML = ToXML & "/>"
            End If
            ToXML = ToXML & "<ow:date>" & FormatDateISO8601(vTimestamp) & "</ow:date>" & vbCrLf
            gLastModified = vTimestamp
            If vAction <> "" Then
                ToXML = ToXML & "<ow:action>" & PCDATAEncode(vAction) & "</ow:action>"
            End If
            ToXML = ToXML & "</ow:attachmentchange>"
        End Function
    End Class
End Namespace