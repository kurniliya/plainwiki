Namespace Openwiki
    Public Class Attachment
        Private vName As String
        Private vRevision As Integer
        Private vHidden As Integer
        Private vDeprecated As Integer
        Private vFilename As String
        Private vTimestamp As Date
        Private vFilesize As Long
        Private vBy As String
        Private vByAlias As String
        Private vComment As String

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

        Public Property Hidden() As Integer
            Get
                Return vHidden
            End Get
            Set(ByVal pHidden As Integer)
                vHidden = pHidden
            End Set
        End Property

        Public Property Deprecated() As Integer
            Get
                Return vDeprecated
            End Get
            Set(ByVal pDeprecated As Integer)
                vDeprecated = pDeprecated
            End Set
        End Property

        Public Property Filename() As String
            Get
                Return vFilename
            End Get
            Set(ByVal pFilename As String)
                vFilename = pFilename
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

        Public Property Filesize() As Long
            Get
                Return vFilesize
            End Get
            Set(ByVal pFilesize As Long)
                vFilesize = pFilesize
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

        Public Property Comment() As String
            Get
                Return vComment
            End Get
            Set(ByVal pComment As String)
                vComment = pComment
            End Set
        End Property

        Private Function GetIcon() As String
            Dim vPos As Integer, vExtension As String
            vPos = InStrRev(vName, ".")
            If vPos > 0 Then
                vExtension = Mid(vName, vPos + 1)
                If Not m(vExtension, "(" & gDocExtensions & ")", True, True) Then
                    vExtension = "empty"
                End If
            Else
                vExtension = "empty"
            End If
            GetIcon = vExtension
        End Function

        Private Function FormatSize(ByVal pSize As Long) As String
            FormatSize = CStr((pSize / 1000) + 1)
            If CInt(FormatSize) >= 1000000 Then
                FormatSize = CStr((CInt(FormatSize) / 1000000) + 1) & "," & (CInt((CInt(FormatSize) / 1000) + 1) - CInt((CInt(FormatSize) / 1000000) + 1)) & "," & (CInt(FormatSize) Mod 1000)
            ElseIf CInt(FormatSize) >= 1000 Then
                FormatSize = CInt((CInt(FormatSize) / 1000) + 1) & "," & (CInt(FormatSize) Mod 1000)
            End If
        End Function

        Public Function ToLinkXML(ByVal pHref As String, ByVal pText As String) As String
            ToLinkXML = "<ow:link name=""" & CDATAEncode(vName) & """" _
                      & " href=""" & pHref & " "" date=""" & FormatDateISO8601(vTimestamp) & """" _
                      & " attachment=""true"">" _
                      & PCDATAEncode(pText) & "</ow:link>" & vbCrLf
        End Function

        Public Function ToXML(ByVal pPagename As String, ByVal pText As String) As String
            Dim vAttachmentLink As String, vIsImage As String

            ToXML = "<ow:attachment name='" & CDATAEncode(vName) & "' revision='" & vRevision & "' hidden='"
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
            vAttachmentLink = GetAttachmentLink(pPagename, vFilename)
            If m(vAttachmentLink, "\.(" & gImageExtensions & ")$", True, True) Then
                vIsImage = "true"
            Else
                vIsImage = "false"
            End If
            ToXML = ToXML & "<ow:file icon='" & GetIcon() & "' size='" & FormatSize(vFilesize) & "' href='" & CDATAEncode(vAttachmentLink) & "' image='" & vIsImage & "'>" & PCDATAEncode(vName) & "</ow:file>"
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
            ToXML = ToXML & PCDATAEncode(pText)
            ToXML = ToXML & "</ow:attachment>"
        End Function
    End Class
End Namespace