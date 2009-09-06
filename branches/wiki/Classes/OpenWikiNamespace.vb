Imports ADODB
Imports System.Text.RegularExpressions


Namespace Openwiki
    Public Class OpenWikiNamespace
        Private vConn As ADODB.Connection, vRS As ADODB.Recordset, vQuery As String
        Private vIndexSchemes As IndexSchemes
        Private vCachedPages As Scripting.Dictionary

        Public Sub New()
            If OPENWIKI_DB = "" Then
                cAllowAttachments = 0
                cWikiLinks = 0
                cCacheXML = 0
            Else
                'vConn = HttpContext.Current.Server.CreateObject("ADODB.Connection")
                vConn = New ADODB.Connection
                vConn.Open(OPENWIKI_DB)
                vRS = New ADODB.Recordset
            End If
            vIndexSchemes = New IndexSchemes
            '            vCachedPages = HttpContext.Current.Server.CreateObject("Scripting.Dictionary")
            vCachedPages = New Scripting.Dictionary
        End Sub

        Protected Overrides Sub Finalize()
            '            On Error Resume Next
            vConn.Close()
            vConn = Nothing
            vRS = Nothing
            vIndexSchemes = Nothing
            vCachedPages = Nothing
        End Sub

        Sub BeginTrans(ByVal pConn As ADODB.Connection)
            If OPENWIKI_DB_SYNTAX <> DB_MYSQL Then
                pConn.BeginTrans()
            End If
        End Sub

        Sub CommitTrans(ByVal pConn As ADODB.Connection)
            If OPENWIKI_DB_SYNTAX <> DB_MYSQL Then
                pConn.CommitTrans()
            End If
        End Sub

        Sub RollbackTrans(ByVal pConn As ADODB.Connection)
            If OPENWIKI_DB_SYNTAX <> DB_MYSQL Then
                pConn.RollbackTrans()
            End If
        End Sub

        Private Function CreatePageKey(ByVal pPageName As String _
            , ByVal pRevision As Integer _
            , ByVal pIncludeText As Integer _
            , ByVal pIncludeAllChangeRecords As Boolean) _
        As String
            CreatePageKey = pRevision & "_" & pIncludeText & "_" & pIncludeAllChangeRecords & "_" & pPageName
        End Function

        Private Function GetCachedPage(ByVal pPageName As String _
            , ByVal pRevision As Integer _
            , ByVal pIncludeText As Integer _
            , ByVal pIncludeAllChangeRecords As Boolean) _
        As WikiPage
            'Dim vKey As String
            'vKey = CreatePageKey(pPageName, pRevision, pIncludeText, pIncludeAllChangeRecords)
            'If Not (vCachedPages.Item(vKey) Is Nothing) Then
            '    GetCachedPage = CStr(vCachedPages.Item(vKey))
            'Else
            '    GetCachedPage = Nothing
            'End If

            GetCachedPage = Nothing
        End Function

        Private Sub SetCachedPage(ByVal pPageName As String _
            , ByVal pRevision As Integer _
            , ByVal pIncludeText As Integer _
            , ByVal pIncludeAllChangeRecords As Boolean _
            , ByVal vPage As WikiPage)
            'Dim vKey As String
            'vKey = CreatePageKey(pPageName, pRevision, pIncludeText, pIncludeAllChangeRecords)
            'vCachedPages.Add(vKey, vPage)
        End Sub

        Public Function GetIndexSchemes() As IndexSchemes
            GetIndexSchemes = vIndexSchemes
        End Function

        Function GetPageAndAttachments(ByVal pPageName As String _
            , ByVal pRevision As Integer _
            , ByVal pIncludeText As Integer _
            , ByVal pIncludeAllChangeRecords As Boolean) _
        As WikiPage
            Dim vPage As WikiPage

            vPage = GetCachedPage(pPageName, pRevision, pIncludeText, pIncludeAllChangeRecords)
            If TypeName(vPage) = TypeName(Nothing) Then
                vPage = GetPage(pPageName, pRevision, pIncludeText, False)
                If (cAllowAttachments = 1) Then
                    GetAttachments(vPage, pRevision, pIncludeAllChangeRecords)
                End If
            ElseIf (cAllowAttachments = 1) And Not vPage.AttachmentsLoaded Then
                GetAttachments(vPage, pRevision, pIncludeAllChangeRecords)
            End If
            GetPageAndAttachments = vPage
        End Function

        Function GetPage(ByVal pPageName As String _
            , ByVal pRevision As Integer _
            , ByVal pIncludeText As Integer _
            , ByVal pIncludeAllChangeRecords As Boolean) _
        As WikiPage
            Dim vPage As WikiPage, vChange As Change

            If cWikiLinks = 0 Then
                GetPage = New WikiPage
                GetPage.AddChange()
                GetPage.Name = "FrontPage"
                GetPage.Text = "Please provide a value for {{{OPENWIKI_DB}}} in your owconfig.asp file."
                Exit Function
            End If
            vPage = GetCachedPage(pPageName, pRevision, pIncludeText, pIncludeAllChangeRecords)
            If TypeName(vPage) = TypeName(Nothing) Then
                'Response.Write("LOAD PAGE: " & pPageName & "<br />")
                vPage = New WikiPage
                If (pIncludeText = 1) Then
                    vQuery = "SELECT * "
                Else
                    vQuery = "SELECT wpg_name, wpg_changes, wpg_lastminor, wpg_lastmajor, wrv_revision, wrv_status, wrv_timestamp, wrv_minoredit, wrv_by, wrv_byalias, wrv_comment "
                End If
                vQuery = vQuery & " FROM openwiki_pages, openwiki_revisions WHERE wpg_name = '" & Replace(pPageName, "'", "''") & "' AND wrv_name = wpg_name"
                If pRevision > 0 Then
                    vQuery = vQuery & " AND wrv_revision = " & pRevision
                ElseIf pIncludeAllChangeRecords Then
                    vQuery = vQuery & " ORDER BY wrv_revision DESC"
                Else
                    vQuery = vQuery & " AND wrv_current = 1"
                End If

                '                On Error Resume Next
                vRS.Open(vQuery, vConn, ADODB.CursorTypeEnum.adOpenForwardOnly)
                If Err.Number <> 0 Then
                    If Err.Number = -2147467259 Then
                        HttpContext.Current.Response.Write("<h2>Error:</h2>")
                        HttpContext.Current.Response.Write("Cannot find the data sources or the data sources are locked by another application.")
                        HttpContext.Current.Response.Write("Make sure you've set the constant <code><b>OPENWIKI_DB</b></code> correctly in your config file, pointing it to your data sources.<br /><br /><br />")
                    Else
                        HttpContext.Current.Response.Write(Err.Number & "<br />" & Err.Description)
                    End If
                    HttpContext.Current.Response.End()
                End If
                '                On Error GoTo 0

                If vRS.EOF Then
                    If pRevision = 0 Then
                        vPage.Name = pPageName
                        vPage.AddChange()
                    Else
                        ' TODO: addMessage("Revision " & pRevision & " not available (showing current version instead)"
                        vRS.Close()
                        GetPage = GetPage(pPageName, 0, pIncludeText, pIncludeAllChangeRecords)
                        Exit Function
                    End If
                Else
                    vPage.Name = CStr(vRS("wpg_name").Value)
                    vPage.Changes = CInt(vRS("wpg_changes").Value)
                    vPage.LastMinor = CInt(vRS("wpg_lastminor").Value)
                    vPage.LastMajor = CInt(vRS("wpg_lastmajor").Value)
                    If (pIncludeText = 1) Then
                        vPage.Text = CStr(vRS("wrv_text").Value)
                    End If
                    If CInt(vRS("wpg_lastminor").Value) = CInt(vRS("wrv_revision").Value) Then
                        ' wrv_current = 1
                        ' vPage.Revision = vRS("wrv_revision") ??? ---> No! Because of the xsl script.
                        vPage.Revision = 0
                    ElseIf pRevision > 0 Then
                        vPage.Revision = pRevision
                    End If
                    Do While Not vRS.EOF
                        vChange = vPage.AddChange

                        If Not IsDBNull(vRS("wrv_revision").Value) Then
                            vChange.Revision = CInt(vRS("wrv_revision").Value)
                        End If

                        If Not IsDBNull(vRS("wrv_status").Value) Then
                            vChange.Status = CStr(vRS("wrv_status").Value)
                        End If

                        If Not IsDBNull(vRS("wrv_minoredit").Value) Then
                            vChange.MinorEdit = CInt(vRS("wrv_minoredit").Value)
                        End If

                        If Not IsDBNull(vRS("wrv_timestamp").Value) Then
                            vChange.Timestamp = CDate(vRS("wrv_timestamp").Value)
                        End If

                        If Not IsDBNull(vRS("wrv_by").Value) Then
                            vChange.By = CStr(vRS("wrv_by").Value)
                        End If

                        If Not IsDBNull(vRS("wrv_byalias").Value) Then
                            vChange.ByAlias = CStr(vRS("wrv_byalias").Value)
                        End If

                        If Not Convert.ToString(vRS("wrv_comment").Value) = Convert.ToString(DBNull.Value) Then
                            vChange.Comment = CStr(vRS("wrv_comment").Value)
                        End If

                        vRS.MoveNext()
                    Loop
                End If
                vRS.Close()

                ' TODO: move this out of this method
                ' If this is the RecentChanges page, then force the presence of the
                ' <RecentChanges> element in the page.
                If vPage.Name = OPENWIKI_RCNAME Then
                    vPage.Text = s(vPage.Text, "\<RecentChanges\>", "<RecentChangesLong>", True, True)
                    If Not m(vPage.Text, "\<RecentChangesLong\>", True, True) Then
                        vPage.Text = vPage.Text & "<RecentChangesLong>"
                    End If
                End If

                SetCachedPage(pPageName, pRevision, pIncludeText, pIncludeAllChangeRecords, vPage)
            End If

            GetPage = vPage
        End Function

        Function GetPageCount() As Integer
            vQuery = "SELECT COUNT(*) FROM openwiki_pages"
            vRS.Open(vQuery, vConn, ADODB.CursorTypeEnum.adOpenForwardOnly)
            GetPageCount = CInt(vRS(0).Value)
            vRS.Close()
        End Function

        Function GetRevisionsCount() As Integer
            vQuery = "SELECT COUNT(*) FROM openwiki_revisions"
            vRS.Open(vQuery, vConn, ADODB.CursorTypeEnum.adOpenForwardOnly)
            GetRevisionsCount = CInt(vRS(0).Value)
            vRS.Close()
        End Function

        Function ToXML(ByVal pXmlStr As String) As String
            Dim vProtection As String

            If (cReadOnly = 1) Then
                vProtection = "readonly"
            ElseIf gEditPassword <> "" And m(gPage, OPENWIKI_PROTECTEDPAGES, False, False) Then
                vProtection = "password"
            ElseIf (cUseRecaptcha = 1) Then
                vProtection = "captcha"
            Else
                vProtection = "none"
            End If
            ToXML = "<ow:wiki version='" & OPENWIKI_XMLVERSION & "' xmlns:ow='" & OPENWIKI_NAMESPACE & "' encoding='" & OPENWIKI_ENCODING & "' mode='" & gAction & "'>" _
                  & "<ow:useragent>" & PCDATAEncode(HttpContext.Current.Request.ServerVariables("HTTP_USER_AGENT")) & "</ow:useragent>" _
                  & "<ow:location>" & PCDATAEncode(gServerRoot) & "</ow:location>" _
                  & "<ow:scriptname>" & PCDATAEncode(gScriptName) & "</ow:scriptname>" _
                  & "<ow:imagepath>" & PCDATAEncode(OPENWIKI_IMAGEPATH) & "</ow:imagepath>" _
                  & "<ow:iconpath>" & PCDATAEncode(OPENWIKI_ICONPATH) & "</ow:iconpath>" _
                  & "<ow:about>" & PCDATAEncode(gServerRoot & gScriptName & "?" & HttpContext.Current.Request.ServerVariables("QUERY_STRING")) & "</ow:about>" _
                  & "<ow:protection>" & vProtection & "</ow:protection>" _
                  & "<ow:title>" & PCDATAEncode(OPENWIKI_TITLE) & "</ow:title>" _
                  & "<ow:frontpage name='" & CDATAEncode(OPENWIKI_FRONTPAGE) & "' href='" & gScriptName & "?" & HttpContext.Current.Server.UrlEncode(OPENWIKI_FRONTPAGE) & "'>" & PCDATAEncode(PrettyWikiLink(OPENWIKI_FRONTPAGE)) & "</ow:frontpage>"
            If cEmbeddedMode = 0 Then
                If cAllowAttachments = 1 Then
                    ToXML = ToXML & "<ow:allowattachments/>"
                End If
                If HttpContext.Current.Request("redirect") <> "" Then
                    ToXML = ToXML & "<ow:redirectedfrom name='" & CDATAEncode(URLDecode(HttpContext.Current.Request("redirect"))) & "'>" & PCDATAEncode(PrettyWikiLink(URLDecode(HttpContext.Current.Request("redirect")))) & "</ow:redirectedfrom>"
                End If
                ToXML = ToXML & getUserPreferences() & GetCookieTrail()
            End If
            ToXML = ToXML & pXmlStr & "</ow:wiki>"
        End Function

        Private Function isValidDocument(ByVal pText As String) As Boolean
            'On Error Resume Next
            Dim vXmlStr As String
            Dim vXmlDoc As MSXML2.FreeThreadedDOMDocument60

            vXmlStr = "<ow:wiki xmlns:ow='x'>" & Wikify(pText) & "</ow:wiki>"

            'If MSXML_VERSION <> 3 Then
            '    vXmlDoc = HttpContext.Current.Server.CreateObject("Msxml2.FreeThreadedDOMDocument." & MSXML_VERSION & ".0")
            '    'vXslDoc.ResolveExternals = True
            '    'vXslDoc.setProperty("AllowXsltScript", True)
            'Else
            '    vXmlDoc = HttpContext.Current.Server.CreateObject("Msxml2.FreeThreadedDOMDocument")
            'End If

            vXmlDoc = New MSXML2.FreeThreadedDOMDocument60

            vXmlDoc.async = False
            If vXmlDoc.loadXML(vXmlStr) Then
                isValidDocument = True
            Else
                isValidDocument = False
                HttpContext.Current.Response.Write("<h1>Error occured</h1>")
                HttpContext.Current.Response.Write("<b>Your input did not validate to well-formed and valid Wiki format.<br />")
                HttpContext.Current.Response.Write("Please go back and correct. The XML output attempt was:</b><br /><br />")
                HttpContext.Current.Response.Write("<pre>" & vbCrLf & HttpContext.Current.Server.HtmlEncode(vXmlStr) & vbCrLf & "</pre>" & vbCrLf)
            End If
        End Function

        Function SavePage(ByVal pRevision As Integer _
            , ByVal pMinorEdit As Integer _
            , ByVal pComment As String _
            , ByVal pText As String) _
        As Boolean
            Dim vRevision As Integer
            Dim vStatus As Integer
            Dim vHost As String
            Dim vUserAgent As String
            Dim vBy As String
            Dim vByAlias As String
            Dim vReplacedTS As Date ', vRevsDeleted

            pText = pText & ""
            If Not isValidDocument(pText) Then
                SavePage = False
                HttpContext.Current.Response.End()
            End If

            vHost = GetRemoteHost()
            vUserAgent = HttpContext.Current.Request.ServerVariables("HTTP_USER_AGENT")
            vBy = GetRemoteUser()
            If vBy = "" Then
                vBy = vHost
            End If
            vByAlias = GetRemoteAlias()

            Dim conn As ADODB.Connection = New ADODB.Connection
            conn.Open(OPENWIKI_DB)
            BeginTrans(conn)
            vQuery = "SELECT * FROM openwiki_revisions WHERE wrv_name = '" & Replace(gPage, "'", "''") & "' AND wrv_current = 1"
            vRS.Open(vQuery, conn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, adCmdText)
            If vRS.EOF Then
                If Trim(pText) = "" Then
                    RollbackTrans(conn)
                    conn.Close()
                    conn = Nothing
                    SavePage = True
                    Exit Function
                End If
                vRevision = 1
                vStatus = 1  ' new
            ElseIf CStr(vRS("wrv_text").Value) = pText Then
                RollbackTrans(conn)
                conn.Close()
                conn = Nothing
                SavePage = True
                Exit Function
            Else
                If (CInt(vRS("wrv_revision").Value) <> (pRevision - 1)) Then
                    If ((CStr(vRS("wrv_by").Value) <> vBy) _
                        Or (CStr(vRS("wrv_host").Value) <> vHost) _
                        Or (CStr(vRS("wrv_agent").Value) <> vUserAgent)) _
                    Then
                        RollbackTrans(conn)
                        conn.Close()
                        conn = Nothing
                        SavePage = False
                        Exit Function
                    End If
                End If
                vRevision = CInt(vRS("wrv_revision").Value) + 1
                If ((CStr(vRS("wrv_by").Value) = vBy) _
                    And (CStr(vRS("wrv_host").Value) = vHost) _
                    And (CStr(vRS("wrv_agent").Value) = vUserAgent)) Then
                    vStatus = CInt(vRS("wrv_status").Value)
                Else
                    vStatus = 2  ' updated
                End If
            End If

            If InStr(pText, "#DEPRECATED") = 1 Then
                vStatus = 3  ' deleted
            ElseIf vStatus = 3 Then
                vStatus = 2  ' updated
            End If

            If vRS.EOF Then
                vQuery = "INSERT INTO openwiki_pages (wpg_name, wpg_lastminor, wpg_changes, wpg_lastmajor) VALUES " _
                       & "('" & Replace(gPage, "'", "''") & "'," & vRevision & " ,1 ," & vRevision & ")"
                conn.Execute(vQuery)
            Else
                vQuery = "UPDATE openwiki_pages " _
                       & "SET wpg_changes = wpg_changes + 1" _
                       & ",   wpg_lastminor = " & vRevision
                If pMinorEdit = 0 Then
                    vQuery = vQuery & ", wpg_lastmajor = " & vRevision
                End If
                vQuery = vQuery & " WHERE wpg_name = '" & Replace(gPage, "'", "''") & "'"
                conn.Execute(vQuery)

                vQuery = "UPDATE openwiki_revisions SET wrv_current = 0 WHERE wrv_name = '" & Replace(gPage, "'", "''") & "' AND wrv_current = 1"
                conn.Execute(vQuery)
            End If
            vRS.Close()

            vRS.Open("openwiki_revisions", conn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, adCmdTable)
            vRS.AddNew()
            vRS("wrv_name").Value = gPage
            vRS("wrv_revision").Value = vRevision
            vRS("wrv_current").Value = 1
            vRS("wrv_status").Value = vStatus
            vRS("wrv_timestamp").Value = Now()
            vRS("wrv_minoredit").Value = pMinorEdit
            vRS("wrv_host").Value = vHost
            vRS("wrv_agent").Value = vUserAgent
            vRS("wrv_by").Value = vBy
            vRS("wrv_byalias").Value = vByAlias
            vRS("wrv_comment").Value = pComment
            vRS("wrv_text").Value = pText
            vRS.Update()
            vRS.Close()

            ' delete old revisions
            vQuery = "SELECT wrv_revision, wrv_timestamp FROM openwiki_revisions WHERE wrv_name = '" & Replace(gPage, "'", "''") & "' ORDER BY wrv_revision DESC"
            vRS.Open(vQuery, conn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, adCmdText)
            If Not vRS.EOF Then
                ' this is the current revision
                vRS.MoveNext()
                If Not vRS.EOF Then
                    vReplacedTS = CDate(vRS("wrv_timestamp").Value)
                    ' keep at least one old revision
                    vRS.MoveNext()
                    Do While Not vRS.EOF
                        ' check the timestamp of revision that replaced this revision
                        If vReplacedTS < (Now().AddDays(-1 * OPENWIKI_DAYSTOKEEP)) Then
                            vQuery = "DELETE FROM openwiki_revisions WHERE wrv_name = '" & Replace(gPage, "'", "''") & "' AND wrv_revision <= " & CInt(vRS("wrv_revision").Value)
                            conn.Execute(vQuery)
                            vRS.Close()
                            vQuery = "SELECT COUNT(*) FROM openwiki_revisions WHERE wrv_name = '" & Replace(gPage, "'", "''") & "'"
                            vRS.Open(vQuery, conn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, adCmdText)
                            vQuery = "UPDATE openwiki_pages SET wpg_changes = " & CInt(vRS(0).Value) & " WHERE wpg_name = '" & Replace(gPage, "'", "''") & "'"
                            conn.Execute(vQuery)
                            Exit Do
                        Else
                            vReplacedTS = CDate(vRS("wrv_timestamp").Value)
                        End If
                        vRS.MoveNext()
                    Loop
                End If
            End If
            vRS.Close()

            ' throw out the bath and the bathwater. TODO: keep the bath
            ClearDocumentCache(conn)

            CommitTrans(conn)
            conn.Close()

            conn = Nothing

            SavePage = True
        End Function


        ' returns the name of the file as you should save it
        ' pStatus : 0 = normal, 1 = hidden, 2 = deprecated
        Function SaveAttachmentMetaData(ByVal pFilename As String _
            , ByVal pFilesize As Integer _
            , ByVal pAddLink As String _
            , ByVal pHidden As Integer _
            , ByVal pComment As String) _
        As String
            Dim vHost As String
            Dim vUserAgent As String
            Dim vBy As String
            Dim vByAlias As String
            Dim vPageRevision As Integer
            Dim vFileRevision As Integer
            Dim vFilename As String
            Dim vPos As Integer

            pFilename = Replace(pFilename, " ", "_")

            'If pHidden = "" Then
            '    pHidden = 0
            'End If

            vHost = GetRemoteHost()
            vUserAgent = HttpContext.Current.Request.ServerVariables("HTTP_USER_AGENT")
            vBy = GetRemoteUser()
            If vBy = "" Then
                vBy = vHost
            End If
            vByAlias = GetRemoteAlias()

            vQuery = "SELECT wpg_lastminor FROM openwiki_pages WHERE wpg_name = '" & Replace(gPage, "'", "''") & "'"
            vRS.Open(vQuery, vConn, ADODB.CursorTypeEnum.adOpenForwardOnly)
            If vRS.EOF Then
                vPageRevision = 1 ' page doesn't exist yet
            Else
                vPageRevision = CInt(vRS(0).Value)
            End If
            vRS.Close()
            vQuery = "SELECT MAX(att_revision) FROM openwiki_attachments WHERE att_wrv_name = '" & Replace(gPage, "'", "''") & "' AND att_name = '" & Replace(pFilename, "'", "''") & "'"
            vRS.Open(vQuery, vConn, ADODB.CursorTypeEnum.adOpenForwardOnly)
            If IsDBNull(vRS(0)) Then
                vFileRevision = 1
            Else
                vFileRevision = CInt(vRS(0).Value) + 1
            End If
            vRS.Close()

            vPos = InStrRev(pFilename, ".")
            If vPos > 0 Then
                vFilename = Left(pFilename, vPos - 1) & "-" & vFileRevision & Mid(pFilename, vPos)
            Else
                vFilename = pFilename & "-" & vFileRevision
            End If
            vFilename = SafeFileName(vFilename)

            BeginTrans(vConn)
            vRS.Open("openwiki_attachments", vConn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, adCmdTable)
            vRS.AddNew()
            vRS("att_wrv_name").Value = gPage
            vRS("att_wrv_revision").Value = vPageRevision
            vRS("att_name").Value = pFilename
            vRS("att_revision").Value = vFileRevision
            vRS("att_hidden").Value = pHidden
            vRS("att_deprecated").Value = 0
            vRS("att_filename").Value = vFilename
            vRS("att_timestamp").Value = Now()
            vRS("att_filesize").Value = pFilesize
            vRS("att_host").Value = vHost
            vRS("att_agent").Value = vUserAgent
            vRS("att_by").Value = vBy
            vRS("att_byalias").Value = vByAlias
            vRS("att_comment").Value = pComment
            vRS.Update()
            vRS.Close()

            SaveAttachmentLog(vConn, pFilename, vFileRevision, "uploaded")

            ClearDocumentCache(vConn)
            'Call ClearDocumentCache2(vConn, gPage)

            If pAddLink <> "" Then
                If OPENWIKI_DB_SYNTAX = DB_MYSQL Then
                    vQuery = "UPDATE openwiki_revisions SET wrv_text = CONCAT(wrv_text, '" & vbCrLf & vbCrLf & "  * " & Replace(pFilename, "'", "''") & "') WHERE wrv_name = '" & Replace(gPage, "'", "''") & "' AND wrv_current = 1"
                    vConn.Execute(vQuery)
                Else
                    vQuery = "SELECT wrv_text FROM openwiki_revisions WHERE wrv_name = '" & Replace(gPage, "'", "''") & "' AND wrv_current = 1"
                    vRS.Open(vQuery, vConn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, adCmdText)
                    If Not vRS.EOF Then
                        vRS("wrv_text").Value = CStr(vRS("wrv_text").Value) & vbCrLf & vbCrLf & "  * " & pFilename
                        vRS.Update()
                    End If
                    vRS.Close()
                End If
            End If

            CommitTrans(vConn)

            SaveAttachmentMetaData = vFilename
        End Function


        Sub HideAttachmentMetaData(ByVal pName As String _
            , ByVal pRevision As Integer _
            , ByVal pHide As Integer)
            BeginTrans(vConn)
            vConn.Execute("UPDATE openwiki_attachments SET att_hidden = " & pHide & " WHERE att_wrv_name = '" & Replace(gPage, "'", "''") & "' AND att_name = '" & Replace(pName, "'", "''") & "' AND att_revision = " & pRevision)
            If pHide = 1 Then
                Call SaveAttachmentLog(vConn, pName, pRevision, "hidden")
            Else
                Call SaveAttachmentLog(vConn, pName, pRevision, "made visible")
            End If
            ClearDocumentCache(vConn)
            'Call ClearDocumentCache2(vConn, gPage)
            CommitTrans(vConn)
        End Sub


        Sub TrashAttachmentMetaData(ByVal pName As String _
            , ByVal pRevision As Integer _
            , ByVal pTrash As Integer)
            BeginTrans(vConn)
            vConn.Execute("UPDATE openwiki_attachments SET att_deprecated = " & pTrash & " WHERE att_wrv_name = '" & Replace(gPage, "'", "''") & "' AND att_name = '" & Replace(pName, "'", "''") & "'")
            If pTrash = 1 Then
                Call SaveAttachmentLog(vConn, pName, pRevision, "deprecated")
            Else
                Call SaveAttachmentLog(vConn, pName, pRevision, "restored")
            End If
            Call ClearDocumentCache(vConn)
            'Call ClearDocumentCache2(vConn, gPage)
            CommitTrans(vConn)
        End Sub


        Sub SaveAttachmentLog(ByVal pConn As ADODB.Connection _
            , ByVal pName As String _
            , ByVal pFileRevision As Integer _
            , ByVal pAction As String)
            Dim vHost As String
            Dim vUserAgent As String
            Dim vBy As String
            Dim vByAlias As String
            Dim pPagename As String
            Dim pPageRevision As Integer

            vHost = GetRemoteHost()
            vUserAgent = HttpContext.Current.Request.ServerVariables("HTTP_USER_AGENT")
            vBy = GetRemoteUser()
            If vBy = "" Then
                vBy = vHost
            End If
            vByAlias = GetRemoteAlias()

            vQuery = "SELECT att_wrv_name, att_wrv_revision FROM openwiki_attachments WHERE att_wrv_name = '" & Replace(gPage, "'", "''") & "' AND att_name = '" & Replace(pName, "'", "''") & "' AND att_revision = " & pFileRevision
            vRS.Open(vQuery, pConn, ADODB.CursorTypeEnum.adOpenForwardOnly)
            If vRS.EOF Then
                vRS.Close()
                Exit Sub
            End If
            pPagename = CStr(vRS("att_wrv_name").Value)
            pPageRevision = CInt(vRS("att_wrv_revision").Value)
            vRS.Close()

            vQuery = "SELECT wrv_timestamp FROM openwiki_revisions WHERE wrv_name = '" & Replace(pPagename, "'", "''") & "' AND wrv_revision = " & pPageRevision
            vRS.Open(vQuery, pConn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, adCmdText)
            If vRS.EOF Then
                vRS.Close()
                Exit Sub
            End If
            vRS("wrv_timestamp").Value = Now()
            vRS.Update()
            vRS.Close()

            vRS.Open("openwiki_attachments_log", pConn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, adCmdTable)
            vRS.AddNew()
            vRS("ath_wrv_name").Value = pPagename
            vRS("ath_wrv_revision").Value = pPageRevision
            vRS("ath_name").Value = pName
            vRS("ath_revision").Value = pFileRevision
            vRS("ath_timestamp").Value = Now()
            vRS("ath_agent").Value = vUserAgent
            vRS("ath_by").Value = vBy
            vRS("ath_byalias").Value = vByAlias
            vRS("ath_action").Value = pAction
            vRS.Update()
            vRS.Close()
        End Sub


        ' Convert the filename to a filename with an extension that is safe
        ' to be served by the webserver.
        Function SafeFileName(ByVal pFilename As String) As String
            Dim vPos As Integer
            Dim vExtension As String

            SafeFileName = pFilename
            vPos = InStrRev(pFilename, ".")
            If vPos > 0 Then
                vExtension = Mid(pFilename, vPos + 1)
                If gNotAcceptedExtensions = "" Then
                    ' accept nothing, except the ones enumerated in gDocExtensions
                    If Not InStr("|" & gDocExtensions & "|", "|" & vExtension & "|") > 0 Then
                        SafeFileName = SafeFileName & ".safe"
                    End If
                Else
                    ' accept everything, except the ones enumerated in gNotAcceptedExtensions
                    If InStr("|" & gNotAcceptedExtensions & "|", "|" & vExtension & "|") > 0 Then
                        SafeFileName = SafeFileName & ".safe"
                    End If
                End If
            End If
        End Function


        Sub GetAttachments(ByVal pPage As WikiPage _
            , ByVal pRevision As Integer _
            , ByVal pIncludeAllChangeRecords As Boolean)
            Dim vAttachment As Attachment
            Dim vMaxRevision As Integer

            If pIncludeAllChangeRecords Then
                ' show all file revisions
                vQuery = "SELECT att_name, att_revision, att_hidden, att_deprecated, att_filename, att_timestamp, att_filesize, att_by, att_byalias, att_comment" _
                & " FROM openwiki_attachments" _
                & " WHERE att_wrv_name = '" & Replace(pPage.Name, "'", "''") & "'" _
                & " AND   att_name = '" & Replace(HttpContext.Current.Request("file"), "'", "''") & "'" _
                & " ORDER BY att_revision DESC"
            Else
                ' show last file revision relative to page revision
                vQuery = "SELECT MAX(att_wrv_revision) FROM openwiki_attachments WHERE att_wrv_name = '" & Replace(pPage.Name, "'", "''") & "'"
                If pRevision > 0 Then
                    vQuery = vQuery & " AND att_wrv_revision <= " & pRevision
                End If
                vRS.Open(vQuery, vConn, ADODB.CursorTypeEnum.adOpenForwardOnly)
                If IsDBNull(vRS(0)) Then
                    vMaxRevision = 0
                Else
                    vMaxRevision = CInt(vRS(0).Value)
                End If
                vRS.Close()
                vQuery = "SELECT att_name, att_revision, att_hidden, att_deprecated, att_filename, att_timestamp, att_filesize, att_by, att_byalias, att_comment" _
                & " FROM openwiki_attachments" _
                & " WHERE att_wrv_name = '" & Replace(pPage.Name, "'", "''") & "'" _
                & " AND   att_wrv_revision <= " & vMaxRevision _
                & " ORDER BY att_name ASC, att_revision DESC"
            End If
            vRS.Open(vQuery, vConn, ADODB.CursorTypeEnum.adOpenForwardOnly)
            Do While Not vRS.EOF
                vAttachment = New Attachment
                vAttachment.Name = CStr(vRS("att_name").Value)
                vAttachment.Revision = CInt(vRS("att_revision").Value)
                vAttachment.Hidden = CInt(vRS("att_hidden").Value)
                vAttachment.Deprecated = CInt(vRS("att_deprecated").Value)
                vAttachment.Filename = CStr(vRS("att_filename").Value)
                vAttachment.Timestamp = CDate(vRS("att_timestamp").Value)
                vAttachment.Filesize = CLng(vRS("att_filesize").Value)
                vAttachment.By = CStr(vRS("att_by").Value)

                If Not IsDBNull(vRS("wrv_byalias").Value) Then
                    vAttachment.ByAlias = CStr(vRS("att_byalias").Value)
                End If

                vAttachment.Comment = CStr(vRS("att_comment").Value)
                pPage.AddAttachment(vAttachment, Not pIncludeAllChangeRecords)
                vRS.MoveNext()
            Loop
            vRS.Close()
            pPage.AttachmentsLoaded = True
        End Sub


        ' pFilter --> 0=All, 1=NoMinorEdit, 2=OnlyMinorEdit
        Function TitleSearch(ByVal pPattern As String _
            , ByVal pDays As Integer _
            , ByVal pFilter As Integer _
            , ByVal pOrderBy As Integer _
            , ByVal pIncludeAttachmentChanges As Integer) _
        As Vector
            '            Dim vTitle As String
            Dim vList As Vector
            Dim vPage As WikiPage
            Dim vChange As Change = New Change
            Dim vCurPage As String = ""
            Dim vAttachmentChange As AttachmentChange
            Dim sTemp As String

            vList = New Vector
            'vRegEx = New Regexp
            'vRegEx.IgnoreCase = True
            'vRegEx.Global = True
            'vRegEx.Pattern = EscapePattern(pPattern)
            vQuery = "SELECT wpg_name, wpg_changes, wrv_revision, wrv_status, wrv_timestamp, wrv_minoredit, wrv_by, wrv_byalias, wrv_comment "
            If (cAllowAttachments = 1) And (pIncludeAttachmentChanges = 1) Then
                vQuery = vQuery & ", ath_name, ath_revision, ath_timestamp, ath_by, ath_byalias, ath_action "
                If OPENWIKI_DB_SYNTAX = DB_ORACLE Then
                    vQuery = vQuery & " FROM   openwiki_pages, openwiki_revisions, openwiki_attachments_log " _
                                    & " WHERE  wpg_name = wrv_name " _
                                    & " AND    wrv_name = ath_wrv_name (+) " _
                                    & " AND    wrv_revision = ath_wrv_revision (+)"
                Else
                    vQuery = vQuery & " FROM (openwiki_pages LEFT JOIN openwiki_revisions ON openwiki_pages.wpg_name = openwiki_revisions.wrv_name) LEFT JOIN openwiki_attachments_log ON (openwiki_revisions.wrv_name = openwiki_attachments_log.ath_wrv_name) AND (openwiki_revisions.wrv_revision = openwiki_attachments_log.ath_wrv_revision) WHERE 1 = 1 "
                End If
            Else
                vQuery = vQuery & "FROM openwiki_pages, openwiki_revisions " _
                                & "WHERE wrv_name = wpg_name "
            End If
            If pDays > 0 Then
                ' is there a database independent way to test the current date?
                'vQuery = vQuery & " AND wpg_timestamp >
            End If
            If pFilter = 0 Then
                vQuery = vQuery & " AND wpg_lastminor = wrv_revision"
            ElseIf pFilter = 1 Then
                vQuery = vQuery & " AND wpg_lastmajor = wrv_revision"
            ElseIf pFilter = 2 Then
                vQuery = vQuery & " AND wpg_lastminor = wrv_revision AND wrv_minoredit = 1"
            End If
            If pOrderBy = 1 Then
                vQuery = vQuery & " ORDER BY wrv_timestamp DESC"
            ElseIf pOrderBy = 2 Then
                vQuery = vQuery & " ORDER BY wrv_timestamp"
            Else
                vQuery = vQuery & " ORDER BY wpg_name"
            End If
            If (cAllowAttachments = 1) And (pIncludeAttachmentChanges = 1) Then
                vQuery = vQuery & ", ath_timestamp DESC"
            End If
            vRS.Open(vQuery, vConn, ADODB.CursorTypeEnum.adOpenForwardOnly)
            Do While Not vRS.EOF
                If Regex.IsMatch(CStr(vRS("wpg_name").Value), pPattern, RegexOptions.IgnoreCase) Then
                    If vCurPage <> CStr(vRS("wpg_name").Value) Then
                        vCurPage = CStr(vRS("wpg_name").Value)
                        vPage = New WikiPage
                        vPage.Name = CStr(vRS("wpg_name").Value)
                        vPage.Changes = CInt(vRS("wpg_changes").Value)
                        vChange = vPage.AddChange
                        vChange.Revision = CInt(vRS("wrv_revision").Value)
                        vChange.Status = CStr(vRS("wrv_status").Value)
                        vChange.MinorEdit = CInt(vRS("wrv_minoredit").Value)
                        vChange.Timestamp = CDate(vRS("wrv_timestamp").Value)
                        vChange.By = CStr(vRS("wrv_by").Value)

                        If Not IsDBNull(vRS("wrv_byalias").Value) Then
                            vChange.ByAlias = CStr(vRS("wrv_byalias").Value)
                        End If

                        If Not Convert.ToString(vRS("wrv_comment").Value) = Convert.ToString(DBNull.Value) Then
                            vChange.Comment = CStr(vRS("wrv_comment").Value)
                        End If

                        vList.Push(vPage)
                    End If
                    If (cAllowAttachments = 1) And (pIncludeAttachmentChanges = 1) Then
                        If (CStr(vRS("ath_name").Value) <> "") And (CDate(vRS("ath_timestamp").Value) > DateAdd("h", -24, Now())) Then
                            vAttachmentChange = New AttachmentChange
                            vAttachmentChange.Name = CStr(vRS("ath_name").Value)
                            vAttachmentChange.Revision = CInt(vRS("ath_revision").Value)
                            vAttachmentChange.Timestamp = CDate(vRS("ath_timestamp").Value)
                            vAttachmentChange.By = CStr(vRS("ath_by").Value)

                            If Not IsDBNull(vRS("wrv_byalias").Value) Then
                                vAttachmentChange.ByAlias = CStr(vRS("ath_byalias").Value)
                            End If

                            vAttachmentChange.Action = CStr(vRS("ath_action").Value)
                            vChange.AddAttachmentChange(vAttachmentChange)
                        End If
                    End If
                End If
                vRS.MoveNext()
            Loop
            vRS.Close()
            TitleSearch = vList
        End Function


        Function FullSearch(ByVal pPattern As String _
            , ByVal pIncludeTitles As Integer) _
        As Vector
            '            Dim vTitle As String
            Dim vList As Vector
            Dim vPage As WikiPage
            Dim vChange As Change
            Dim vFound As Boolean
            Dim vPattern As String
            Dim vPattern2 As String = ""

            pPattern = EscapePattern(pPattern)
            vList = New Vector
            'vRegEx = New Regexp
            'vRegEx.IgnoreCase = True
            'vRegEx.Global = True
            If HttpContext.Current.Request("fromtitle") = "true" Then
                vPattern = Replace(pPattern, "_", " ")
            Else
                vPattern = pPattern
            End If
            If (pIncludeTitles = 1) Then
                'vRegEx2 = New Regexp
                'vRegEx2.IgnoreCase = True
                'vRegEx2.Global = True
                'vRegEx2.Pattern = pPattern
                vPattern2 = pPattern
            End If
            vQuery = "SELECT * FROM openwiki_pages, openwiki_revisions WHERE wrv_name = wpg_name AND wrv_current = 1 AND wrv_text IS NOT NULL ORDER BY wpg_name"
            vRS.Open(vQuery, vConn, ADODB.CursorTypeEnum.adOpenForwardOnly)
            Do While Not vRS.EOF
                vFound = False
                If (pIncludeTitles = 1) Then
                    If Regex.IsMatch(CStr(vRS("wpg_name").Value), vPattern2, RegexOptions.IgnoreCase) Then
                        vFound = True
                    End If
                End If
                If Not vFound Then
                    If Regex.IsMatch(CStr(vRS("wrv_text").Value), vPattern, RegexOptions.IgnoreCase) Then
                        vFound = True
                    End If
                End If
                If vFound Then
                    vPage = New WikiPage
                    vPage.Name = CStr(vRS("wpg_name").Value)
                    vPage.Changes = CInt(vRS("wpg_changes").Value)
                    vChange = vPage.AddChange
                    vChange.Revision = CInt(vRS("wrv_revision").Value)
                    vChange.Status = CStr(vRS("wrv_status").Value)
                    vChange.MinorEdit = CInt(vRS("wrv_minoredit").Value)
                    vChange.Timestamp = CDate(vRS("wrv_timestamp").Value)
                    vChange.By = CStr(vRS("wrv_by").Value)

                    If Not IsDBNull(vRS("wrv_byalias").Value) Then
                        vChange.ByAlias = CStr(vRS("wrv_byalias").Value)
                    End If

                    If Not Convert.ToString(vRS("wrv_comment").Value) = Convert.ToString(DBNull.Value) Then
                        vChange.Comment = CStr(vRS("wrv_comment").Value)
                    End If

                    vList.Push(vPage)
                End If
                vRS.MoveNext()
            Loop
            vRS.Close()
            FullSearch = vList
        End Function

        Function EquationSearch(ByVal pPattern As String _
            , ByVal pIncludeTitles As Integer _
            , ByVal pOrderBy As Integer) _
        As Vector
            '            Dim vTitle As String
            Dim vList As Vector
            Dim vPage As WikiPage
            Dim vChange As Change
            Dim vFound As Boolean
            Dim vText As String
            Dim vPattern As String
            Dim vPattern2 As String = ""

            pPattern = EscapePattern(pPattern)
            vList = New Vector
            'vRegEx = New Regexp
            'vRegEx.IgnoreCase = True
            'vRegEx.Global = True
            If HttpContext.Current.Request("fromtitle") = "true" Then
                vPattern = Replace(pPattern, "_", " ")
            Else
                vPattern = pPattern
            End If
            If (pIncludeTitles = 1) Then
                'vRegEx2 = New Regexp
                'vRegEx2.IgnoreCase = True
                'vRegEx2.Global = True
                'vRegEx2.Pattern = pPattern
                vPattern2 = pPattern
            End If
            vQuery = "SELECT * FROM openwiki_pages, openwiki_revisions WHERE wrv_name = wpg_name AND wrv_current = 1 AND wrv_text IS NOT NULL"
            If pOrderBy = 1 Then
                vQuery = vQuery & " ORDER BY wrv_timestamp DESC"
            ElseIf pOrderBy = 2 Then
                vQuery = vQuery & " ORDER BY wrv_timestamp"
            Else
                vQuery = vQuery & " ORDER BY wpg_name"
            End If
            vRS.Open(vQuery, vConn, ADODB.CursorTypeEnum.adOpenForwardOnly)
            Do While Not vRS.EOF
                vText = CStr(vRS("wrv_text").Value)
                vFound = False
                If (pIncludeTitles = 1) Then
                    If Regex.IsMatch(CStr(vRS("wpg_name").Value), vPattern2, RegexOptions.IgnoreCase) Then
                        vFound = True
                    End If
                End If
                If Not vFound Then
                    If Regex.IsMatch(vText, vPattern, RegexOptions.IgnoreCase) Then
                        vFound = True
                    End If
                End If
                If vFound And (pIncludeTitles = 1) And (cUseSpecialPagesPrefix = 1) And m(CStr(vRS("wpg_name").Value), "^" & gSpecialPagesPrefix, False, False) Then
                    vFound = False
                End If
                If vFound Then
                    vPage = New WikiPage
                    vPage.Name = CStr(vRS("wpg_name").Value)
                    s(vText, "<math>([\s\S]*?)<\/math>", "&CutEquation($1)", False, False)
                    vPage.Text = gEquation
                    vPage.Changes = CInt(vRS("wpg_changes").Value)
                    vChange = vPage.AddChange
                    vChange.Revision = CInt(vRS("wrv_revision").Value)
                    vChange.Status = CStr(vRS("wrv_status").Value)
                    vChange.MinorEdit = CInt(vRS("wrv_minoredit").Value)
                    vChange.Timestamp = CDate(vRS("wrv_timestamp").Value)
                    vChange.By = CStr(vRS("wrv_by").Value)

                    If Not IsDBNull(vRS("wrv_byalias").Value) Then
                        vChange.ByAlias = CStr(vRS("wrv_byalias").Value)
                    End If

                    If Not Convert.ToString(vRS("wrv_comment").Value) = Convert.ToString(DBNull.Value) Then
                        vChange.Comment = CStr(vRS("wrv_comment").Value)
                    End If

                    vList.Push(vPage)
                End If
                vRS.MoveNext()
            Loop
            vRS.Close()
            EquationSearch = vList
        End Function

        Function GetPreviousRevision(ByVal pDiffType As Integer _
            , ByVal pDiffTo As Integer) _
        As Integer
            Dim vBy As String = ""
            Dim vHost As String = ""
            Dim vAgent As String = ""

            GetPreviousRevision = 0
            If pDiffTo <= 0 Then
                pDiffTo = 99999999
            End If
            vQuery = "SELECT wrv_revision, wrv_minoredit, wrv_by, wrv_host, wrv_agent FROM openwiki_revisions WHERE wrv_name = '" & Replace(gPage, "'", "''") & "' AND wrv_revision <= " & pDiffTo
            vQuery = vQuery & " ORDER BY wrv_revision DESC"
            vRS.Open(vQuery, vConn, ADODB.CursorTypeEnum.adOpenForwardOnly)
            If Not vRS.EOF Then
                vBy = CStr(vRS("wrv_by").Value)
                vHost = CStr(vRS("wrv_host").Value)
                vAgent = CStr(vRS("wrv_agent").Value)
            End If
            Do While Not vRS.EOF
                GetPreviousRevision = CInt(vRS("wrv_revision").Value)
                If pDiffType = 0 Then
                    ' previous major
                    If CInt(vRS("wrv_minoredit").Value) = 0 Then
                        vRS.MoveNext()
                        If Not vRS.EOF Then
                            GetPreviousRevision = CInt(vRS("wrv_revision").Value)
                        End If
                        Exit Do
                    End If
                ElseIf pDiffType = 1 Then
                    ' previous minor
                    vRS.MoveNext()
                    If Not vRS.EOF Then
                        GetPreviousRevision = CInt(vRS("wrv_revision").Value)
                    End If
                    Exit Do
                Else
                    ' previous author
                    If CStr(vRS("wrv_by").Value) <> vBy Or CStr(vRS("wrv_host").Value) <> vHost Or CStr(vRS("wrv_agent").Value) <> vAgent Then
                        Exit Do
                    End If
                End If
                vRS.MoveNext()
            Loop
            vRS.Close()
        End Function


        Function InterWiki() As String
            Dim vTemp As String = ""

            vQuery = "SELECT wik_name, wik_url FROM openwiki_interwikis ORDER BY wik_name"
            vRS.Open(vQuery, vConn, ADODB.CursorTypeEnum.adOpenForwardOnly)
            Do While Not vRS.EOF
                vTemp = vTemp & "<ow:interlink>" _
                 & "<ow:name>" & PCDATAEncode(CStr(vRS("wrv_name").Value)) & "</ow:name>" _
                 & "<ow:href>" & CDATAEncode(CStr(vRS("wrv_url").Value)) & "</ow:href>" _
                 & "<ow:class>" & CDATAEncode(LCase(Trim(CStr(vRS("wrv_name").Value)))) & "</ow:class>" _
                 & "</ow:interlink>"
                vRS.MoveNext()
            Loop
            vRS.Close()
            InterWiki = "<ow:interlinks>" & vTemp & "</ow:interlinks>"
        End Function

        Function ListRedirects() As String
            Dim vTemp As String = ""
            '            Dim vTempPage
            Dim vText As String
            Dim vPageFrom As String
            Dim vPageTo As String
            Dim vPos As Integer
            Dim vBuffFrom As Vector
            Dim vBuffTo As Vector
            Dim i As Integer

            vBuffFrom = New Vector
            vBuffTo = New Vector

            If OPENWIKI_DB_SYNTAX = DB_ACCESS Then
                vQuery = "SELECT * FROM openwiki_pages, openwiki_revisions " _
                 & "WHERE wrv_name = wpg_name AND wrv_current = 1 AND wrv_text LIKE '[#]REDIRECT %' " _
                 & "ORDER BY wpg_name"
            Else
                vQuery = "SELECT * FROM openwiki_pages, openwiki_revisions " _
          & "WHERE wrv_name = wpg_name AND wrv_current = 1 AND wrv_text LIKE '\#REDIRECT %' ESCAPE '\' " _
          & "ORDER BY wpg_name"
            End If
            vRS.Open(vQuery, vConn, ADODB.CursorTypeEnum.adOpenForwardOnly)
            Do While Not vRS.EOF
                vText = CStr(vRS("wrv_text").Value)
                vPageFrom = CStr(vRS("wpg_name").Value)
                If m(vText, "^#REDIRECT\s+", False, False) Then
                    vPos = InStr(Len("#REDIRECT "), vText, vbCr)
                    If vPos > 0 Then
                        vPageTo = Trim(Mid(vText, Len("#REDIRECT "), vPos - Len("#REDIRECT ")))
                    Else
                        vPageTo = Trim(Mid(vText, Len("#REDIRECT ")))
                    End If
                    vBuffFrom.Push(vPageFrom)
                    vBuffTo.Push(vPageTo)
                End If
                vRS.MoveNext()
            Loop
            vRS.Close()

            If Not vBuffFrom.IsEmpty() Then
                For i = 0 To vBuffFrom.Count - 1
                    vPageFrom = CStr(vBuffFrom.ElementAt(i))
                    vPageTo = CStr(vBuffTo.ElementAt(i))
                    vTemp = vTemp & "<ow:redirect>" _
                     & "<ow:from>" & GetWikiLink("", vPageFrom, "") & "</ow:from>" _
                     & "<ow:to>" & GetWikiLink("", vPageTo, "") & "</ow:to>" _
                     & "</ow:redirect>"
                Next
                ListRedirects = "<ow:redirectlinks>" & vTemp & "</ow:redirectlinks>"
            End If

            vBuffFrom = Nothing
            vBuffTo = Nothing
        End Function

        Function GetInterWiki(ByVal pName As String) As String
            If OPENWIKI_DB <> "" Then
                If pName = "This" Then
                    GetInterWiki = gScriptName & "?p="
                Else
                    vQuery = "SELECT wik_url FROM openwiki_interwikis WHERE wik_name = '" & Replace(pName, "'", "''") & "'"
                    vRS.Open(vQuery, vConn, ADODB.CursorTypeEnum.adOpenForwardOnly)
                    If Not vRS.EOF Then
                        GetInterWiki = CStr(vRS("wik_url").Value)
                    End If
                    vRS.Close()
                End If
            End If
        End Function


        Function GetRSSFromCache(ByVal pURL As String _
            , ByVal pRefreshRate As Integer _
            , ByVal pFreshlyFromRemoteSite As Boolean _
            , ByVal pRetryLater As Boolean) _
        As String
            Dim conn As ADODB.Connection = New ADODB.Connection
            Dim vRS As ADODB.Recordset = New ADODB.Recordset
            Dim vLast As Date
            Dim vNext As Date
            Dim vRefreshRate As Integer

            conn.Open(OPENWIKI_DB)
            vQuery = "SELECT rss_last, rss_next, rss_refreshrate, rss_cache FROM openwiki_rss WHERE rss_url = '" & Replace(pURL, "'", "''") & "'"
            vRS.Open(vQuery, conn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, adCmdText)
            If vRS.EOF Then
                GetRSSFromCache = "notexists"
            Else
                vLast = CDate(vRS("rss_last").Value)
                vNext = CDate(vRS("rss_next").Value)
                vRefreshRate = CInt(vRS("rss_refreshrate").Value)
                If vRefreshRate <> pRefreshRate Then
                    vNext = DateAdd("n", pRefreshRate, vLast)
                    vRS("rss_next").Value = vNext
                    vRS("rss_refreshrate").Value = pRefreshRate
                    vRS.Update()
                ElseIf pRetryLater Then
                    ' retry a minute from now
                    vNext = DateAdd("n", 1, Now())
                    vRS("rss_next").Value = vNext
                    vRS.Update()
                End If

                If pFreshlyFromRemoteSite Or (DateDiff("n", vNext, Now()) < 0) Then
                    GetRSSFromCache = "<ow:feed href='" & Replace(pURL, "&", "&amp;") & "' "
                    If pFreshlyFromRemoteSite Then
                        GetRSSFromCache = GetRSSFromCache & "fresh='true' "
                    Else
                        GetRSSFromCache = GetRSSFromCache & "fresh='false' "
                    End If
                    GetRSSFromCache = GetRSSFromCache & "last='" & FormatDateISO8601(vLast) & "' "
                    GetRSSFromCache = GetRSSFromCache & "next='" & FormatDateISO8601(vNext) & "' "
                    GetRSSFromCache = GetRSSFromCache & "refreshrate='" & pRefreshRate & "'>"
                    GetRSSFromCache = GetRSSFromCache & CStr(vRS("rss_cache").Value)
                    GetRSSFromCache = GetRSSFromCache & "</ow:feed>"
                End If

            End If
            vRS.Close()
            conn.Close()
            vRS = Nothing
            conn = Nothing
        End Function

        Sub SaveRSSToCache(ByVal pURL As String _
            , ByVal pRefreshRate As Integer _
            , ByVal pCache As String)
            Dim conn As ADODB.Connection = New ADODB.Connection
            Dim vRS As ADODB.Recordset = New ADODB.Recordset

            conn.Open(OPENWIKI_DB)
            vQuery = "SELECT * FROM openwiki_rss WHERE rss_url = '" & Replace(pURL, "'", "''") & "'"

            vRS.Open(vQuery, conn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, adCmdText)
            If vRS.EOF Then
                vRS.Close()
                vRS.Open("openwiki_rss", conn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, adCmdTable)
                vRS.AddNew()
                vRS("rss_url").Value = pURL
            End If
            vRS("rss_last").Value = Now()
            If pCache = "" Then
                vRS("rss_next").Value = DateAdd("n", 30, Now())   ' 30 minutes from now
            Else
                vRS("rss_next").Value = DateAdd("n", pRefreshRate, Now())
            End If
            vRS("rss_refreshrate").Value = pRefreshRate
            vRS("rss_cache").Value = pCache
            vRS.Update()
            vRS.Close()
            conn.Close()
            vRS = Nothing
            conn = Nothing
        End Sub

        Sub Aggregate(ByVal pURL As String _
            , ByVal pXmlDoc As MSXML2.FreeThreadedDOMDocument60)
            Dim conn As ADODB.Connection = New ADODB.Connection
            Dim vRS As ADODB.Recordset = New ADODB.Recordset
            Dim vRoot As MSXML2.IXMLDOMElement
            Dim vItems As MSXML2.IXMLDOMNodeList
            Dim vItem As MSXML2.IXMLDOMElement
            Dim vXmlIsland As String
            Dim vAgXmlIsland As String
            Dim vNow As Date
            Dim i As Integer
            Dim vRssLink As String
            Dim vRdfResource As String
            Dim vRdfTimestamp As String
            Dim vDcDate As String

            '            On Error Resume Next
            'Response.Write("<p />Aggregating " & pURL & "<br />")

            vRoot = pXmlDoc.documentElement

            If vRoot.nodeName = "rss" Then
                vItems = vRoot.selectNodes("channel/item")
            ElseIf CStr(vRoot.getAttribute("xmlns")) = "http://my.netscape.com/rdf/simple/0.9/" Then
                vItems = vRoot.selectNodes("item")
            ElseIf CStr(vRoot.getAttribute("xmlns")) = "http://purl.org/rss/1.0/" Then
                vItems = vRoot.selectNodes("item")
            Else
                Exit Sub
            End If

            vNow = Now()
            i = 0

            ' TODO: find workaround for bug in MSXML v4
            If Not vRoot.selectSingleNode("channel/wiki:interwiki") Is Nothing Then
                vAgXmlIsland = "<ag:source><rdf:Description wiki:interwiki=""" & vRoot.selectSingleNode("channel/wiki:interwiki").text & """><rdf:value>" & PCDATAEncode(vRoot.selectSingleNode("channel/title").text) & "</rdf:value></rdf:Description></ag:source>"
            Else
                vAgXmlIsland = "<ag:source>" & PCDATAEncode(vRoot.selectSingleNode("channel/title").text) & "</ag:source>"
            End If
            vAgXmlIsland = vAgXmlIsland & "<ag:sourceURL>" & PCDATAEncode(vRoot.selectSingleNode("channel/link").text) & "</ag:sourceURL>"

            '            conn = HttpContext.Current.Server.CreateObject("ADODB.Connection")
            conn.Open(OPENWIKI_DB)
            'vRS = HttpContext.Current.Server.CreateObject("ADODB.Recordset")

            ' walk trough all item elements and store them in the openwiki_rss_aggregations table
            For Each vItem In vItems
                vRssLink = vItem.selectSingleNode("link").text

                vRdfResource = CStr(vItem.getAttribute("rdf:about"))
                If IsNothing(vRdfResource) Then
                    vRdfResource = vRssLink
                End If

                If vItem.selectSingleNode("ag:timestamp") Is Nothing Then
                    vRdfTimestamp = CStr(DateAdd("s", i, vNow))
                Else
                    vRdfTimestamp = vItem.selectSingleNode("ag:timestamp").text
                    s(vRdfTimestamp, gTimestampPattern, "&ToDateTime($1,$2,$3,$4,$5,$6,$7,$8,$9)", False, False)
                    If DateDiff("d", vNow, sReturn) > 1 Then
                        ' we cannot take this date serious, it's too far in the future
                        vRdfTimestamp = CStr(DateAdd("s", i, vNow))
                    Else
                        vRdfTimestamp = sReturn
                        vAgXmlIsland = vItem.selectSingleNode("ag:source").xml & vItem.selectSingleNode("ag:sourceURL").xml
                    End If
                End If
                i = i - 1

                vXmlIsland = "<title>" & PCDATAEncode(vItem.selectSingleNode("title").text) & "</title><link>" & PCDATAEncode(vItem.selectSingleNode("link").text) & "</link>"
                If Not vItem.selectSingleNode("description") Is Nothing Then
                    vXmlIsland = vXmlIsland & "<description>" & PCDATAEncode(vItem.selectSingleNode("description").text) & "</description>"
                End If
                If Not vItem.selectSingleNode("dc:creator") Is Nothing Then
                    vXmlIsland = vXmlIsland & vItem.selectSingleNode("dc:creator").xml
                End If
                If Not vItem.selectSingleNode("dc:contributor") Is Nothing Then
                    vXmlIsland = vXmlIsland & vItem.selectSingleNode("dc:contributor").xml
                End If
                If vItem.selectSingleNode("dc:date") Is Nothing Then
                    vDcDate = ""
                Else
                    vDcDate = vItem.selectSingleNode("dc:date").text
                    vXmlIsland = vXmlIsland & "<dc:date>" & vItem.selectSingleNode("dc:date").text & "</dc:date>"
                End If
                If Not vItem.selectSingleNode("wiki:version") Is Nothing Then
                    vXmlIsland = vXmlIsland & "<wiki:version>" & vItem.selectSingleNode("wiki:version").text & "</wiki:version>"
                End If
                If Not vItem.selectSingleNode("wiki:status") Is Nothing Then
                    vXmlIsland = vXmlIsland & "<wiki:status>" & vItem.selectSingleNode("wiki:status").text & "</wiki:status>"
                End If
                If Not vItem.selectSingleNode("wiki:importance") Is Nothing Then
                    vXmlIsland = vXmlIsland & "<wiki:importance>" & vItem.selectSingleNode("wiki:importance").text & "</wiki:importance>"
                End If
                If Not vItem.selectSingleNode("wiki:diff") Is Nothing Then
                    vXmlIsland = vXmlIsland & vItem.selectSingleNode("wiki:diff").xml
                End If
                If Not vItem.selectSingleNode("wiki:history") Is Nothing Then
                    vXmlIsland = vXmlIsland & vItem.selectSingleNode("wiki:history").xml
                End If
                vXmlIsland = vXmlIsland & vAgXmlIsland & "<ag:timestamp>" _
                    & FormatDateISO8601(CDate(vRdfTimestamp)) & "</ag:timestamp>"

                vXmlIsland = "<item rdf:about='" & PCDATAEncode(vRdfResource) & "'>" & vXmlIsland & "</item>"

                ' TODO: erm... this is actually inefficient.. use better ADO techniques
                vQuery = "SELECT * FROM openwiki_rss_aggregations WHERE agr_feed='" & Replace(pURL, "'", "''") & "' AND agr_rsslink = '" & Replace(vRssLink, "'", "''") & "'"
                vRS.Open(vQuery, conn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, adCmdText)
                If vRS.EOF Then
                    vRS.Close()
                    vRS.Open("openwiki_rss_aggregations", conn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, adCmdTable)
                    vRS.AddNew()
                    vRS("agr_feed").Value = pURL
                    vRS("agr_resource").Value = vRdfResource
                    vRS("agr_rsslink").Value = vRssLink
                    vRS("agr_timestamp").Value = vRdfTimestamp
                    vRS("agr_dcdate").Value = vDcDate
                    vRS("agr_xmlisland").Value = vXmlIsland
                    vRS.Update()
                ElseIf CStr(vRS("agr_dcdate").Value) <> vDcDate Then
                    vRS("agr_resource").Value = vRdfResource
                    vRS("agr_timestamp").Value = vRdfTimestamp
                    vRS("agr_dcdate").Value = vDcDate
                    vRS("agr_xmlisland").Value = vXmlIsland
                    vRS.Update()
                End If
                vRS.Close()
            Next

            conn.Close()
            vRS = Nothing
            conn = Nothing

            'Response.Write("<p />Done aggregating " & pURL & "<br />")
        End Sub

        Function GetAggregation(ByVal pURLs As Vector) As String
            Dim vRdfSeq As String = ""
            Dim vItems As String = ""
            Dim vTemp As String
            Dim i As Integer

            vQuery = ""
            Do While Not pURLs.IsEmpty
                vQuery = vQuery & "'" & Replace(CStr(pURLs.Pop()), "'", "''") & "'"
                If pURLs.Count > 0 Then
                    vQuery = vQuery & ","
                End If
            Loop
            vQuery = "SELECT * FROM openwiki_rss_aggregations WHERE agr_feed IN (" & vQuery & ") ORDER BY agr_timestamp DESC"
            vRS.Open(vQuery, vConn, ADODB.CursorTypeEnum.adOpenForwardOnly)
            i = 0
            If OPENWIKI_MAXNROFAGGR <= 0 Then
                OPENWIKI_MAXNROFAGGR = 100
            End If
            Do While Not vRS.EOF
                i = i + 1
                If i > OPENWIKI_MAXNROFAGGR Then
                    Exit Do
                End If
                vTemp = CDATAEncode(CStr(vRS("agr_resource").Value))
                vRdfSeq = vRdfSeq & "<rdf:li rdf:resource='" & vTemp & "'/>"
                vItems = vItems & CStr(vRS("agr_xmlisland").Value)
                vRS.MoveNext()
            Loop
            vRS.Close()
            GetAggregation = "<?xml version='1.0' encoding='ISO-8859-1'?>" & vbCrLf _
                          & "<!-- All Your Wiki Are Belong To Us -->" & vbCrLf _
                          & "<rdf:RDF xmlns='http://purl.org/rss/1.0/' xmlns:rdf='http://www.w3.org/1999/02/22-rdf-syntax-ns#' xmlns:dc='http://purl.org/dc/elements/1.1/' xmlns:wiki='http://purl.org/rss/1.0/modules/wiki/' xmlns:ag='http://purl.org/rss/1.0/modules/aggregation/'>" _
                          & "<channel rdf:about='" & CDATAEncode(gServerRoot & gScriptName & "?p=" & gPage & "&a=rss") & "'>" _
                          & "<title>" & PCDATAEncode(OPENWIKI_TITLE & " -- " & PrettyWikiLink(gPage)) & "</title>" _
                          & "<link>" & PCDATAEncode(gServerRoot & gScriptName & "?" & gPage) & "</link>" _
                          & "<description>" & PCDATAEncode(OPENWIKI_TITLE & " -- " & PrettyWikiLink(gPage)) & "</description>" _
                          & "<image rdf:about='" & CDATAEncode(gServerRoot & "ow/images/aggregator.gif") & "'/>" _
                          & "<items><rdf:Seq>" _
                          & vRdfSeq _
                          & "</rdf:Seq></items>" _
                          & "</channel>" _
                          & "<image rdf:about='" & CDATAEncode(gServerRoot & "ow/images/aggregator.gif") & "'>" _
                          & "<title>" & PCDATAEncode(OPENWIKI_TITLE) & "</title>" _
                          & "<link>" & CDATAEncode(gServerRoot & gScriptName & "?p=" & gPage) & "</link>" _
                          & "<url>" & PCDATAEncode(gServerRoot & "ow/images/logo_aggregator.gif") & "</url>" _
                          & "</image>" _
                          & vItems _
                          & "</rdf:RDF>"
        End Function


        Private Function CreateDocKey(ByVal pSubKey As String) As String
            CreateDocKey = pSubKey & gFS _
                         & gCookieHash & gFS _
                         & gRevision & gFS _
                         & HttpContext.Current.Request.Cookies(gCookieHash & "?up")("pwl") & gFS _
                         & HttpContext.Current.Request.Cookies(gCookieHash & "?up")("new") & gFS _
                         & HttpContext.Current.Request.Cookies(gCookieHash & "?up")("emo")
            CreateDocKey = CStr(Hash(CStr(CreateDocKey)))
        End Function

        Function GetDocumentCache(ByVal pSubKey As String) As String
            vQuery = "SELECT chc_xmlisland FROM openwiki_cache WHERE chc_name = '" & Replace(gPage, "'", "''") & "' AND chc_hash = " & CreateDocKey(pSubKey)
            vRS.Open(vQuery, vConn, ADODB.CursorTypeEnum.adOpenForwardOnly)
            If vRS.EOF Then
                GetDocumentCache = ""
            Else
                GetDocumentCache = CStr(vRS("chc_xmlisland").Value)
            End If
            vRS.Close()
        End Function

        Sub SetDocumentCache(ByVal pSubKey As String, ByVal pXmlStr As String)
            Dim vKey As String
            vKey = CreateDocKey(pSubKey)
            vQuery = "SELECT chc_xmlisland FROM openwiki_cache WHERE chc_name = '" & Replace(gPage, "'", "''") & "' AND chc_hash = " & vKey
            vRS.Open(vQuery, vConn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, adCmdText)
            If vRS.EOF Then
                vRS.Close()
                vRS.Open("openwiki_cache", vConn, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, adCmdTable)
                vRS.AddNew()
                vRS("chc_name").Value = gPage
                vRS("chc_hash").Value = vKey
            End If
            vRS("chc_xmlisland").Value = pXmlStr
            vRS.Update()
            vRS.Close()
        End Sub

        Sub ClearDocumentCache(ByVal pConn As ADODB.Connection)
            pConn.Execute("DELETE FROM openwiki_cache")
        End Sub

        Sub ClearDocumentCache2(ByVal pConn As ADODB.Connection, ByVal pPagename As String)
            If pConn Is Nothing Then
                pConn = vConn
            End If
            pConn.Execute("DELETE FROM openwiki_cache WHERE chc_name = '" & Replace(pPagename, "'", "''") & "'")
        End Sub

    End Class
End Namespace