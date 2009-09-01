Namespace Openwiki
    Module Attach
        Sub ActionAttach()
            ActionView()
        End Sub


        Sub ActionUpload()
            'HttpContext.Current.Response.Expires = -10000
            'HttpContext.Current.Server.ScriptTimeout = OPENWIKI_UPLOADTIMEOUT
            'Dim theForm, theField, vFilename
            'On Error Resume Next
            'Err.Number = 0
            'theForm = HttpContext.Current.Server.CreateObject("ABCUpload4.XForm")
            'If Err.Number <> 0 Then
            '    HttpContext.Current.Response.Write("<b>Error</b>: Missing component ABCUpload4. You can download this component from <a href='http://www.websupergoo.com/downloadftp.htm'>websupergoo.com</a>")
            '    HttpContext.Current.Response.End()
            'End If
            'On Error GoTo 0
            'theForm.MaxUploadSize = OPENWIKI_MAXUPLOADSIZE
            'theForm.Overwrite = True
            'theForm.AbsolutePath = False
            '' TODO: maybe implement pop-up progress-bar
            ''theForm.ID = HttpContext.Current.Request.QueryString("ID")

            ''On Error Resume Next
            'theField = theForm("file")(1)
            'If theField.FileExists Then
            '    ' If you want to store your files as BLOBs in the database, then you should
            '    ' comment the next line
            '    CreateFolders()

            '    vFilename = theField.SafeFileName
            '    vFilename = gNamespace.SaveAttachmentMetaData(vFilename, theField.Length, theForm("link"), theForm("hide"), theForm("comment"))

            '    ' Save to filesystem.
            '    theField.Save(OPENWIKI_UPLOADDIR & gPage & "/" & vFilename)
            'End If
            'HttpContext.Current.Response.Redirect(gScriptName & "?p=" & HttpContext.Current.Server.UrlEncode(gPage) & "&a=attach")
            'HttpContext.Current.Response.End()
        End Sub


        Sub ActionHidefile()
            'gNamespace.HideAttachmentMetaData(HttpContext.Current.Request("file"), HttpContext.Current.Request("rev"), 1)
            'HttpContext.Current.Response.Redirect(gScriptName & "?p=" & HttpContext.Current.Server.UrlEncode(gPage) & "&a=attach")
            'HttpContext.Current.Response.End()
        End Sub

        Sub ActionUndohidefile()
            'gNamespace.HideAttachmentMetaData(HttpContext.Current.Request("file"), HttpContext.Current.Request("rev"), 0)
            'HttpContext.Current.Response.Redirect(gScriptName & "?p=" & HttpContext.Current.Server.UrlEncode(gPage) & "&a=attach")
            'HttpContext.Current.Response.End()
        End Sub

        Sub ActionTrashfile()
            'gNamespace.TrashAttachmentMetaData(HttpContext.Current.Request("file"), HttpContext.Current.Request("rev"), 1)
            'HttpContext.Current.Response.Redirect(gScriptName & "?p=" & HttpContext.Current.Server.UrlEncode(gPage) & "&a=attach")
            'HttpContext.Current.Response.End()
        End Sub

        Sub ActionUndotrashfile()
            'gNamespace.TrashAttachmentMetaData(HttpContext.Current.Request("file"), HttpContext.Current.Request("rev"), 0)
            'HttpContext.Current.Response.Redirect(gScriptName & "?p=" & HttpContext.Current.Server.UrlEncode(gPage) & "&a=attach")
            'HttpContext.Current.Response.End()
        End Sub

        Sub ActionAttachchanges()
            'gTransformer.Transform(gNamespace.GetPageAndAttachments(gPage, 0, False, True).ToXML(0))
            'gActionReturn = True
        End Sub



        ' If you want to store your files as BLOBs in the database, then you'd need
        ' to change this function.
        '
        ' IN:
        '   pPagename : page that has the attachment
        '   pFilename : filename of the attachment
        ' RETURN: full URL to view/download the attachment
        Function GetAttachmentLink(ByVal pPagename As String, ByVal pFilename As String) As String
            GetAttachmentLink = gServerRoot & OPENWIKI_UPLOADDIR & pPagename & "/" & pFilename
        End Function


        ' Create all the subfolders if they do not exist yet.
        Sub CreateFolders()
            'Dim vFSO As Scripting.FileSystemObject, vPosBegin As Integer, vPosEnd As Integer, vPath As String
            'vFSO = HttpContext.Current.Server.CreateObject("Scripting.FileSystemObject")
            'If Not vFSO.FolderExists(HttpContext.Current.Server.MapPath(OPENWIKI_UPLOADDIR & gPage & "/")) Then
            '    vPosBegin = 1
            '    vPath = HttpContext.Current.Server.MapPath(OPENWIKI_UPLOADDIR)
            '    Do While True
            '        vPosEnd = InStr(vPosBegin, gPage, "/")
            '        If vPosEnd > vPosBegin Then
            '            vPath = vPath & "\" & Mid(gPage, vPosBegin, vPosEnd - vPosBegin)
            '            If Not vFSO.FolderExists(vPath) Then
            '                Call vFSO.CreateFolder(vPath)
            '            End If
            '            vPosBegin = vPosEnd + 1
            '        Else
            '            vPath = vPath & "\" & Mid(gPage, vPosBegin)
            '            If Not vFSO.FolderExists(vPath) Then
            '                Call vFSO.CreateFolder(vPath)
            '            End If
            '            Exit Do
            '        End If
            '    Loop
            'End If
        End Sub
    End Module
End Namespace