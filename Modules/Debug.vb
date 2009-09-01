Namespace Openwiki
    Module Debug
        Sub WriteDebug(ByVal pText As String, ByVal pParam As String, ByVal pDebugLevel As Integer)
            'Dim LogFile, FileExists, objPI, objRoot, objRecord, objField, objAttr

            'If OPENWIKI_DEBUGLEVEL >= pDebugLevel Then
            '    '		Response.Write("in panic<br>")
            '    'Instantiate the Microsoft XMLDOM
            '    LogFile = HttpContext.Current.Server.CreateObject("Msxml2.DOMDocument.6.0")
            '    LogFile.PreserveWhiteSpace = True

            '    FileExists = LogFile.Load(OPENWIKI_DEBUGPATH)
            '    '		Response.Write("FileExists " & FileExists & "<br>")

            '    If FileExists = True Then
            '        '			Response.Write("file exists<br>")
            '        'If the file loaded set the objRoot Object equal to the root element
            '        'of the XML document
            '        objRoot = LogFile.DocumentElement
            '    Else
            '        '			Response.Write("file not exists<br>")
            '        '			Response.Write(LogFile.parseError.errorCode)

            '        'Create root element and append it to the XML document
            '        objRoot = LogFile.createElement("ow:panic")

            '        'Declare default namespace
            '        objAttr = LogFile.CreateAttribute("xmlns")
            '        objAttr.Text = "http://www.w3.org/1999/xhtml"
            '        objRoot.SetAttributeNode(objAttr)

            '        'Declare OpenWiki namespace
            '        objAttr = LogFile.CreateAttribute("xmlns:ow")
            '        objAttr.Text = "http://openwiki.com/2001/OW/Wiki"
            '        objRoot.SetAttributeNode(objAttr)

            '        LogFile.AppendChild(objRoot)
            '    End If

            '    'Create the new container element for the new record
            '    objRecord = LogFile.CreateElement("ow:debug")

            '    objField = LogFile.CreateElement("ow:date")
            '    objField.Text = Now()
            '    objRecord.AppendChild(objField)

            '    objField = LogFile.CreateElement("ow:text")
            '    objField.Text = pText
            '    objRecord.AppendChild(objField)

            '    objField = LogFile.CreateElement("ow:value")
            '    objField.Text = HttpContext.Current.Server.HtmlEncode(pParam)
            '    objRecord.AppendChild(objField)

            '    objRoot.AppendChild(objRecord)

            '    'Check once again to see if the file loaded successfully. If it did
            '    'not, that means we are creating a new document and need to be sure to
            '    'insert the XML processing instruction.
            '    If FileExists = False Then
            '        objPI = LogFile.CreateProcessingInstruction("xml", "version='1.0'")
            '        LogFile.InsertBefore(objPI, LogFile.childNodes(0))
            '    End If

            '    'Save the XML document
            '    LogFile.Save(OPENWIKI_DEBUGPATH)

            '    'Collect garbage
            '    LogFile = Nothing
            '    objRoot = Nothing
            '    objPI = Nothing
            '    objField = Nothing
            '    objAttr = Nothing
            'End If
        End Sub

    End Module
End Namespace