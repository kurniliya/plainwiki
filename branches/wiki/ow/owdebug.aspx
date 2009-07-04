

Sub WriteDebug(ByRef pText As Object, ByRef pParam As Object, ByRef pDebugLevel As Object)
	Dim LogFile As MSXML2.DOMDocument60
	Dim objField, objRoot, FileExists, objPI, objRecord, objAttr As Object
	
	If OPENWIKI_DEBUGLEVEL >= pDebugLevel Then
		'		Response.Write("in panic<br>")
		'Instantiate the Microsoft XMLDOM
		LogFile = New MSXML2.DOMDocument60
		LogFile.PreserveWhiteSpace = True
		
		FileExists = LogFile.Load(OPENWIKI_DEBUGPATH)
		'		Response.Write("FileExists " & FileExists & "<br>")
		
		If FileExists = True Then
			'			Response.Write("file exists<br>")
			'If the file loaded set the objRoot Object equal to the root element
			'of the XML document
			objRoot = LogFile.DocumentElement
		Else
			'			Response.Write("file not exists<br>")
			'			Response.Write(LogFile.parseError.errorCode)
			
			'Create root element and append it to the XML document
			objRoot = LogFile.CreateElement("ow:panic")
			
			'Declare default namespace
			objAttr = LogFile.CreateAttribute("xmlns")
			objAttr.Text = "http://www.w3.org/1999/xhtml"
			objRoot.SetAttributeNode(objAttr)
			
			'Declare OpenWiki namespace
			objAttr = LogFile.CreateAttribute("xmlns:ow")
			objAttr.Text = "http://openwiki.com/2001/OW/Wiki"
			objRoot.SetAttributeNode(objAttr)
			
			LogFile.AppendChild(objRoot)
		End If
		
		'Create the new container element for the new record
		objRecord = LogFile.CreateElement("ow:debug")
		
		objField = LogFile.CreateElement("ow:date")
		objField.Text = Now
		objRecord.AppendChild(objField)
		
		objField = LogFile.CreateElement("ow:text")
		objField.Text = pText
		objRecord.AppendChild(objField)
		
		objField = LogFile.CreateElement("ow:value")
		objField.Text = Server.HTMLEncode(pParam)
		objRecord.AppendChild(objField)
		
		objRoot.AppendChild(objRecord)
		
		'Check once again to see if the file loaded successfully. If it did
		'not, that means we are creating a new document and need to be sure to
		'insert the XML processing instruction.
		If FileExists = False Then
			objPI = LogFile.CreateProcessingInstruction("xml", "version='1.0'")
			LogFile.InsertBefore(objPI, LogFile.childNodes(0))
		End If
		
		'Save the XML document
		LogFile.Save(OPENWIKI_DEBUGPATH)
		
		'Collect garbage
		'UPGRADE_NOTE: Object LogFile may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
		LogFile = Nothing
		'UPGRADE_NOTE: Object objRoot may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
		objRoot = Nothing
		'UPGRADE_NOTE: Object objPI may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
		objPI = Nothing
		'UPGRADE_NOTE: Object objField may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
		objField = Nothing
		'UPGRADE_NOTE: Object objAttr may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
		objAttr = Nothing
	End If
End Sub

