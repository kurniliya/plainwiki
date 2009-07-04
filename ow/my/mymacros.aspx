<script language="VB" runat="Server">


<%	'#$>$#C:\Sources\plainwiki\trunk\wiki\ow\my\mymacros.asp|%>
<%	'#$>$#C:\Sources\plainwiki\trunk\wiki\ow\owall.asp|%>


<%	
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
	
	OwProcessRequest()
	%><%	'#$>$#C:\Sources\plainwiki\trunk\wiki\ow.asp|%>

	
End Sub

'#$<$#C:\Sources\plainwiki\trunk\wiki\ow\my\mymacros.asp|


' Examples of custom build macros.
'
' When you create a new macro add the letter P to the name of
' the sub for each parameter you define. A macro can take at
' most 2 parameters.
'
' A macro should return the value that is supposed to be
' substituted in the text by setting the global variable
' gMacroReturn.

' For each macro you add below, you must add it's name to the
' return value of this function. Seperate the names by the
' pipe (|) character.
'
' If you want to redefine all available macros set the
' variable gMacros (see also owpatterns.asp).
Function MyMacroPatterns() As Object
	If cEmbeddedMode = 0 Then
		'gMacros = "BR|RecentChanges|TitleSearch|FullSearch|TextSearch|TableOfContents|WordIndex|TitleIndex|GoTo|RandomPage|InterWiki|SystemInfo|Include|PageCount|UserPreferences|Icon|Anchor|Date|Time|DateTime|Syndicate|Aggregate|Footnote"
		MyMacroPatterns = "Glossary|Files"
	End If
End Function



' code by Dan Rawsthorne
' taken from http://openwiki.com/?OpenWiki/Suggestions
Sub MacroGlossaryP(ByRef pParams As Object)
	gMacroReturn = GetGlossaryP(pParams)
End Sub


Public Function GetGlossaryP(ByRef pPattern As Object) As Object
	Dim i, vList, vCount, vResult As Object
	vList = gNamespace.FullSearch(pPattern, False)
	vCount = vList.Count - 1
	For i = 0 To vCount
		vResult = vResult & vList.ElementAt(i).ToXML(False)
	Next 
	GetGlossaryP = "<ow:titleindex>" & vResult & "</ow:titleindex>"
End Function



' original code by Leopold Faschalek
' modified by Dave Cantrell
' modified by Laurens Pit
' taken from http://openwiki.com/?OpenWiki/Suggestions
Sub MacroFilesP(ByRef pPath As Object)
	Call MacroFilesPP(pPath, "[\s\S]*")
End Sub


Sub MacroFilesPP(ByRef pPath As Object, ByRef pWild As Object)
	Dim oFso As Scripting.FileSystemObject
	Dim oFiles, oFolder, oFile As Object
	oFso = New Scripting.FileSystemObject
	Dim sLocalPathPrefix As Object
	Dim sUncPath As Object
	If oFso.FolderExists(pPath) Then
		oFolder = oFso.GetFolder(pPath)
		oFiles = oFolder.Files
		'parses path and converts to UNC path so files can be retrieved from server across network
		sLocalPathPrefix = Left(oFolder.Path, 1)
		sUncPath = "file:\\\" & pPath & "\"
		'sUncPath = "\\mymachine\" & Lcase( sLocalPathPrefix ) & "$" & Right( oFolder.Path, Len( oFolder.Path ) - 2 ) & "\"  '"
		'sUncPath = "http:\\www.mysite.com\" & Lcase( sLocalPathPrefix ) & "$" & Right( oFolder.Path, Len( oFolder.Path ) - 2 ) & "\"  '"
		gMacroReturn = "<b>" & sUncPath & "</b><ul>"
		For	Each oFile In oFiles
			If m(oFile.Name, pWild, False, True) Then
				gMacroReturn = gMacroReturn & "<li><a href='" & sUncPath & oFile.Name & "' target='_blank'>" & oFile.Name & "</a></li>"
			End If
		Next oFile
		gMacroReturn = gMacroReturn & "</ul>"
	Else
		gMacroReturn = "<ow:error>error in path: " & pPath & "</ow:error>"
	End If
	'UPGRADE_NOTE: Object oFile may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
	oFile = Nothing
	'UPGRADE_NOTE: Object oFiles may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
	oFiles = Nothing
	'UPGRADE_NOTE: Object oFolder may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
	oFolder = Nothing
	'UPGRADE_NOTE: Object oFso may not be destroyed until it is garbage collected. Copy this link in your browser for more: 'http://msdn.microsoft.com/library/en-us/vbcon/html/vbup1029.asp'
	oFso = Nothing
	
	' prevent the caching of pages wherein this macro is used
	cCacheXML = False
End Sub




Sub MacroPagechangedP(ByRef pParam As Object)
	Dim vPage, vTimestamp As Object
	vPage = gNamespace.GetPage(pParam, 0, False, False)
	If vPage.Exists Then
		vTimestamp = vPage.GetLastChange().Timestamp()
		gMacroReturn = FormatDate(vTimestamp)
	End If
End Sub

</script>
