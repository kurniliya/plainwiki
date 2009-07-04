<script language="VB" runat="Server">

Const OPENWIKI_VERSION As String = "0.78"

Const OPENWIKI_REVISION As String = "$Revision: 1.2 $"

Const OPENWIKI_XMLVERSION As String = "0.91"

Const OPENWIKI_NAMESPACE As String = "http://openwiki.com/2001/OW/Wiki"


' possible values for OPENWIKI_DB_SYNTAX
Const DB_ACCESS As Short = 0

Const DB_SQLSERVER As Short = 0

Const DB_ORACLE As Short = 1

Const DB_MYSQL As Short = 2

Const DB_POSTGRESQL As Short = 3


' declare 'constants'
Dim OPENWIKI_DB As String
Dim OPENWIKI_DB_SYNTAX As Short

Dim OPENWIKI_TITLE As String
Dim OPENWIKI_ENCODING As String
Dim OPENWIKI_IMAGEPATH As String
Dim OPENWIKI_FRONTPAGE As String
Dim OPENWIKI_ICONPATH As String

Dim OPENWIKI_MAXINCLUDELEVEL As Byte
Dim OPENWIKI_MAXTEXT As Integer
Dim OPENWIKI_SCRIPTNAME As String
Dim OPENWIKI_STYLESHEETS As String

Dim OPENWIKI_MAXTRAIL As Byte
Dim OPENWIKI_RCDAYS As Byte
Dim OPENWIKI_RCNAME As String
Dim OPENWIKI_TEMPLATES As String

Dim OPENWIKI_TIMEZONE As String
Dim OPENWIKI_SCRIPTTIMEOUT As Byte
Dim OPENWIKI_MAXWEBGETS As Byte
Dim OPENWIKI_MAXNROFAGGR As Byte

Dim OPENWIKI_DAYSTOKEEP_DEPRECATED As Byte
Dim OPENWIKI_DAYSTOKEEP As Byte

Dim OPENWIKI_STOPWORDS As String

Dim OPENWIKI_UPLOADTIMEOUT As Short
Dim OPENWIKI_MAXUPLOADSIZE As Integer
Dim OPENWIKI_UPLOADDIR As String

Dim OPENWIKI_RECAPTCHAPRIVATEKEY As String

Dim OPENWIKI_DEBUGLEVEL As Byte
Dim OPENWIKI_DEBUGPATH As Object

Dim OPENWIKI_PROTECTEDPAGES As String


Dim MSXML_VERSION As Byte


' declare options
Dim gDefaultBookmarks As String
Dim gReadPassword As String
Dim gEditPassword As String
Dim gAdminPassword As String

Dim cEmbeddedMode As Byte
Dim cAcronymLinks As Byte
Dim cDirectEdit As Byte
Dim cMathML As Byte
Dim cFreeLinks As Byte
Dim cHtmlTags As Byte
Dim cReadOnly As Byte
Dim cCacheXSL As Byte
Dim cTemplateLinking As Byte
Dim cRawHtml As Byte
Dim cNakedView As Byte
Dim cCacheXML As Byte
Dim cWikiLinks As Byte
Dim cUseSubpage As Byte

Dim cSimpleLinks As Byte
Dim cBracketWiki As Byte
Dim cHtmlLinks As Byte
Dim cStripNTDomain As Byte
Dim cUseLookup As Byte
Dim cNetworkFile As Byte
Dim cOldSkool As Byte
Dim cAllowCharRefs As Byte
Dim cUseHeadings As Byte
Dim cWikifyHeaders As Byte
Dim cNonEnglish As Byte
Dim cLinkImages As Byte
Dim cNumTOC As Byte
Dim cFreeUpper As Byte
Dim cMaskIPAddress As Byte
Dim cBracketText As Byte
Dim cBracketIndex As Byte
Dim cNewSkool As Byte

Dim cExternalOut As Byte
Dim cPrettyLinks As Byte
Dim cUseLinkIcons As Byte
Dim cEmoticons As Byte

Dim cAllowAggregations As Byte
Dim cAllowNewSyndications As Byte
Dim cNTAuthentication As Byte
Dim cShowBrackets As Byte
Dim cAllowRSSExport As Byte

Dim cAllowAttachments As Byte

Dim cUseXhtmlHttpHeaders As Byte

Dim cUseSpecialPagesPrefix As Byte

Dim cUseRecaptcha As Byte


' global variables
Dim gSubpagePattern As Object
Dim gLinkPattern As Object
Dim gInterSitePattern As Object
Dim gStopWords As Object
Dim gDocExtensions As Object
Dim gFreeLinkPattern As Object
Dim gUrlPattern As Object
Dim gTimestampPattern As Object
Dim gUrlProtocols As Object
Dim gHeaderPattern As Object
Dim gImagePattern As Object
Dim gImageExtensions As Object
Dim gMacros As Object
Dim gMailPattern As Object
Dim gNotAcceptedExtensions As Object
Dim gISBNPattern As Object
Dim gInterLinkPattern As Object

Dim gSpecialPagesPrefix As String

Dim gCategoryMarkPattern As String

Dim gFS As String
Dim gIndentLimit As Byte


' incoming parameters
Dim gPage As Object ' page to be worked on

Dim gRevision As Object ' revision of page to be worked on

Dim gAction As Object ' action

Dim gTxt As Object ' text value passed to input boxes


Dim gLastModified As Object ' last-modified date of page to be worked on

Dim gServerRoot As Object ' URL path to script

Dim gScriptName As Object ' Name of this script

Dim gTransformer As Object ' transformer of XML data

Dim gNamespace As Object ' namespace data

Dim gRaw As Object ' vector or raw data used by Wikify function

Dim gBracketIndices As Object ' keep track of the bracketed indices

Dim gTOC As Object ' table of contents

Dim gCategories As Object ' categories of page

Dim gIncludeLevel As Object ' recursive level of included pages

Dim gCurrentWorkingPages As Object ' stack of pages currently working on when including pages

Dim gIncludingAsTemplate As Object ' including subpages as template

Dim gNrOfRSSRetrievals As Object ' nr of remote calls performed to retrieve an RSS feed

Dim gAggregateURLs As Object ' URL's to RSS feeds that need to be aggregated for this page

Dim gCookieHash As Object ' Hash value to use in cookie names

Dim gTemp As Object ' temporary value that may be used at all times

Dim gActionReturn As Object ' return value used by actions

Dim gMacroReturn As Object ' return value used by macros

</script>
<%gFS = Chr(179) ' The FS character is a superscript "3"
gIndentLimit = 20 ' maximum indent level for bulleted/numbered items

If (ScriptEngineMajorVersion < 5) Or (ScriptEngineMajorVersion = 5 And ScriptEngineMinorVersion < 5) Then
	Response.Write("<h2>Error: Missing VBScript v5.5</h2>")
	Response.Write("In order for this script to work correctly the component " & "VBScript v5.5 " & "or a higher version needs to be installed on the server. You can download this component from " & "<a href=""http://msdn.microsoft.com/scripting/"">http://msdn.microsoft.com/scripting/</a>.")
	Response.End()
End If

'Dim c, i
'c = Request.ServerVariables.Count
'For i = 1 To c
'    Response.Write(Request.ServerVariables.Key(i) & " ==> " & Request.ServerVariables.Item(i) & "<br>")
'Next
'Response.End

%>
