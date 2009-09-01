Namespace Openwiki
    Module Preamble
        Public Const OPENWIKI_VERSION As String = "0.78"
        Public Const OPENWIKI_REVISION As String = "$Revision: 1.2 $"
        Public Const OPENWIKI_XMLVERSION As String = "0.91"
        Public Const OPENWIKI_NAMESPACE As String = "http://openwiki.com/2001/OW/Wiki"

        ' possible values for OPENWIKI_DB_SYNTAX
        Public Const DB_ACCESS As Integer = 0
        Public Const DB_SQLSERVER As Integer = 0
        Public Const DB_ORACLE As Integer = 1
        Public Const DB_MYSQL As Integer = 2
        Public Const DB_POSTGRESQL As Integer = 3

        ' declare 'constants'
        'Public OPENWIKI_DB
        'Public OPENWIKI_DB_SYNTAX
        'Public OPENWIKI_ICONPATH
        'Public OPENWIKI_IMAGEPATH
        'Public OPENWIKI_ENCODING
        'Public OPENWIKI_TITLE
        'Public OPENWIKI_FRONTPAGE
        'Public OPENWIKI_SCRIPTNAME
        'Public OPENWIKI_STYLESHEETS
        'Public OPENWIKI_MAXTEXT
        'Public OPENWIKI_MAXINCLUDELEVEL
        'Public OPENWIKI_RCNAME
        'Public OPENWIKI_RCDAYS
        'Public OPENWIKI_MAXTRAIL
        'Public OPENWIKI_TEMPLATES
        'Public OPENWIKI_TIMEZONE
        'Public OPENWIKI_MAXNROFAGGR
        'Public OPENWIKI_MAXWEBGETS
        'Public OPENWIKI_SCRIPTTIMEOUT
        'Public OPENWIKI_DAYSTOKEEP
        'Public OPENWIKI_DAYSTOKEEP_DEPRECATED
        'Public OPENWIKI_STOPWORDS
        'Public OPENWIKI_UPLOADDIR
        'Public OPENWIKI_MAXUPLOADSIZE
        'Public OPENWIKI_UPLOADTIMEOUT
        'Public OPENWIKI_RECAPTCHAPRIVATEKEY
        'Public OPENWIKI_DEBUGLEVEL
        'Public OPENWIKI_DEBUGPATH
        'Public OPENWIKI_PROTECTEDPAGES

        'Public MSXML_VERSION

        ' declare options
        'Public gReadPassword
        'Public gEditPassword
        'Public gDefaultBookmarks
        'Public gAdminPassword
        'Public cReadOnly
        'Public cNakedView
        'Public cUseSubpage
        'Public cFreeLinks
        'Public cWikiLinks
        'Public cAcronymLinks
        'Public cTemplateLinking
        'Public cRawHtml
        'Public cMathML
        'Public cHtmlTags
        'Public cCacheXSL
        'Public cCacheXML
        'Public cDirectEdit
        'Public cEmbeddedMode
        'Public cSimpleLinks
        'Public cNonEnglish
        'Public cNetworkFile
        'Public cBracketText
        'Public cBracketIndex
        'Public cHtmlLinks
        'Public cBracketWiki
        'Public cFreeUpper
        'Public cLinkImages
        'Public cUseHeadings
        'Public cUseLookup
        'Public cStripNTDomain
        'Public cMaskIPAddress
        'Public cOldSkool
        'Public cNewSkool
        'Public cNumTOC
        'Public cAllowCharRefs
        'Public cWikifyHeaders
        'Public cEmoticons
        'Public cUseLinkIcons
        'Public cPrettyLinks
        'Public cExternalOut
        'Public cAllowRSSExport
        'Public cAllowNewSyndications
        'Public cAllowAggregations
        'Public cNTAuthentication
        'Public cShowBrackets
        'Public cAllowAttachments
        'Public cUseXhtmlHttpHeaders
        'Public cUseSpecialPagesPrefix
        'Public cUseRecaptcha

        ' global variables
        Public gLinkPattern As String
        Public gSubpagePattern As String
        Public gStopWords As String
        Public gTimestampPattern As String
        Public gUrlProtocols As String
        Public gUrlPattern As String
        Public gMailPattern As String
        Public gInterSitePattern As String
        Public gInterLinkPattern As String
        Public gFreeLinkPattern As String
        Public gImageExtensions As String
        Public gImagePattern As String
        Public gDocExtensions As String
        Public gNotAcceptedExtensions As String
        Public gISBNPattern As String
        Public gHeaderPattern As String
        Public gMacros As String
        'Public gSpecialPagesPrefix
        'Public gCategoryMarkPattern
        Public gFS As Char = Chr(179)           ' The FS character is a superscript "3"
        Public gIndentLimit As Integer = 20        ' maximum indent level for bulleted/numbered items

        ' incoming parameters
        Public gPage As String                ' page to be worked on
        Public gRevision As Integer            ' revision of page to be worked on
        Public gAction As String              ' action
        Public gTxt As String                 ' text value passed to input boxes

        Public gLastModified As Date        ' last-modified date of page to be worked on
        Public gServerRoot As String          ' URL path to script
        Public gScriptName As String          ' Name of this script
        Public gTransformer As Transformer         ' transformer of XML data
        Public gNamespace As OpenWikiNamespace          ' namespace data
        Public gRaw As Vector                 ' vector or raw data used by Wikify function
        Public gBracketIndices As Vector      ' keep track of the bracketed indices
        Public gTOC As TableOfContents                 ' table of contents
        Public gCategories As Vector          ' categories of page
        Public gIncludeLevel As Integer        ' recursive level of included pages
        Public gCurrentWorkingPages As Vector ' stack of pages currently working on when including pages
        Public gIncludingAsTemplate As Boolean ' including subpages as template
        Public gNrOfRSSRetrievals As Integer   ' nr of remote calls performed to retrieve an RSS feed
        Public gAggregateURLs As Vector       ' URL's to RSS feeds that need to be aggregated for this page
        Public gCookieHash As String          ' Hash value to use in cookie names
        '        Public gTemp                ' temporary value that may be used at all times
        Public gActionReturn As Boolean        ' return value used by actions
        Public gMacroReturn As String         ' return value used by macros

        'If (ScriptEngineMajorVersion < 5) Or (ScriptEngineMajorVersion = 5 And ScriptEngineMinorVersion < 5) Then
        '    HttpContext.Current.Response.Write("<h2>Error: Missing VBScript v5.5</h2>")
        '    HttpContext.Current.Response.Write("In order for this script to work correctly the component " _
        '                 & "VBScript v5.5 " _
        '                 & "or a higher version needs to be installed on the HttpContext.Current.Server. You can download this component from " _
        '                 & "<a href=""http://msdn.microsoft.com/scripting/"">http://msdn.microsoft.com/scripting/</a>.")
        '    HttpContext.Current.Response.End
        'End If

        'Dim c, i
        'c = HttpContext.Current.Request.ServerVariables.Count
        'For i = 1 To c
        '    HttpContext.Current.Response.Write(Request.ServerVariables.Key(i) & " ==> " & HttpContext.Current.Request.ServerVariables.Item(i) & "<br>")
        'Next
        'Response.End

    End Module
End Namespace