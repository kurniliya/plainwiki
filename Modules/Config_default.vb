Namespace Openwiki
    Module Config_default
        ' Following are all the configuration items with default values set.
        ' Override them if you want in a separate file, see e.g. /web1/ow.asp.

        ' "The Truth about MS Access" : http://www.15seconds.com/Issue/010514.htm
        ' OPENWIKI_DB = "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & HttpContext.Current.Server.MapPath("OpenWikiDist.mdb")
        Public OPENWIKI_DB As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & HttpContext.Current.Server.MapPath("OpenWikiDist.mdb")
        ' OPENWIKI_DB = "Driver={SQL Server};server=mymachine;uid=openwiki;pwd=openwiki;database=OpenWiki"
        ' OPENWIKI_DB = "Driver={Microsoft ODBC for Oracle};Server=OW;Uid=laurens;Pwd=aphex2twin;"
        ' OPENWIKI_DB = "MySystemDSName"
        ' OPENWIKI_DB = "MySQLOpenWiki"
        ' OPENWIKI_DB = "PostgreSQLOpenWiki"
        ' OPENWIKI_DB = "OpenWikiDist"

        Public Const OPENWIKI_DB_SYNTAX As Integer = DB_ACCESS               ' see owpreamble.asp for possible values

        Public Const OPENWIKI_IMAGEPATH As String = "ow/images"        ' path to images directory
        Public Const OPENWIKI_ICONPATH As String = "ow/images/icons"  ' path to icons directory
        Public Const OPENWIKI_ENCODING As String = "utf-8"            ' character encoding to use
        Public Const OPENWIKI_TITLE As String = "NEQwiki, the nonlinear equations encyclopedia"          ' title of your wiki
        Public Const OPENWIKI_FRONTPAGE As String = "FrontPage"        ' name of your front page.
        Public Const OPENWIKI_SCRIPTNAME As String = ""           ' "mydir/ow.asp" : in case the auto-detected scriptname isn't correct
        Public Const OPENWIKI_STYLESHEETS As String = "ow/xsl/"          ' the subdirectory where the stylesheet files (*.xsl) are located
        Public Const OPENWIKI_MAXTEXT As Integer = 204800             ' Maximum 200K texts
        Public Const OPENWIKI_MAXINCLUDELEVEL As Integer = 5                  ' Maximum depth of Include's
        Public Const OPENWIKI_RCNAME As String = gSpecialPagesPrefix & "RecentChanges"    ' Name of recent changes page (change space to _)
        Public Const OPENWIKI_RCDAYS As Integer = 30                 ' Default number of RecentChanges days
        Public Const OPENWIKI_MAXTRAIL As Integer = 0                  ' Maximum number of links in the trail
        Public Const OPENWIKI_STOPWORDS As String = "StopWords"        ' Name of page containing stop words (change space to _). Stop words are words that won't be hyperlinked. Use empty string "" if you do not want to support stop words.
        Public Const OPENWIKI_TEMPLATES As String = "Template$"        ' Pattern for templates usable when creating a new page
        Public OPENWIKI_TIMEZONE As String = "+03:00"           ' Timezone of the server running this wiki, valid values are e.g. "+04:00", "-09:00", etc.
        Public OPENWIKI_MAXNROFAGGR As Integer = 150                ' Maximum number of rows to show in an aggregated feed
        Public Const OPENWIKI_MAXWEBGETS As Integer = 3                  ' Maximum number of RSS feeds that may be refreshed from a remote server for one user HttpContext.Current.Request.
        Public Const OPENWIKI_SCRIPTTIMEOUT As Integer = 120                ' Maximum amount of seconds to wait for RSS feeds to be syndicated, if set to 0 the default timeout value of ASP is used.
        Public Const OPENWIKI_DAYSTOKEEP As Integer = 30                 ' Number of days to keep old revisions
        Public Const OPENWIKI_DAYSTOKEEP_DEPRECATED As Integer = 30           ' Number of days to keep deprecated pages and attachments
        Public Const OPENWIKI_UPLOADDIR As String = "attachments/"     ' The virtual directory where uploads are stored
        Public Const OPENWIKI_MAXUPLOADSIZE As Integer = 8388608            ' Use to limit the size of uploads, in bytes (default = 8,388,608)
        Public Const OPENWIKI_UPLOADTIMEOUT As Integer = 300                ' Timeout in seconds (upload must succeed within this time limit)
        Public Const OPENWIKI_RECAPTCHAPRIVATEKEY As String = "6Lea0wYAAAAAANFrNX75pLVzS95BJXuJrGIIALeP"
        'Public Const OPENWIKI_DEBUGLEVEL As Integer = 0                  ' Set positive value to enable debug logging
        'Public OPENWIKI_DEBUGPATH As String = HttpContext.Current.Server.MapPath("/cgi-bin/owdebug.xml")  ' Path for storing logs
        Public Const OPENWIKI_PROTECTEDPAGES As String = "FrontPage"                 ' Pattern of wiki pages password protected from editing

        Public MSXML_VERSION As Integer = 6   ' specify version of MSXML installed. Version 3 should be supported everywhere

        Public Const gReadPassword As String = ""    ' use empty string "" if anyone may read
        Public gEditPassword As String = ""    ' use empty string "" if anyone may edit
        Public Const gAdminPassword As String = "adminpw"   ' use empty string "" if anyone may administer this Wiki
        ' In case you want more sophisticated security, then you should
        ' rely on the Integrated Windows authentication feature of IIS.

        Public Const gDefaultBookmarks As String = ""

        ' Major system options
        Public Const cUseXhtmlHttpHeaders As Integer = 1        ' 1 = application/xhtml+xml 0 = text/html
        Public cReadOnly As Integer = 0        ' 1 = readonly wiki         0 = editable wiki
        Public Const cNakedView As Integer = 0        ' 1 = run in naked mode     0 = show headers/footers
        Public Const cUseSubpage As Integer = 1        ' 1 = use /subpages         0 = do not use /subpages
        Public Const cFreeLinks As Integer = 1        ' 1 = use [[word]] links    0 = LinkPattern only
        Public cWikiLinks As Integer = 1        ' 1 = use LinkPattern       0 = possibly allow [[word]] only
        Public Const cAcronymLinks As Integer = 0        ' 1 = link acronyms         0 = do not link 3 or more capitalized characters
        Public Const cTemplateLinking As Integer = 1        ' 1 = allow TemplateName->WikiLink   0 = don't do template linking
        Public Const cRawHtml As Integer = 1        ' 1 = allow <html> tag      0 = no raw HTML in pages
        Public Const cMathML As Integer = 1        ' 1 = allow <math> tag      0 = no raw math in pages
        Public Const cHtmlTags As Integer = 1        ' 1 = "unsafe" HTML tags    0 = only minimal tags
        Public Const cCacheXSL As Integer = 0        ' 1 = cache stylesheet      0 = don't cache stylesheet
        Public cCacheXML As Integer = 0        ' 1 = cache partial results 0 = do not cache partial results
        Public Const cAllowRSSExport As Integer = 1        ' 1 = allow RSS feed        0 = do not export your pages to RSS
        Public Const cAllowNewSyndications As Integer = 1        ' 1 = allow new URLs to be syndicated    0 = only allow syndication of the URLs in the database table openwiki_rss
        Public Const cAllowAggregations As Integer = 1        ' 1 = allow aggregation of syndications (note: you MUST use MSXML v3 sp2 for this to work)   0 = do not allow aggregrations
        Public Const cEmbeddedMode As Integer = 0        ' 1 = embed the wiki into another app    0 = process browser request
        Public cAllowAttachments As Integer = 0        ' 1 = allow attachments     0 = do not allow attachments (WARNING: Allowing attachments poses a security risk!! See file owattach.asp)
        Public Const cUseSpecialPagesPrefix As Integer = 1           ' 1 = use gSpecialPagesPrefix in gLinkPattern
        Public Const gSpecialPagesPrefix As String = "Special:"
        Public Const cUseRecaptcha As Integer = 0        ' 1 = use reCAPTHCA when edit pages if no password protection defined
        Public Const gCategoryMarkPattern As String = "\[\[:Category([\w]*)\]\]"       ' Pattern used to find category marks on wikipages

        ' Minor system options
        Public Const cSimpleLinks As Integer = 0        ' 1 = only letters,         0 = allow _ and numbers
        Public Const cNonEnglish As Integer = 1        ' 1 = extra link chars,     0 = only A-Za-z chars
        Public Const cNetworkFile As Integer = 1        ' 1 = allow remote file:    0 = no file:// links
        Public Const cBracketText As Integer = 1        ' 1 = allow [URL text]      0 = no link descriptions
        Public Const cBracketIndex As Integer = 1        ' 1 = [URL] -> [<index>]    0 = [URL] -> [URL]
        Public Const cHtmlLinks As Integer = 1        ' 1 = allow A HREF links    0 = no raw HTML links
        Public Const cBracketWiki As Integer = 1        ' 1 = [WikiLnk txt] link    0 = no local descriptions
        Public Const cShowBrackets As Integer = 0        ' 1 = keep brackets         0 = remove brackets when it's an external link
        Public Const cFreeUpper As Integer = 1        ' 1 = force upper case      0 = do not force case for free links
        Public Const cLinkImages As Integer = 1        ' 1 = display image         0 = display link to image
        Public Const cUseHeadings As Integer = 1        ' 1 = allow = h1 text =     0 = no header formatting
        Public Const cUseLookup As Integer = 1        ' 1 = lookup host names     0 = skip lookup (IP only)
        Public Const cStripNTDomain As Integer = 1        ' 1 = strip NT domainname   0 = keep NT domainname in remote username
        Public Const cMaskIPAddress As Integer = 1        ' 1 = mask last part of IP  0 = show full IP address in RecentChanges list, etc.
        Public Const cOldSkool As Integer = 1        ' 1 = use '' and '''        0 = don't use '' and ''' for italic and bold, and use Wiki''''''Link to escape WikiLink
        Public Const cNewSkool As Integer = 1        ' 1 = use //, **, -- and __ 0 = don't use //, **, -- and __ for italic, bold, strikethrough and underline and use ~WikiLink to escape WikiLink
        Public Const cNumTOC As Integer = 1        ' 1 = TOC numbered          0 = TOC just indented text
        Public Const cNTAuthentication As Integer = 1        ' 1 = Use NT username       0 = blank username in preferences
        Public Const cDirectEdit As Integer = 1        ' 1 = go direct to edit     0 = go to blank page first
        Public Const cAllowCharRefs As Integer = 1        ' 1 = allow char refs       0 = no character references allowed (like &copy; or &#151;)
        Public Const cWikifyHeaders As Integer = 1        ' 1 = wikify headers        0 = do not apply wiki formatting within headers

        ' User options
        Public cEmoticons As Integer = 1        ' 1 = use emoticons         0 = don't show feelings
        Public Const cUseLinkIcons As Integer = 1        ' 1 = icons for ext links   0 = no icon images for external links
        Public cPrettyLinks As Integer = 1        ' 1 = display Words Smashed Together     0 = display WordsSmashedTogether
        Public cExternalOut As Integer = 1        ' 1 = external links open in new window, 0 = open in same window

    End Module
End Namespace