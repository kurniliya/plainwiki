<%

' Following are all the configuration items with default values set.
' Override them if you want in a separate file, see e.g. /web1/ow.asp.

' "The Truth about MS Access" : http://www.15seconds.com/Issue/010514.htm
' OPENWIKI_DB = "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & Server.MapPath("OpenWikiDist.mdb")
' OPENWIKI_DB = "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & Server.MapPath("/cgi-bin/OpenWikiDist.mdb")
 OPENWIKI_DB = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("/cgi-bin/OpenWikiDist.mdb")
' OPENWIKI_DB = "Driver={SQL Server};server=mymachine;uid=openwiki;pwd=openwiki;database=OpenWiki"
' OPENWIKI_DB = "Driver={Microsoft ODBC for Oracle};Server=OW;Uid=laurens;Pwd=aphex2twin;"
' OPENWIKI_DB = "MySystemDSName"
' OPENWIKI_DB = "MySQLOpenWiki"
' OPENWIKI_DB = "PostgreSQLOpenWiki"
' OPENWIKI_DB = "OpenWikiDist"

'OPENWIKI_DB_SYNTAX = DB_ACCESS               ' see owpreamble.asp for possible values

OPENWIKI_IMAGEPATH       = "ow/images"        ' path to images directory
OPENWIKI_ICONPATH        = "ow/images/icons"  ' path to icons directory
OPENWIKI_ENCODING        = "utf-8"            ' character encoding to use
OPENWIKI_TITLE           = "NEQwiki, the nonlinear equations encyclopedia"          ' title of your wiki
OPENWIKI_FRONTPAGE       = "FrontPage"        ' name of your front page.
OPENWIKI_SCRIPTNAME      = "ow.asp"           ' "mydir/ow.asp" : in case the auto-detected scriptname isn't correct
OPENWIKI_STYLESHEETS     = "ow/xsl/"          ' the subdirectory where the stylesheet files (*.xsl) are located
OPENWIKI_MAXTEXT         = 204800             ' Maximum 200K texts
OPENWIKI_MAXINCLUDELEVEL = 5                  ' Maximum depth of Include's
OPENWIKI_RCNAME          = gSpecialPagesPrefix & "RecentChanges"    ' Name of recent changes page (change space to _)
OPENWIKI_RCDAYS          = 30                 ' Default number of RecentChanges days
OPENWIKI_MAXTRAIL        = 0                  ' Maximum number of links in the trail
OPENWIKI_STOPWORDS       = "StopWords"        ' Name of page containing stop words (change space to _). Stop words are words that won't be hyperlinked. Use empty string "" if you do not want to support stop words.
OPENWIKI_TEMPLATES       = "Template$"        ' Pattern for templates usable when creating a new page
OPENWIKI_TIMEZONE        = "+03:00"           ' Timezone of the server running this wiki, valid values are e.g. "+04:00", "-09:00", etc.
OPENWIKI_MAXNROFAGGR     = 150                ' Maximum number of rows to show in an aggregated feed
OPENWIKI_MAXWEBGETS      = 3                  ' Maximum number of RSS feeds that may be refreshed from a remote server for one user request.
OPENWIKI_SCRIPTTIMEOUT   = 120                ' Maximum amount of seconds to wait for RSS feeds to be syndicated, if set to 0 the default timeout value of ASP is used.
OPENWIKI_DAYSTOKEEP      = 30                 ' Number of days to keep old revisions
OPENWIKI_DAYSTOKEEP_DEPRECATED = 30           ' Number of days to keep deprecated pages and attachments
OPENWIKI_UPLOADDIR       = "attachments/"     ' The virtual directory where uploads are stored
OPENWIKI_MAXUPLOADSIZE   = 8388608            ' Use to limit the size of uploads, in bytes (default = 8,388,608)
OPENWIKI_UPLOADTIMEOUT   = 300                ' Timeout in seconds (upload must succeed within this time limit)
OPENWIKI_RECAPTCHAPRIVATEKEY = "6Lea0wYAAAAAANFrNX75pLVzS95BJXuJrGIIALeP"
OPENWIKI_DEBUGLEVEL      = 0                  ' Set positive value to enable debug logging
OPENWIKI_DEBUGPATH       = Server.MapPath("/cgi-bin/owdebug.xml")	' Path for storing logs
OPENWIKI_PROTECTEDPAGES  = "FrontPage"        ' Pattern of wiki pages password protected from editing
OPENWIKI_ENGINEUPGRADEDATE = CDate("24/01/2010 00:00:00")	' Date of last engine upgrade. Locale dependent!

MSXML_VERSION = 6   ' specify version of MSXML installed. Version 3 should be supported everywhere

gReadPassword = ""    ' use empty string "" if anyone may read
gEditPassword = "1111"    ' use empty string "" if anyone may edit
gAdminPassword = "adminpw"   ' use empty string "" if anyone may administer this Wiki
' In case you want more sophisticated security, then you should
' rely on the Integrated Windows authentication feature of IIS.

gDefaultBookmarks = ""

' Major system options
cUseXhtmlHttpHeaders   = 1        ' 1 = application/xhtml+xml 0 = text/html
cReadOnly              = 0        ' 1 = readonly wiki         0 = editable wiki
cNakedView             = 0        ' 1 = run in naked mode     0 = show headers/footers
cUseSubpage            = 1        ' 1 = use /subpages         0 = do not use /subpages
cFreeLinks             = 1        ' 1 = use [[word]] links    0 = LinkPattern only
cWikiLinks             = 1        ' 1 = use LinkPattern       0 = possibly allow [[word]] only
cAcronymLinks          = 0        ' 1 = link acronyms         0 = do not link 3 or more capitalized characters
cTemplateLinking       = 1        ' 1 = allow TemplateName->WikiLink   0 = don't do template linking
cRawHtml               = 1        ' 1 = allow <html> tag      0 = no raw HTML in pages
cMathML                = 1        ' 1 = allow <math> tag      0 = no raw math in pages
cHtmlTags              = 1        ' 1 = "unsafe" HTML tags    0 = only minimal tags
cCacheXSL              = 0        ' 1 = cache stylesheet      0 = don't cache stylesheet
cCacheXML              = 0        ' 1 = cache partial results 0 = do not cache partial results
cAllowRSSExport        = 1        ' 1 = allow RSS feed        0 = do not export your pages to RSS
cAllowNewSyndications  = 1        ' 1 = allow new URLs to be syndicated    0 = only allow syndication of the URLs in the database table openwiki_rss
cAllowAggregations     = 1        ' 1 = allow aggregation of syndications (note: you MUST use MSXML v3 sp2 for this to work)   0 = do not allow aggregrations
cEmbeddedMode          = 0        ' 1 = embed the wiki into another app    0 = process browser request
cAllowAttachments      = 0        ' 1 = allow attachments     0 = do not allow attachments (WARNING: Allowing attachments poses a security risk!! See file owattach.asp)
cUseSpecialPagesPrefix = 1 		  ' 1 = use gSpecialPagesPrefix in gLinkPattern
gSpecialPagesPrefix    = "Special:"
cUseRecaptcha          = 1        ' 1 = use reCAPTHCA when edit pages if no password protection defined
gCategoryMarkPattern   = "\[\[:Category([\w]*)\]\]"       ' Pattern used to find category marks on wikipages

' Minor system options
cSimpleLinks          = 0        ' 1 = only letters,         0 = allow _ and numbers
cNonEnglish           = 1        ' 1 = extra link chars,     0 = only A-Za-z chars
cNetworkFile          = 1        ' 1 = allow remote file:    0 = no file:// links
cBracketText          = 1        ' 1 = allow [URL text]      0 = no link descriptions
cBracketIndex         = 1        ' 1 = [URL] -> [<index>]    0 = [URL] -> [URL]
cHtmlLinks            = 1        ' 1 = allow A HREF links    0 = no raw HTML links
cBracketWiki          = 1        ' 1 = [WikiLnk txt] link    0 = no local descriptions
cShowBrackets         = 0        ' 1 = keep brackets         0 = remove brackets when it's an external link
cFreeUpper            = 1        ' 1 = force upper case      0 = do not force case for free links
cLinkImages           = 1        ' 1 = display image         0 = display link to image
cUseHeadings          = 1        ' 1 = allow = h1 text =     0 = no header formatting
cUseLookup            = 1        ' 1 = lookup host names     0 = skip lookup (IP only)
cStripNTDomain        = 1        ' 1 = strip NT domainname   0 = keep NT domainname in remote username
cMaskIPAddress        = 1        ' 1 = mask last part of IP  0 = show full IP address in RecentChanges list, etc.
cOldSkool             = 1        ' 1 = use '' and '''        0 = don't use '' and ''' for italic and bold, and use Wiki''''''Link to escape WikiLink
cNewSkool             = 1        ' 1 = use //, **, -- and __ 0 = don't use //, **, -- and __ for italic, bold, strikethrough and underline and use ~WikiLink to escape WikiLink
cNumTOC               = 1        ' 1 = TOC numbered          0 = TOC just indented text
cNTAuthentication     = 1        ' 1 = Use NT username       0 = blank username in preferences
cDirectEdit           = 1        ' 1 = go direct to edit     0 = go to blank page first
cAllowCharRefs        = 1        ' 1 = allow char refs       0 = no character references allowed (like &copy; or &#151;)
cWikifyHeaders        = 1        ' 1 = wikify headers        0 = do not apply wiki formatting within headers

' User options
cEmoticons            = 1        ' 1 = use emoticons         0 = don't show feelings
cUseLinkIcons         = 1        ' 1 = icons for ext links   0 = no icon images for external links
cPrettyLinks          = 1        ' 1 = display Words Smashed Together     0 = display WordsSmashedTogether
cExternalOut          = 1        ' 1 = external links open in new window, 0 = open in same window

%>