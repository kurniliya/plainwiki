<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:msxsl="urn:schemas-microsoft-com:xslt"
                xmlns:ow="http://openwiki.com/2001/OW/Wiki"
                xmlns="http://www.w3.org/1999/xhtml"
                extension-element-prefixes="msxsl ow"
                exclude-result-prefixes=""
                version="1.0">
<xsl:output method="xml" indent="no" omit-xml-declaration="yes"/>

<xsl:variable name="brandingText">NEQwiki - ecnyclopedia of nonlinear differential equations.</xsl:variable>

<xsl:variable name="mainPageHeading">NEQwiki</xsl:variable>

<xsl:template name="head">
	<head>
		<meta http-equiv="Content-Type" content="application/xhtml+xml; charset={@encoding};" />
		<meta name="generator" content="OpenWiki 0.78" />
		<meta name="keywords" content="math, partial differential equations, mephi" />
		<meta name="description" content="NEQwiki - encyclopedia of nonlinear differential equations" />
		<meta name="ROBOTS" content="INDEX,FOLLOW" />
		<meta name="MSSmartTagsPreventParsing" content="true" />
		<title>
			<xsl:value-of select="ow:title"/> - <xsl:value-of select="ow:page/ow:link"/>
		</title>

<!--
	<link rel="stylesheet" type="text/css" href="ow/css/ow.css" />
-->
		<link rel="stylesheet" href="ow/css/common/shared.css?207xx" type="text/css" media="screen" />
		<link rel="stylesheet" href="ow/css/common/commonPrint.css?207xx" type="text/css" media="print" />
		<link rel="stylesheet" href="ow/css/monobook/main.css?207xx" type="text/css" media="screen" />
		<link rel="stylesheet" href="ow/css/chick/main.css?207xx" type="text/css" media="handheld" />
		
		<xsl:text disable-output-escaping="yes">&lt;!--[if lt IE 5.5000]&gt;</xsl:text>
		<link rel="stylesheet" href="ow/css/monobook/IE50Fixes.css?207xx" type="text/css" media="screen" />			
		<xsl:text disable-output-escaping="yes">&lt;![endif]--&gt;
		</xsl:text>
		
		<xsl:text disable-output-escaping="yes">&lt;!--[if IE 5.5000]&gt;</xsl:text>
			<link rel="stylesheet" href="ow/css/monobook/IE55Fixes.css?207xx" type="text/css" media="screen" />
		<xsl:text disable-output-escaping="yes">&lt;![endif]--&gt;
		</xsl:text>
		
		<xsl:text disable-output-escaping="yes">&lt;!--[if IE 6]&gt;</xsl:text>
			<link rel="stylesheet" href="ow/css/monobook/IE60Fixes.css?207xx" type="text/css" media="screen" />
		<xsl:text disable-output-escaping="yes">&lt;![endif]--&gt;
		</xsl:text>
		
		<xsl:text disable-output-escaping="yes">&lt;!--[if IE 7]&gt;</xsl:text>
			<link rel="stylesheet" href="ow/css/monobook/IE70Fixes.css?207xx" type="text/css" media="screen" />
		<xsl:text disable-output-escaping="yes">&lt;![endif]--&gt;
		</xsl:text>
		
		<xsl:text disable-output-escaping="yes">&lt;!--[if lt IE 7]&gt;</xsl:text>
			<script type="text/javascript" src="ow/js/common/IEFixes.js?207xx"></script>
			<meta http-equiv="imagetoolbar" content="no" />
		<xsl:text disable-output-escaping="yes">&lt;![endif]--&gt;
		</xsl:text>

		<script type= "text/javascript">
			<xsl:text disable-output-escaping="yes">
				/*&lt;![CDATA[*/
			</xsl:text>
			var skin = "monobook";
			var stylepath = "/skins-1.5";
			var wgArticlePath = "/wiki/$1";
			var wgScriptPath = "/w";
			var wgScript = "/w/index.php";
			var wgVariantArticlePath = false;
			var wgActionPaths = {};
			var wgServer = "http://en.wikipedia.org";
			var wgCanonicalNamespace = "";
			var wgCanonicalSpecialPageName = false;
			var wgNamespaceNumber = 0;
			var wgPageName = "Agrippina_(opera)";
			var wgTitle = "Agrippina (opera)";
			var wgAction = "view";
			var wgArticleId = "1257935";
			var wgIsArticle = true;
			var wgUserName = null;
			var wgUserGroups = null;
			var wgUserLanguage = "en";
			var wgContentLanguage = "en";
			var wgBreakFrames = false;
			var wgCurRevisionId = 283719792;
			var wgVersion = "1.15alpha";
			var wgEnableAPI = true;
			var wgEnableWriteAPI = true;
			var wgSeparatorTransformTable = ["", ""];
			var wgDigitTransformTable = ["", ""];
			var wgMWSuggestTemplate = "http://en.wikipedia.org/w/api.php?action=opensearch\x26search={searchTerms}\x26namespace={namespaces}\x26suggest";
			var wgDBname = "enwiki";
			var wgSearchNamespaces = [0];
			var wgMWSuggestMessages = ["with suggestions", "no suggestions"];
			var wgRestrictionEdit = [];
			var wgRestrictionMove = ["sysop"];
			<xsl:text disable-output-escaping="yes">
				/*]]&gt;*/
			</xsl:text>
		</script>

		<script type="text/javascript" src="ow/js/wikibits.js?207xx">
		</script>
	</head>
</xsl:template>

<xsl:template name="brandingImage">
    <a href="{/ow:wiki/ow:frontpage/@href}"><img src="{/ow:wiki/ow:imagepath}/logo.gif" align="right" border="0" alt="NEQwiki" /></a>
</xsl:template>

<xsl:template name="poweredBy">
	<div id="f-poweredbyico">
		<a href="http://openwiki.com"><img src="{/ow:wiki/ow:imagepath}/poweredby.gif" width="88" height="31" alt="Powered by OpenWiki" /></a>
    </div>
</xsl:template>

<xsl:template name="validatorButtons">
    <a href="http://validator.w3.org/check/referer"><img src="{/ow:wiki/ow:imagepath}/valid-xhtml10.gif" alt="Valid XHTML 1.0!" width="88" height="31" border="0" /></a>
    <a href="http://jigsaw.w3.org/css-validator/validator?uri={/ow:wiki/ow:location}ow.css"><img src="{/ow:wiki/ow:imagepath}/valid-css.gif" alt="Valid CSS!" width="88" height="31" border="0" /></a>
</xsl:template>

<xsl:template name="menu_column">

	<div id="p-cactions" class="portlet">
		<h5>Views</h5>
		<div class="pBody">
			<ul>
				<xsl:call-template name="menu_section_cactions" />
			</ul>
		</div>
	</div>

	<div class='generated-sidebar portlet' id='p-navigation'>
		<h5>Navigation</h5>
		<div class='pBody'>
			<ul>
				<xsl:call-template name="menu_section_navigation" />
			</ul>
		</div>
	</div>

	<div id="p-search" class="portlet">
		<h5><label for="searchInput">Search</label></h5>
		<div id="searchBody" class="pBody">
			<xsl:call-template name="menu_section_search" />
		</div>
	</div>
	
	<div class='generated-sidebar portlet' id='p-interaction'>
		<h5>Interaction</h5>
		<div class='pBody'>
			<ul>	
				<xsl:call-template name="menu_section_interaction" />
			</ul>
		</div>
	</div>

	<div class="portlet" id="p-tb">
		<h5>Toolbox</h5>
		<div class="pBody">
			<ul>
				<xsl:call-template name="menu_section_toolbox" />
			</ul>
		</div>
	</div>
</xsl:template>

<xsl:template name="menu_section_cactions">
	<li id="ca-nstab-main" class="selected">
		<a> 
			<xsl:attribute name="href"><xsl:value-of select="/ow:wiki/ow:scriptname"/>?p=<xsl:value-of select="$name"/></xsl:attribute>
			<xsl:attribute name="title">View the content page [c]</xsl:attribute>
			<xsl:attribute name="accesskey">c</xsl:attribute>
			Article
		</a>
	</li>
	<li id="ca-edit">
		<a>
			<xsl:attribute name="href"><xsl:value-of select="/ow:wiki/ow:scriptname"/>?p=<xsl:value-of select="$name"/>&amp;a=edit<xsl:if test="@revision">&amp;revision=<xsl:value-of select="@revision"/></xsl:if></xsl:attribute>
			<xsl:attribute name="title">You can edit this page. &#10;Please use the preview button before saving. [e]</xsl:attribute>
			<xsl:attribute name="accesskey">e</xsl:attribute>
			Edit this page
		</a>
	</li>	
	<li id="ca-history">
		<a>
			<xsl:attribute name="href"><xsl:value-of select="/ow:wiki/ow:scriptname"/>?p=<xsl:value-of select="$name"/>&amp;a=changes</xsl:attribute>
			<xsl:attribute name="title">Past versions of this page [h]</xsl:attribute>
			<xsl:attribute name="accesskey">h</xsl:attribute>
			History
		</a>
	</li>		
</xsl:template>

<xsl:template name="menu_section_navigation">
	<li id="n-mainpage-description">
		<a>
			<xsl:attribute name="href"><xsl:value-of select="/ow:wiki/ow:frontpage/@href"/></xsl:attribute>
			Main page
		</a>
	</li>
	<li id="n-titleindex">
		<a>
			<xsl:attribute name="href">ow.asp?TitleIndex</xsl:attribute>
			Title index
		</a>
	</li>	
	<li id="n-randompage">
		<a>
			<xsl:attribute name="href">ow.asp?RandomPage</xsl:attribute>
			Random article
		</a>
	</li>	
</xsl:template>

<xsl:template name="menu_section_search">
	<form method="get" id="searchform">
		<xsl:attribute name="action">
			<xsl:value-of select="/ow:wiki/ow:scriptname"/>
		</xsl:attribute>
		<div>
			<input type="hidden" name="a" value="fullsearch" />
            <input  type="text" name="txt"  class="searchButton" id="searchInput" size="30" ondblclick='event.cancelBubble=true;' /> 
            <input type="submit"  class="searchButton" id="mw-searchButton"  value="Search"/>
        </div>
    </form>
</xsl:template>

<xsl:template name="menu_section_interaction">
	<li id="n-recentchanges">
		<a>
			<xsl:attribute name="href"><xsl:value-of select="/ow:wiki/ow:scriptname"/>?p=RecentChanges</xsl:attribute>
			<xsl:attribute name="title">The list of recent changes in the wiki [r]</xsl:attribute>
			<xsl:attribute name="accesskey">r	</xsl:attribute>
			Recent changes
		</a>
	</li>
	<li id="n-help">
		<a>
			<xsl:attribute name="href"><xsl:value-of select="/ow:wiki/ow:scriptname"/>?p=Help</xsl:attribute>
			<xsl:attribute name="title">Guidance on how to use and edit this wiki</xsl:attribute>
			Help
		</a>
	</li>
</xsl:template>

<xsl:template name="menu_section_toolbox">
	<li id="t-whatlinkshere">
		<a href="{ow:scriptname}?a=fullsearch&amp;txt={$name}&amp;fromtitle=true">
			<xsl:attribute name="title">List of all pages containing links to this page [j]</xsl:attribute>
			<xsl:attribute name="accesskey">j</xsl:attribute>
			What links here
		</a>
	</li>
	<li id="t-print">
		<a href="/w/index.php?title=Libretto&amp;printable=yes">
			<xsl:attribute name="href"><xsl:value-of select="/ow:wiki/ow:scriptname"/>?p=<xsl:value-of select="$name"/>&amp;a=print&amp;revision=<xsl:value-of select="ow:change/@revision"/></xsl:attribute>
			<xsl:attribute name="rel">alternate</xsl:attribute>
			<xsl:attribute name="title">Printable version of this page [p]</xsl:attribute>
			<xsl:attribute name="accesskey">p</xsl:attribute>
			Printable version
		</a>
	</li>	
	<li id="t-viewxml">
		<a>
			<xsl:attribute name="href">
				<xsl:value-of select="/ow:wiki/ow:scriptname"/>?p=<xsl:value-of select="$name"/>&amp;a=xml&amp;revision=<xsl:value-of select="ow:change/@revision"/>
			</xsl:attribute>
			View XML
		</a>
	</li>
</xsl:template>

<xsl:template name="footer_list">
	<ul id="f-list">
		<xsl:call-template name="footer_list_lastmod" />
	</ul>
</xsl:template>

<xsl:template name="footer_list_lastmod">
	<li id="lastmod">
        <xsl:if test="not(ow:page/@changes='0')">
              This page was last modified on  <xsl:value-of select="ow:formatLongDate(string(ow:page/ow:change/ow:date))"/>, at <xsl:value-of select="ow:formatTime(string(ow:page/ow:change/ow:date))"/>
        </xsl:if>
	</li>
</xsl:template>

</xsl:stylesheet>