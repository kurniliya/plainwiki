<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:msxsl="urn:schemas-microsoft-com:xslt"
                xmlns:ow="http://openwiki.com/2001/OW/Wiki"
                xmlns="http://www.w3.org/1999/xhtml"
                extension-element-prefixes="msxsl ow"
                exclude-result-prefixes=""
                version="1.0">

<xsl:output method="xml" indent="no" omit-xml-declaration="yes"/>

<xsl:template name="nofollow_head">
	<head>
		<meta http-equiv="Content-Type" content="application/xhtml+xml; charset={@encoding};" />
		<meta name="generator" content="OpenWiki 0.78" />
		<meta name="robots" content="noindex,nofollow" />
		<meta name="keywords" content="math, partial differential equations, mephi" />
		<meta name="description" content="NEQwiki - the encyclopedia of nonlinear differential equations" />
		<meta name="MSSmartTagsPreventParsing" content="true" />
		<link rel="shortcut icon" href="{/ow:wiki/ow:imagepath}/favicon.ico" />		
		<title>
			<xsl:value-of select="ow:page/ow:link"/> - <xsl:value-of select="ow:title"/>
		</title>
		
		<link rel="alternate" type="application/rss+xml" href="{/ow:wiki/ow:scriptname}?a=rss" title="Recent changes" />		
		
		<!--<link rel="stylesheet" href="ow/css/common/shared.css?207xx" type="text/css" media="screen" />-->
		<link rel="stylesheet" href="ow/css/common/commonPrint.css?207xx" type="text/css" media="print" />
		<!--<link rel="stylesheet" href="ow/css/monobook/main.css?207xx" type="text/css" media="screen" />-->
		<link rel="stylesheet" href="ow/css/ow.css" type="text/css" media="screen" />
		
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
		<xsl:text disable-output-escaping="yes">
			&lt;script type="text/javascript" src="ow/js/common/IEFixes.js?207xx">
			&lt;/script>
		</xsl:text>
		<meta http-equiv="imagetoolbar" content="no" />
		<xsl:text disable-output-escaping="yes">&lt;![endif]--&gt;
		</xsl:text>
	</head>
</xsl:template>

</xsl:stylesheet>