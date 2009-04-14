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
  <meta name="keywords" content="math, partial differential equations, mephi"/>
  <meta name="description" content="NEQwiki - encyclopedia of nonlinear differential equations"/>
  <meta name="ROBOTS" content="INDEX,FOLLOW"/>
  <meta name="MSSmartTagsPreventParsing" content="true"/>
  <title><xsl:value-of select="ow:title"/> - <xsl:value-of select="ow:page/ow:link"/></title>
<!--
  <link rel="stylesheet" type="text/css" href="ow/css/ow.css" />
-->
  <link rel="stylesheet" type="text/css" href="ow/css/monobook/main.css?207xx" />
  <script type="text/javascript" src="ow/js/wikibits.js?207xx">
	<!-- wikibits js -->
  </script>
  </head>
</xsl:template>

<xsl:template name="brandingImage">
    <a href="{/ow:wiki/ow:frontpage/@href}"><img src="{/ow:wiki/ow:imagepath}/logo.gif" align="right" border="0" alt="NEQwiki" /></a>
</xsl:template>

<xsl:template name="poweredBy">
    <a href="http://openwiki.com"><img src="{/ow:wiki/ow:imagepath}/poweredby.gif" width="88" height="31" border="0" alt="" /></a>
</xsl:template>

<xsl:template name="validatorButtons">
    <a href="http://validator.w3.org/check/referer"><img src="{/ow:wiki/ow:imagepath}/valid-xhtml10.gif" alt="Valid XHTML 1.0!" width="88" height="31" border="0" /></a>
    <a href="http://jigsaw.w3.org/css-validator/validator?uri={/ow:wiki/ow:location}ow.css"><img src="{/ow:wiki/ow:imagepath}/valid-css.gif" alt="Valid CSS!" width="88" height="31" border="0" /></a>
</xsl:template>

<xsl:template name="menu_column">
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

	<div class="portlet" id="p-tb">
		<h5>Toolbox</h5>
		<div class="pBody">
			<ul>
				<xsl:call-template name="menu_section_toolbox" />
			</ul>
		</div>
	</div>
</xsl:template>

<xsl:template name="menu_section_navigation">
	<li id="n-mainpage-description">
		<a>
			<xsl:attribute name="href">
				<xsl:value-of select="/ow:wiki/ow:frontpage/@href"/>
			</xsl:attribute>
			Main page
		</a>
	</li>
	<li id="n-randompage">
		<a>
			<xsl:attribute name="href">
				ow.asp?RandomPage
			</xsl:attribute>
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

<xsl:template name="menu_section_toolbox">
	<li id="t-viewxml">
		<a>
			<xsl:attribute name="href">
				<xsl:value-of select="/ow:wiki/ow:scriptname"/>?p=<xsl:value-of select="$name"/>&amp;a=xml&amp;revision=<xsl:value-of select="ow:change/@revision"/>
			</xsl:attribute>
			View XML
		</a>
	</li>
</xsl:template>

</xsl:stylesheet>