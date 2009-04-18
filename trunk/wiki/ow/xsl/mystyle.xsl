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

<div class="portlet" id="p-logo">
	<a style="background-image: url(ow/images/logo.gif);" href="http://www.primat.mephi.ru" title="Visit the main page [z]" accesskey="z">
	</a>
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
			<xsl:attribute name="href"><xsl:value-of select="/ow:wiki/ow:scriptname"/>?p=<xsl:value-of select="$name"/>&amp;a=edit<xsl:if test="ow:page/@revision">&amp;revision=<xsl:value-of select="ow:page/@revision"/></xsl:if></xsl:attribute>
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
	<li id="n-todolist">
		<a>
			<xsl:attribute name="href"><xsl:value-of select="/ow:wiki/ow:scriptname"/>?ToDo</xsl:attribute>
			To do list
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
		<a>
			<xsl:attribute name="href">
				<xsl:value-of select="/ow:wiki/ow:scriptname"/>?p=<xsl:value-of select="$name"/>&amp;a=print&amp;revision=<xsl:value-of select="ow:page/@revision"/>
			</xsl:attribute>
			<xsl:attribute name="rel">alternate</xsl:attribute>
			<xsl:attribute name="title">Printable version of this page [p]</xsl:attribute>
			<xsl:attribute name="accesskey">p</xsl:attribute>
			Printable version
		</a>
	</li>	
	<li id="t-viewxml">
		<a>
			<xsl:attribute name="href">
				<xsl:value-of select="/ow:wiki/ow:scriptname"/>?p=<xsl:value-of select="$name"/>&amp;a=xml&amp;revision=<xsl:value-of select="ow:page/@revision"/>
			</xsl:attribute>
			View XML
		</a>
	</li>
	<li id="t-sandbox">
		<a>
			<xsl:attribute name="href">
				<xsl:value-of select="/ow:wiki/ow:scriptname"/>?Sandbox
			</xsl:attribute>
			Sandbox
		</a>
	</li>
	<li id="t-sandbox">
		<a>
			<xsl:attribute name="href">
				<xsl:value-of select="/ow:wiki/ow:scriptname"/>?CreatePage
			</xsl:attribute>
			Create new page
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
		<xsl:if test="ow:page/@changes">
			<xsl:if test="not(ow:page/@changes='0')">
				  This page was last modified on  <xsl:value-of select="ow:formatLongDate(string(ow:page/ow:change/ow:date))"/>, at <xsl:value-of select="ow:formatTime(string(ow:page/ow:change/ow:date))"/>
			</xsl:if>
		</xsl:if>
	</li>
</xsl:template>

</xsl:stylesheet>