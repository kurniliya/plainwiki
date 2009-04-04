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

<!-- if editOnDblCklick='1' then double click on page will change it into edit mode. -->
<xsl:variable name="editOnDblCklick" select="'0'" />

<!-- if showThirdLineInFooter='1' then Print this page, View XML and Find page links will be shown in page footer. -->
<xsl:variable name="showThirdLineInFooter" select="'0'" />

<!-- if showBookmarksInFooter='1' then bookmarks are shown also in footer. -->
<xsl:variable name="showBookmarksInFooter" select="'0'" />

<!-- if showEditLinkOnTop='1' then Edit this page link is shown on top. -->
<xsl:variable name="showEditLinkOnTop" select="'0'" />

<xsl:template name="head">
  <head>
  <meta http-equiv="Content-Type" content="application/xhtml+xml; charset={@encoding};" />
  <meta name="keywords" content="math, partial differential equations, mephi"/>
  <meta name="description" content="NEQwiki - encyclopedia of nonlinear differential equations"/>
  <meta name="ROBOTS" content="INDEX,FOLLOW"/>
  <meta name="MSSmartTagsPreventParsing" content="true"/>
  <title><xsl:value-of select="ow:title"/> - <xsl:value-of select="ow:page/ow:link"/></title>
  <link rel="stylesheet" type="text/css" href="ow/css/ow.css" />
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


</xsl:stylesheet>