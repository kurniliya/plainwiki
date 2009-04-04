<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:msxsl="urn:schemas-microsoft-com:xslt"
                xmlns:ow="http://openwiki.com/2001/OW/Wiki"
                xmlns="http://www.w3.org/1999/xhtml"
                extension-element-prefixes="msxsl ow"
                exclude-result-prefixes=""
                version="1.0">
<xsl:output method="xml" indent="no" omit-xml-declaration="yes"/>

<!-- if editOnDblCklick='1' then double click on page will change it into edit mode. -->
<xsl:variable name="editOnDblCklick" select="'0'" />

<!-- if showThirdLineInFooter='1' then Print this page, View XML and Find page links will be shown in page footer. -->
<xsl:variable name="showThirdLineInFooter" select="'0'" />

<!-- if showBookmarksInFooter='1' then bookmarks are shown also in footer. -->
<xsl:variable name="showBookmarksInFooter" select="'0'" />

<!-- if showEditLinkOnTop='1' then Edit this page link is shown on top. -->
<xsl:variable name="showEditLinkOnTop" select="'0'" />
<xsl:variable name="showValidatorButtons" select="'1'" />
<xsl:variable name="showPoweredBy" select="'1'" />

</xsl:stylesheet>