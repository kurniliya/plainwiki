<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:msxsl="urn:schemas-microsoft-com:xslt"
                xmlns:ow="http://openwiki.com/2001/OW/Wiki"
                xmlns="http://www.w3.org/1999/xhtml"
                extension-element-prefixes="msxsl ow"
                exclude-result-prefixes=""
                version="1.0">
<xsl:output method="xml" indent="yes" omit-xml-declaration="yes"/>

<xsl:include href="owedittoolbar.xsl"/>
<xsl:include href="owrecaptcha.xsl"/>

<xsl:template name="ExternalJS">
	<xsl:comment>External JSs</xsl:comment>
		<xsl:text disable-output-escaping="yes">
			&lt;script type="text/javascript" src="ow/js/wikibits.js">
			&lt;/script>
		</xsl:text>
		<xsl:text disable-output-escaping="yes"> 
			&lt;script type="text/javascript" src="ow/js/infobar.js">
			&lt;/script>
		</xsl:text>		
	<xsl:comment>End of external JSs</xsl:comment>
	
<!--	<xsl:comment>Inline JSs</xsl:comment>	
		<script type="text/javascript" charset="{/ow:wiki/@encoding}">
			<xsl:text disable-output-escaping="yes">
				/*&lt;![CDATA[*/
				if (window.toggleToc) { showTocToggle(); } 
				/*]]&gt;*/ 
			</xsl:text>
		</script>	
	<xsl:comment>End of inline JSs</xsl:comment>-->
</xsl:template>

<xsl:template name="EditJS">
	<xsl:comment>Edit page JSs</xsl:comment>
		<xsl:text disable-output-escaping="yes">
			&lt;script type="text/javascript" src="ow/js/recaptcha_documentwrite.js">
			&lt;/script>
			&lt;script type="text/javascript" src="ow/js/edit.js?">
			&lt;/script>
		</xsl:text>
		<xsl:call-template name="edit_buttons_toolbar"/>
		<xsl:call-template name="recaptchaJS"/>
	<xsl:comment>End of edit page JSs</xsl:comment>
</xsl:template>

</xsl:stylesheet>