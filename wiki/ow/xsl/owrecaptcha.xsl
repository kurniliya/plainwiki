<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:msxsl="urn:schemas-microsoft-com:xslt"
                xmlns:ow="http://openwiki.com/2001/OW/Wiki"               
                xmlns="http://www.w3.org/1999/xhtml"
                extension-element-prefixes="msxsl ow"
                exclude-result-prefixes=""
                version="1.0">
<xsl:output method="xml" indent="no" omit-xml-declaration="yes"/>

<xsl:variable name="recaptcha_publickey" select="'6Lea0wYAAAAAADggQdYAgmRdguWVCRWeqYbvumWG'" />

<!--<xsl:template name="showRecapthca">
	<xsl:if test="/ow:wiki/ow:protection='captcha'">
		<div id="recaptcha_holder" />
	</xsl:if>		
</xsl:template>-->

<xsl:template name="recaptchaJS">
	<xsl:if test="/ow:wiki/ow:protection='captcha'">
		<script type="text/javascript"
   src="http://api.recaptcha.net/challenge?k={$recaptcha_publickey}&amp;error={ow:captcha_error}"></script>
	</xsl:if>
</xsl:template>

</xsl:stylesheet>