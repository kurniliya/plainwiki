<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:msxsl="urn:schemas-microsoft-com:xslt"
                xmlns:ow="http://openwiki.com/2001/OW/Wiki"
                xmlns="http://www.w3.org/1999/xhtml"
                extension-element-prefixes="msxsl ow"
                exclude-result-prefixes=""
                version="1.0">
<xsl:output method="xml" indent="no" omit-xml-declaration="yes"/>

<xsl:template name="GoogleAnalytics">
	<script src="http://www.google-analytics.com/ga.js" type="text/javascript">
	</script>
	<script type="text/javascript">
		try {
			var pageTracker = _gat._getTracker("UA-12156025-1");
			pageTracker._trackPageview();
		} catch(err) {}
	</script>	
</xsl:template>

</xsl:stylesheet>