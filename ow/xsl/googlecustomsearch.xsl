<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:msxsl="urn:schemas-microsoft-com:xslt"
                xmlns:ow="http://openwiki.com/2001/OW/Wiki"
                xmlns="http://www.w3.org/1999/xhtml"
                extension-element-prefixes="msxsl ow"
                exclude-result-prefixes=""
                version="1.0">
<xsl:output method="xml" indent="no" omit-xml-declaration="yes"/>

<xsl:template name="GoogleCustomSearch">
<!--
	<xsl:text disable-output-escaping="yes">
		&lt;object data="./static/googlecse/googlecse.html" type="text/html" width="100%" height="100%" >
		&lt;/object>
	</xsl:text>
-->
		<form action="http://www.google.com/cse" id="cse-search-box" target = "_blank">
		  <div>
		    <input type="hidden" name="cx" value="007191275890417134042:xxroumdgyhy" />
		    <input type="hidden" name="ie" value="UTF-8" />
		    <input type="text" name="q" id="q" class="searchButton" size="31" />
		    <input type="submit" name="sa" id="sa" class="searchButton" value="Google" />
		  </div>
		</form>
		
		<script type="text/javascript" src="http://www.google.com/jsapi"></script>
		<script type="text/javascript">google.load("elements", "1", {packages: "transliteration"});</script>
		<script type="text/javascript" src="http://www.google.com/coop/cse/t13n?form=cse-search-box&amp;t13n_langs=en"></script>
		
		<script type="text/javascript" src="http://www.google.com/coop/cse/brand?form=cse-search-box&amp;lang=en"></script>		
</xsl:template>

</xsl:stylesheet>