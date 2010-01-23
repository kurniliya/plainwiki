<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:msxsl="urn:schemas-microsoft-com:xslt"
                xmlns:ow="http://openwiki.com/2001/OW/Wiki"
                xmlns="http://www.w3.org/1999/xhtml"
                extension-element-prefixes="msxsl ow"
                exclude-result-prefixes=""
                version="1.0">
<xsl:output method="xml" indent="yes" omit-xml-declaration="yes"/>

<xsl:template name="StatCounter">
	<xsl:comment>Start of StatCounter Code</xsl:comment>
	
	<script type="text/javascript">
		<xsl:text disable-output-escaping="yes">
			/*&lt;![CDATA[*/</xsl:text>
			var sc_project=5410615; 
			var sc_invisible=1; 
			var sc_partition=47; 
			var sc_click_stat=1; 
			var sc_security="ba59c5fd"; 
		<xsl:text disable-output-escaping="yes">
			/*]]&gt;*/</xsl:text>
	</script>
	
	<script type="text/javascript" src="http://www.statcounter.com/counter/counter_xhtml.js">
	</script>
	<noscript>
		<div	class="statcounter">
			<a title="hits counter" class="statcounter" href="http://www.statcounter.com/">
				<img	class="statcounter" src="http://c.statcounter.com/5410615/0/ba59c5fd/1/"	alt="hits counter" />
			</a>
		</div>
	</noscript>
	
	<xsl:comment>End of StatCounter Code</xsl:comment>
</xsl:template>

</xsl:stylesheet>