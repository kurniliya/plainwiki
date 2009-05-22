<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:msxsl="urn:schemas-microsoft-com:xslt"
                xmlns:ow="http://openwiki.com/2001/OW/Wiki"
                extension-element-prefixes="msxsl ow"
                exclude-result-prefixes=""
                version="1.0">
<xsl:output method="xml" indent="no" omit-xml-declaration="yes"/>

<xsl:template name="pi">
  <xsl:text disable-output-escaping="yes">&lt;?xml version="1.0" encoding="</xsl:text><xsl:value-of select="@encoding"/><xsl:text disable-output-escaping="yes">"?>
&lt;?xml-stylesheet type="text/xsl" href="ow/xsl/pmathml.xsl"?&gt;
&lt;!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1 plus MathML 2.0 plus SVG 1.1//EN" 
		"http://www.w3.org/2002/04/xhtml-math-svg/xhtml-math-svg-flat.dtd">
</xsl:text>
</xsl:template>            
                
</xsl:stylesheet>
