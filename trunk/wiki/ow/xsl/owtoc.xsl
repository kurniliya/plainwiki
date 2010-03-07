<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:msxsl="urn:schemas-microsoft-com:xslt"
                xmlns:ow="http://openwiki.com/2001/OW/Wiki"               
                xmlns="http://www.w3.org/1999/xhtml"
                extension-element-prefixes="msxsl ow"
                exclude-result-prefixes=""
                version="1.0">
<xsl:output method="xml" indent="no" omit-xml-declaration="yes" />

<!-- ==================== handles the openwiki-toc element ==================== -->
<xsl:template match="ow:toc_root">
	<xsl:choose>
		<xsl:when test="@align='right'">
			<table cellspacing="0" cellpadding="0" style="clear: right; margin-bottom: .5em; float: right; padding: .5em 0 .8em 1.4em; background: none; width: auto;">
				<tr>
					<td>
						<table id="toc" class="toc" summary="Contents">
							<tr>
								<td>
									 <div id="toctitle">
										<h2>Contents</h2> 
									</div>
									<xsl:apply-templates select="./ow:toc" />
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</xsl:when>
        <xsl:otherwise>
			<table id="toc" class="toc" summary="Contents">
				<tr>
					<td>
						 <div id="toctitle">
							<h2>Contents</h2> 
						</div>
						<xsl:apply-templates select="./ow:toc" />
					</td>
				</tr>
			</table>
		</xsl:otherwise>
	</xsl:choose>
</xsl:template>

<xsl:template match="ow:toc" name="toc">
	<xsl:choose>
		<xsl:when test="@mode='indented'">
			<ul>
				<xsl:for-each select="./*">
					<xsl:choose>
						<xsl:when test="number">
							<li class="toclevel-{level}">
								<a href="#h{number}">
									<span class="toctext">
										<xsl:value-of select="text" disable-output-escaping="yes" />
									</span>	
								</a>
							</li>
						</xsl:when>
						<xsl:otherwise>
							<li>
								<xsl:call-template name="toc" />
							</li>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:for-each>
			</ul>
		</xsl:when>
		<xsl:otherwise>
			<ul>
				<xsl:for-each select="./*">
					<xsl:choose>
						<xsl:when test="number">
							<li class="toclevel-{level}">
								<a href="#h{number}">
									<span class="tocnumber">
										<xsl:value-of select="number_trace" disable-output-escaping="yes" />
										<xsl:text> </xsl:text>
									</span>					
									<span class="toctext">
										<xsl:value-of select="text" disable-output-escaping="yes" />
									</span>	
								</a>
							</li>
						</xsl:when>
						<xsl:otherwise>
							<li>
								<xsl:call-template name="toc" />
							</li>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:for-each>
			</ul>
		</xsl:otherwise>
	</xsl:choose>
</xsl:template>   

</xsl:stylesheet>