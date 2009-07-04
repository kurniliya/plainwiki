<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:msxsl="urn:schemas-microsoft-com:xslt"
                xmlns:ow="http://openwiki.com/2001/OW/Wiki"               
                xmlns="http://www.w3.org/1999/xhtml"
                extension-element-prefixes="msxsl ow"
                exclude-result-prefixes=""
                version="1.0">
<xsl:output method="xml" indent="no" omit-xml-declaration="yes" />

<xsl:template name="edit_warning">
	<div id="mw-anon-edit-warning">
		<table id="anoneditwarning" class="plainlinks fmbox fmbox-editnotice" style="">	
			<tbody>
			<tr>
				<td class="mbox-image">
					<a href="/wiki/File:Imbox_notice.png" class="image" title="Imbox notice.png">
						<img alt="" src="ow/images/icons/Imbox_notice.png" />
					</a>
				</td>
				<td class="mbox-text" style="">
					You have not provided your <a>
			<xsl:attribute name="href"><xsl:value-of select="/ow:wiki/ow:scriptname"/>?Special:UserPreferences</xsl:attribute>
			<xsl:attribute name="title">You are encouraged to log in; however, it is not mandatory. [o]</xsl:attribute>
			<xsl:attribute name="accesskey">o</xsl:attribute>
			username</a>. Saving revision now will cause your IP address to be recorded publicly in <a>
			<xsl:attribute name="href"><xsl:value-of select="/ow:wiki/ow:scriptname"/>?p=<xsl:value-of select="$name"/>&amp;a=changes</xsl:attribute>
			<xsl:attribute name="title">Past versions of this page [h]</xsl:attribute>
			<xsl:attribute name="accesskey">h</xsl:attribute>
			this page's edit history</a>. 
				</td>
			</tr>
			</tbody>
		</table>
	</div>
</xsl:template>

</xsl:stylesheet>