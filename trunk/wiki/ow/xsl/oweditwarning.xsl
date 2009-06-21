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
					You have not provided your username. Saving revision now will cause your IP address to be recorded publicly in this page's edit history. 
				</td>
			</tr>
			</tbody>
		</table>
	</div>
</xsl:template>

</xsl:stylesheet>