<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:msxsl="urn:schemas-microsoft-com:xslt"
                xmlns:ow="http://openwiki.com/2001/OW/Wiki"               
                xmlns="http://www.w3.org/1999/xhtml"
                extension-element-prefixes="msxsl ow"
                exclude-result-prefixes=""
                version="1.0">
<xsl:output method="xml" indent="no" omit-xml-declaration="yes" />

<xsl:template name="edit_buttons_toolbar">
	<div id='toolbar'>
		<script type='text/javascript'>
			<xsl:text disable-output-escaping="yes">
				&lt;![CDATA[</xsl:text>
			addButton("ow/skins/common/images/button_bold.png","Bold text","\'\'\'","\'\'\'","Bold text","mw-editbutton-bold");
			addButton("ow/skins/common/images/button_italic.png","Italic text","\'\'","\'\'","Italic text","mw-editbutton-italic");
			addButton("ow/skins/common/images/button_link.png","Internal link","[","]","WikiPageName LinkTitle","mw-editbutton-link");
			addButton("ow/skins/common/images/button_extlink.png","External link (remember http:// prefix)","[","]","http://www.example.com link title","mw-editbutton-extlink");
			addButton("ow/skins/common/images/button_headline.png","Level 2 headline","\n== "," ==\n","Headline text","mw-editbutton-headline");
//			addButton("ow/skins/common/images/button_image.png","Embedded file","[[File:","]]","Example.jpg","mw-editbutton-image");
//			addButton("ow/skins/common/images/button_media.png","File link","[[Media:","]]","Example.ogg","mw-editbutton-media");
			addButton("ow/skins/common/images/button_math.png","Mathematical formula (MathML)","\x3cmath\x3e","\x3c/math\x3e","Insert formula here","mw-editbutton-math");
			addButton("ow/skins/common/images/button_nowiki.png","Ignore wiki formatting","\x3cnowiki\x3e","\x3c/nowiki\x3e","Insert non-formatted text here","mw-editbutton-nowiki");
//			addButton("ow/skins/common/images/button_sig.png","Your signature with timestamp","--~~~~","","","mw-editbutton-signature");
			addButton("ow/skins/common/images/button_hr.png","Horizontal line (use sparingly)","\n----\n","","","mw-editbutton-hr");
			<xsl:text disable-output-escaping="yes">
				]]&gt;</xsl:text>
		</script>
	</div>
</xsl:template>	

</xsl:stylesheet>