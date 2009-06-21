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
			addButton("ow/skins/common/images/button_bold.png","Bold text","**","**","Bold text","mw-editbutton-bold");
			addButton("ow/skins/common/images/button_italic.png","Italic text","//","//","Italic text","mw-editbutton-italic");
			addButton("ow/skins/common/images/button_underline.png","Underlined text","__","__","Underlined text","mw-editbutton-underline");
			addButton("ow/skins/common/images/button_strike.png","Strike through text","--","--","Struck out text","mw-editbutton-strike");
			addButton("ow/skins/common/images/button_small.png","Small text","\x3csmall\x3e","\x3c/small\x3e","Small text","mw-editbutton-small");
			addButton("ow/skins/common/images/button_big.png","Big text","\x3cbig\x3e","\x3c/big\x3e","Big text","mw-editbutton-big");
			addButton("ow/skins/common/images/button_tt.png","Teletype text","\x3ctt\x3e","\x3c/tt\x3e","Teletype text", "mw-editbutton-tt");			
			addButton("ow/skins/common/images/button_sup_letter.png","Superscript","^^","^^","Superscript text","mw-editbutton-sup-letter");
			addButton("ow/skins/common/images/button_sub_letter.png","Subscript","vv","vv","Subscript text","mw-editbutton-sub-letter");			
			addButton("ow/skins/common/images/button_link.png","Internal link","[","]","WikiPageName LinkTitle","mw-editbutton-link");
			addButton("ow/skins/common/images/button_extlink.png","External link (remember http:// prefix)","[","]","http://www.example.com link title","mw-editbutton-extlink");
			addButton("ow/skins/common/images/button_headline.png","Level 2 headline","\n== "," ==\n","Headline text","mw-editbutton-headline");
			addButton("ow/skins/common/images/button_headline2.png","Level 3 headline","\n=== "," ===\n","Headline text","mw-editbutton-headline2");
//			addButton("ow/skins/common/images/button_image.png","Embedded file","[[File:","]]","Example.jpg","mw-editbutton-image");
//			addButton("ow/skins/common/images/button_media.png","File link","[[Media:","]]","Example.ogg","mw-editbutton-media");
			addButton("ow/skins/common/images/button_math.png","Mathematical formula (MathML)","\x3cmath\x3e","\x3c/math\x3e","Insert formula here","mw-editbutton-math");
			addButton("ow/skins/common/images/button_nowiki.png","Ignore wiki formatting","\x3cnowiki\x3e","\x3c/nowiki\x3e","Insert non-formatted text here","mw-editbutton-nowiki");
//			addButton("ow/skins/common/images/button_sig.png","Your signature with timestamp","--~~~~","","","mw-editbutton-signature");
			addButton("ow/skins/common/images/button_hr.png","Horizontal line (use sparingly)","\n----\n","","","mw-editbutton-hr");
			addButton("ow/skins/common/images/button_enter.png","Line break","\x3cbr /\x3e","","","mw-editbutton-enter");
			addButton("ow/skins/common/images/button_definition_list.png","Definition list","  ;Term : ", "", "Definition", "mw-editbutton-definition-list");
			addButton("ow/skins/common/images/button_numbered_list.png","Numbered list","  1. ", "", "List item", "mw-editbutton-numbered-list");
			addButton("ow/skins/common/images/button_bulleted_list.png","Bulleted list","  * ", "", "List item", "mw-editbutton-bulleted-list");
			addButton("ow/skins/common/images/button_shifting.png","Indent text","  : ", "", "Indented text", "mw-editbutton-shifting");
			addButton("ow/skins/common/images/button_align_left.png","Align left","\x3cdiv style=\"text-align: left; direction: ltr;\"\x3e","\x3c/div\x3e","Left-aligned text","mw-editbutton-alignleft");
			addButton("ow/skins/common/images/button_center.png","Centred text","\x3cdiv style=\"text-align: center;\"\x3e","\x3c/div\x3e","Centred text","mw-editbutton-center");
			addButton("ow/skins/common/images/button_align_right.png","Align right","\x3cdiv style=\"text-align: right; direction: ltr; margin-left: 1em;\"\x3e","\x3c/div\x3e","Right-aligned text","mw-editbutton-alignright");			
			addButton("ow/skins/common/images/button_font_color.png","Coloured text","\x3cspan style=\"color: some-colour;\"\x3e","\x3c/span\x3e","Coloured text","mw-editbutton-fontcolor");
			addButton("ow/skins/common/images/button_comment.png","Comment","\x3c!--","--\x3e","Comment", "mw-editbutton-comment");
			addButton("ow/skins/common/images/button_code.png","Insert code","\x3ccode\x3e\n","\n\x3c/code\x3e\n","Code", "mw-editbutton-code");
			addButton("ow/skins/common/images/button_pre.png","Pre formatted text","\x3cpre\x3e","\x3c/pre\x3e","Pre formatted text", "mw-editbutton-pre");
			addButton("ow/skins/common/images/button_blockquote.png","Block quote text","\x3cblockquote style=\"margin: 1em 8em 1em 2em;\"\x3e\n\x3cp\x3e","\x3c/p\x3e\n\x3cp style=\"margin-left: 2em;\"\x3e\x3ccite style=\"font-style: normal;\"\x3e—Author\x3c/cite\x3e\x3c/p\x3e\n\x3c/blockquote\x3e", "Citation", "mw-editbutton-blockquote");
			addButton("ow/skins/common/images/button_redirect.png","Redirect","#REDIRECT ","","WikiName","mw-editbutton-redirect");
			addButton("ow/skins/common/images/button_table_row.png","Table row","|| "," ||","Table row", "mw-editbutton-table-row");
			addButton("ow/skins/common/images/button_nbsp.png","Non breaking space","&#160;","","","mw-editbutton-nbsp");
			addButton("ow/skins/common/images/button_category03.png","Category","[[:Category","]]", "SampleName", "mw-editbutton-category03");
			<xsl:text disable-output-escaping="yes">
				]]&gt;</xsl:text>
		</script>
	</div>
</xsl:template>	

</xsl:stylesheet>