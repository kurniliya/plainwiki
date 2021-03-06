<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:msxsl="urn:schemas-microsoft-com:xslt"
                xmlns:ow="http://openwiki.com/2001/OW/Wiki"               
                xmlns="http://www.w3.org/1999/xhtml"
                extension-element-prefixes="msxsl ow"
                exclude-result-prefixes=""
                version="1.0">
<xsl:output method="xml" indent="yes" omit-xml-declaration="yes" />

<xsl:include href="owpi.xsl"/>
<xsl:include href="owinc.xsl"/>
<xsl:include href="owattach.xsl"/>
<xsl:include href="owconfig.xsl"/>
<xsl:include href="mystyle.xsl"/>
<xsl:include href="owhead.xsl"/>
<xsl:include href="owtoc.xsl"/>
<!--<xsl:include href="owrecaptcha.xsl"/>-->
<!--<xsl:include href="owedittoolbar.xsl"/>-->
<xsl:include href="oweditwarning.xsl"/>
<!--<xsl:include href="googleanalytics.xsl"/>-->
<xsl:include href="statcounter.xsl"/>
<xsl:include href="owjs.xsl"/>

<xsl:variable name="name" select="ow:urlencode(string(/ow:wiki/ow:page/@name))" />

<xsl:template match="*">
  <xsl:element name="{name()}">
    <xsl:copy-of select="@*"/>
    <xsl:apply-templates/>
  </xsl:element>
</xsl:template>

<xsl:template match="processing-instruction()|comment()|text()">
  <xsl:copy>
    <xsl:apply-templates/>
  </xsl:copy>
</xsl:template>

<!-- ridiculous! IE processes <br></br> differently compared to <br /> ! -->
<xsl:template match="br">
  <br />
</xsl:template>

<xsl:template match="big">
  <b><big><xsl:apply-templates/></big></b>
</xsl:template>

<xsl:template match="table">
  <table cellspacing="0" cellpadding="2" border="1" width="100%">
    <xsl:apply-templates/>
  </table>
</xsl:template>

<!-- ==================== used to do client-side transformation ==================== -->
<xsl:template match="/ow:wiki">
  <xsl:choose>
    <xsl:when test="@mode='view'">
      <xsl:apply-templates select="." mode="view"/>
    </xsl:when>
    <xsl:when test="@mode='edit'">
      <xsl:apply-templates select="." mode="edit"/>
    </xsl:when>
    <xsl:when test="@mode='print'">
      <xsl:apply-templates select="." mode="print"/>
    </xsl:when>
    <xsl:when test="@mode='naked'">
      <xsl:apply-templates select="." mode="naked"/>
    </xsl:when>
    <xsl:when test="@mode='diff'">
      <xsl:apply-templates select="." mode="diff"/>
    </xsl:when>
    <xsl:when test="@mode='changes'">
      <xsl:apply-templates select="." mode="changes"/>
    </xsl:when>
    <xsl:when test="@mode='titlesearch'">
      <xsl:apply-templates select="." mode="titlesearch"/>
    </xsl:when>
    <xsl:when test="@mode='fullsearch'">
      <xsl:apply-templates select="." mode="fullsearch"/>
    </xsl:when>
    <xsl:when test="@mode='login'">
      <xsl:apply-templates select="." mode="login"/>
    </xsl:when>
    <xsl:when test="@mode='attach'">
      <xsl:apply-templates select="." mode="attach"/>
    </xsl:when>
    <xsl:when test="@mode='attachchanges'">
      <xsl:apply-templates select="." mode="attachchanges"/>
    </xsl:when>
    <xsl:when test="@mode='embedded'">
      <xsl:apply-templates select="." mode="embedded"/>
    </xsl:when>
  </xsl:choose>
</xsl:template>

<xsl:template match="/ow:wiki" mode="view">
	<xsl:call-template name="pi"/>
	<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" dir="ltr">
		<xsl:call-template name="head"/>
		<body class="mediawiki ltr ns-0 ns-subject skin-monobook" onload="window.defaultStatus='{$brandingText}'">
			<div id="globalWrapper">
				<xsl:if test="$editOnDblCklick='1'">
					<xsl:attribute name="ondblclick">location.href='<xsl:value-of select="ow:scriptname"/>?p=<xsl:value-of select="$name"/>&amp;a=edit<xsl:if test='ow:page/@revision'>&amp;revision=<xsl:value-of select="ow:page/@revision"/></xsl:if>'</xsl:attribute>
				</xsl:if>        
<!--				
				<xsl:call-template name="brandingImage"/>
				<h1>
					<a href="{/ow:wiki/ow:frontpage/@href}"><xsl:value-of select="$mainPageHeading"/></a>
				</h1>
-->
				<div id="column-content">
					<div id="content">
						<xsl:apply-templates select="ow:page"/>
					</div>
				</div>
				<div id="column-one">
					<xsl:call-template name="menu_column" />					
				</div>
				<div class="visualClear"></div>
				<div id="footer">
					<xsl:call-template name="poweredBy" />
					<xsl:call-template name="copyright_ico" />					
					<xsl:call-template name="footer_list" />
				</div>
			</div>
			
			<xsl:call-template name="ExternalJS" />	
			<xsl:call-template name="StatCounter" />			
		 </body>
	 </html>
</xsl:template>

<xsl:template match="ow:page">
	<a id="top"></a>
    <xsl:if test="/ow:wiki/ow:userpreferences/ow:editlinkontop">

      <xsl:if test="$showEditLinkOnTop='1'">    
          <a class="same"><xsl:attribute name="href"><xsl:value-of select="/ow:wiki/ow:scriptname"/>?p=<xsl:value-of select="$name"/>&amp;a=edit<xsl:if test="@revision">&amp;revision=<xsl:value-of select="@revision"/></xsl:if></xsl:attribute>Edit</a> this page
<!--
          <xsl:if test="not(@changes='0')">
              <font size="-2">(This page was last modified on  <xsl:value-of select="ow:formatLongDate(string(ow:change/ow:date))"/>)</font>
          </xsl:if>
-->
          <br />
      </xsl:if>        
    </xsl:if>
<!--    
    <xsl:if test="/ow:wiki/ow:userpreferences/ow:bookmarksontop">
      <xsl:if test="not(/ow:wiki/ow:userpreferences/ow:bookmarks='None')"> 
        <xsl:apply-templates select="/ow:wiki/ow:userpreferences/ow:bookmarks"/>
      </xsl:if>
    </xsl:if>

    <hr noshade="noshade" size="1" />
    
    <xsl:apply-templates select="../ow:trail"/>
-->

    <xsl:if test="../ow:redirectedfrom">
        <b>Redirected from <a title="Edit this page"><xsl:attribute name="href"><xsl:value-of select="/ow:wiki/ow:scriptname"/>?a=edit&amp;p=<xsl:value-of select="ow:urlencode(string(../ow:redirectedfrom/@name))"/></xsl:attribute><xsl:value-of select="../ow:redirectedfrom/text()"/></a></b>
        <p />
    </xsl:if>
    <xsl:if test="@revision">
        <b>Showing revision <xsl:value-of select="@revision"/></b>
    </xsl:if>

	<div id="bodyContent">	
		<h3 id="siteSub">From Neqwiki, the nonlinear equations encyclopedia</h3>
		<xsl:apply-templates select="ow:body"/>
	</div>

<!--
    <form name="f" method="get">
    <xsl:attribute name="action"><xsl:value-of select="/ow:wiki/ow:scriptname"/></xsl:attribute>
    <hr noshade="noshade" size="1" />
    <table cellspacing="0" cellpadding="0" border="0" width="100%">

      <xsl:if test="$showBookmarksInFooter='1'">
        <xsl:if test="not(/ow:wiki/ow:userpreferences/ow:bookmarks='None')">
          <tr>
            <td align="left" class="n">
              <xsl:apply-templates select="/ow:wiki/ow:userpreferences/ow:bookmarks"/>
            </td>
          </tr>
        </xsl:if>        
      </xsl:if>

      <tr>
        <td align="left" class="n">
        
            <a class="same"><xsl:attribute name="href"><xsl:value-of select="/ow:wiki/ow:scriptname"/>?p=<xsl:value-of select="$name"/>&amp;a=edit<xsl:if test='@revision'>&amp;revision=<xsl:value-of select="@revision"/></xsl:if></xsl:attribute>Edit <xsl:if test='@revision'>revision <xsl:value-of select="@revision"/> of</xsl:if> this page</a>
            <xsl:if test="@revision or (ow:change and not(ow:change/@revision = 1))">
                |
                <a class="same"><xsl:attribute name="href"><xsl:value-of select="/ow:wiki/ow:scriptname"/>?p=<xsl:value-of select="$name"/>&amp;a=changes</xsl:attribute>View other revisions</a>
            </xsl:if>
            <xsl:if test='@revision'>

                |
                <a class="same"><xsl:attribute name="href"><xsl:value-of select="/ow:wiki/ow:scriptname"/>?p=<xsl:value-of select="$name"/></xsl:attribute>View current revision</a>
            </xsl:if>

            <xsl:if test="/ow:wiki/ow:allowattachments">
                |
                <a class="same"><xsl:attribute name="href"><xsl:value-of select="/ow:wiki/ow:scriptname"/>?p=<xsl:value-of select="$name"/>&amp;a=attach</xsl:attribute>Attachments</a> (<xsl:value-of select="count(ow:attachments/ow:attachment[@deprecated='false'])"/>)
            </xsl:if>
            
                |
                <a class="same" href="{ow:scriptname}?a=fullsearch&amp;txt={$name}&amp;fromtitle=true">
                  Referencing pages
                </a>

                |
                <a class="same"><xsl:attribute name="href"><xsl:value-of select="/ow:wiki/ow:scriptname"/>?p=<xsl:value-of select="$name"/>&amp;a=print&amp;revision=<xsl:value-of select="ow:change/@revision"/></xsl:attribute>Printable version
                </a>

        </td>
      </tr>
      <tr>
        <td align="left" class="n">
        
            <xsl:if test="$showThirdLineInFooter='1'">
                <a class="same"><xsl:attribute name="href"><xsl:value-of select="/ow:wiki/ow:scriptname"/>?p=<xsl:value-of select="$name"/>&amp;a=xml&amp;revision=<xsl:value-of select="ow:change/@revision"/></xsl:attribute>View XML</a>
              <br />
              
              <a class="same"><xsl:attribute name="href"><xsl:value-of select="/ow:wiki/ow:scriptname"/>?p=FindPage&amp;txt=<xsl:value-of select="$name"/></xsl:attribute>Find page</a> by browsing, searching or an index
              <br />
            </xsl:if>
            
            <xsl:if test="not(@changes='0')">
            
                This page was last modified on <xsl:value-of select="ow:formatLongDate(string(ow:change/ow:date))"/>

                <xsl:text> </xsl:text>
                <a class="same"><xsl:attribute name="href"><xsl:value-of select="/ow:wiki/ow:scriptname"/>?p=<xsl:value-of select="$name"/><xsl:if test="@revision">&amp;difffrom=<xsl:value-of select="@revision"/></xsl:if>&amp;a=diff</xsl:attribute>(diff)</a>
                <br />
            </xsl:if>
            
            <input type="hidden" name="a" value="fullsearch" />
            <input type="text" name="txt" size="30" ondblclick='event.cancelBubble=true;' /> <input type="submit" value="Search"/>

        </td>
      </tr>
    </table>
    </form>
-->  
</xsl:template>

<!-- ==================== wiki link to an existing page ==================== -->
<xsl:template match="ow:link">
    <xsl:choose>
        <xsl:when test="@date">
            <a href="{@href}{@anchor}" title="Last changed: {ow:formatLongDate(string(@date))}"><xsl:value-of select="text()"/></a>
        </xsl:when>
        <xsl:otherwise>
			<a>
				<xsl:attribute name="class">new</xsl:attribute>
				<xsl:attribute name="href"><xsl:value-of select="@href"/></xsl:attribute>
				<xsl:attribute name="title">Describe this page</xsl:attribute>
				<xsl:value-of select="text()"/>
			</a>
        </xsl:otherwise>
    </xsl:choose>
</xsl:template>

<!-- ==================== bookmarks from the user preferences ==================== -->
<xsl:template match="ow:bookmarks">
    <xsl:for-each select="ow:link">
        <a class="userBookmark" href="{@href}"><xsl:value-of select="text()"/></a>
        <xsl:if test="not(position()=last())"> | </xsl:if>
    </xsl:for-each>
</xsl:template>

<!-- ==================== the trail, the last visited wiki pages ==================== -->
<xsl:template match="ow:trail">
    <xsl:if test="count(ow:link) &gt; 1 and ../ow:userpreferences/ow:trailontop">
        <small>
            <xsl:for-each select="ow:link">
                <xsl:choose>
                    <xsl:when test="../../ow:page/ow:link/@href=@href">
                        »<xsl:value-of select="text()"/>
                    </xsl:when>
                    <xsl:otherwise>
                        »<a href="{@href}"><xsl:value-of select="text()"/></a>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:for-each>
        </small>
        <hr noshade="noshade" size="1" />
    </xsl:if>
</xsl:template>

<!-- ==================== actual body of a page ==================== -->
<xsl:template match="ow:body">
    <xsl:if test=".='' and not(/ow:wiki/@mode='embedded')">
        <br />
        <a><xsl:attribute name="href"><xsl:value-of select="/ow:wiki/ow:scriptname"/>?p=<xsl:value-of select="$name"/>&amp;a=edit</xsl:attribute>Describe <xsl:value-of select="../ow:link/text()"/> here</a>
        <xsl:apply-templates select="../../ow:templates"/>
    </xsl:if>
    <xsl:if test="./ow:deprecated">
        <font color="#ff0000"><b>This page will be permanently destroyed.</b></font>
        <p />
    </xsl:if>
    <xsl:apply-templates select="text() | *"/>
    <xsl:apply-templates select="../ow:attachments">
        <xsl:with-param name="showhidden">false</xsl:with-param>
        <xsl:with-param name="showactions">false</xsl:with-param>
    </xsl:apply-templates>
</xsl:template>

<!-- ==================== templates one can use to create a new page ==================== -->
<xsl:template match="ow:templates">
    <p/>
    <br />
    <br />
    Alternatively, create this page using one of these templates:
    <ul>
    <xsl:apply-templates select="ow:page"/>
    </ul>
    To create your own template add a page with a name ending in Template.
</xsl:template>

<!-- ==================== template one can use to create a new page ==================== -->
<xsl:template match="ow:templates/ow:page">
    <li>
      <a><xsl:attribute name="href"><xsl:value-of select="/ow:wiki/ow:scriptname"/>?p=<xsl:value-of select="$name"/>&amp;a=edit&amp;template=<xsl:value-of select="ow:urlencode(string(@name))"/></xsl:attribute><xsl:value-of select="ow:link/text()"/></a>
      &#160;
      (<a onclick="return !window.open(this.href)"><xsl:attribute name="href"><xsl:value-of select="/ow:wiki/ow:scriptname"/>?<xsl:value-of select="ow:urlencode(string(@name))"/></xsl:attribute>view template</a>
       <a onclick="return !window.open(this.href)"><xsl:attribute name="href"><xsl:value-of select="/ow:wiki/ow:scriptname"/>?<xsl:value-of select="ow:urlencode(string(@name))"/></xsl:attribute><img src="ow/images/popup.gif" width="15" height="9" alt="" /></a>)
    </li>
</xsl:template>

<!-- ==================== handles the openwiki-html element ==================== -->
<xsl:template match="ow:html">
  <xsl:value-of select="." disable-output-escaping="yes" />
</xsl:template>

<!-- ==================== handles the openwiki-math element ==================== -->
<xsl:template match="ow:math">
	<xsl:choose>
		<xsl:when test="./ow:display='inline'">
			<math xmlns="http://www.w3.org/1998/Math/MathML">
				<mstyle mathsize="150%">
					<xsl:value-of select="text()" disable-output-escaping="yes" />
				</mstyle>
			</math>
		</xsl:when>	
		<xsl:otherwise>
			<p style="text-align: center; padding: 1em;">
				<xsl:if test="@id">
					<xsl:text disable-output-escaping="yes">&lt;a id="</xsl:text>
					<xsl:value-of select="string(@id)"/>
					<xsl:text disable-output-escaping="yes">"&gt;&lt;/a&gt;</xsl:text>
				</xsl:if>
				
				<xsl:if test="@number">
					<span style="float: right;">
						(<xsl:value-of select="string(@number)"/>)
					</span>
				</xsl:if>
				<math xmlns="http://www.w3.org/1998/Math/MathML" display="block">
					<mstyle mathsize="150%">
						<xsl:value-of select="." disable-output-escaping="yes" />
					</mstyle>
				</math>
			</p>
		</xsl:otherwise>
	</xsl:choose>
</xsl:template>

<!-- ==================== handles the openwiki-redirectlinks element ==================== -->
<xsl:template match="ow:redirectlinks">
	<div class="mw-spcontent">
		<ol class="special" start="1">
			<xsl:apply-templates select="./ow:redirect"/>
		</ol>
	</div>
</xsl:template>

<xsl:template match="ow:redirect">
	<li>
		<a>
			<xsl:attribute name="class">mw-redirect</xsl:attribute>
			<xsl:attribute name="title"><xsl:value-of select="./ow:from/ow:link/@name" /></xsl:attribute>
			<xsl:attribute name="href"><xsl:value-of select="./ow:from/ow:link/@href" /></xsl:attribute>
			<xsl:value-of select="./ow:from/ow:link/text()" />
		</a>
		<xsl:text> →‎ </xsl:text>
		<a>
			<xsl:attribute name="class">mw-redirect</xsl:attribute>
			<xsl:attribute name="title"><xsl:value-of select="./ow:to/ow:link/@name" /></xsl:attribute>
			<xsl:attribute name="href"><xsl:value-of select="./ow:to/ow:link/@href" /></xsl:attribute>
			<xsl:value-of select="./ow:to/ow:link/text()" />
		</a>		
	</li>
</xsl:template>

<!-- ==================== handles the openwiki-category element ==================== -->
<xsl:template match="ow:categories">
	<div id="catlinks" class="catlinks">
		<div id="mw-normal-catlinks">
			<a> 
				<xsl:attribute name="title">CategoryCategory</xsl:attribute>
				<xsl:attribute name="href"><xsl:value-of select="/ow:wiki/ow:scriptname" />?CategoryCategory</xsl:attribute>
				Categories</a>			
			<xsl:text>: </xsl:text>
			<xsl:apply-templates select="./ow:category"/>
		</div>
	</div>
</xsl:template>

<xsl:template match="ow:category">
	<span dir="ltr">
		<xsl:choose>
			<xsl:when test="./ow:link/@date">
				<a>
					<xsl:attribute name="title">Last changed: <xsl:value-of select="ow:formatLongDate(string(./ow:link/@date))" /></xsl:attribute>
					<xsl:attribute name="href">
						<xsl:value-of select="./ow:link/@href" /><xsl:value-of select="./ow:link/@anchor" />
					</xsl:attribute>
					<xsl:value-of select="./name" />
				</a>
			</xsl:when>
			<xsl:otherwise>
				<a>
					<xsl:attribute name="class">new</xsl:attribute>
					<xsl:attribute name="href"><xsl:value-of select="./ow:link/@href"/></xsl:attribute>
					<xsl:attribute name="title">Describe this page</xsl:attribute>
					<xsl:value-of select="./name" />
				</a>
			</xsl:otherwise>
		</xsl:choose>				
	</span>
<!--	
    <xsl:choose>
        <xsl:when test="@date">
            <a href="{@href}{@anchor}" title="Last changed: {ow:formatLongDate(string(@date))}"><xsl:value-of select="text()"/></a>
        </xsl:when>
        <xsl:otherwise>
			<a>
				<xsl:attribute name="class">new</xsl:attribute>
				<xsl:attribute name="href"><xsl:value-of select="@href"/></xsl:attribute>
				<xsl:attribute name="title">Describe this page</xsl:attribute>
				<xsl:value-of select="text()"/>
			</a>
        </xsl:otherwise>
    </xsl:choose>	
-->	
	<xsl:if test="not(position()=last())">
		<xsl:text> | </xsl:text>
	</xsl:if>
</xsl:template>


<!-- ==================== handles the openwiki-infobox element ==================== -->
<xsl:template match="ow:infobox">
	<table class="infobox vcard"  style="width:22em; font-size:90%; text-align:left;">
		<tr style="text-align:center;">
			<th colspan="2" style="text-align:center; font-size:larger; background-color:SkyBlue; color:#000;" class="fn summary"><xsl:apply-templates select="ow:infobox_name/text()" /></th>
		</tr>
		<xsl:for-each select="ow:infobox_row">
			<tr>
				<th><xsl:apply-templates select="ow:param_name/text()|ow:param_name/*" /></th>
				<td><xsl:apply-templates select="ow:param_val/text()|ow:param_val/*" /></td>
			</tr>
		</xsl:for-each>
	</table>
</xsl:template>

<!-- ==================== inclusion of another wikipage in this wikipage ==================== -->
<xsl:template match="ow:body/ow:page">
    <xsl:apply-templates select="ow:body"/>
    <div style="float: right"><small>[goto <xsl:apply-templates select="ow:link"/>]</small></div>
    <p/>
</xsl:template>

<!-- ==================== shows an error message ==================== -->
<xsl:template match="ow:error">
    <li style="color:red;"><xsl:value-of select="."/></li>
</xsl:template>

<!-- ==================== shows footnotes ==================== -->
<xsl:template match="ow:footnotes">
    <p></p>
    ____
    <xsl:apply-templates select="ow:footnote" />
</xsl:template>

<xsl:template match="ow:footnote">
    <br /><a name="#footnote{@index}"></a><sup>&#160;&#160;&#160;<xsl:value-of select="@index"/>&#160;</sup><xsl:apply-templates />
</xsl:template>

<!-- ==================== show an RSS feed ==================== -->
<xsl:template match="ow:feed">
    <xsl:apply-templates/>
    <small>
    <br />
    last update: <xsl:value-of select="ow:formatLongDateTime(string(@last))"/>
    <br />
    <a href="{@href}" onclick="return !window.open(this.href)"><img src="ow/images/xml.gif" width="36" height="14" border="0" alt="" /></a> |
    <a href="{/ow:wiki/ow:scriptname}?p={/ow:wiki/ow:page/ow:link/@name}&amp;a=refresh&amp;refreshurl={ow:urlencode(string(@href))}">refresh</a> |
    <a href="{/ow:wiki/ow:scriptname}?p={/ow:wiki/ow:page/ow:link/@name}&amp;a=refresh">refresh all</a>
    </small>
</xsl:template>

<!-- ==================== show an aggregated RSS feed ==================== -->
<xsl:template match="ow:aggregation">
    <xsl:apply-templates/>
    <small>
    <br />
    last update: <xsl:value-of select="ow:formatLongDateTime(string(@last))"/>
    <br />
    <a href="{@href}" onclick="return !window.open(this.href)"><img src="ow/images/xml.gif" width="36" height="14" border="0" alt="" /></a> |
    <a href="{@refreshURL}">refresh</a>
    </small>
</xsl:template>

<!-- ==================== holds interwiki elements ==================== -->
<xsl:template match="ow:interlinks">
    <script type="text/javascript" charset="{/ow:wiki/@encoding}">
      <xsl:text disable-output-escaping="yes">
        function ask(pURL) {
            var x = prompt("Enter the word you're searching for:", "");
            if (x != null) {
                var pos = pURL.indexOf("$1");
                if (pos > 0) {
                    top.location.assign(pURL.substring(0, pos) + x + pURL.substring(pos + 2, pURL.length));
                } else {
                    top.location.assign(pURL + x);
                }
            }
        }
	  </xsl:text>
    </script>
    <dl>
		<xsl:for-each select="ow:interlink">
			<dt><xsl:value-of select="ow:name"/></dt>
			<dd><a href="#" onclick="javascript:ask('{ow:href}');" class="external {ow:class}"><xsl:value-of select="ow:href"/></a></dd>
		</xsl:for-each>
    </dl>
</xsl:template>

<xsl:template match="/ow:wiki" mode="edit">
  <xsl:call-template name="pi"/>
  <html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" dir="ltr">
  <xsl:call-template name="nofollow_head"/>
    <body class="mediawiki ltr ns-0 ns-subject skin-monobook" onload="window.defaultStatus='{$brandingText}'">
        <xsl:attribute name="onload">document.getElementById('editform').text.focus();</xsl:attribute>

        <script type="text/javascript" charset="{@encoding}">
          <xsl:text disable-output-escaping="yes">/*&lt;![CDATA[*/
            function openw(pURL)
            {
                var w = window.open(pURL, "openw", "width=680,height=560,resizable=1,statusbar=1,scrollbars=1");
                w.focus();
            }

            function preview()
            {
                var w = window.open("", "preview", "width=680,height=560,resizable=1,statusbar=1,scrollbars=1");
                w.focus();

                var body = '&lt;html&gt;&lt;head&gt;&lt;meta http-equiv="Content-Type" content="text/html; charset=</xsl:text><xsl:value-of select="@encoding"/><xsl:text disable-output-escaping="yes">;" />&lt;/head&gt;&lt;body&gt;&lt;form id="pvw" name="pvw" method="post" action="</xsl:text><xsl:value-of select="/ow:wiki/ow:location"/><xsl:value-of select="/ow:wiki/ow:scriptname"/><xsl:text disable-output-escaping="yes">" /&gt;';
                body += '&lt;input type="hidden" name="a" value="preview" /&gt;';
                body += '&lt;input type="hidden" name="p" value="</xsl:text><xsl:value-of select="$name"/><xsl:text disable-output-escaping="yes">" /&gt;';
                body += '&lt;input id="text" type="hidden" name="text"/&gt;&lt;/form&gt;&lt;/body&gt;&lt;/html&gt;';

                w.document.open();
                w.document.write(body);
                w.document.close();

                w.document.getElementById('pvw').elements['text'].value = window.document.getElementById('editform').elements['text'].value;
                w.document.getElementById('pvw').submit();
            }

            function saveDocumentCheck(evt) {
                    var desiredKeyState = evt.ctrlKey &amp;&amp; !evt.altKey &amp;&amp; !evt.shiftKey;
                    var key = evt.keyCode;
                    var charS = 83;
                    if ( desiredKeyState &amp;&amp; key == charS ) {
                            window.document.getElementById('editform').elements['save'][0].click();
                            evt.returnValue = false;
                    }
            }

            function theTextAreaValue() {
                return window.document.getElementById('editform').elements['text'].value;
            }

            savedValue = 'Empty';
            function checkChanged() {
                    currentValue = theTextAreaValue();
                    if (currentValue != savedValue) {
                            event.returnValue = 'Text changed without saving.';
                    }
            }
            function saveText(v) {
                    if (savedValue == 'Empty') {
                            setText(v);
                    }
                    window.onbeforeunload = checkChanged;
            }
            function setText(v) {
                    savedValue = v;
            }

          /*]]&gt;*/</xsl:text>
        </script>
		<div id="globalWrapper">
		<div id="column-content">
			<div id="content">
				<a id="top"></a>
					<h1 id="firstHeading" class="firstHeading">Editing <xsl:if test="ow:page/@revision">revision <xsl:value-of select="ow:page/@revision"/> of </xsl:if><xsl:value-of select="ow:page/@name"/></h1>
					<div id="bodyContent">
						<h3 id="siteSub">From Neqwiki, the nonlinear equations encyclopedia</h3>
						<xsl:if test="ow:page/@revision">
							<b>Editing old revision <xsl:value-of select="ow:page/@revision"/>. Saving this page will replace the latest revision with this text.</b>
						</xsl:if>
						<xsl:if test="count(ow:error) &gt; 0">
							<ul>
								<xsl:apply-templates select="ow:error"/>
							</ul>
						</xsl:if>
				
						<xsl:if test="ow:textedits">
							<p>
								The text you edited is shown below.
								The text in the textarea box shows the latest version of this page.
							</p>
							<hr />
							<pre><xsl:value-of select="ow:textedits"/></pre>
							<hr />
						</xsl:if>

						<xsl:if test="not(/ow:wiki/ow:userpreferences/ow:username/text())">
							<xsl:call-template name="edit_warning" />
						</xsl:if>
						
						<!--<xsl:call-template name="edit_buttons_toolbar"/>-->
						<div id='toolbar'>
							<xsl:comment>Edit toolbar</xsl:comment>
						</div>
				
						<form id="editform" method="post" onsubmit="setText(theTextAreaValue()); return true;">
							<xsl:attribute name="action"><xsl:value-of select="/ow:wiki/ow:scriptname"/>?a=edit#preview</xsl:attribute>
							<fieldset style="border:none;">
								<xsl:if test="not(/ow:wiki/ow:protection='captcha')">
									<input type="submit" name="save" value="Save" />
									&#160;
								</xsl:if>
								<input type="button" name="prev1" value="Preview" onclick="javascript:preview();" />
								&#160;
								<input type="button" name="cancel1" value="Cancel" onclick="javascript:window.location='{/ow:wiki/ow:scriptname}?p={$name}';" />
								<br />
								<br />

								<textarea id="text" name="text" style="overflow:auto;" onfocus="saveText(this.value)" onkeydown="saveDocumentCheck(event);"><xsl:attribute name="rows"><xsl:value-of select="/ow:wiki/ow:userpreferences/ow:rows"/></xsl:attribute><xsl:attribute name="cols"><xsl:value-of select="/ow:wiki/ow:userpreferences/ow:cols"/></xsl:attribute><xsl:value-of select="ow:page/ow:raw/text()"/></textarea><br />
								<input type="checkbox" name="rc" value="1">
								  <xsl:if test="ow:page/ow:change/@minor='false' and not(starts-with(ow:page/ow:raw/text(), '#MINOREDIT'))">
									<xsl:attribute name="checked">checked</xsl:attribute>
								  </xsl:if>
								</input>
								Include page in
								<a href="{/ow:wiki/ow:scriptname}?p=RecentChanges" onclick="javascript:openw('{/ow:wiki/ow:scriptname}?p=RecentChanges&amp;a=print'); return false;">Recent Changes</a>
								<a href="{/ow:wiki/ow:scriptname}?p=RecentChanges" onclick="javascript:openw('{/ow:wiki/ow:scriptname}?p=RecentChanges&amp;a=print'); return false;"><img src="ow/images/popup.gif" width="15" height="9" alt="" /></a>
								list.
								<br />
								<br />
								Optional comment about this change:
								<br />
								<input type="text" name="comment" style="color:#333333; width:100%" maxlength="1000"><xsl:attribute name="size"><xsl:value-of select="/ow:wiki/ow:userpreferences/ow:cols"/></xsl:attribute><xsl:attribute name="value"><xsl:value-of select="ow:page/ow:change/ow:comment/text()"/></xsl:attribute></input>
								<br />
								<input type="hidden" name="revision" value="{ow:page/@revision}" />
								<input type="hidden" name="newrev" value="{ow:page/ow:change/@revision}" />
								<input type="hidden" name="p" value="{$name}" />
								<!--<xsl:call-template name="showRecapthca" />-->
								<div id="recaptcha_holder" />
								<input type="submit" name="save" value="Save" />
								&#160;
								<input type="button" name="prev2" value="Preview" onclick="javascript:preview();" />
								<!-- <input type="submit" name="preview" value="Preview" /> -->
								&#160;
								<input type="button" name="cancel2" value="Cancel" onclick="javascript:window.location='{/ow:wiki/ow:scriptname}?p={$name}';" />
							</fieldset>
						</form>
					</div>
				</div>
			</div>
			<div id="column-one">
				<xsl:call-template name="menu_column" />					
			</div>
			<div class="visualClear"></div>
			<div id="footer">
				<xsl:call-template name="poweredBy" />
				<xsl:call-template name="footer_list" />
			</div>
		</div>
		<xsl:call-template name="ExternalJS" />
		<xsl:call-template name="EditJS" />		
    </body>
  </html>
</xsl:template>

<xsl:template match="/ow:wiki" mode="print">
  <xsl:call-template name="pi"/>
  <html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" dir="ltr">
  <xsl:call-template name="nofollow_head_print"/>
    <body class="mediawiki ltr ns-0 ns-subject skin-monobook" onload="window.defaultStatus='{$brandingText}'">
		<div id="globalWrapper">
			<div id="column-content">
				<div id="content">
					<a id="top"></a>
					<div id="bodyContent">	
						<h3 id="siteSub">From Neqwiki, the nonlinear equations encyclopedia</h3>
						<xsl:apply-templates select="ow:page/ow:body"/>
					</div>
				</div>
			</div>
			<div id="column-one">
				<xsl:call-template name="menu_column" />					
			</div>
			<div class="visualClear"></div>
			<div id="footer">
				<xsl:call-template name="poweredBy" />
				<xsl:call-template name="footer_list" />
			</div>
		</div>
		<xsl:call-template name="ExternalJS" />
    </body>
  </html>
</xsl:template>

<xsl:template match="/ow:wiki" mode="naked">
<xsl:call-template name="pi"/>
	<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" dir="ltr">
		<xsl:call-template name="nofollow_head"/>
		<body class="mediawiki ltr ns-0 ns-subject skin-monobook" onload="window.defaultStatus='{$brandingText}'">
			<div id="globalWrapper">
				<xsl:apply-templates select="ow:page/ow:body"/>
			</div>
			<xsl:call-template name="ExternalJS" />
		</body>
	</html>
</xsl:template>

<xsl:template match="/ow:wiki" mode="embedded">
    <xsl:apply-templates select="ow:page/ow:body"/>
</xsl:template>

<xsl:template match="ow:diff">
    <pre class="diff">
        <xsl:apply-templates/>
    </pre>
</xsl:template>

<xsl:template match="/ow:wiki" mode="diff">
	<xsl:call-template name="pi"/>
	<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" dir="ltr">
		<xsl:call-template name="nofollow_head"/>
		<body class="mediawiki ltr ns-0 ns-subject skin-monobook" onload="window.defaultStatus='{$brandingText}'">
			<div id="globalWrapper">
				<div id="column-content">
					<div id="content">
						<a id="top"></a>			
						<h1>
						  <a class="same" href="{ow:scriptname}?a=fullsearch&amp;txt={$name}&amp;fromtitle=true" title="Do a full text search for {ow:page/ow:link/text()}">
							<xsl:value-of select="ow:page/ow:link/text()"/>
						  </a>
						</h1>
						<xsl:choose>
							<xsl:when test="ow:diff = ''">
								<b>No difference available. This is the first <xsl:value-of select="ow:diff/@type"/> revision.</b>
								<hr noshade="noshade" size="1"/>
								<xsl:apply-templates select="ow:trail"/>
								<xsl:if test='ow:page/@revision'>
									<b>Showing revision <xsl:value-of select="ow:page/@revision"/></b>
									<p></p>
								</xsl:if>
								<xsl:apply-templates select="ow:page/ow:body"/>
							</xsl:when>
							<xsl:otherwise>
								<xsl:if test="not(ow:diff/@type='selected')">
									<b>Difference from prior <xsl:value-of select="ow:diff/@type"/>
									revision<xsl:if test="not(ow:diff/@to = ow:page/@lastminor)"> relative to revision
									<xsl:value-of select="ow:diff/@to"/>
									</xsl:if>.</b>
								</xsl:if>
								<xsl:if test="ow:diff/@type='selected'">
									<b>Difference from revision <xsl:value-of select="ow:diff/@from"/> to
									<xsl:choose>
										<xsl:when test="ow:diff/@to = ow:page/@lastminor">
											the current revision.
										</xsl:when>
										<xsl:otherwise>
											revision <xsl:value-of select="ow:diff/@to"/>.
										</xsl:otherwise>
									</xsl:choose>
									</b>
								</xsl:if>
								<br />
								<xsl:if test="not(ow:diff/@type='major')">
									<a><xsl:attribute name="href"><xsl:value-of select="/ow:wiki/ow:scriptname"/>?p=<xsl:value-of select="$name"/>&amp;a=diff</xsl:attribute>major diff</a>
									<xsl:text> </xsl:text>
								</xsl:if>
								<xsl:if test="not(ow:diff/@type='minor')">
									<a><xsl:attribute name="href"><xsl:value-of select="/ow:wiki/ow:scriptname"/>?p=<xsl:value-of select="$name"/>&amp;a=diff&amp;diff=1</xsl:attribute>minor diff</a>
									<xsl:text> </xsl:text>
								</xsl:if>
								<xsl:if test="not(ow:diff/@type='author')">
									<a><xsl:attribute name="href"><xsl:value-of select="/ow:wiki/ow:scriptname"/>?p=<xsl:value-of select="$name"/>&amp;a=diff&amp;diff=2</xsl:attribute>author diff</a>
									<xsl:text> </xsl:text>
								</xsl:if>
								<a><xsl:attribute name="href"><xsl:value-of select="/ow:wiki/ow:scriptname"/>?p=<xsl:value-of select="$name"/><xsl:if test="ow:diff/@to">&amp;revision=<xsl:value-of select="ow:diff/@to"/></xsl:if></xsl:attribute>hide diff</a>
								<p></p>
								<xsl:apply-templates select="ow:diff"/>
							</xsl:otherwise>
						</xsl:choose>
				
						<form name="f" method="get">
						<xsl:attribute name="action"><xsl:value-of select="/ow:wiki/ow:scriptname"/></xsl:attribute>
						<hr />
				
						<xsl:if test="$showBookmarksInFooter='1'">
						  <xsl:apply-templates select="ow:userpreferences/ow:bookmarks"/>
						</xsl:if>				
						<xsl:if test="not(ow:page/@changes='0')">
							This page was last modified on <xsl:value-of select="ow:formatLongDate(string(ow:page/ow:change/ow:date))"/>
							<xsl:text> </xsl:text>
							<a><xsl:attribute name="href"><xsl:value-of select="/ow:wiki/ow:scriptname"/>?p=<xsl:value-of select="$name"/><xsl:if test="ow:diff/@to">&amp;revision=<xsl:value-of select="ow:diff/@to"/></xsl:if></xsl:attribute>(hide diff)</a>
							<br />
						</xsl:if>
						<input type="hidden" name="a" value="fullsearch"/>
						<input type="text" name="txt" size="30"/> <input type="submit" value="Search"/>
						</form>
					</div>
				</div>
				<div id="column-one">
					<xsl:call-template name="menu_column" />					
				</div>
				<div class="visualClear"></div>
				<div id="footer">
					<xsl:call-template name="poweredBy" />
					<xsl:call-template name="footer_list" />
				</div>
			 </div>
			 <xsl:call-template name="ExternalJS" />
		</body>
	</html>
</xsl:template>

<xsl:template match="ow:recentchanges" mode="shortversion">
    <table cellspacing="0" cellpadding="2" border="0">
    <xsl:for-each select="ow:page">
        <tr>
        <xsl:choose>
            <xsl:when test='not(substring-before(./preceding-sibling::*[position()=1]/ow:change/ow:date, "T") = substring-before(ow:change/ow:date, "T"))'>
                <td class="rc" style="width:1%; white-space: nowrap;"><xsl:value-of select="ow:formatShortDate(string(ow:change/ow:date))"/></td>
            </xsl:when>
            <xsl:otherwise>
                <td class="rc" style="width:1%;">&#160;</td>
            </xsl:otherwise>
        </xsl:choose>
        <td class="rc">
        <xsl:value-of select="ow:formatTime(string(ow:change/ow:date))"/>
        -
        <xsl:apply-templates select="ow:link"/>&#160;<xsl:if test="ow:change/@status='new'"><span class="rcnew">new</span></xsl:if><xsl:if test="ow:change/@status='deleted'"><span class="deprecated">deprecated</span></xsl:if>
        </td>
        </tr>
    </xsl:for-each>
    </table>
</xsl:template>

<xsl:template match="ow:recentchanges">
    <xsl:choose>
        <xsl:when test="@short='true'">
            <xsl:apply-templates select="." mode="shortversion"/>
        </xsl:when>
        <xsl:otherwise>
            <table cellspacing="0" cellpadding="2" width="100%" border="0">
            <xsl:for-each select="ow:page">
                <xsl:if test='not(substring-before(./preceding-sibling::*[position()=1]/ow:change/ow:date, "T") = substring-before(ow:change/ow:date, "T"))'>
                    <tr class="rc">
                        <td colspan="4">&#160;</td>
                    </tr>
                    <tr class="rc">
                        <td colspan="4"><b><xsl:value-of select="ow:formatLongDate(string(ow:change/ow:date))"/></b></td>
                    </tr>
                </xsl:if>
                <tr class="rc">
                    <td align="left" style="width:1%;"><xsl:value-of select="ow:formatTime(string(ow:change/ow:date))"/></td>
                    <td align="left" style="width:25%; white-space: nowrap;"><xsl:if test="@changes > 1">[<a><xsl:attribute name="href"><xsl:value-of select="/ow:wiki/ow:scriptname"/>?p=<xsl:value-of select="ow:urlencode(string(@name))"/>&amp;a=diff</xsl:attribute>diff</a>] <xsl:text> </xsl:text> [<xsl:value-of select="@changes"/>&#160;<a><xsl:attribute name="href"><xsl:value-of select="/ow:wiki/ow:scriptname"/>?p=<xsl:value-of select="ow:urlencode(string(@name))"/>&amp;a=changes</xsl:attribute>changes</a>]</xsl:if>&#160;</td>
                    <td align="left"><a><xsl:attribute name="href"><xsl:value-of select="/ow:wiki/ow:scriptname"/>?<xsl:value-of select="ow:urlencode(string(@name))"/></xsl:attribute><xsl:value-of select="ow:link/text()"/></a>&#160;<xsl:if test="ow:change/@status='new'"><span class="rcnew">new</span></xsl:if><xsl:if test="ow:change/@status='deleted'"><span class="deprecated">deprecated</span></xsl:if></td>

                    <xsl:choose>
                      <xsl:when test="ow:change/ow:by/@alias">
                        <td align="left"><a><xsl:attribute name="href"><xsl:value-of select="/ow:wiki/ow:scriptname"/>?<xsl:value-of select="ow:urlencode(string(ow:change/ow:by/@alias))"/></xsl:attribute><xsl:value-of select="ow:change/ow:by/text()"/></a></td>
                      </xsl:when>
                      <xsl:otherwise>
                        <td align="left"><xsl:value-of select="ow:change/ow:by/@name"/></td>
                      </xsl:otherwise>
                    </xsl:choose>

                </tr>
                <xsl:if test="ow:change/ow:comment">
                    <tr class="rc">
                        <td align="left" colspan="2">&#160;</td>
                        <td align="left" colspan="2" class="comment"><xsl:value-of select="ow:change/ow:comment"/></td>
                    </tr>
                </xsl:if>

                <xsl:for-each select="ow:change/ow:attachmentchange">
                    <tr class="rc">
                        <td colspan="4">
                            <xsl:apply-templates select="."/>
                        </td>
                    </tr>
                </xsl:for-each>
            </xsl:for-each>
            </table>
        </xsl:otherwise>
    </xsl:choose>
</xsl:template>

<xsl:template match="ow:recentchanges_original">
    <ul>
    <xsl:for-each select="ow:page">
        <xsl:if test='not(substring-before(./preceding-sibling::*[position()=1]/ow:change/ow:date, "T") = substring-before(ow:change/ow:date, "T"))'>
          <xsl:text disable-output-escaping="yes">&lt;/ul&gt;</xsl:text>
            <b><xsl:value-of select="ow:formatLongDate(string(ow:change/ow:date))"/></b>
          <xsl:text disable-output-escaping="yes">&lt;ul&gt;</xsl:text>
        </xsl:if>
        <li>
            <xsl:value-of select="ow:formatTime(string(ow:change/ow:date))"/>
            -
            <a><xsl:attribute name="href"><xsl:value-of select="/ow:wiki/ow:scriptname"/>?<xsl:value-of select="@name"/></xsl:attribute><xsl:value-of select="ow:link/text()"/></a>
            <xsl:if test="ow:change/@status='new'">
              <xsl:text> </xsl:text>
              <span class="rcnew">new</span>
            </xsl:if>
            <xsl:text> </xsl:text>
            <xsl:if test="@changes > 1">
                (<a><xsl:attribute name="href"><xsl:value-of select="/ow:wiki/ow:scriptname"/>?p=<xsl:value-of select="@name"/>&amp;a=diff</xsl:attribute>diff</a>)
                (<xsl:value-of select="@changes"/>&#160;<a><xsl:attribute name="href"><xsl:value-of select="/ow:wiki/ow:scriptname"/>?p=<xsl:value-of select="@name"/>&amp;a=changes</xsl:attribute>changes</a>)
            </xsl:if>
            <xsl:if test="ow:change/ow:comment">
                <xsl:text> </xsl:text>
                <b>[<xsl:value-of select="ow:change/ow:comment"/>]</b>
            </xsl:if>
            . . . . . .
            <xsl:choose>
              <xsl:when test="ow:change/ow:by/@alias">
                <a><xsl:attribute name="href"><xsl:value-of select="/ow:wiki/ow:scriptname"/>?<xsl:value-of select="ow:change/ow:by/@alias"/></xsl:attribute><xsl:value-of select="ow:change/ow:by/text()"/></a>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="ow:change/ow:by/@name"/>
              </xsl:otherwise>
            </xsl:choose>
        </li>
    </xsl:for-each>
    </ul>
</xsl:template>

<xsl:template match="ow:wiki" mode="changes">
  <xsl:call-template name="pi"/>
  <html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" dir="ltr">
  <xsl:call-template name="nofollow_head"/>
    <body class="mediawiki ltr ns-0 ns-subject skin-monobook" onload="window.defaultStatus='{$brandingText}'">
		<div id="globalWrapper">
			<div id="column-content">
				<div id="content">
					<a id="top"></a>
					<h1 id="firstHeading" class="firstHeading">Revision history of <xsl:value-of select="ow:page/ow:link/text()"/></h1>
					<div id="bodyContent">
						<h3 id="siteSub">From Neqwiki, the nonlinear equations encyclopedia</h3>
						<ul>
						<xsl:for-each select="ow:page/ow:change">
							<li>
								Revision:
								<xsl:value-of select="@revision"/>
								. .
								<xsl:value-of select="ow:formatLongDate(string(ow:date))"/>
								<xsl:text> </xsl:text>
								<xsl:value-of select="ow:formatTime(string(ow:date))"/>
								<xsl:text> </xsl:text>
								<a><xsl:attribute name="href"><xsl:value-of select="/ow:wiki/ow:scriptname"/>?p=<xsl:value-of select="$name"/>&amp;revision=<xsl:value-of select="@revision"/></xsl:attribute>View</a>
								<xsl:if test="position() > 1">
									(<a><xsl:attribute name="href"><xsl:value-of select="/ow:wiki/ow:scriptname"/>?p=<xsl:value-of select="$name"/>&amp;a=diff&amp;difffrom=<xsl:value-of select="@revision"/></xsl:attribute>diff</a>)
								</xsl:if>
								. . . . . .
								<xsl:choose>
								  <xsl:when test="ow:by/@alias">
									<a><xsl:attribute name="href"><xsl:value-of select="/ow:wiki/ow:scriptname"/>?<xsl:value-of select="ow:urlencode(string(ow:by/@alias))"/></xsl:attribute><xsl:value-of select="ow:by/text()"/></a>
								  </xsl:when>
								  <xsl:otherwise>
									<xsl:value-of select="ow:by/@name"/>
								  </xsl:otherwise>
								</xsl:choose>
								<xsl:if test="ow:comment">
									<br />
									<xsl:text> </xsl:text>
									<span class="comment"><xsl:value-of select="ow:comment"/></span>
								</xsl:if>
							</li>
						</xsl:for-each>
						</ul>
					</div>
				</div>
			</div>
			<div id="column-one">
				<xsl:call-template name="menu_column" />					
			</div>
			<div class="visualClear"></div>
			<div id="footer">
				<xsl:call-template name="poweredBy" />
				<xsl:call-template name="footer_list" />
			</div>					
		</div>
		<xsl:call-template name="ExternalJS" />
      </body>
  </html>
</xsl:template>

<xsl:template match="ow:titleindex">
	<div style="text-align:center">
		<xsl:for-each select="ow:page">
			<xsl:if test="not(substring(./preceding-sibling::*[position()=1]/@name, 1, 1) = substring(@name, 1, 1))">
				<a><xsl:attribute name="href">#<xsl:value-of select="substring(@name, 1, 1)"/></xsl:attribute><xsl:value-of select="substring(@name, 1, 1)"/></a>
				<xsl:text> </xsl:text>
			</xsl:if>
		</xsl:for-each>
	</div>
    <xsl:for-each select="ow:page">
        <xsl:if test="not(substring(./preceding-sibling::*[position()=1]/@name, 1, 1) = substring(@name, 1, 1))">
            <br />
            <a><xsl:attribute name="id"><xsl:value-of select="substring(@name, 1, 1)"/></xsl:attribute></a>
            <b><xsl:value-of select="substring(@name, 1, 1)"/></b>
            <br />
        </xsl:if>
        <xsl:apply-templates select="ow:link"/>
        <br />
    </xsl:for-each>
</xsl:template>

<xsl:template match="ow:wordindex">
    <center>
    <xsl:for-each select="ow:word">
        <xsl:if test="not(substring(./preceding-sibling::*[position()=1]/@value, 1, 1) = substring(@value, 1, 1))">
            <a><xsl:attribute name="href">#<xsl:value-of select="substring(@value, 1, 1)"/></xsl:attribute><xsl:value-of select="substring(@value, 1, 1)"/></a>
        </xsl:if>
        <xsl:text> </xsl:text>
    </xsl:for-each>
    </center>
    <xsl:text disable-output-escaping="yes">&lt;ul&gt;</xsl:text>
    <xsl:for-each select="ow:word">
        <xsl:if test="not(substring(./preceding-sibling::*[position()=1]/@value, 1, 1) = substring(@value, 1, 1))">
            <xsl:text disable-output-escaping="yes">&lt;/ul&gt;</xsl:text>
            <a><xsl:attribute name="name"><xsl:value-of select="substring(@value, 1, 1)"/></xsl:attribute></a>
            <b><xsl:value-of select="substring(@value, 1, 1)"/></b>
            <xsl:text disable-output-escaping="yes">&lt;ul&gt;</xsl:text>
        </xsl:if>
        <xsl:if test="not(./preceding-sibling::*[position()=1]/@value = @value)">
            <xsl:text disable-output-escaping="yes">&lt;/ul&gt;</xsl:text>
            <b><xsl:value-of select="@value"/></b>
            <xsl:text disable-output-escaping="yes">&lt;ul&gt;</xsl:text>
        </xsl:if>
        <li><xsl:apply-templates select="ow:page/ow:link"/></li>
    </xsl:for-each>
    <xsl:text disable-output-escaping="yes">&lt;/ul&gt;</xsl:text>
</xsl:template>

<xsl:template match="ow:randompages">
    <xsl:choose>
        <xsl:when test='count(ow:page)=1'>
            <xsl:apply-templates select="ow:page/ow:link"/>
        </xsl:when>
        <xsl:otherwise>
            <ul>
            <xsl:for-each select="ow:page">
                <li><xsl:apply-templates select="ow:link"/></li>
            </xsl:for-each>
            </ul>
        </xsl:otherwise>
    </xsl:choose>
</xsl:template>

<xsl:template match="ow:titlesearch">
	<xsl:if test="count(ow:page)>0">
		<ul>
			<xsl:for-each select="ow:page">
				<li>
					<xsl:if test="contains(@name, '/')">
						....
					</xsl:if>
					<xsl:apply-templates select="ow:link"/>
				</li>
			</xsl:for-each>
		</ul>
	</xsl:if>
</xsl:template>

<xsl:template match="/ow:wiki" mode="titlesearch">
  <xsl:call-template name="pi"/>
  <html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" dir="ltr">
  <xsl:call-template name="nofollow_head"/>
    <body bgcolor="#ffffff" onload="window.defaultStatus='{$brandingText}'">
        <xsl:call-template name="brandingImage"/>
        <h1>Title search for "<xsl:value-of select="ow:titlesearch/@value"/>"</h1>
        <xsl:apply-templates select="ow:userpreferences/ow:bookmarks"/>
        <hr />
        <xsl:apply-templates select="ow:titlesearch"/>
        <xsl:value-of select="count(ow:titlesearch/ow:page)"/> hits out of
        <xsl:value-of select="ow:titlesearch/@pagecount"/> pages searched.

        <form name="f" method="get">
        <xsl:attribute name="action"><xsl:value-of select="/ow:wiki/ow:scriptname"/></xsl:attribute>
        <hr />
        <xsl:apply-templates select="ow:userpreferences/ow:bookmarks"/>
        <br />
        <input type="hidden" name="a" value="fullsearch"/>
        <input type="text" name="txt" size="30"><xsl:attribute name="value"><xsl:value-of select="ow:titlesearch/@value"/></xsl:attribute></input> <input type="submit" value="Search"/>
        </form>
        <xsl:call-template name="ExternalJS" />
    </body>
  </html>
</xsl:template>

<xsl:template match="ow:fullsearch">
	<xsl:if test="count(ow:page)>0">
		<ul>
			<xsl:for-each select="ow:page">
			  <li>
				<xsl:if test="contains(@name, '/')">
					....
				</xsl:if>
				<xsl:apply-templates select="ow:link"/>
			  </li>
			</xsl:for-each>
		</ul>
    </xsl:if>
</xsl:template>

<xsl:template match="ow:equationsearch">
	<xsl:if test="count(ow:page)>0">
		<ul>
			<xsl:for-each select="ow:page">
			  <li>
				<xsl:if test="contains(@name, '/')">
					....
				</xsl:if>
				<xsl:apply-templates select="ow:link"/>
				<xsl:apply-templates select="ow:equation/*"/>
			  </li>
			</xsl:for-each>
		</ul>
    </xsl:if>
</xsl:template>

<xsl:template match="/ow:wiki" mode="fullsearch">
	<xsl:call-template name="pi"/>
	<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" dir="ltr">
		<xsl:call-template name="nofollow_head"/>
		<body class="mediawiki ltr ns--1 ns-special skin-monobook" onload="window.defaultStatus='{$brandingText}'">
			<div id="globalWrapper">
				<div id="column-content">
					<div id="content">
						<a id="top"></a>
						<h1 id="firstHeading" class="firstHeading">Search results for "<xsl:value-of select="ow:fullsearch/@value"/>"</h1>
						<div id="bodyContent">
							<xsl:apply-templates select="ow:fullsearch"/>
							<xsl:value-of select="count(ow:fullsearch/ow:page)"/> hits out of
							<xsl:value-of select="ow:fullsearch/@pagecount"/> pages searched.
							<p>
								<form method="get">
								<xsl:attribute name="action"><xsl:value-of select="/ow:wiki/ow:scriptname"/></xsl:attribute>
								<br />
								<input type="hidden" name="a" value="fullsearch"/>
								<input type="text" name="txt" size="30"><xsl:attribute name="value"><xsl:value-of select="ow:fullsearch/@value"/></xsl:attribute></input> <input type="submit" value="Search"/>
								</form>
							</p>
						</div>
					</div>
				</div>
				<div id="column-one">
					<xsl:call-template name="menu_column" />					
				</div>
				<div class="visualClear"></div>
				<div id="footer">
					<xsl:call-template name="poweredBy" />
					<xsl:call-template name="footer_list" />
				</div>
			</div>
			<xsl:call-template name="ExternalJS" />
		</body>
	</html>
</xsl:template>

<xsl:template match="ow:message">
    <xsl:if test="@code='userpreferences_saved'">
      <b>User preferences saved successfully.</b>
    </xsl:if>
    <xsl:if test="@code='userpreferences_cleared'">
      <b>User preferences cleared successfully.</b>
    </xsl:if>
</xsl:template>

<xsl:template match="ow:userpreferences">
	<form id="f" method="post">
	<xsl:attribute name="action"><xsl:value-of select="/ow:wiki/ow:scriptname"/></xsl:attribute>
		<fieldset>
			<legend>User preferences:</legend>
			Username: <input type="text" name="username" ondblclick="event.cancelBubble=true;"><xsl:attribute name="value"><xsl:value-of select="/ow:wiki/ow:userpreferences/ow:username"/></xsl:attribute></input><br />
			Bookmarks: <input type="text" name="bookmarks" size="60" ondblclick="event.cancelBubble=true;"><xsl:attribute name="value"><xsl:for-each select="/ow:wiki/ow:userpreferences/ow:bookmarks/ow:link"><xsl:value-of select="@name"/><xsl:text> </xsl:text></xsl:for-each></xsl:attribute></input><br />
			Edit form columns: <input type="text" name="cols" size="3" ondblclick="event.cancelBubble=true;"><xsl:attribute name="value"><xsl:value-of select="/ow:wiki/ow:userpreferences/ow:cols"/></xsl:attribute></input> rows: <input type="text" name="rows" size="3" ondblclick="event.cancelBubble=true;"><xsl:attribute name="value"><xsl:value-of select="/ow:wiki/ow:userpreferences/ow:rows"/></xsl:attribute></input><br />
			<input type="checkbox" name="prettywikilinks" value="1">
			<xsl:if test="/ow:wiki/ow:userpreferences/ow:prettywikilinks"><xsl:attribute name="checked">checked</xsl:attribute></xsl:if>
			</input>
			Show pretty wiki links
			<br />
			<input type="checkbox" name="bookmarksontop" value="1">
			<xsl:if test="/ow:wiki/ow:userpreferences/ow:bookmarksontop"><xsl:attribute name="checked">checked</xsl:attribute></xsl:if>
			</input>
			Show bookmarks on top
			<br />
			<input type="checkbox" name="editlinkontop" value="1">
			<xsl:if test="/ow:wiki/ow:userpreferences/ow:editlinkontop"><xsl:attribute name="checked">checked</xsl:attribute></xsl:if>
			</input>
			Show edit link on top
			<br />
			<input type="checkbox" name="trailontop" value="1">
			<xsl:if test="/ow:wiki/ow:userpreferences/ow:trailontop"><xsl:attribute name="checked">checked</xsl:attribute></xsl:if>
			</input>
			Show trail on top
			<br />
			<input type="checkbox" name="opennew" value="1">
			<xsl:if test="/ow:wiki/ow:userpreferences/ow:opennew"><xsl:attribute name="checked">checked</xsl:attribute></xsl:if>
			</input>
			Open external links in new window
			<br />
			<input type="checkbox" name="emoticons" value="1">
			<xsl:if test="/ow:wiki/ow:userpreferences/ow:emoticons"><xsl:attribute name="checked">checked</xsl:attribute></xsl:if>
			</input>
			Show emoticons in text <small>(goto <a href="?HelpOnEmoticons">HelpOnEmoticons</a>)</small>
			<br />
			<input type="submit" name="save" value="Save Preferences"/>
			&#160;&#160;
			<input type="submit" name="clear" value="Clear Preferences"/>
			<br />
			<input type="hidden" name="p"><xsl:attribute name="value"><xsl:value-of select="/ow:wiki/ow:page/@name"/></xsl:attribute></input>
			<input type="hidden" name="a" value="userpreferences"/>
		</fieldset>
	</form>
</xsl:template>

<xsl:template match="/ow:wiki" mode="login">
	<xsl:call-template name="pi"/>
	<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" dir="ltr">
		<xsl:call-template name="nofollow_head"/>
		<body class="mediawiki ltr ns--1 ns-special skin-monobook" onload="document.getElementById('pwd').focus();">
			<div id="globalWrapper">
				<div id="column-content">
					<div id="content">
						<a id="top"></a>
						<h1 id="firstHeading" class="firstHeading">Log in</h1>
						<div id="bodyContent">
							<xsl:if test="count(ow:error) &gt; 0">
								<ul>
									<xsl:apply-templates select="ow:error"/>
								</ul>
							</xsl:if>
							<form id="f" method="post" action="{/ow:wiki/ow:scriptname}?a=login&amp;mode={ow:login/@mode}">
								<fieldset>
									<xsl:if test="ow:login/@mode='edit'">
										<legend>Enter password to edit content:</legend>
									</xsl:if>	
									Password: <input type="password" id="pwd" name="pwd" size="10"/>
									<xsl:text> </xsl:text>
									<input type="submit" name="submit" value="let me in!"/>
									<br />
									<b>Note</b>: cookies and JavaScript must be enabled!<br />
									<input type="checkbox" name="r" value="1">
										<xsl:if test="ow:login/ow:rememberme='false'">
											<xsl:attribute name="checked">checked</xsl:attribute>
										</xsl:if>
									</input>
									Remember me
									<input type="hidden" name="backlink">
										<xsl:attribute name="value"><xsl:value-of select="ow:login/ow:backlink"/></xsl:attribute>
									</input>
								</fieldset>
							</form>
						</div>
					</div>
				</div>
				<div id="column-one">
					<xsl:call-template name="menu_column" />					
				</div>
				<div class="visualClear"></div>
				<div id="footer">
					<xsl:call-template name="poweredBy" />
					<xsl:call-template name="footer_list" />
				</div>
			</div>
			<xsl:call-template name="ExternalJS" />
		</body>
	</html>
</xsl:template>

</xsl:stylesheet>