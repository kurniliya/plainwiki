// MediaWiki JavaScript support functions

var clientPC = navigator.userAgent.toLowerCase(); // Get client info
var is_gecko = /gecko/.test( clientPC ) &&
	!/khtml|spoofer|netscape\/7\.0/.test(clientPC);
var webkit_match = clientPC.match(/applewebkit\/(\d+)/);
if (webkit_match) {
	var is_safari = clientPC.indexOf('applewebkit') != -1 &&
		clientPC.indexOf('spoofer') == -1;
	var is_safari_win = is_safari && clientPC.indexOf('windows') != -1;
	var webkit_version = parseInt(webkit_match[1]);
}
var is_khtml = navigator.vendor == 'KDE' ||
	( document.childNodes && !document.all && !navigator.taintEnabled );
// For accesskeys; note that FF3+ is included here!
var is_ff2 = /firefox\/[2-9]|minefield\/3/.test( clientPC );
var is_ff2_ = /firefox\/2/.test( clientPC );
// These aren't used here, but some custom scripts rely on them
var is_ff2_win = is_ff2 && clientPC.indexOf('windows') != -1;
var is_ff2_x11 = is_ff2 && clientPC.indexOf('x11') != -1;
if (clientPC.indexOf('opera') != -1) {
	var is_opera = true;
	var is_opera_preseven = window.opera && !document.childNodes;
	var is_opera_seven = window.opera && document.childNodes;
	var is_opera_95 = /opera\/(9.[5-9]|[1-9][0-9])/.test( clientPC );
}

// set up the words in your language
var NavigationBarHide = '[hide]';
var NavigationBarShow = '[show]';

// Global external objects used by this script.
/*extern ta, stylepath, skin */

// add any onload functions in this hook (please don't hard-code any events in the xhtml source)
var doneOnloadHook;

if (!window.onloadFuncts) {
	var onloadFuncts = [];
}

function addOnloadHook(hookFunct) {
	// Allows add-on scripts to add onload functions
	if(!doneOnloadHook) {
		onloadFuncts[onloadFuncts.length] = hookFunct;
	} else {
		hookFunct();  // bug in MSIE script loading
	}
}

function hookEvent(hookName, hookFunct) {
	addHandler(window, hookName, hookFunct);
}

var loadedScripts = {}; // included-scripts tracker
function importScriptURI(url) {
	if (loadedScripts[url]) {
		return null;
	}
	loadedScripts[url] = true;
	var s = document.createElement('script');
	s.setAttribute('src',url);
	s.setAttribute('type','text/javascript');
	document.getElementsByTagName('head')[0].appendChild(s);
	return s;
}
 
function importStylesheet(page) {
	return importStylesheetURI(wgScript + '?action=raw&ctype=text/css&title=' + encodeURIComponent(page.replace(/ /g,'_')));
}
 
function importStylesheetURI(url) {
	return document.createStyleSheet ? document.createStyleSheet(url) : appendCSS('@import "' + url + '";');
}
 
function appendCSS(text) {
	var s = document.createElement('style');
	s.type = 'text/css';
	s.rel = 'stylesheet';
	if (s.styleSheet) s.styleSheet.cssText = text //IE
	else s.appendChild(document.createTextNode(text + '')) //Safari sometimes borks on null
	document.getElementsByTagName('head')[0].appendChild(s);
	return s;
}

function showTocToggle() {
	if (document.createTextNode) {
		// Uses DOM calls to avoid document.write + XHTML issues

		var linkHolder = document.getElementById("toctitle");
		if (!linkHolder) {
			return;
		}

		if (!document.getElementById("togglelink")){

			var toggleLink;
		        if (typeof document.createElementNS != 'undefined') {
		        	toggleLink = document.createElementNS("http://www.w3.org/1999/xhtml","a");
		        } else {
		        	toggleLink = document.createElement("a");
		        }
			
			toggleLink.id = "togglelink";
			toggleLink.className = "internal";
			toggleLink.href = "javascript:toggleToc()";
			toggleLink.appendChild(document.createTextNode(NavigationBarHide));
			
			var outerSpan;
		        if (typeof document.createElementNS != 'undefined') {
		        	outerSpan = document.createElementNS("http://www.w3.org/1999/xhtml","span");
		        } else {
		        	outerSpan = document.createElement("span");
		        }			
			
			outerSpan.className = "toctoggle";
	
			//outerSpan.appendChild(document.createTextNode("["));
			outerSpan.appendChild(toggleLink);
			//outerSpan.appendChild(document.createTextNode("]"));
	
			linkHolder.appendChild(document.createTextNode(" "));
			linkHolder.appendChild(outerSpan);
		}

		var cookiePos = document.cookie.indexOf("hidetoc=");
		if (cookiePos > -1 && document.cookie.charAt(cookiePos + 8) == 1) {
			toggleToc();
		}
		
	}
}

function changeText(el, newText) {
	// Safari work around
	if (el.innerText) {
		el.innerText = newText;
	} else if (el.firstChild && el.firstChild.nodeValue) {
		el.firstChild.nodeValue = newText;
	}
}

function toggleToc() {
	var toc = document.getElementById("toc").getElementsByTagName("ul")[0];
	var toggleLink = document.getElementById("togglelink");

	if (toc && toggleLink && toc.style.display == "none") {
		changeText(toggleLink, NavigationBarHide);
		toc.style.display = "block";
		document.cookie = "hidetoc=0";
	} else {
		changeText(toggleLink, NavigationBarShow);
		toc.style.display = "none";
		document.cookie = "hidetoc=1";
	}
}

var mwEditButtons = [];
var mwCustomEditButtons = []; // eg to add in MediaWiki:Common.js

function runOnloadHook() {
	// don't run anything below this for non-dom browsers
	if (doneOnloadHook || !(document.getElementById && document.getElementsByTagName)) {
		return;
	}

	// set this before running any hooks, since any errors below
	// might cause the function to terminate prematurely
	doneOnloadHook = true;

	// Run any added-on functions
	for (var i = 0; i < onloadFuncts.length; i++) {
		onloadFuncts[i]();
	}
}

/**
 * Add an event handler to an element
 *
 * @param Element element Element to add handler to
 * @param String attach Event to attach to
 * @param callable handler Event handler callback
 */
function addHandler( element, attach, handler ) {
	if( window.addEventListener ) {
		element.addEventListener( attach, handler, false );
	} else if( window.attachEvent ) {
		element.attachEvent( 'on' + attach, handler );
	}
}

//note: all skins should call runOnloadHook() at the end of html output,
//      so the below should be redundant. It's there just in case.
hookEvent("load", runOnloadHook);



//<source lang="javascript">
/* Former common.js scripts */
 
/* Scripts specific to Internet Explorer */
 
if (navigator.appName == "Microsoft Internet Explorer")
{
    /** Internet Explorer bug fix **************************************************
     *
     *  Description: Fixes IE horizontal scrollbar bug
     *  Maintainers: [[User:Tom-]]?
     */
 
    var oldWidth;
    var docEl = document.documentElement;
 
    function fixIEScroll()
    {
        if (!oldWidth || docEl.clientWidth > oldWidth)
            doFixIEScroll();
        else
            setTimeout(doFixIEScroll, 1);
 
        oldWidth = docEl.clientWidth;
    }
 
    function doFixIEScroll() {
        docEl.style.overflowX = (docEl.scrollWidth - docEl.clientWidth < 4) ? "hidden" : "";
    }
 
    document.attachEvent("onreadystatechange", fixIEScroll);
    document.attachEvent("onresize", fixIEScroll);
 
 
    /**
     * Remove need for CSS hacks regarding MSIE and IPA.
     */
    if (document.createStyleSheet) {
        document.createStyleSheet().addRule('.IPA', 'font-family: "Doulos SIL", "Charis SIL", Gentium, "DejaVu Sans", Code2000, "TITUS Cyberbit Basic", "Arial Unicode MS", "Lucida Sans Unicode", "Chrysanthi Unicode";');
    }
 
    // In print IE (7?) does not like line-height
    appendCSS( '@media print { sup, sub, p, .documentDescription { line-height: normal; }}');
 
    // IE overflow bug
    appendCSS('div.overflowbugx { overflow-x: scroll !important; overflow-y: hidden !important; } div.overflowbugy { overflow-y: scroll !important; overflow-x: hidden !important; }');
 
    //Import scripts specific to Internet Explorer 6
    if (navigator.appVersion.substr(22, 1) == "6") {
        importScript("MediaWiki:Common.js/IE60Fixes.js")
    }
}
 
 
/* Test if an element has a certain class **************************************
 *
 * Description: Uses regular expressions and caching for better performance.
 * Maintainers: [[User:Mike Dillon]], [[User:R. Koot]], [[User:SG]]
 */
 
var hasClass = (function () {
    var reCache = {};
    return function (element, className) {
        return (reCache[className] ? reCache[className] : (reCache[className] = new RegExp("(?:\\s|^)" + className + "(?:\\s|$)"))).test(element.className);
    };
})();
 
 
/** Dynamic Navigation Bars (experimental) *************************************
 *
 *  Description: See [[Wikipedia:NavFrame]].
 *  Maintainers: UNMAINTAINED
 */
 
// shows and hides content and picture (if available) of navigation bars
// Parameters:
//     indexNavigationBar: the index of navigation bar to be toggled
function toggleNavigationBar(indexNavigationBar)
{
    var NavToggle = document.getElementById("NavToggle" + indexNavigationBar);
    var NavFrame = document.getElementById("NavFrame" + indexNavigationBar);
 
    if (!NavFrame || !NavToggle) {
        return false;
    }
 
    // if shown now
    if (NavToggle.firstChild.data == NavigationBarHide) {
        for (var NavChild = NavFrame.firstChild; NavChild != null; NavChild = NavChild.nextSibling) {
            if (hasClass(NavChild, 'NavContent') || hasClass(NavChild, 'NavPic')) {
                NavChild.style.display = 'none';
            }
        }
    NavToggle.firstChild.data = NavigationBarShow;
 
    // if hidden now
    } else if (NavToggle.firstChild.data == NavigationBarShow) {
        for (var NavChild = NavFrame.firstChild; NavChild != null; NavChild = NavChild.nextSibling) {
            if (hasClass(NavChild, 'NavContent') || hasClass(NavChild, 'NavPic')) {
                NavChild.style.display = 'block';
            }
        }
        NavToggle.firstChild.data = NavigationBarHide;
    }
}
 
// adds show/hide-button to navigation bars
function createNavigationBarToggleButton()
{
    var indexNavigationBar = 0;
    // iterate over all < div >-elements 
    var divs = document.getElementsByTagName("div");
    for (var i = 0; NavFrame = divs[i]; i++) {
        // if found a navigation bar
        if (hasClass(NavFrame, "NavFrame")) {
 
            indexNavigationBar++;
            var NavToggle = document.createElement("a");
            NavToggle.className = 'NavToggle';
            NavToggle.setAttribute('id', 'NavToggle' + indexNavigationBar);
            NavToggle.setAttribute('href', 'javascript:toggleNavigationBar(' + indexNavigationBar + ');');
 
            var isCollapsed = hasClass( NavFrame, "collapsed" );
            /*
             * Check if any children are already hidden.  This loop is here for backwards compatibility:
             * the old way of making NavFrames start out collapsed was to manually add style="display:none"
             * to all the NavPic/NavContent elements.  Since this was bad for accessibility (no way to make
             * the content visible without JavaScript support), the new recommended way is to add the class
             * "collapsed" to the NavFrame itself, just like with collapsible tables.
             */
            for (var NavChild = NavFrame.firstChild; NavChild != null && !isCollapsed; NavChild = NavChild.nextSibling) {
                if ( hasClass( NavChild, 'NavPic' ) || hasClass( NavChild, 'NavContent' ) ) {
                    if ( NavChild.style.display == 'none' ) {
                        isCollapsed = true;
                    }
                }
            }
            if (isCollapsed) {
                for (var NavChild = NavFrame.firstChild; NavChild != null; NavChild = NavChild.nextSibling) {
                    if ( hasClass( NavChild, 'NavPic' ) || hasClass( NavChild, 'NavContent' ) ) {
                        NavChild.style.display = 'none';
                    }
                }
            }
            var NavToggleText = document.createTextNode(isCollapsed ? NavigationBarShow : NavigationBarHide);
            NavToggle.appendChild(NavToggleText);
 
            // Find the NavHead and attach the toggle link (Must be this complicated because Moz's firstChild handling is borked)
            for(var j=0; j < NavFrame.childNodes.length; j++) {
                if (hasClass(NavFrame.childNodes[j], "NavHead")) {
                    NavFrame.childNodes[j].appendChild(NavToggle);
                }
            }
            NavFrame.setAttribute('id', 'NavFrame' + indexNavigationBar);
        }
    }
}
 
addOnloadHook( showTocToggle );
addOnloadHook( createNavigationBarToggleButton );
 
/***********************************************

* Animated Information Bar- by JavaScript Kit (www.javascriptkit.com)
* This notice must stay intact for usage
* Visit JavaScript Kit at http://www.javascriptkit.com/ for this script and 100s more

***********************************************/

function informationbar(){
	this.displayfreq = "always"
	this.content = '<a href="javascript:infobar.close()"><img src="./ow/images/close.gif" style="width: 14px; height: 14px; float: right; border: 0; margin-right: 5px" /></a>'
}

informationbar.prototype.setContent = function(data){
	this.content = this.content + data
	var pos = document.getElementById('globalWrapper');
        var div 
        if (typeof document.createElementNS != 'undefined') {
        	div = document.createElementNS("http://www.w3.org/1999/xhtml","div");
        } else {
        	div = document.createElement("div");
        }
        div.innerHTML = '<div id="informationbar" style="top: -500px">' + this.content + '</div>';
        var nodes = div.childNodes;
	while (nodes.length)
		pos.appendChild(nodes[0]);
}

informationbar.prototype.animatetoview = function(){
	var barinstance = this
	if (parseInt(this.barref.style.top) < 0){
		this.barref.style.top = parseInt(this.barref.style.top) + 5 + "px"
		setTimeout(function(){barinstance.animatetoview()}, 50)
	}
	else{
		if (document.all && !window.XMLHttpRequest)
		this.barref.style.setExpression("top", 'document.compatMode=="CSS1Compat"? document.documentElement.scrollTop+"px" : body.scrollTop+"px"')
	else
		this.barref.style.top = 0
	}
}

informationbar.prototype.close = function(){
	document.getElementById("informationbar").style.display = "none"
	if (this.displayfreq == "session")
		document.cookie = "infobarshown=1; path=/"
}

informationbar.prototype.setfrequency = function(type){
	this.displayfreq = type
}

informationbar.prototype.initialize = function(){
	if (this.displayfreq == "session" && document.cookie.indexOf("infobarshown") == -1 || this.displayfreq == "always"){
		this.barref = document.getElementById("informationbar")
		this.barheight = parseInt(this.barref.offsetHeight)
		this.barref.style.top = this.barheight*(-1) + "px"
		this.animatetoview()
	}
}

window.onunload = function(){
	this.barref = null	
}

function isChrome(){
	try{
		return navigator.userAgent.indexOf('Chrome') != -1
	} catch(e){
		return false;
	}
}

function isSafari(){
	try{
		return 	navigator.vendor.indexOf('Apple') != -1
	} catch(e){
		return false;
	}		
}

function isOpera(){
	try{
		return 	window.opera
	} catch(e){
		return false;
	}		
}

function isIE(){
	try{
		return 	navigator.userAgent.indexOf('MSIE') != -1
	} catch(e){
		return false;
	}		
}

var infobar = new informationbar()


if(isChrome() || isSafari()|| isOpera() || isIE()){
	infobar.setContent('For proper formulae rendering it\'s recommended to use <a href="http://www.mozilla.com/">Firefox browser</a>.')
	infobar.setfrequency('session') //Uncomment this line to set information bar to only display once per browser session!
	infobar.initialize()
} 
 
 //</source>