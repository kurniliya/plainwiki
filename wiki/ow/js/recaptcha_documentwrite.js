// John Resig's document.write() replacement for XHTML compatibility 
// working with Firefox 1.5+, Opera 9, and Safari 2+
// a fix for document.write() usage with XHTML documents served with the doctype "application/xhtml+xml".

document.write = function(str){
    var moz = !window.opera && !/Apple/.test(navigator.vendor);
       
    // Watch for writing out closing tags, we just
    // ignore these (as we auto-generate our own)
    // if ( str.match(/^<\//) ) return;
    
    // Forced to work only with reCAPTCHA APIs
    // In order not to mask possible problems in other scripts
    if ( !str.match(/recaptcha/i) ) return;

    // Make sure & are formatted properly, but Opera
    // messes this up and just ignores it
    if ( !window.opera )
        str = str.replace(/&(?![#a-z0-9]+;)/g, "&amp;");

    // Watch for when no closing tag is provided
    // (Only does one element, quite weak)
    // kurniliya: works improperly with reCAPTCHA API
    // cause initial regex fails in case: '<script ...></script>' adding one more closing tag
    //    str = str.replace(/<([a-z]+)(.*[^\/])>$/, "<$1$2></$1>");
    str = str.replace(/<([a-z]+)([^>\/]*[^\/])>$/, "<$1$2></$1>");
       
    // Mozilla assumes that everything in XHTML innerHTML
    // is actually XHTML - Opera and Safari assume that it's XML
    if ( !moz )
        str = str.replace(/(<[a-z]+)/g, "$1 xmlns='http://www.w3.org/1999/xhtml'");
       
    // The HTML needs to be within a XHTML element
    var div;
    if (typeof document.createElementNS != 'undefined') {
    	div = document.createElementNS("http://www.w3.org/1999/xhtml","div");
    } else {
	div = document.createElement("div");
    }
    
    
    div.innerHTML = str;
       
    // Find the element in the document where to "write"
    var pos;
       
    // Opera and Safari treat getElementsByTagName("*") accurately
    // always including the last element on the page
    // if ( !moz ) {
    //    pos = document.getElementsByTagName("*");
    //    pos = pos[pos.length - 1];
               
        // Mozilla does not, we have to traverse manually
    //} else {
    //    pos = document;
    //    while ( pos.lastChild && pos.lastChild.nodeType == 1 )
    //        pos = pos.lastChild;
    //}
    
    pos = document.getElementById( 'recaptcha_holder' );
       
    // Add all the nodes in that position
    var nodes = div.childNodes;
    while ( nodes.length )
//        pos.parentNode.appendChild( nodes[0] );
        pos.appendChild( nodes[0] );
};