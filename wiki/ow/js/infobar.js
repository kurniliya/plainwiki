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