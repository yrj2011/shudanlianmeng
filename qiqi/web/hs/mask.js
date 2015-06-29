var $ = function(id){
 return document.getElementById(id)||eval(id);
}
var ie = ie6 = ie7 = ff = 0;
if( typeof ActiveXObject != "undefined"){
	ie = 1;
	if (window.navigator.userAgent.indexOf("MSIE 6")>0)
	{
		ie6 = 1; 
	}
}
function getPageSize(){
        var body = document.documentElement;
        var bodyOffsetWidth = 0;
        var bodyOffsetHeight = 0;
        var bodyScrollWidth = 0;
        var bodyScrollHeight = 0;
        var pageDimensions = [0,0];
       var test = [];
        pageDimensions[0]=body.clientHeight; 
        pageDimensions[1]=body.clientWidth; 
        bodyOffsetWidth=body.offsetWidth;
        bodyOffsetHeight=body.offsetHeight;
        bodyScrollWidth=body.scrollWidth;
        bodyScrollHeight=body.scrollHeight;
        if(bodyOffsetHeight > pageDimensions[0])
        {
            pageDimensions[0]=bodyOffsetHeight;
        }    
       
        if(bodyOffsetWidth > pageDimensions[1])
        {
            pageDimensions[1]=bodyOffsetWidth;
        }    
       
        if(bodyScrollHeight > pageDimensions[0])
        {
            pageDimensions[0]=bodyScrollHeight; 
        }     
       
        if(bodyScrollWidth > pageDimensions[1])
        {
            pageDimensions[1]=bodyScrollWidth;
        }   
		if (ie6)
		{
			return [pageDimensions[0]-4, pageDimensions[1]-21]
		}
        return pageDimensions; 
}
function setCss(elem , style){
	var oldCss = {};
	for(var p in style){
		oldCss[p] = elem.style[p];
		elem.style[p] = style[p];
	}
	return oldCss;
}
function getHeight(elem){
	if (elem.style.display && elem.style.display != "none" && elem.style.display != "")
	{
		return elem.offsetHeight||parseInt(elem.style.height);
	}else{
		var oldStyle = setCss(elem,
		{
			visibility:"hidden"
		}
		);
		var height = elem.offsetHeight||parseInt(elem.style.height);
		setCss(elem , oldStyle);
		return height;
	}
}
function getWidth(elem){
	if (elem.style.display != "none" && elem.style.display != "")
	{
		return elem.clientWidth||parseInt(elem.style.width);
	}else{
		var oldStyle = setCss(elem,
		{
			visibility:"hidden"
		}
		);
		var width = elem.clientWidth||parseInt(elem.style.width);
		setCss(elem , oldStyle);
		return width;
	}
}

var eventHandle = {
  addEvent:function(obj, eventType, func){
	 if(ie){
		if(typeof func == "function"){
			obj.attachEvent("on"+eventType, func);
		}else{
			eval("obj.on"+eventType+" = function(){"+func+"}");
		}
	 }else{
		if(typeof func == "function"){
			obj.addEventListener(eventType, func, false);
		}else{
			eval("obj.on"+eventType+" = func()");
		}
	 }
  },
  deleteEvent:function(obj, eventType, func){
	if(ie){
		obj.detachEvent("on"+eventType, func);
	}else{
		obj.removeEventListener(eventType, func , false);
	}
     
  }
}
Object.extend = function(d,s){
	for(p in s){
		d[p] = s[p];
	}
	return d;
}
var Mask = {
	maskDiv:null,
	winDiv:null,
	winTitle:null,
	winContent:null,
	timer:null,
	tempOpa:0,
	initInfo:[],
	aSelect:[],
	closable:1,
	doMask:function(opt){
		var aSelect = document.getElementsByTagName("select");
		for(var i = 0; i < aSelect.length; i++){
			if (aSelect[i].style.display != "hidden" || aSelect[i].style.display != "none"){
				this.aSelect.push(aSelect[i]);
				aSelect[i].style.visibility = "hidden";
			}
		}
		if (!Mask.maskDiv)
		{		
		 var maskDiv = document.createElement("div");
		 this.maskDiv = maskDiv;
		 maskDiv.id = "maskDiv";
		}
		Mask.initInfo = getPageSize();
		var style = {
			position:"absolute",
			background:"#ccc",
			margin:"0px",
			padding:"0px",
			opacity:"0",
			filter:"alpha(opacity=0)",
			height:Mask.initInfo[0]+"px",
			width:Mask.initInfo[1]+"px",
			left:"0px",
			top:"0px",
			display:"block",
			zIndex:"10"
		};
		Mask.fadeIn()
		Object.extend(Mask.maskDiv.style, style);
		document.body.appendChild(Mask.maskDiv);
	},
	fadeIn:function(){
		if (Mask.tempOpa >= 50)
		{
			clearTimeout(Mask.timer);
			Mask.tempOpa = 0;
		}else{
			var style = {
				opacity:(Mask.tempOpa)/100,
				filter:"alpha(opacity="+Mask.tempOpa+")"
			};
			Mask.tempOpa += 5;
			Object.extend(Mask.maskDiv.style, style);
			Mask.timer = setTimeout(function(){Mask.fadeIn()}, 10);
		}
	},
	buildWindow:function(opt){
		Mask.doMask();
		if (!this.winDiv)
		{
		var winDiv = document.createElement("div");
		winDiv.id = "winDiv";
		this.winDiv = winDiv;
		var winTitle = document.createElement("div");	
		winTitle.id = "winTitle";
		var titleHtml = "<div style=\"float:left;\">"+opt.title+"</div>";
		
		if(this.closable){
			titleHtml += "<div style=\"float:right;width:40px; text-align:right\" ><span id=\"win-close\" style=\"cursor:pointer\"></span></div>";
		}
		winTitle.innerHTML =  titleHtml;
		
		this.winTitle = winTitle;
		var winBody = document.createElement("div");
		winBody.id = "winBody";
		var winContent = document.createElement("div");
		this.winContent = winContent;
		winContent.id = "winContent";
		winBody.appendChild(winContent);
		winDiv.appendChild(winTitle);
		winDiv.appendChild(winBody);
		document.body.appendChild(winDiv);
		}
		this.winDiv.style.display = "block";
		this.bindEvent();
		if (opt.width)
		{
			this.winDiv.style.width = opt.width+"px";
		}
		
		this.setContent(opt.content);
	},
	setContent:function(content){
		if (typeof content == "object")
		{
			content.style.display = "block";
			this.winContent.innerHTML = "";
			this.winContent.appendChild(content);
		}else{
			this.winContent.innerHTML = content;
		}
		Mask.fixLocation();
	},
	fixLocation:function(){
		var height = getHeight(Mask.winDiv);
		var top = (document.documentElement.clientHeight-height)/2+document.documentElement.scrollTop;
		Mask.winDiv.style.top = top+"px";
		
		var width = getWidth(Mask.winDiv);
		var left = (document.documentElement.clientWidth-width)/2;
		Mask.winDiv.style.left = left+"px";
	},
	bindEvent:function(){
		$("win-close").onclick = function(){
			Mask.undo();
		}
	},
	undo:function(){
		for(var i = 0; i < this.aSelect.length;i++){
			this.aSelect[i].style.visibility  = "visible";
		}
		setCss(this.winDiv, {display:"none"});
		Mask.maskDiv.style.display = "none";
		document.body.style.overflow = "auto";
	}
}