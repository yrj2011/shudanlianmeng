var isSubmitFlag = false;
var isAlert = true;
var userAgent = navigator.userAgent.toLowerCase();
var is_opera = userAgent.indexOf('opera') != -1 && opera.version();
var is_moz = (navigator.product == 'Gecko') && userAgent.substr(userAgent.indexOf('firefox') + 8, 3);
var is_ie = (userAgent.indexOf('msie') != -1 && !is_opera) && userAgent.substr(userAgent.indexOf('msie') + 5, 3);
var is_ie7 = userAgent.indexOf('msie 7.0') != -1;

String.prototype.trim = function(){
    return this.replace( /(^\s*)|(\s*$)/g, "" );
}

function copy(text) {
    if(is_ie) {
		clipboardData.setData('Text', text);
		alert("复制成功");
	} else if(prompt('请你使用 Ctrl+C 复制到剪贴板', text)) {
		alert("复制成功");
	}
	return false;
}

function avoidReSubmit(id) {	
	dialog(400,250,"正在提交数据","","<div class='submiting'>正在提交数据，请耐心等待</div>");
	if (id)
	    getObj(id).disabled = true;
	allReadonly();
	if (isSubmitFlag) {
		alert("正在提交数据，请耐心等待，无需重复提交");
		return false;
	} else {
		isSubmitFlag = true;
	}
	return true;
}

function dialog(w, h, title, url, content) {
		if (getObj("shadow")) 
			return ;
    var shadow = document.createElement("div");
    shadow.id = "shadow";
		shadow.className = "shadow";
    shadow.style.height = document.body.scrollHeight + "px";
    var dialog = getObj("dialog");
    if (dialog)
        document.body.removeChild(dialog);
    dialog = document.createElement("div");
    dialog.id="dialog";
		dialog.className = "dialog";
    dialog.style.marginLeft = "-" + w/2 + "px";
    dialog.style.width = w + "px";
    dialog.style.height = h + "px";
    if (is_ie7 || is_moz) {
			dialog.style.marginTop = "-" + h/2 + "px";
			dialog.style.position = "fixed";
    } else {
			var sScrollTop=parent.document.body.scrollTop+parent.document.documentElement.scrollTop;
			var iTop=sScrollTop + 80;
			var sTop=iTop>0?iTop:0;
			dialog.style.top = sTop + "px";
    }
    var strHtml = '<div class="dialog_t"></div><table width="100%" border="0" cellspacing="0" cellpadding="0"><tr><td class="dialog_s"></td>'
        + '<td><div class="dialog_r"><div class="dialog_bar"><span style="float:left; padding-left:10px;" id="dialog_title">'
        + title + '</span><span onclick="doCut();" id="ff_ss" style="cursor:pointer;float:right; padding-right:10px;">×</span></div>'
        + '<div id="dialogBody" style="background:#FFF; height:' + (h-16) + 'px;">';
    if (url) {
        if (url.indexOf("?") < 0)
            url += "?";
        url += "&ttime=" + Math.random();
        strHtml += '<iframe src="' + url + '" width="' + (w-16) + 'px" height="' + (h-16) + 'px" frameborder="0" scrolling="auto"></iframe>';
    } else {
        strHtml += content;
    }
    strHtml += '</div></div></td><td class="dialog_s"></td></tr></table><div class="dialog_t"></div>';
    dialog.innerHTML = strHtml;
    document.body.appendChild(shadow);
    document.body.appendChild(dialog);
    
    this.doCut = function(){
        dialog.style.display = "none";
        document.body.removeChild(shadow);
		//document.body.onselectstart = function(){return true;}
        //document.body.oncontextmenu = function(){return true;}
    };
    this.doCut2 = function(url) { reflesh(url); };
}

function reflesh(url) {
    if (url) {
        window.location.href = url;
    } else {
		window.location.reload();
        //url = window.location.href;
        //if (url.indexOf("?") < 0) url += "?";
        //window.location.href = url + "&ttime=" + Math.random();
    }
}

function allReadonly() {
	var objs = document.getElementsByTagName("input");
	for (var i=0; i<objs.length; i++)	
		objs[i].readOnly = true;
	objs = document.getElementsByTagName("textarea");
	for (var i=0; i<objs.length; i++) 
		objs[i].readOnly = true;
}

function goBack() {
	window.location.href=document.referrer;
	return false;
}

function getObj(id) {
	return (typeof (id)=='object')?id:document.getElementById(id);
}

function getValue(id) {
	return getObj(id).value;
}

function getRV(name) {
	var val = "";
	var rs = document.getElementsByName(name);
	for (var i=0; i<rs.length; i++) {
		if (rs[i].checked)
			val = rs[i].value
	}
	return val;
}

function setValue(id, value) {
	getObj(id).value = value;
}

function hide(tid){
	var obj = getObj(tid);
	if (obj)
	  obj.style.display = 'none';
}

function show(tid) {
	var obj = getObj(tid);
	if (obj) 
	  obj.style.display = '';
}

function PageQuery(qs) {
	if(qs.length > 1)
		this.q = qs.substr(1);
	else 
		this.q = null;
	this.keyValuePairs = new Array();
	if(this.q) {
		for(var i=0; i < this.q.split("&").length; i++) {
			this.keyValuePairs[i] = this.q.split("&")[i];
		}
	}
	this.getKeyValuePairs = function() { 
		return this.keyValuePairs; 
	}
	this.getValue = function(s) {
		for(var j=0; j < this.keyValuePairs.length; j++) {
			if(this.keyValuePairs[j].split("=")[0] == s)
				return this.keyValuePairs[j].split("=")[1];
		}
		return false;
	}
	this.getParameters = function() {
		var a = new Array(this.getLength());
		for(var j=0; j < this.keyValuePairs.length; j++) {
			a[j] = this.keyValuePairs[j].split("=")[0];
		}
		return a;
	}
	this.getLength = function() { 
		return this.keyValuePairs.length; 
	} 
}

function setQS() {
	var page = new PageQuery(window.location.search);
	var id = "";
	for (var i=0; i<page.getLength(); i++) {
	    id = page.getParameters()[i];
		if (getObj(id)) {
		    getObj(id).value = unescape( decodeURI(page.getValue(id)) );
		}
	}
}

function getQuery(name) {
	var aParams = document.location.search.substr(1).split('&') ;
	for (i=0 ; i < aParams.length ; i++) {
		var aParam = aParams[i].split('=') ;
		if (aParam.length > 1 && aParam[0] == name) {
			return aParam[1];
		}
	}
	return "";
}

function addClass(obj, className) {
    var str = obj.className.trim();
    if (str.indexOf(className) < 0)
        obj.className = str + " " + className;
}

function removeClass(obj, className) {
    var str = obj.className.trim();
    if (str.indexOf(className) >= 0)
        obj.className = str.replace(className, "");
}

function showQQ(qq) {
    if (qq)
        document.write("<a href='http://wpa.qq.com/msgrd?v=3&uin="+qq+"&site=qq&menu=yes'><img width='25' height='17' border='0'  src='http://wpa.qq.com/pa?p=1:"+qq+":17' alt='' /></a>");
}

/*************字符串校验*************/
function doCheck(checks) {
	var result = true;
	isAlert = true;
	for (var i=0; i<checks.length; i++) {
		var check = checks[i];
		var str = "";
		try {
		    for (var j=0; j<check.length; j++) {
		        if (j == 0)
		            str = check[0] + "('";
		        if (j == (check.length - 1))
		            str += check[j] + "')";
		        else if (j > 0)
		            str += check[j] + "','";
		    }
		    result = eval(str);
		} catch (e) {
			alert(str);
		}
		if (!result && isAlert)
			isAlert = false;
	}
	result = result && isAlert;
	isAlert = true;
	return result;
}

function isMatch(pattern, input){
	return pattern.test(input.trim());
}

function doAlert(msg, obj) {
	addClass(obj, "txt_fail");
	obj.onkeyup = function(){ if(event.keyCode > 30) removeClass(obj, "txt_fail");};
	obj.title = msg;
    if (isAlert) {
	    alert(msg);
	    try {obj.focus();} catch(e) {}
	}
	return false;
}

function isEmpty(id, cname) {
	if (arguments.length == 1) {
		if( id.trim() == "" )
			return false;
		else
			return true;
	} else {	    
		var obj = getObj(id);
		if (obj.value.trim() == "")
			return doAlert(cname + "  不能为空", obj);
		else 
			return true;
	}
}

function isLength(id, cname, min, max){
	var obj = getObj(id);
	if(obj.value.trim().length < min || obj.value.trim().length > max){
		return doAlert(cname + "  长度范围为 " + min+ "～" + max, obj);
	} else {
		return true;
	}
}

function isEqual(id1, id2, cname1, cname2, isFlag){
	if (arguments.length == 2) {
		return id1.trim() == id2.trim();
	} else {
		var obj1 = getObj(id1);
		var obj2 = getObj(id2);
		if (obj1.value.trim() == obj2.value.trim()) {
		    if (isFlag)
		        return doAlert(cname1 + " 和 " + cname2 + "  不允许相同", obj2);
		    else
		        return true;
		} else {
		    if (isFlag)
			    return true;
			else
			    return doAlert(cname1 + " 和 " + cname2 + "  不一致", obj2);
		}
	}
}

function isRange(id, cname, min, max) {
    var obj = getObj(id);
	if(parseFloat(obj.value.trim()) < min || parseFloat(obj.value.trim()) > max){
		return doAlert(cname + "  数值范围为 " + min+ "～" + max, obj);		
	} else {
		return true;
	}
}

function isNum(id, cname) {
	if (arguments.length == 1) {
		return isMatch(/^\d*$/, id.trim());
	} else {
		var obj = getObj(id);
		if (!isMatch(/^\d*$/, obj.value)) 
			return doAlert(cname + "  必须为数字", obj)			
		else 
			return true;
	}
}

function isNumber(id, cname) {
	if (arguments.length == 1) {
		return isMatch(/^\d*$/, id.trim());
	} else {
	    var obj = getObj(id);
		if (!isMatch(/^\d+$/, obj.value)) 
			return doAlert(cname + "  必须为整数", obj);
		else 
			return true;
	}
}

function isEmail(id, cname, isMust){
	if (arguments.length == 1) {
		if( isEmpty(id) ) return false;
		var pattern = /^([a-zA-Z0-9_-])+@([a-zA-Z0-9_-])+(\.[a-zA-Z0-9_-])+/;
		return isMatch(pattern, id);
	} else {
		var obj = getObj(id);
		if (!isMust)
		    if (!isEmpty(obj.value)) return true;
		var pattern = /^([a-zA-Z0-9_-])+@([a-zA-Z0-9_-])+(\.[a-zA-Z0-9_-])+/;
		if (!isMatch(pattern, obj.value))
			return doAlert(cname + "  不是有效的电子邮件地址", obj);
		else 
			return true;			
	}
}

function isStock(id, cname) {
	if (arguments.length == 1) {
		return isMatch( /^\d+(\.\d{1,2})?$/, id.trim() );
	} else {
		var obj = getObj(id);
		if (!isMatch(/^\d+(\.\d{1,2})?$/, obj.value))
			return doAlert(cname + "  小数点后只允许两位", obj);
		else 
			return true;
	}
}

function isMoney(id, cname) {
	if (arguments.length == 1) {
		return isMatch( /^[+-]?\d*(,\d{3})*(\.\d+)?$/g, id.trim() );
	} else {
		var obj = getObj(id);
		if (!isMatch(/^[+-]?\d*(,\d{3})*(\.\d+)?$/g, obj.value))
			return doAlert(cname + "  不是有效的数字", obj);
		else 
			return true;
	}
}

function checkAll(form, prefix, checkall) {
	var checkall = checkall ? checkall : 'chkall';
	for(var i = 0; i < form.elements.length; i++) {
		var e = form.elements[i];
		if(e.name && e.name != checkall && (!prefix || (prefix && e.name.match(prefix)))) {
			e.checked = form.elements[checkall].checked;
		}
	}
}

function thumbImg(obj, w) {
    if (obj.width > w) {
        obj.height = obj.height * w / obj.width;
        obj.width = w;
        obj.title = "新窗口打开预览";
        obj.style.cursor = "pointer";
        obj.onclick = function() {
           openDynaWin("图片预览", "<img src='" + obj.src + "' />");
        };
    }
}

function openDynaWin(theTitle,theContent){
    var str = "<html><head><title>" + theTitle + " - 网</title></head><body><table align='center' width='100%'><tr><td align='center' >";
        str += theContent + "</td></tr><tr><td align='center'><input type='button' style='font-size:9pt' value='关闭窗口' onclick='javascript:window.close()'></td></tr></table></body></html>";
    var w=window.open();
    w.document.write(str);
    w.document.close();
}
