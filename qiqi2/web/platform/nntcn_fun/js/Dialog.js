//------------------------------------------------------
//作用：显示信息对话框
//参数：
//	   msgStr：			要显示的信息内容
//	   infoType：		信息(图片)类别: 0 为消息,1 为是否(将显示取消按钮)
//	   actType：		确定按钮返回的操作:
//						0 为无操作
//						1 为关闭窗口
//						2 为返回上一页
//						3 为本框架转到连接地址
//						4 为新窗口打开连接地址
//						5 为父框架执行连接地址
//						6 为子框架执行连接地址
//						7 为本框架转向自定义地址
//						8 为新窗口转向自定义地址
//						9 为父框架转向自定义地址
//						10 为子框架转向自定义地址
//						11 为刷新当前窗口
//						12 为刷新父窗口
//						13 为刷新父窗口,关闭本窗口
//						其他 或者可填写其他自定义函数
//		obj:			页面上传来的对象,直接写this, 如果对象不是连接则请写连接地址
//------------------------------------------------------
var objPage;
var obj;
function viewTaskList(obj,url) {
var d = new Date();
var e = document.getElementById("showtask"); 
e.style.position = "absolute"; 
e.style.left = getL(obj)+"px";
e.style.top = getT(obj)+document.getElementById(obj).offsetHeight+"px";

if(document.all["showtask"].style.display=="none")
{
document.all["showtask"].style.display="";
document.all.contentframe.src=url+"&rndid="+d.getMinutes()+d.getSeconds();
}else{
document.all["showtask"].style.display="none";
document.all.contentframe.src="about:blank";
}
}

function viewTaskTingList(obj,url) {
var d = new Date();
var e = document.getElementById("showtask"); 
e.style.position = "absolute"; 
e.style.left = (getL(obj)-520)+"px";
e.style.top = (getT(obj)+document.getElementById(obj).offsetHeight-80)+"px";

if(document.all["showtask"].style.display=="none")
{
document.all["showtask"].style.display="";
document.all.contentframe.src=url+"&rndid="+d.getMinutes()+d.getSeconds();
}else{
document.all["showtask"].style.display="none";
document.all.contentframe.src="about:blank";
}
}
 
function hide(){  
document.all["showtask"].style.display="none";
document.all.contentframe.src="about:blank";
}  
 
function getL(id){  
var e=document.getElementById(id);
var  l=e.offsetLeft;  
while(e=e.offsetParent){  
l+=e.offsetLeft;  
}  
return  l  
}  
 
function getT(id){  
var e=document.getElementById(id);
var  t=e.offsetTop;  
while(e=e.offsetParent){  
t+=e.offsetTop;  
}  
return  t  
}
function SetCwinHeight()
{
var cwin=document.getElementById("contentframe");
if (document.getElementById)
{
if (cwin && !window.opera)
{
if (cwin.contentDocument && cwin.contentDocument.body.offsetHeight)
cwin.height = cwin.contentDocument.body.offsetHeight; 
else if(cwin.Document && cwin.Document.body.scrollHeight)
cwin.height = cwin.Document.body.scrollHeight;
}
}
}
function showDialog(msgStr, infoType, actType, obj)
{
	msgStr = msgStr ? msgStr : "系统提示信息";
	infoType = infoType ? infoType : 0;
	actType = actType ? actType : 0;
	obj = obj ? obj : "";
	objPage = obj;
	
	if(infoType == 0)
	{
		alert(msgStr);
		DoActionFromAlert(actType, obj);
	}
	else
	{
		if(confirm(msgStr))
		{
			DoActionFromAlert(actType, obj);
		}
	}
	
	return false;
}

function DoActionFromAlert(actType, obj)
{
	switch (actType)
	{
		case 0:
			break;
		
		case 1:
			window.top.opener = null;window.top.close();
			break;
		
		case 2:
			history.go(-1);
			break;
		
		case 3:
			document.location.href = obj.href;
			break;
		
		case 4:
			window.open(obj.href);
			break;
		
		case 5:
			parent.document.location.href = obj.href;
			break;
		
		case 6:
			frameSrc(obj.href)
			break;
		
		case 7:
			document.location.href = obj;
			break;
		
		case 8:
			window.top.document.location.href = obj;
			break;
		
		case 9:
			parent.document.location.href = obj;
			break;
		
		case 10:
			frameSrc(obj);
			break;
		
		case 11:
			document.location.reload();
			break;
		
		case 12:
			parent.document.location.reload();
			break;
		
		case 13:
			window.top.opener.document.location.reload();
			window.top.opener = null;window.top.close();
			break;
		
		default:
			eval(actType);
			break;
	}
	
	return false;
}

//创建一个隐藏的框架
function createFrame()
{
	var oFrame = document.createElement("iframe");
	oFrame.id = "OptionFrame";
	oFrame.name = "OptionFrame";
	oFrame.width = "0";
	oFrame.height = "0";
	oFrame.src = "";
	document.body.appendChild(oFrame);
	document.getElementById('OptionFrame')['style']['display'] = 'none';
}

//隐藏框架内执行连接地址
function frameSrc(hrefSrc)
{
	document.getElementById('OptionFrame') ? function(){} : createFrame();
	document.getElementById('OptionFrame').src = hrefSrc;
}

    function jsCopy(obj_id){ 
    alert('已复制，请在ie浏览器粘贴，打开进入');
    var objinput=obj_id;
        var e=document.getElementById(objinput);//对象是content 
        window.clipboardData.setData("Text",e.value);
        //window.location.href='/platform/GudeShowNotice.htm';
		var winname = window.open('', '', '');
		winname.opener = null;
		//winname.document.open('text/html', 'replace');
		winname.document.close();
    }
    function jsOnlyCopy(obj_id){ 
    alert('已复制，请黏贴搜索');
    var objinput=obj_id;
        var e=document.getElementById(objinput);//对象是content 
        window.clipboardData.setData("Text",e.value);
        //window.location.href='/platform/GudeShowNotice.htm';
		//var winname = window.open('', '', '');
		//winname.opener = null;
		//winname.document.open('text/html', 'replace');
		//winname.document.close();
    }
function Showfirm(tourl,str)
{//利用对话框返回的值（true 或者 false）
    if(confirm(str))
    {//如果是true ，那么就把页面转向
         location.href=tourl;
     }
    else
    {//否则说明下了，赫赫
     }
}