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
		case 14:
			parent.document.location.reload();
			window.close();
			break;
		case 15:
			window.top.document.location.href = obj;
			window.close();
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

function jsCopy(obj_id)
{ 

    var objinput=obj_id;   
    var e=document.getElementById(objinput);//对象是content 
	window.clipboardData.setData("Text",e.value);
	alert("该商品地址已复制到您的粘贴板，请您在新窗口的地址栏粘贴此链接打开该商品！");
	var winname = window.open('', '', '');
	winname.opener = null;
	winname.document.open('text/html', 'replace');
} 


	
function jsWindon(url)
{
    var winname = window.open('', '', '');
	winname.opener = null;
	winname.document.open('text/html', 'replace');
	winname.document.write("<SCRIPT>window.location.replace(\""+url+"\");</SCRIPT>");
	winname.document.close();
}
function jsOpen(url)
{
    window.location.replace(url);
}