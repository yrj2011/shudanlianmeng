//------------------------------------------------------
//���ã���ʾ��Ϣ�Ի���
//������
//	   msgStr��			Ҫ��ʾ����Ϣ����
//	   infoType��		��Ϣ(ͼƬ)���: 0 Ϊ��Ϣ,1 Ϊ�Ƿ�(����ʾȡ����ť)
//	   actType��		ȷ����ť���صĲ���:
//						0 Ϊ�޲���
//						1 Ϊ�رմ���
//						2 Ϊ������һҳ
//						3 Ϊ�����ת�����ӵ�ַ
//						4 Ϊ�´��ڴ����ӵ�ַ
//						5 Ϊ�����ִ�����ӵ�ַ
//						6 Ϊ�ӿ��ִ�����ӵ�ַ
//						7 Ϊ�����ת���Զ����ַ
//						8 Ϊ�´���ת���Զ����ַ
//						9 Ϊ�����ת���Զ����ַ
//						10 Ϊ�ӿ��ת���Զ����ַ
//						11 Ϊˢ�µ�ǰ����
//						12 Ϊˢ�¸�����
//						13 Ϊˢ�¸�����,�رձ�����
//						���� ���߿���д�����Զ��庯��
//		obj:			ҳ���ϴ����Ķ���,ֱ��дthis, �����������������д���ӵ�ַ
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
	msgStr = msgStr ? msgStr : "ϵͳ��ʾ��Ϣ";
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

//����һ�����صĿ��
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

//���ؿ����ִ�����ӵ�ַ
function frameSrc(hrefSrc)
{
	document.getElementById('OptionFrame') ? function(){} : createFrame();
	document.getElementById('OptionFrame').src = hrefSrc;
}

    function jsCopy(obj_id){ 
    alert('�Ѹ��ƣ�����ie�����ճ�����򿪽���');
    var objinput=obj_id;
        var e=document.getElementById(objinput);//������content 
        window.clipboardData.setData("Text",e.value);
        //window.location.href='/platform/GudeShowNotice.htm';
		var winname = window.open('', '', '');
		winname.opener = null;
		//winname.document.open('text/html', 'replace');
		winname.document.close();
    }
    function jsOnlyCopy(obj_id){ 
    alert('�Ѹ��ƣ����������');
    var objinput=obj_id;
        var e=document.getElementById(objinput);//������content 
        window.clipboardData.setData("Text",e.value);
        //window.location.href='/platform/GudeShowNotice.htm';
		//var winname = window.open('', '', '');
		//winname.opener = null;
		//winname.document.open('text/html', 'replace');
		//winname.document.close();
    }
function Showfirm(tourl,str)
{//���öԻ��򷵻ص�ֵ��true ���� false��
    if(confirm(str))
    {//�����true ����ô�Ͱ�ҳ��ת��
         location.href=tourl;
     }
    else
    {//����˵�����ˣ��պ�
     }
}