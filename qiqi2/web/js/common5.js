function $(a){return document.getElementById(a);}

function createxmlhttp()
{
	var xmlhttp;
	try
	{
		xmlhttp = new ActiveXObject('Msxml2.XMLHTTP');//IE5.5+
	}
	catch (e)
	{
		xmlhttp = new XMLHttpRequest();//FF等其它浏览器
	}
	return xmlhttp;
} 

function urlcheck(obj){
	var flag=0;
	var arr = ["ProUrl1","ProUrl2","ProUrl3","ProUrl4","ProUrl5","ProUrl6"];
	//alert(arr);
	for(var i=0; i<arr.length; i++){
		if(obj.value!="" && obj.value == $(arr[i]).value){
			flag += 1; 
		}
	}
	if(flag>=2){ 
	alert('商品地址有重复')
	obj.value="";
	obj.focus();
	}
}
function totalprice(){
	var arr = ["Price1","Price2","Price3","Price4","Price5","Price6"],total=0;
	for(var i=0; i<arr.length; i++){
		//alert(parseFloat($(arr[i]).value));
		if($(arr[i]).value !=""){
		total += parseFloat($(arr[i]).value); 
		}
	}
	//alert(total);
	$("Price7").value=total;	
}


function check(obj)
{
		urlcheck(obj);
		var nickname	= $( "nick_name" );
		var zurl,znickname;
        if(obj.value==""){
            //alert("请输入商品链接");
			//obj.focus();
            return;
        }else if(obj.value.search(/taobao/i)==-1){
			alert("你输入的是非淘宝网链接，请重新输入")
			obj.value="";
			obj.focus();
			return;			
		}else{
            zurl = obj.value;
        }
        znickname = nickname.value;
		//提交
		//var url="?go=check",post="url="+escape(zurl)+"&nickname="+znickname
		var url="PushMissionvip.asp?go=check&url="+escape(zurl)+"&nickname="+znickname
		//alert(post);
		//return;

	//AJAX处理
	var xmlhttp=new createxmlhttp();
		xmlhttp.open("POST",url,true);
		xmlhttp.setRequestHeader("Cache-Control","no-cache");
		xmlhttp.setRequestHeader("Content-Type", "application/x-www-form-urlencoded; charset=gb2312");
		xmlhttp.send();
		xmlhttp.onreadystatechange=function (){
		if(xmlhttp.readyState==4&&xmlhttp.status==200){
			var vback=xmlhttp.responseText;
			//alert(vback);
			if (vback=="no"){
				//alert(post)
			alert("你输入的淘宝地址掌柜名与你填写的不符，请重新填写");
			obj.value="";
			obj.focus();
			return;
			}
			//注销
			xmlhttp=null;
		}else{
			
		}
	}
}
