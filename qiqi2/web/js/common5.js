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
		xmlhttp = new XMLHttpRequest();//FF�����������
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
	alert('��Ʒ��ַ���ظ�')
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
            //alert("��������Ʒ����");
			//obj.focus();
            return;
        }else if(obj.value.search(/taobao/i)==-1){
			alert("��������Ƿ��Ա������ӣ�����������")
			obj.value="";
			obj.focus();
			return;			
		}else{
            zurl = obj.value;
        }
        znickname = nickname.value;
		//�ύ
		//var url="?go=check",post="url="+escape(zurl)+"&nickname="+znickname
		var url="PushMissionvip.asp?go=check&url="+escape(zurl)+"&nickname="+znickname
		//alert(post);
		//return;

	//AJAX����
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
			alert("��������Ա���ַ�ƹ���������д�Ĳ�������������д");
			obj.value="";
			obj.focus();
			return;
			}
			//ע��
			xmlhttp=null;
		}else{
			
		}
	}
}
