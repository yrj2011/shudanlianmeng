


//����XMLHttpRequest����
var xmlhttp;
try{
    xmlhttp= new ActiveXObject('Msxml2.XMLHTTP');
}catch(e){
    try{
        xmlhttp= new ActiveXObject('Microsoft.XMLHTTP');
    }catch(e){
        try{
            xmlhttp= new XMLHttpRequest();
        }catch(e){}
    }
}

function getPart(url,pageno1,s){
//document.getElementById("partdiv").innerHTML = "���������С�������������";
var pn=pageno1;
var ss=s;
var urll=url+"?"+"pageno="+pn+"&sort="+ss;
//alert(urll);
    xmlhttp.open("get",urll,true);
    xmlhttp.onreadystatechange = function(){

	  
	    if(xmlhttp.readyState!=4)
		{
			document.getElementById("jiazai").style.display="";
		}
        if(xmlhttp.readyState == 4)
        {   
		
            if(xmlhttp.status == 200)
            {
				document.getElementById("jiazai").style.display="none";
                if(xmlhttp.responseText!=""){

                    document.getElementById("partdiv").innerHTML =xmlhttp.responseText;        
                }
            }
            else{
                document.getElementById("partdiv").innerHTML = "�������볬ʱ,�����µ�½";
            }
        }
    }
    xmlhttp.setRequestHeader("If-Modified-Since","0");
    xmlhttp.send(null);
}


