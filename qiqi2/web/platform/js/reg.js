var http_request = false;

function makeRequest(url) {

	document.all.chkusername.innerHTML = '&nbsp;&nbsp;ϵͳ���ڼ��....'; 
	http_request = false;

	if (window.XMLHttpRequest) { 
		http_request = new XMLHttpRequest();
		if (http_request.overrideMimeType) {
			http_request.overrideMimeType('text/xml');
		}
	} 
	else if (window.ActiveXObject) { 
		try {
			http_request = new ActiveXObject("Msxml2.XMLHTTP"); 
			http_request.setrequestheader("content-type","application/x-www-form-urlencoded; charset=utf-8"); 
		} catch (e) {
			try {
				http_request = new ActiveXObject("Microsoft.XMLHTTP"); 

			} catch (e) {}
		}
	}

	if (!http_request) {
		document.all.chkusername.innerHTML = '&nbsp;&nbsp;�޷���ȡѡ������....'; 
		return false;
	}
	http_request.onreadystatechange = alertContents;
	http_request.open('GET', url, true);
	http_request.send(null);

}

function alertContents() {

	if (http_request.readyState == 4) {
		if (http_request.status == 200) {
			var alertmsg = http_request.responseText 
			if ( alertmsg.indexOf("error1") != -1 ) 
				{
					document.all.chkusername.innerHTML = '&nbsp;&nbsp;�����������2-6���ַ�֮��';
				}
			else
				{
					 if (alertmsg.indexOf("error2") != -1 ) 
						{
							document.all.chkusername.innerHTML = '&nbsp;&nbsp;<font color=red>���û��Ѵ���</font>';
						}
					else if (alertmsg.indexOf("yes") != -1 )  
						{
							document.all.chkusername.innerHTML = '&nbsp;&nbsp;���û���ע��';
						}
				}
				
		} 
		else {
			document.all.chkusername.innerHTML = '&nbsp;&nbsp;��ȡ����,������...';
		}
	}

}
