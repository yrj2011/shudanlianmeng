<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="checklogin.asp"-->
<!--#INCLUDE FILE="../background.asp"-->
<%
Function SendSms(UserName, UserPass, DstMobile, SmsMsg) 

	set http = Server.CreateObject("Microsoft.XMLHTTP")
	http.Open "GET", "http://www.china-sms.com/send/gsend.asp?name="&UserName&"&pwd="&UserPass&"&dst="&DstMobile&"&msg="&SmsMsg&"", false 
	http.setRequestHeader "Content-type:", "text/xml;charset=GB2312"
	'http.setRequestHeader "Charset", "GB2312"
	http.Send
	msg=http.ResponseText
	set	http = nothing
	
	if left(msg,4)<>"id=0" then 
      '���Ͷ��ųɹ�,�ɽ�һ�����ɹ���������
      SendSms=true
      exit Function
    else
      response.write msg
      SendSms=false
	  exit Function
    end if
End Function
if vip="1" then
money=stupvipuser
else
money=stuppuuser
end if

action=HtmlEncode(trim(request.form("action")))
if action="ok" then
username=HtmlEncode(trim(request.form("username")))
message=HtmlEncode(trim(request.form("message")))
mslen=((len(message))\ 70)+1
mslen=formatnumber(mslen,2,-1)
price=mslen*money/100
price=formatnumber(price,2,-1)

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select * From "&jieducm&"_register where username='"&username&"' and dst<>0" ,Conn,3,3  
if rs.eof then
 Response.Write("<script>alert('�޴˻�Ա��Է���û�н����ֻ���֤��');history.back();</script>")
 response.End()
else
if rs("username")=session("useridname") then
 Response.Write("<script>alert('���ܸ��Լ����ͣ�');history.back();</script>")
 response.End()
end if
jdst=rs("dst")
end if	


if formatnumber(price,2,-1)>formatnumber(cunkuan,2,-1) then
    Response.Write("<script>alert('�Բ�����Ĵ���!');history.back();</script>")	  
    response.End()
end if


if SendSms(smsname, smspwdphone, jdst, message) then


Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select cunkuan From "&jieducm&"_register where username='"&session("useridname")&"' " ,Conn,3,3  
if rs.eof then
 Response.Write("<script>alert('���ܳ�ʱ�����');history.back();</script>")
 response.End()
else
rs("cunkuan")=rs("cunkuan")-price
rs.update
rs.close
end if	
 
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_sms " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("susername")=username	
		rs("content")=message
        rs("now")=now()
    	rs.update
    	rs.close
		
     num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=num
		rs("class")="�����ֻ�����"
		rs("content")=session("useridname")&"��"&username&"�����ֻ�����"
		rs("price")="-"&price
		rs("jiegou")="���ͳɹ�"
    	rs.update
    	rs.close	
    Response.Write("<script>alert('���ͳɹ�!');window.location.href='sms.asp';</script>")	  
else
   response.write "����ʧ��!"
end if
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">				
<head>
<title>��Ա����-<%=stupname%> -<%=stuptitle%></title>
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<link href="../css/style.css" rel="stylesheet" type="text/css" />
<LINK href="../css/center.css" type=text/css rel=stylesheet>
 
<script type="text/javascript" src="../js/common.js"></script>
<SCRIPT src="../js/Dialog.js" type=text/javascript></SCRIPT>

</head>
<script type="text/javascript">
function checkForm() {
   var checks = [
		["isEmpty", "username", "������"],
		["isLength", "message", "��������", 1, 700] ];
	var result = doCheck(checks);
	if (result)
		return avoidReSubmit();
	return result;
}

function sltMsg(obj) {
	setValue("message", obj.value);
}

function turnMsg(id) {
    setValue("message", getObj(id).innerHTML);
    alert("ת�����������룬����д������");
    getObj("username").focus();
}

function countMsg() {
	var str = getValue("message");
	var len = str.length;
	getObj("msg_num").innerHTML = len;
	var i = parseInt(len / 70);
	if (len % 70 > 0)
		i = i + 1;
	getObj("msg_cost").innerHTML = i * <%=money%>;
}

function sltAll(state) {
    var boxs = document.getElementsByTagName("input");
	for (var i=0; i<boxs.length; i++) {
		if (boxs[i].name.indexOf("msg_")==0) {
			boxs[i].checked = state;
		}
	}
}

function doMsg(act) {
	var str = "";
    var boxs = document.getElementsByTagName("input");
    for (var i=0; i<boxs.length; i++) {
		if (boxs[i].name.indexOf("msg_")==0 && boxs[i].checked) {
	        str += boxs[i].value + ",";
		}
	}
	if (str == "") {
	    alert("����û�й�ѡ�κ���Ϣ");
	} else {
		reflesh("?ids=" + str);
	}
	return false;
}
</script>
<body>
<!--#include file=../inc/jieducm_top.asp-->
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
<tr>
	    <td>
        <table width="910"  border="0" cellspacing="0" cellpadding="0">
	          <tr>
	          
	            <td width="140" valign="top">
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td height="5"> </td>
    </tr>
</table>
				<table id="menuToolBar" width="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
    <td valign="top" bgcolor="FFFFFF">
    <!--#include file="userleft.asp"--></TD>
</TR>
</table></td>
	            <td width="15">&nbsp;</td>
	          
	            <td valign="top">
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td height="5"></td>
    </tr>
</table>
				<div class=pp9>
                  <div style="PADDING-BOTTOM: 15px; WIDTH: 97%">
                    <div class=pp7>�����ڵ�λ���ǣ�<%=stupname%> &gt;&gt; ��Ա���� &gt;&gt;���Ͷ���&gt;&gt; </div>
                    <div class=pp8><strong>���Ͷ���</strong></div>
					<DIV style="MARGIN-TOP: 10px; LINE-HEIGHT: 200%">
      <UL>
        <LI>* �ڶ���ƽ̨�����Ը���Ҫ��ϵ���˻���ƽ̨�û����Ͷ��ţ�</LI>
        <LI>* ��ͨ�û�ÿ��<%=stuppuuser%>��Ǯ; ����VIPÿ��ֻҪ<%=stupvipuser%>��Ǯ��ʡǮ��ʡʱ��Ŷ�� </LI>
      </UL></DIV>
                    <div class="box_main">
        <form name="myForm" method="post" action="SMS.asp" id="myForm" onsubmit="return checkForm();">
<input name="action" type="hidden" value="ok" />

          <table class="tbl_edit" width="100%" border="0" cellspacing="0" cellpadding="4">
            <tr>
              <td width="120" align="right">�����ˣ�</td>
              <td><input name="username" type="text" class="text_normal" id="username" />  ƽ̨�û���</td>
            </tr>
            <tr>
              <td align="right">�����Ϣ��</td>
              <td><select name="sltQuick" id="sltQuick" onchange="sltMsg(this);">
                  <option value="">��ѡ��</option>
                  
                  <option value="�벻Ҫʹ��������ϵ">�벻Ҫʹ��������ϵ</option>
                  
                  <option value="��Ϊ�����ʱлл">��Ϊ�����ʱлл</option>
                  
                  <option value="����Ҿ����������ƽ̨֧��������Ʒ">����Ҿ����������ƽ̨֧��������Ʒ</option>
                  
                  <option value="�������޸���Ʒ���˷Ѽ۸�������һ��">�������޸���Ʒ���˷Ѽ۸�������һ��</option>
                  
                  <option value="���Ѹ�������Ҿ�����������ƽ̨����">���Ѹ�������Ҿ�����������ƽ̨����</option>
                  
                  <option value="���ѷ������������������ƽ̨ȷ���ջ�">���ѷ������������������ƽ̨ȷ���ջ�</option>
                  
                  <option value="���ѷ�����������ڰ�Сʱ����ȷ���ջ�">���ѷ�����������ڰ�Сʱ����ȷ���ջ�</option>
                  
                  <option value="�����ջ���������ƽ̨ȷ������">�����ջ���������ƽ̨ȷ������</option>
                  
                  <option value="�ѵ��ں������������������ƽ̨�ύ����">�ѵ��ں������������������ƽ̨�ύ����</option>
                  
                </select>
                �����ÿ�ݶ��ѡ���ֱ����ʾ�ڷ���������</td>
            </tr>
            <tr>
              <td align="right" valign="top">�������ݣ�</td>
              <td><span class="f_gray">����ʮ��������һ������ʮһ����һ����ʮ�����������Դ����ƣ�</span><br />
                <textarea name="message" cols="60" rows="5" id="message" onpropertychange="countMsg();"></textarea>
                <br />
                ���Ѿ�����<span class="f_b_org" id="msg_num">0</span>���֣� ���żƷѣ�<span class="f_b_org" id="msg_cost">0</span>��</td>
            </tr>
          </table>
          <div class="btn_d">
            <input name="btnSend" type="submit" class="btn_1" id="btnSend" value="������Ϣ" />
          </div>
        </form>

	<div class=pp8><strong>�ѷ��͵���Ϣ</strong></div>
	 <FORM name=form method=post action="smsdel.asp">
    <table class="tbl_lst" width="100%" border="0" cellspacing="0" cellpadding="3">
      <thead>
       
        <tr>
          <td width="60" class="tbl_s" style="font-weight:normal;">  <input id=chkAll 
                  onClick=CheckAll(this.form) type=checkbox 
                  value=checkbox name=chkAll>
            ѡ��</td>
          <td>������</td>
          <td>��������</td>
          <td width="130">����ʱ��</td>
          <td width="40">����</td>
        </tr> 
		  <%	
   	sql="SELECT * FROM "&jieducm&"_sms where username='"&session("useridname")&"' order by id desc"
	Set Rs = Server.CreateObject("Adodb.RecordSet")
	Rs.Open Sql,conn,1,1
	if not rs.eof then
	dim maxperpage,url,PageNo
	 url="sms.asp"
	 rs.pagesize=10
	 PageNo=REQUEST("PageNo")
	 if PageNo="" or PageNo=0 then PageNo=1
	 RS.AbsolutePage=PageNo
	 TSum=rs.pagecount
	 maxperpage=rs.pagesize
	 RowCount=rs.PageSize
	   PageNo=PageNo+1
	   PageNo=PageNo-1
	 if CINT(PageNo)>1 then
	    if CINT(PageNo)>CINT(TSum) then
		  response.Write("�Բ���û������Ҫ��ҳ��")
          Response.End
	    end if
	 end if
		end if    
     if PageNo<0 then
	    response.Write("û����һҳ!")
		Response.End
	 End if
	
	if PageNo=1 then i=rs.RECORDCOUNT else i=rs.RECORDCOUNT-(30*(PageNo-1))
 DO WHILE NOT rs.EOF AND RowCount>0%>
		  
		  <tr>
          <td class="tbl_s" style="font-weight:normal;"><input type="checkbox" name="adid" value="<%=rs("id")%>" onClick=Checked(form)></td>
          <td><%=rs("susername")%></td>
          <td><%=rs("content")%></td>
          <td><%=rs("Now")%></td>
          <td><a href="#1" onClick="javascript:showDialog('��ȷ��Ҫɾ������Ϣ��',1,7,'smsdel.asp?adid=<%=rs("id")%>')" title="��ȷ��Ҫɾ������Ϣ��">ɾ��</a></td>
        </tr>
		
     <%
      rs.MoveNext
	  RowCount = RowCount - 1 
      Loop%>
	  <tr>
		    <td colspan="5" class="tbl_s" style="font-weight:normal;"><INPUT title=ɾ�� onclick=DelAll() type=button value=ɾ�� name=Submit><div align="center">
		      <% call showpage(url,rs.RECORDCOUNT,maxperpage,false,true,"������") %>
		      </div></td>
		    </tr>
      </thead>
      <tbody>
      </tbody>
    </table>
	</FORM>
      </div>
				  </div>
	            </div></td>
	          </tr>
	        </table>	    </td>
	  </tr>
</table>
  <!--#include file=../inc/jieducm_bottom.asp-->
<%rs.close
set rs=nothing
conn.close
set conn=nothing%>
</body>
</html> <SCRIPT language=javascript>
function CheckAll(form)
{
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.Name != "chkAll")
       e.checked = form.chkAll.checked;
    }
}
function Checked()
{
	var j = 0
	for(i=0;i < document.form.elements.length;i++){
		if(document.form.elements[i].name == "adid"){
			if(document.form.elements[i].checked){
				j++;
			}
		}
	}
	return j;
}

function DelAll()
{
	if(Checked()  <= 0){
		alert("������ѡ��1����Ϣ!");
	}	
	else{
		if(confirm("ȷ��Ҫɾ��ѡ�����Ϣ��\n�˲��������Իָ���")){
			form.action="smsdel.asp";
			form.submit();
		}
	}
}

</SCRIPT>