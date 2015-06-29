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
      '发送短信成功,可进一步做成功后续处理
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
 Response.Write("<script>alert('无此会员或对方还没有进行手机认证！');history.back();</script>")
 response.End()
else
if rs("username")=session("useridname") then
 Response.Write("<script>alert('不能给自己发送！');history.back();</script>")
 response.End()
end if
jdst=rs("dst")
end if	


if formatnumber(price,2,-1)>formatnumber(cunkuan,2,-1) then
    Response.Write("<script>alert('对不起你的存款不足!');history.back();</script>")	  
    response.End()
end if


if SendSms(smsname, smspwdphone, jdst, message) then


Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select cunkuan From "&jieducm&"_register where username='"&session("useridname")&"' " ,Conn,3,3  
if rs.eof then
 Response.Write("<script>alert('可能超时引起的');history.back();</script>")
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
		rs("class")="发送手机短信"
		rs("content")=session("useridname")&"给"&username&"发送手机短信"
		rs("price")="-"&price
		rs("jiegou")="发送成功"
    	rs.update
    	rs.close	
    Response.Write("<script>alert('发送成功!');window.location.href='sms.asp';</script>")	  
else
   response.write "发送失败!"
end if
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">				
<head>
<title>会员中心-<%=stupname%> -<%=stuptitle%></title>
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
		["isEmpty", "username", "收信人"],
		["isLength", "message", "发送内容", 1, 700] ];
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
    alert("转发内容已输入，请填写收信人");
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
	    alert("您还没有勾选任何信息");
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
                    <div class=pp7>您现在的位置是：<%=stupname%> &gt;&gt; 会员中心 &gt;&gt;发送短信&gt;&gt; </div>
                    <div class=pp8><strong>发送短信</strong></div>
					<DIV style="MARGIN-TOP: 10px; LINE-HEIGHT: 200%">
      <UL>
        <LI>* 在短信平台您可以给想要联系的人或者平台用户发送短信，</LI>
        <LI>* 普通用户每条<%=stuppuuser%>分钱; 尊享VIP每条只要<%=stupvipuser%>分钱，省钱又省时间哦！ </LI>
      </UL></DIV>
                    <div class="box_main">
        <form name="myForm" method="post" action="SMS.asp" id="myForm" onsubmit="return checkForm();">
<input name="action" type="hidden" value="ok" />

          <table class="tbl_edit" width="100%" border="0" cellspacing="0" cellpadding="4">
            <tr>
              <td width="120" align="right">收信人：</td>
              <td><input name="username" type="text" class="text_normal" id="username" />  平台用户名</td>
            </tr>
            <tr>
              <td align="right">快捷信息：</td>
              <td><select name="sltQuick" id="sltQuick" onchange="sltMsg(this);">
                  <option value="">请选择</option>
                  
                  <option value="请不要使用旺旺联系">请不要使用旺旺联系</option>
                  
                  <option value="请为任务加时谢谢">请为任务加时谢谢</option>
                  
                  <option value="请买家尽快在网店和平台支付任务商品">请买家尽快在网店和平台支付任务商品</option>
                  
                  <option value="请卖家修改商品加运费价格与任务一致">请卖家修改商品加运费价格与任务一致</option>
                  
                  <option value="我已付款，请卖家尽快在网店与平台发货">我已付款，请卖家尽快在网店与平台发货</option>
                  
                  <option value="我已发货，请买家在网店与平台确认收货">我已发货，请买家在网店与平台确认收货</option>
                  
                  <option value="我已发货，请买家在八小时后再确认收货">我已发货，请买家在八小时后再确认收货</option>
                  
                  <option value="我已收货，请卖家平台确认任务">我已收货，请卖家平台确认任务</option>
                  
                  <option value="已到期好评，请买家在网店与平台提交好评">已到期好评，请买家在网店与平台提交好评</option>
                  
                </select>
                任务常用快捷短语，选择后直接显示在发送内容里</td>
            </tr>
            <tr>
              <td align="right" valign="top">发送内容：</td>
              <td><span class="f_gray">（七十字以内算一条，七十一字至一百四十字算两条…以此类推）</span><br />
                <textarea name="message" cols="60" rows="5" id="message" onpropertychange="countMsg();"></textarea>
                <br />
                您已经输入<span class="f_b_org" id="msg_num">0</span>个字； 短信计费：<span class="f_b_org" id="msg_cost">0</span>分</td>
            </tr>
          </table>
          <div class="btn_d">
            <input name="btnSend" type="submit" class="btn_1" id="btnSend" value="发送信息" />
          </div>
        </form>

	<div class=pp8><strong>已发送的信息</strong></div>
	 <FORM name=form method=post action="smsdel.asp">
    <table class="tbl_lst" width="100%" border="0" cellspacing="0" cellpadding="3">
      <thead>
       
        <tr>
          <td width="60" class="tbl_s" style="font-weight:normal;">  <input id=chkAll 
                  onClick=CheckAll(this.form) type=checkbox 
                  value=checkbox name=chkAll>
            选择</td>
          <td>收信人</td>
          <td>短信内容</td>
          <td width="130">发送时间</td>
          <td width="40">操作</td>
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
		  response.Write("对不起没有您想要的页数")
          Response.End
	    end if
	 end if
		end if    
     if PageNo<0 then
	    response.Write("没有这一页!")
		Response.End
	 End if
	
	if PageNo=1 then i=rs.RECORDCOUNT else i=rs.RECORDCOUNT-(30*(PageNo-1))
 DO WHILE NOT rs.EOF AND RowCount>0%>
		  
		  <tr>
          <td class="tbl_s" style="font-weight:normal;"><input type="checkbox" name="adid" value="<%=rs("id")%>" onClick=Checked(form)></td>
          <td><%=rs("susername")%></td>
          <td><%=rs("content")%></td>
          <td><%=rs("Now")%></td>
          <td><a href="#1" onClick="javascript:showDialog('你确认要删除此信息吗！',1,7,'smsdel.asp?adid=<%=rs("id")%>')" title="你确认要删除此信息吗！">删除</a></td>
        </tr>
		
     <%
      rs.MoveNext
	  RowCount = RowCount - 1 
      Loop%>
	  <tr>
		    <td colspan="5" class="tbl_s" style="font-weight:normal;"><INPUT title=删除 onclick=DelAll() type=button value=删除 name=Submit><div align="center">
		      <% call showpage(url,rs.RECORDCOUNT,maxperpage,false,true,"条主题") %>
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
		alert("您至少选择1条消息!");
	}	
	else{
		if(confirm("确定要删除选择的消息吗？\n此操作不可以恢复！")){
			form.action="smsdel.asp";
			form.submit();
		}
	}
}

</SCRIPT>