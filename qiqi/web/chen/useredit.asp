<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">

<!--#include file="inc/conn.asp"-->
<!--#INCLUDE FILE="inc/md5.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/function.asp"-->
<%
'------------------------------------------------------------------
'***************************淘宝信誉互刷系统***********************
'本程序专为淘宝免费互刷打造，是目前市面上功能比较齐全的平台
'程序开发：捷度传媒
'官方主页：http://www.jieducm.cn
'技术支持：robin 
'QQ;616591415
'电话：13598006137
'版权：版权属于捷度传媒信息技术有限公司，不得拷贝、修改，侵权必究。
'Copyright (C) 2008－2011 捷度传媒信息技术有限公司 版权所有
'*****************************************************************
'------------------------------------------------------------------
%>
<LINK href="../images/Default.css" type=text/css rel=stylesheet>
<LINK href="../images/Bbs.css" type=text/css rel=stylesheet>
<LINK href="../images/style.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2900.2180" name=GENERATOR>
<style type="text/css">
<!--
.STYLE1 {font-size: 16px}
-->
</style>
</HEAD>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">



<SCRIPT language=javascript>
function  save_onclick()
{
  var strTemp = document.form.password.value;
  var strTemp1 = document.form.password2.value;
  if (strTemp!= strTemp1 )
  {
      alert("两次填写密码不一致！");
      document.form.password.focus();
      return false;
  }
    var qq1=document.form.QQ.value;
  if(qq1=="")
  {
    alert("请输入您的QQ！");
	document.form.QQ.focus();
	return false;
	}
	    var Email1=document.form.Email.value;
  if(Email1=="")
  {
    alert("请输入您的Email！");
	document.form.Email.focus();
	return false;
	}
 
  var strTemp = document.form.faceheight.value;
  if ((strTemp <1) || (strTemp >130))
  {
      alert("头像长度小于1或大于130");
      document.form.faceheight.focus();
      return false;
  }
  var strTemp = document.form.facewidth.value;
  if ((strTemp <1) || (strTemp >130))
  {
      alert("头像高度小于1或大于130");
      document.form.facewidth.focus();
      return false;
  }


  
  return true;  
}
</SCRIPT>
<% 
username=request("username")
Set rs=server.createobject("ADODB.RECORDSET")
   rs.open "Select * From "&jieducm&"_register where username='"&username&"'",Conn,3,3
%>
<FORM action="usersave.asp" method="POST" name="form" id="form" onSubmit="return save_onclick()">
<DIV style="WIDTH: 950px">
<DIV class=regheight 
style="PADDING-RIGHT: 15px; PADDING-LEFT: 15px; PADDING-BOTTOM: 15px; PADDING-TOP: 15px"><BR>
  <BR>
<DIV style="COLOR: red"><SPAN id=ctl00_AllBaseContentPlaceHolder_MsgLabel 
style="COLOR: red"></SPAN></DIV>
<DIV>
<P align="left">用户名：</P>
<div align="left"><INPUT id=username maxLength=50 
name=username value="<%=rs("username")%>"  readonly></div>
</DIV>
<div align="left"><BR>
</div>
<DIV>
<P align="left">登录密码：</P>
<div align="left">
  <INPUT id=password type=password 
maxLength=16 name="password"> 
  <FONT 
color=#ff0000>*</FONT>不修改请留空<BR>
</div>
</DIV>
<DIV>
<P align="left">重复密码：</P>
<div align="left">
  <INPUT id=password2 type=password 
maxLength=16 name=password2 > 
  <FONT 
color=#ff0000>*</FONT> 重复上面的密码</div>
</DIV>
<div align="left"><BR>
</div>
<DIV>
<P align="left">QQ：</P>
<div align="left">
  <INPUT id=QQ maxLength=15 
name=QQ value="<%=rs("qq")%>"> 
  <FONT color=#ff0000>*</FONT> 
  刷宝时必用<SPAN id=ctl00_AllBaseContentPlaceHolder_RequiredFieldValidator3 
style="DISPLAY: none; COLOR: red">* QQ不能为空！</SPAN></div>
</DIV>
<div align="left"><BR>
</div>
<DIV>
<P align="left">电子邮箱：</P>
<div align="left">
  <INPUT id=Email maxLength=50 
name=Email value="<%=rs("mail")%>"> 
  <FONT color=#ff0000>*</FONT> 
  取回密码使用，注册后不能更改<SPAN id=ctl00_AllBaseContentPlaceHolder_RequiredFieldValidator4 
style="DISPLAY: none; COLOR: red">* 电子邮箱不能为空！</SPAN><SPAN 
id=ctl00_AllBaseContentPlaceHolder_RegularExpressionValidator1 
style="DISPLAY: none; COLOR: red">* 电子邮箱格式不正确！</SPAN></div>
</DIV>
<div align="left"><BR>
</div>

<DIV>
<P align="left">密码问题：</P>
<div align="left">
  <input name="weiti" type="text" id="weiti" size="30" value="<%=rs("weiti")%>" />
  <FONT color=#ff0000>*</FONT> 
  忘记密码的提示问题</div>
</DIV>

<div align="left"><BR>
</div>

<DIV>
<P align="left">问题答案：</P>
<div align="left">
  <input name="daai" type="text" id="daai" size="30"  value="<%=rs("daai")%>"/>
  <FONT color=#ff0000>*</FONT> 
  忘记密码的问题答案</div>
</DIV>

<div align="left"><BR>
</div>

<DIV>
<P align="left">操作码：</P>
<div align="left">
  <INPUT id=czm maxLength=50 
name=czm value="<%=rs("czm")%>">
  <FONT color=#ff0000>*</FONT> 
  互刷时必用 6-16位* 操作码不能为空！</div>
</DIV>

<BR>
<BR>
<DIV style="COLOR: red">以下为选填</DIV><BR>
<BR>
<DIV>
<P>手机：</P><INPUT id=phone maxLength=20 
name=phone value="<%=rs("phone")%>"> 请填写真实</DIV><BR>
<DIV>
<P>绑定的手机：</P><INPUT id=dst maxLength=20 
name=dst value="<%=rs("dst")%>"> 请填写真实 0表示没有绑定手机</DIV><BR>
<DIV>
<P>推荐人：</P><INPUT id=tjr maxLength=20 
name=tjr value="<%=rs("tjr")%>"> 请填写真实</DIV>
<br>
<DIV>
<P>真实姓名：</P><INPUT id=rname maxLength=20 
name=rname value="<%=rs("rname")%>"> 请填写真实</DIV><BR>
<DIV>
<P>店铺地址：</P><input type="text" name="HomePage" size="30" Value="<%=rs("homepage")%>" > 填写的话，我们可以帮你推广呵</DIV><BR>
<DIV>
<P>店铺描述：</P><INPUT id=shopnote maxLength=200 
name=shopnote value="<%=rs("shopnote")%>"> 填写的话，我们可以帮你推广呵</DIV><BR>
<DIV>
<font 
      color=#000000><b>性别：</b></font>
<span class="sex"><span style="WIDTH: 10px">&nbsp;&nbsp;&nbsp;
<select name="sex" size="1" id="select2">
  <option value="1" <%if rs("sex")=1 then%>selected<%end if%>>男</option>
  <option value="0"<%if rs("sex")=0 then%>selected<%end if%>>女</option>
</select>
</span></span></DIV><br>
<DIV>
<P>联系地址：</P><INPUT id=address maxLength=100 
name=address value="<%=rs("address")%>"></DIV><BR>
<BR>
<BR>
<BR>
<BR>
<BR>
<DIV>
<% Sql2 = "select  count(*) as jiewu  from "&jieducm&"_toushu where username='"&rs("username")&"' "
            Set Rs2 = Server.CreateObject("Adodb.RecordSet")
				Rs2.Open Sql2,conn,1,1
				if NOT rs2.EOF  then
				jiewu=rs2("jiewu")
				end if
				
				  %> 
<P>投诉他人次数：</P><input name="jiewu" type="text" id="jiewu" size="30"  value="<%=jiewu%>" disabled/></DIV>
<BR>
<DIV>
<% Sql2 = "select  count(*) as jiewu1  from "&jieducm&"_toushu where name='"&rs("username")&"' "
            Set Rs2 = Server.CreateObject("Adodb.RecordSet")
				Rs2.Open Sql2,conn,1,1
				if NOT rs2.EOF  then
				jiewu1=rs2("jiewu1")
				end if
				
				  %> 
<P>被投诉次数：</P><input name="jiewu1" type="text" id="jiewu1" size="30"  value="<%=jiewu1%>" disabled/></DIV>
<BR>
<DIV>
<P>设置登陆IP：</P><input name="login_ip" type="text" id="login_ip" size="30"  value="<%=rs("login_ip")%>" />设置为0不检测登陆ip</DIV>


<DIV class=regsubmit><INPUT style="CURSOR: pointer; HEIGHT: 30px"  type=submit value="  确认  " > 
&nbsp;  
<INPUT name="按钮"  type=button style="CURSOR: pointer; HEIGHT: 30px" value="  返回  " onClick="history.back();">
</DIV><BR><BR></DIV></DIV>


</FORM></DIV>
</BODY></HTML>
