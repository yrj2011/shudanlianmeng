<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="inc/index_conn.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="config.asp"-->
<!--#include file="md5.asp"-->
<%
'------------------------------------------------------------------
'***************************淘宝信誉互刷系统***********************
'本程序专为淘宝免费互刷打造，是目前市面上功能比较齐全的平台
'程序开发：七  七 传 媒
'官方主页：http://www.778892.com
'技术支持：xsh
'QQ;859680966
'电话：15889679112
'版权：版权属于七七传媒信息技术有限公司，不得拷贝、修改，侵权必究。
'Copyright (C) 2013 七七传媒信息技术有限公司 版权所有
'*****************************************************************
'------------------------------------------------------------------
promotion=request.QueryString("promotion")
if promotion<>"" then
Sql = "select username from "&jieducm&"_register where id="&promotion&""
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
IF ( Rs.Eof or Rs.Bof) Then
response.write "<script>alert('无此推荐人！');window.location.href='register.asp';</script>"  
response.End()
else
session("tjrname")=rs("username")
end if				
end if

zhuce=request.Form("zhuce")
if zhuce="ok" then
 session("jieducm_ok")="jieducm_reg"
 call register()    
end if
 %>
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD>
<TITLE><%=stupname%>-<%=stuptitle%></title>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<LINK rev=Stylesheet href="img/global.css" type=text/css rel=Stylesheet>
<LINK href="css/login.css" type=text/css rel=stylesheet>
<LINK href="css/top_bottom.css" type=text/css rel=stylesheet>
<LINK href="css/layout.css" type=text/css rel=stylesheet>
<LINK rel=stylesheet type=text/css href="css/global.css"> 
<SCRIPT language=JavaScript src="js/jquery.min.js"></SCRIPT>
<SCRIPT src="js/jieducm_pupu.js" type=text/javascript></SCRIPT>
<script language="JavaScript" type="text/JavaScript" src="js/js.js"></script>
<META content="MSHTML 6.00.2900.2180" name=GENERATOR></HEAD>
<BODY  onload=document.form.UserID.focus()> 

<SCRIPT language=javascript>
function xiechuemail(){
document.getElementById("Email").value=document.getElementById("QQ").value+'@qq.com';
}
function  save_onclick()
{
  var strTemp = document.form.UserID.value;
  if (strTemp.length == 0 )
  {
      alert("请输入用户名！");
      document.form.UserID.focus();
      return false;
  }

if(form.UserID.value.replace(" ","").replace( ).length!=form.UserID.value.length)
{ alert('用户名不能有空格');
  document.form.UserID.focus();
      return false;
}
			
   if (strTemp.length < 3 || strTemp.length>16 )
  {
      alert("用户名的长度必须为3到16个字符，一个中文算两个字符！");
      document.form.UserID.focus();
      return false;
  }
  var strTemp = document.form.password.value;
  if (strTemp.length == 0 )
  {
      alert("请输入密码！");
      document.form.password.focus();
      return false;
  }
  
    if (strTemp.length < 6 || strTemp.length>16 )
  {
      alert("密码必须是6至16位数字或字母！");
      document.form.password.focus();
      return false;
  }
  
  var strTemp1 =document.form.password2.value;
  if (strTemp1.length == 0 )
  {
      alert("请输入重复密码密码！");
      document.form.password2.focus();
      return false;
  }
  
  if (strTemp!=strTemp1)
  {
      alert("两次输入密码不一致！");
      document.form.password.focus();
      return false;
  }
  
  var qq=document.form.QQ.value
  if(qq.length==0)
  {
     alert("QQ不能为空!");
	 document.form.QQ.focus();
	 return false;
	 }
	   var Email=document.form.Email.value
  if(Email.length==0)
  {
     alert("Email不能为空!");
	 document.form.Email.focus();
	 return false;
	 }
	 
	  if(document.form.Email.value.length!=0)
    if (document.form.Email.value.charAt(0)=="." ||        
         document.form.Email.value.charAt(0)=="@"||       
         document.form.Email.value.indexOf('@', 0) == -1 || 
         document.form.Email.value.indexOf('.', 0) == -1 || 
         document.form.Email.value.lastIndexOf("@")==document.form.Email.value.length-1 || 
         document.form.Email.value.lastIndexOf(".")==document.form.Email.value.length-1)
     {
      alert("Email地址格式不正确！");
      document.form.Email.focus();
      return false;
      }
	  
    var czm =document.form.czm.value;
  if (czm.length == 0)
  {
      alert("操作码不能为空！");
      document.form.czm.focus();
      return false;
  }
if (czm.length < 6 || czm.length>16 )
  {
      alert("操作码必须是6至16位数字或字母！");
      document.form.czm.focus();
      return false;
  }
   var czm2 =document.form.czm2.value;
  if (czm2.length == 0)
  {
      alert("确认操作码操作码不能为空！");
      document.form.czm2.focus();
      return false;
  }
  if (czm!=czm2)
  {
      alert("两次操作码输入不一致！");
      document.form.czm.focus();
      return false;
  }
  
  if (strTemp==czm)
  {
      alert("操作码不允许和登录密码相同！");
      document.form.czm.focus();
      return false;
  }
  var phone =document.form.phone.value;
  if (phone.length == 0)
  {
      alert("手机不能为空！");
      document.form.phone.focus();
      return false;
  }
  if (phone.length !=11 )
  {
      alert("请填写正确的手机号码！");
      document.form.phone.focus();
      return false;
  }
  
<%if stupCheckCode="2" then%>
if (document.form.CheckCode)
	{
		if (document.form.CheckCode.value==""){
		   alert ("请输入您的验证码！");
		   document.form.CheckCode.focus();
		   return(false);
		}
	}
<%end if%>
  return true;  
}

	function SwitchDivDisplay(divName) {
		var obj_reg_info = document.getElementById(divName);
		//alert(obj_reg_info);
	   	if(obj_reg_info.style.display=="none") {
	       	obj_reg_info.style.display="inline";
	   	} else {
	   		obj_reg_info.style.display="block";
	   		obj_reg_info.style.display="none";
		}
	}

</SCRIPT>
<!--#include file=jieducm_top.asp-->
<TABLE cellSpacing=0 cellPadding=0 width=960 align=center border=0>
  <TBODY>
  <TR>
    <TD width=20 height=32><IMG src="images/Top_10.gif"></TD>
    <TD align=right width=21 background=images/Top_11.gif><IMG height=32 
      src="images/Top_9.gif" width=10></TD>
    <TD class=K_mttitle align=middle width=120 
      background=images/Top_12.gif>会员注册</TD>
    <TD align=left width=47 background=images/Top_11.gif><IMG height=32 
      src="images/Top_13.gif" width=10></TD>
    <TD align=left width=728 background=images/Top_11.gif></TD>
    <TD align=right width=24 background=images/Top_11.gif><IMG height=32 
      src="images/Top_14.gif" width=6></TD></TR>
  <TR>
    <TD class=K_mtcontent align=left colSpan=6> 
	   <%if  stupzhu="2" then%>

<DIV class=register>
<H2>会员注册<SPAN class=current>1.填写会员信息</SPAN><SPAN>2.通过邮件确认</SPAN><SPAN>3.注册成功</SPAN></H2>
<FORM id=form name=form onSubmit="return save_onclick()" method=post>
<INPUT type=hidden value=ok name=zhuce> 
<UL><li style="line-height:22px;"><span class="font14b2">新用户免费注册-轻松注册10秒钟搞定</span><br />
          请填写以下信息，<span class="red-bcolor">*</span>为必填内容。<br />
          请认真仔细的填写以下信息，真实的个人信息有助于给你使用的服务带来更多的保障以及便捷！&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		  <a href="/news.asp?/1465.html" target="_blank"><strong>提示 非法字符？猛击这里</strong></a></li> 
  <LI onmouseover='this.setAttribute("class","current")' 
  onmouseout='this.setAttribute("class","")'><LABEL for=username>会员名：</LABEL> 
  <DIV class=form-entry><INPUT class=txt id=UserID onblur=name_chk(this) 
  maxLength=12 name=UserID>
  <SCRIPT id=tname></SCRIPT>
  
  <DIV class=form-state id=c0><FONT color=#ff0000>*</FONT> 4-12个字符</DIV></DIV></LI>
  <LI onmouseover='this.setAttribute("class","current")' 
  onmouseout='this.setAttribute("class","")'><LABEL for=password>登录密码：</LABEL> 
  <DIV class=form-entry><INPUT class=txt id=password type=password 
  name=password> 
  <DIV class=form-state><FONT color=#ff0000>*</FONT> 登录时需要使用密码</DIV></DIV></LI>
  <LI onmouseover='this.setAttribute("class","current")' 
  onmouseout='this.setAttribute("class","")'><LABEL 
  for=verify-password>确认密码：</LABEL> 
  <DIV class=form-entry><INPUT class=txt id=password2 type=password 
  name=password2> 
  <DIV class=form-state><FONT color=#ff0000>*</FONT> 重复上面的密码</DIV></DIV></LI>
  <LI onmouseover='this.setAttribute("class","current")' 
  onmouseout='this.setAttribute("class","")'><LABEL for=password2>操作密码：</LABEL> 
  <DIV class=form-entry><INPUT class=txt id=czm name=czm type=password > 
  <DIV class=form-state><FONT color=#ff0000>*</FONT> 发任务，修改密码等操作需要用到</DIV></DIV></LI>
  <LI onmouseover='this.setAttribute("class","current")' 
  onmouseout='this.setAttribute("class","")'><LABEL for=password2>确认操作码：</LABEL> 

  <DIV class=form-entry><INPUT class=txt id=czm2 name=czm2 type=password > 
  <DIV class=form-state><FONT color=#ff0000>*</FONT> 重复上面的操作码</DIV></DIV></LI>
  <LI onmouseover='this.setAttribute("class","current")' 
  onmouseout='this.setAttribute("class","")'><LABEL for=qq>QQ号码：</LABEL> 
  <DIV class=form-entry>
  <input   class=txt name="QQ" type="text" id="QQ" size="30" onKeyUp="if(isNaN(value))execCommand('undo')"  onBlur="xiechuemail()"/>
  <DIV class=form-state><FONT color=#ff0000>*</FONT> QQ号码不能为空，双方联系必备！</DIV></DIV></LI>
  <LI onmouseover='this.setAttribute("class","current")' 
  onmouseout='this.setAttribute("class","")'><LABEL for=email>电子邮箱：</LABEL> 
  <DIV class=form-entry><input class=txt name="Email" type="text" id="Email" size="30" />
  <DIV class=form-state><FONT color=#ff0000>*</FONT> 激活账号和取回密码时需要用到，请检查是否正确！</DIV></DIV></LI>
  <LI onmouseover='this.setAttribute("class","current")' 
  onmouseout='this.setAttribute("class","")'><LABEL for=tjr>推荐人：</LABEL> 
  <DIV class=form-entry><input class=txt name="tjr" type="text" id="tjr" value="<%=session("tjrname")%>" size="30"  <%if session("tjrname")<>"" then%> readonly <%end if%>  />
  <DIV 
  class=form-state>没有可留空！对于被推荐人，没有任务损失，推荐人得到积分奖励</DIV></DIV></LI>
  <LI onmouseover='this.setAttribute("class","current")' onmouseout='this.setAttribute("class","")'><LABEL for=phone>手机号码：</LABEL> 
  <DIV class=form-entry><INPUT class=txt id=phone  onkeyup="if(isNaN(value))execCommand('undo')" name=phone > 
  <DIV class=form-state>可凭手机找回密码</DIV></DIV></LI>
  
  <LI class=random-code><LABEL for=username>验证码：</LABEL> 
  <DIV class=form-entry><INPUT class=txt maxLength=4 id=CheckCode onKeyUp="if(isNaN(value))execCommand('undo')"  name=CheckCode> 
  <img src="jieducm_code.asp" alt="验证码" onClick="this.src='jieducm_code.asp?rnd=' + Math.random();" /> 
  <SPAN>看不清？点击图片换一张</SPAN> </DIV></LI>
  
  <LI class=random-code><LABEL for=checkbox>填写更多：</LABEL> 
  <DIV class=form-entry><INPUT id=checkbox onClick="javascript:SwitchDivDisplay('showinfo_c1')" type=checkbox name=checkbox> 补充更多信息（选填） </DIV></LI></UL>
<UL id=showinfo_c1 style="DISPLAY: none">
  <LI onmouseover='this.setAttribute("class","current")' 
  onmouseout='this.setAttribute("class","")'><LABEL for=xt><SPAN 
  class=font14b2>选填项目：</SPAN></LABEL> 
  <DIV class=form-entry>
  <HR style="COLOR: #ff9933">
  </DIV></LI>
  <LI onmouseover='this.setAttribute("class","current")' 
  onmouseout='this.setAttribute("class","")'><LABEL for=weiti>密码问题：</LABEL> 
  <DIV class=form-entry><INPUT class=txt id=weiti name=weiti>
  <DIV class=form-state>忘记密码的提示问题</DIV></DIV></LI>
  <LI onmouseover='this.setAttribute("class","current")' 
  onmouseout='this.setAttribute("class","")'><LABEL for=daai>问题答案：</LABEL> 
  <DIV class=form-entry><INPUT class=txt id=daai name=daai> 
  <DIV class=form-state>忘记密码的问题答案</DIV></DIV></LI>
  <LI onmouseover='this.setAttribute("class","current")' 
  onmouseout='this.setAttribute("class","")'><LABEL for=rname>真实姓名：</LABEL> 
  <DIV class=form-entry><INPUT class=txt id=rname name=rname> 
  <DIV class=form-state><FONT color=#ff0000>*</FONT> 
  真实姓名</DIV></DIV></LI>
  <LI onmouseover='this.setAttribute("class","current")' 
  onmouseout='this.setAttribute("class","")'><LABEL for=HomePage>店铺地址：</LABEL> 
  <DIV class=form-entry><INPUT class=txt id=HomePage value=http:// 
  name=HomePage> 
  <DIV class=form-state><FONT color=#ff0000>*</FONT> 
  店铺地址</DIV></DIV></LI>
  <LI onmouseover='this.setAttribute("class","current")' 
  onmouseout='this.setAttribute("class","")'><LABEL for=shopnote>店铺描述：</LABEL> 
  <DIV class=form-entry><INPUT class=txt id=shopnote name=shopnote>
  <DIV class=form-state><FONT color=#ff0000>*</FONT> 
  店铺描述</DIV></DIV></LI>
  <LI onmouseover='this.setAttribute("class","current")' 
  onmouseout='this.setAttribute("class","")'><LABEL for=sex>性别：</LABEL> 
  <DIV class=form-entry><INPUT id=radio type=radio CHECKED value=1 name=sex> 男士 
  <INPUT id=radio2 type=radio value=2 name=sex> 女士 </DIV></LI>
  <LI onmouseover='this.setAttribute("class","current")' 
  onmouseout='this.setAttribute("class","")'><LABEL for=address>联系地址：</LABEL> 
  <DIV class=form-entry><INPUT class=txt id=address name=address>
  <DIV class=form-state><FONT color=#ff0000>*</FONT> 
  联系地址</DIV></DIV></LI></UL>
  <div align="center"><a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=<%=stupqq%>&site=qq&menu=yes">联系QQ：<%=stupqq%>  验证 开通号码</a></div>
<UL>
  <LI>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  <INPUT class=submit style="MARGIN-LEFT: 20px; CURSOR: pointer" type="submit" value=确认注册>
  <INPUT class=submit style="MARGIN-LEFT: 20px; CURSOR: pointer" onClick="window.showModelessDialog('images/jieducm_zhidu.htm',window,'scroll:yes;resizable:yes;help:no;dialogWidth:650px;dialogHeight:450px')" type=button value=注册协议>&nbsp;&nbsp;
  </LI>
</UL></FORM></DIV>
	<%elseif  stupzhu="1" then%>
    <table width="900" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td height="100" colspan="3"><div align="center"><span class="font14b2"> 系统暂停注册</span><br />
        </div></td>
        </tr>
        </table>
   
    <%end if%>
	</TD></TR></TBODY></TABLE><br>
<!--#include file=jieducm_bottom.asp-->
<%
rs.close
set rs=nothing
conn.close
set conn=nothing%>
</BODY></HTML>
