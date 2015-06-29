<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#INCLUDE FILE="../config.asp"-->
<!--#INCLUDE FILE="../inc/function.asp"-->
<!--#include file="checklogin.asp"-->
<!--#include file="../md5.asp"-->
<%

'------------------------------------------------------------------
'***************************淘宝信誉互刷系统**************************************
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
if request.Form<>"" then
	 password=HtmlEncode(trim(request.form("password")))
	 mail=HtmlEncode(trim(request.form("email")))
     HomePage=HtmlEncode(trim(request.form("HomePage")))
	 sex=HtmlEncode(trim(request.form("sex")))
	 phone=HtmlEncode(trim(request.form("phone")))
	 rname=HtmlEncode(trim(request.form("rname")))
	 shopnote=HtmlEncode(trim(request.form("shopnote")))
	 address=HtmlEncode(trim(request.form("address")))
	    Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_register where username='" &session("useridname")& "'" ,Conn,3,3  
    	rs("mail")=mail	
		rs("homepage")=homepage
		rs("sex")=sex
		rs("phone")=phone
		rs("rname")=rname
		rs("shopnote")=shopnote
		rs("address")=address
		rs("weiti")=request.form("weiti")
		rs("daai")=request.form("daai")
    	rs.update
    	rs.close
		Response.Write("<script>alert('修改成功!');window.location.href='guest_info.asp';</script>")
		response.End()
set rs=nothing
conn.close
set conn=nothing
end if
%>	

<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD>
<title><%=stupname%> -<%=stuptitle%></title>
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<link href="../css/style.css" rel="stylesheet" type="text/css" />
<LINK href="../css/center.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2900.2180" name=GENERATOR>
</HEAD>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">   
<!--#include file=../inc/jieducm_top.asp-->
<SCRIPT language=javascript>

function  save_onclick()
{

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
  
  return true;  
}

</SCRIPT>
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
                    <div class=pp7>您现在的位置是：<%=stupname%> &gt;&gt; 会员中心 &gt;&gt; 修改资料 &gt;&gt; </div>
                    <div class=pp8><strong>修改资料</strong></div>

 <% 
 Set rs=server.createobject("ADODB.RECORDSET")
 rs.open "Select * From "&jieducm&"_register where username='"&session("useridname")&"'",Conn,3,3
%>
					<FORM name="form" id="form" onSubmit="return save_onclick()" method=post>
<table width="99%"  border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td height="10" colspan="3">&nbsp;</td>
        </tr>
      <tr>
        <td width="146" height="40" align="right" class="font12h">用户名：</td>
        <td width="230"><input name="UserID" type="text" id="UserID"  size="30" maxlength="12" value="<%=session("useridname")%>"  disabled="disabled" /></td>
        <td width="524"></td>
      </tr>
      <tr>
        <td height="40" align="right" class="font12h">Q&nbsp;Q:</td>
        <td><input name="QQ" type="text" id="QQ" size="30" value="<%=rs("qq")%>" disabled="disabled" /></td>
        <td><span class="red-bcolor">*</span> QQ不能修改！</td>
      </tr>
      <tr>
        <td height="40" align="right" class="font12h">电子邮箱：</td>
        <td><input name="Email" type="text" id="Email" size="30" value="<%=rs("mail")%>" /></td>
        <td><span class="red-bcolor">*</span>取回密码使用，注册后不能更改</td>
      </tr>
      
      <tr>
        <td height="40" align="right" class="font12h">密码问题：</td>
        <td><span class="font12h">
          <input name="weiti" type="text" id="weiti" size="30" value="<%=rs("weiti")%>" />
        </span></td>
        <td>忘记密码的提示问题</td>
      </tr>
      <tr>
        <td height="40" align="right" class="font12h">问题答案：</td>
        <td><span class="font12h">
          <input name="daai" type="text" id="daai" size="30"  value="<%=rs("daai")%>"/>
        </span></td>
        <td>忘记密码的问题答案</td>
      </tr>
      
  
      <tr>
        <td height="40" align="right" class="font12h">手机号码：</td>
        <td align="left" class="red-bcolor"><input name="phone" type="text" id="phone" size="30" value="<%=rs("phone")%>" /></td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="40" colspan="3" align="left" class="font12h"><table width="100%" border="0" align="left" cellpadding="0" cellspacing="0"  id=showinfo_c1 >
             <tr>
        <td width="16%" height="40" align="right" class="font14b2">选填项目：</td>
        <td colspan="2"><hr  style="color:#FF9933"/></td>
        </tr>
      
      <tr>
        <td height="40" align="right" class="font12h">真实姓名：</td>
        <td width="80%"><input name="rname" type="text" id="rname" size="30"  value="<%=rs("rname")%>"/></td>
        <td width="4%">&nbsp;</td>
      </tr>
      <tr>
        <td height="40" align="right" class="font12h">店铺地址：</td>
        <td><input name="HomePage" type="text" id="HomePage"  size="30"  value="<%=rs("homepage")%>" /></td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="40" align="right" class="font12h">店铺描述：</td>
        <td><input name="shopnote" type="text" id="shopnote" size="30"  value="<%=rs("shopnote")%>" / ></td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="30" align="right"><span class="font12h">性别：</span></td>
        <td><input name="sex" type="radio" id="radio" value="1" checked="checked"  <%if rs("sex")="1" then%>checked="checked" <%end if%> />
          男士
          <input type="radio" name="sex" id="radio2" value="2"  <%if rs("sex")="2" then%>checked="checked" <%end if%>/>
          女士</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="45" align="right" class="font12h">联系地址：</td>
        <td><input name="address" type="text" id="address" size="30" value="<%=rs("address")%>" /></td>
        <td>&nbsp;</td>
      </tr>
      
        </table></td>
        </tr>
      <tr>
        <td height="40" align="right" class="font12h">&nbsp;</td>
        <td align="center" class="red-bcolor"><input type="submit" name="button" id="button" value="修  改" /></td>
        <td>&nbsp;</td>
      </tr>
    </table>
					</FORM>
                  </div>
                </div></td>
	          </tr>
	        </table>	    </td>
	  </tr>
</table>

<!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
<%
rs.close
set rs=nothing
call closeconn()%>
</BODY></HTML>
