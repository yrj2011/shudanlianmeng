<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#INCLUDE FILE="../inc/function.asp"-->
<!--#INCLUDE FILE="../config.asp"-->
<!--#include file="checklogin.asp"-->

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
sendname=request("sendname")
action=request.form("action")
uid=(trim(request.form("uid")))
classid=request.Form("class")

if action="ok" and classid="1" then

if ismessage=1 then
 Response.Write("<script>alert('你已被管理员禁止了发站内信息的功能!');history.back();</script>")
 response.End()
elseif uid=session("useridname") then
 Response.Write("<script>alert('不能对自己发送!');history.back();</script>")
 response.End()
end if
				

Sql = "select * from "&jieducm&"_register where username='"&uid&"'"
Set Rs1 = Server.CreateObject("Adodb.RecordSet")
Rs1.Open Sql,conn,1,1
if rs1.eof then
  Response.Write("<script>alert('无此用户!');history.back();</script>")
  response.End()
else              
messagelxfs=(trim(request.form("messagelxfs")))
if messagelxfs<>"" then
messagetext=(trim(request.form("messagetext")))
messagename=(trim(request.form("messagename")))
ip = Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
If ip = "" Then 
ip = Request.ServerVariables("REMOTE_ADDR") 
end if
sql="select * from "&jieducm&"_china_message"
set rs2=server.createobject("adodb.recordset")
rs2.open sql,conn,1,3
rs2.addnew
rs2("messagename")=messagename'发送人
rs2("messagelxfs")=messagelxfs'标题
rs2("messagetext")=messagetext'内容
rs2("uid")=uid
rs2("ip")=ip
rs2("zn")="yes"
rs2("messagedate")=now
rs2("hits")=0
rs2.update
rs2.close  
set rs2=nothing 
end if
end if

sql="select * from "&jieducm&"_china"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
rs.addnew
rs("username")=messagename'发送人
rs("title")=messagelxfs'标题
rs("content")=messagetext'内容
rs("uid")=uid
rs("ip")=ip
rs("now")=now()
rs.update
rs.close

elseif action="ok" and classid="2" then
messagelxfs=(trim(request.form("messagelxfs")))
messagetext=(trim(request.form("messagetext")))

Sql = "select * from "&jieducm&"_register where username='"&uid&"'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if rs.eof then
  Response.Write("<script>alert('无此用户!');history.back();</script>")
  response.End()
else
  email=rs("mail")
  call send_email(email,messagelxfs,messagetext)
end if

response.Write "<script language='javascript'>alert('恭喜您发送成功！');window.location.href='index.asp';</script>"
response.End()
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
<SCRIPT LANGUAGE="JavaScript">
<!--//
function checkmessage()
{   
    if (document.form.uid.value.length<1)
	{
        alert("请填写收信人！");
        document.form.uid.focus();
        return false;
    }
	
	    if (document.form.messagelxfs.value.length<1)
	{
        alert("请填写标题！");
        document.form.messagelxfs.focus();
        return false;
    }
	
    if (document.form.messagetext.value.length<1)
	{
        alert("请填写内容！");
        document.form.messagetext.focus();
        return false;
    }
	    if (document.form.messagetext.value.length>1000)
	{
        alert("请把字数控制在1000以内！");
        document.form.messagetext.focus();

        return false;
    }
}
//-->
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
                    <div class=pp7>您现在的位置是：<%=stupname%> &gt;&gt; 会员中心 &gt;&gt; 发送站内消息 &gt;&gt; </div>
                    <div class=pp8><strong>发送站内消息</strong></div><br>
                   <form name="form" method="post" action="" onSubmit="return checkmessage()">
				     <input name="action" type="hidden" value="ok">
      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#799AE1">
  <tr>
    <td height="20" bgcolor="#799AE1" align="center">
      <table width="98%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td align="center"><font color="#FFFFFF" style="font-size:14px">发 送 站 内 消 息</font></td>
          <td width="35">　</td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"> <br>
        <table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#D6DFF7">
          <tr bgcolor="#FFFFFF">
            <td width="27%" height="26" align="right">发送人：</td>
            <td><input name="messagename" type="text" id="messagename" size="20" maxlength="20" value="<%=session("useridname")%>" readonly></td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td height="26" align="right">发送方式：</td>
            <td><label>
              <input name="class" type="radio" value="1" checked>
              站内信
              <input type="radio" name="class" value="2">邮件
            </label></td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td height="26" align="right">接收人：</td>
            <td>
            <input name="uid" type="text" id="uid" size="20" maxlength="20"  value="<%=sendname%>">          </td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td height="26" align="right">标&nbsp; 题：</td>
            <td><input name="messagelxfs" type="text" id="messagelxfs2" size="40" maxlength="50"></td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td align="right" valign="top">内&nbsp; 容：</td>
            <td><textarea name="messagetext" cols="39" rows="5"></textarea>
              <input name="id" type="hidden" value="<%=request("id")%>"></td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td height="30">　</td>
            <td><input type="submit" name="Submit" value="发送消息">
&nbsp;
<input type="button" value="取消发送" name="reset" onClick="window.location.href='jieducm_car.asp'"></td>
          </tr>
        </table>
        <br>
    </td>
  </tr>
</table></form>
                  </div>
                </div></td>
	          </tr>
	        </table>	    </td>
	  </tr>
</table>

<!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
<%call closeconn()%>
</BODY></HTML>
