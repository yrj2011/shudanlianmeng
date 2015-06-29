 <!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="../md5.asp"-->
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
jieducm_ka=request("jieducm_ka")
if session("AdminName")="" then
   response.redirect("admin_login.asp")
end if
if request.form<>"" then
key=request("key")
czm=request("czm")
		
if md5(czm)<>admin_czmpass then
     Set rsr=server.createobject("ADODB.RECORDSET")
	  rsr.open "Select * From "&jieducm&"_recordm " ,Conn,3,3  
	    rsr.addnew
		rsr("username")=session("AdminName")  
		rsr("class")="充值卡密码查询"
		rsr("content")=session("AdminName")&"查询卡号为"&key&"的密码"		
		rsr("jiegou")="失败"
    	rsr.update
    	rsr.close
 Response.Write("<script>alert('操作码有误!请不要重复操作!平台会记录你的行为!');window.history.back();</script>")
  response.end()
end if

Set rsr=server.createobject("ADODB.RECORDSET")
rsr.open "Select * From "&jieducm&"_beifei where ka='"&key&"'" ,Conn,1,1
if not rsr.eof then
kahao=rsr("ka")
pwd=rsr("pwd")
else
 Response.Write("<script>alert('无此卡号!');window.history.back();</script>")
  response.end()
end if

     Set rsr=server.createobject("ADODB.RECORDSET")
	  rsr.open "Select * From "&jieducm&"_recordm " ,Conn,3,3  
	    rsr.addnew
		rsr("username")=session("AdminName")  
		rsr("class")="充值卡密码查询"
		rsr("content")=session("AdminName")&"查询卡号为"&key&"的密码"		
		rsr("jiegou")="成功"
    	rsr.update
    	rsr.close
end if
%>
<html>
<head>
<title>充值卡管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="Inc/Admin.css" type="text/css">

</head>
<body>
<table align="center" width="100%" border="1" cellspacing="0" cellpadding="4" class=dj586_Com_bk style="border-collapse: collapse">
<tr class=dj586_Com_ss> 
<td colspan="6"><div class='bodytitle'><div class='bodytitleleft'></div><div class='bodytitletxt'>充值卡管理 </div></div></td>
</tr>
<tr align="left" class=dj586_Com_ds>
<td colspan="6">  管理导航：<a href=Admin_Okjicar.asp>点卡充值记录</a>| <a href="Admin_Jicar.asp">充值卡管理</a> | </td>      
</tr>
</table><br><form name="myform" method="post" action="">
<table width="619" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="73">请输入卡号：</td>
    <td width="164">
      <input type="text" name="key" id="key" value="<%=jieducm_ka%>">    </td>
    <td width="61">操作码：</td>
    <td width="164"><input type="password" name="czm" id="czm"></td>
    <td width="157"><input type="submit" name="button" id="button" value="查询密码" onClick="return CheckForm();"></td>
  </tr>
</table> 
</form>
<table width="400" border="0">
  <tr>
    <td>卡号</td>
    <td>密码</td>
  </tr>
  
  <tr>
    <td><%=kahao%></td>
    <td><%=pwd%></td>
  </tr>
</table>

</body>
</html>
