
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
id=request("id")
action=request.form("action")
uid=request.Form("uid")
czm=request.Form("czm")
if action="ok" then
if md5(czm)<>admin_czmpass then

   Set rs=server.createobject("ADODB.RECORDSET")
   rs.open "Select * From "&jieducm&"_recordm " ,Conn,3,3  
   rs.addnew
   rs("username")=session("AdminName")
   rs("class")="删除会员"
   rs("content")=session("AdminName")&"管理员删除会员操作码输入错误"
   rs("jiegou")="失败"
   rs.update
   rs.close
   Response.Write("<script>alert('操作码有误!请不要重复操作，平台会记录你的行为！');history.back();</script>")
   response.End()
end if

Sql = "select username from "&jieducm&"_register where id="&uid&""
Set Rsm = Server.CreateObject("Adodb.RecordSet")
Rsm.Open Sql,conn,1,1
IF not Rsm.Eof or  not rsm.bof Then
	username=rsm("username")
end if
				
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "delete  from "&jieducm&"_register where id in("&id&")",conn,3,3

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "delete  from "&jieducm&"_xinyu where username='"&username&"'",conn,3,3

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "delete  from "&jieducm&"_zhongxin where username='"&username&"'",conn,3,3

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "delete  from "&jieducm&"_jieshou where username='"&username&"'",conn,3,3

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "delete  from "&jieducm&"_hei where name='"&username&"'",conn,3,3

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "delete  from "&jieducm&"_record where username='"&username&"'",conn,3,3

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "delete  from "&jieducm&"_keeper where username='"&username&"'",conn,3,3

   Set rs=server.createobject("ADODB.RECORDSET")
   rs.open "Select * From "&jieducm&"_recordm " ,Conn,3,3  
   rs.addnew
   rs("username")=session("AdminName")
   rs("class")="删除会员"
   rs("content")=session("AdminName")&"管理员删除会员"&username&"成功"
   rs("jiegou")="成功"
   rs.update
   rs.close
   
set rs=nothing
conn.close
set conn=nothing
Response.Write("<script>alert('删除会员成功！');window.location.href='usermanage.asp';</script>")
response.End()
end if
%><form id="form1" name="form1" method="post" action="">
<input name="action" type="hidden" value="ok" />
<input name="uid" type="hidden" value="<%=id%>" />

  <label>
  操作码：
  <input type="password" name="czm" />
  
  </label>
  <label>
  <input type="submit" name="Submit" value="确认删除" />
  </label>
</form>