<!--#include file="inc/conn.asp"-->

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
id=request("id")
if id="" then
 Response.Write("<script>alert('请先选择要删除的记录!');history.back();</script>")
 response.End()
end if
 Sql = "select pid from "&jieducm&"_zhongxin where id in ("&id&")"
 Set rs = Server.CreateObject("Adodb.RecordSet")
 rs.Open Sql,conn,1,1
 IF not rs.Eof Then
 Do While Not Rs.Eof
     Set rsj=server.createobject("ADODB.RECORDSET")
     rsj.open "delete  from "&jieducm&"_jieshou where pid = '"&rs("pid")&"' ",conn,3,3
  Rs.MoveNext
  Loop
 End IF
 
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "delete  from "&jieducm&"_zhongxin where id in ("&id&")",conn,3,3


set rs=nothing
conn.close
set conn=nothing
response.Redirect("admin_mission.asp")
%>