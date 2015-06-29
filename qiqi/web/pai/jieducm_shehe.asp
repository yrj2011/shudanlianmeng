<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../taobao/checklogin.asp"-->
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
id=request.QueryString("id")
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "select * from "&jieducm&"_jieshou where pid='"&id&"' and start='1' ",conn,3,3
if not (rs.eof) then
    rs("baohu2")="2"
    rs.update	
	rs.close
    response.Redirect("MyMission.asp")
	response.End()
else
   Response.Write("<script>alert('操作有误，此信息状态已发现变化！');window.location.href='MyMission.asp';</script>")
   response.End()
end if

set rs=nothing
conn.close
set conn=nothing
%>