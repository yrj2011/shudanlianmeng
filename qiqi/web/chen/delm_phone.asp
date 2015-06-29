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

dim theid,thdel
id=request("adid")
action=request("action")
if action="del" then

sql="delete from "&jieducm&"_sms where id in("&id&")"
conn.execute(sql) 

	   Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_recordm " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("AdminName")
		rs("class")="删除手机短信"
		rs("content")="管理员删除已发手机短信"
		rs("jiegou")="成功"
    	rs.update
    	rs.close

conn.close
set conn=nothing
end if                            
response.redirect request.ServerVariables("HTTP_REFERER")
%>