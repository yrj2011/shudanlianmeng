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
rs.open "select * from "&jieducm&"_jieshou where pid='"&id&"' and start='2' and username='"&session("useridname")&"' ",conn,3,3
if (rs.eof) then
    Response.Write("<script>alert('对不起,此信息状态已发生变化!');window.location.href='remission.asp';</script>")
	response.End()
else
  rs("start")="1"
  rs.update
  rs.close
end if

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "select * from "&jieducm&"_zhongxin where pid='"&id&"'",conn,3,3
if not (rs.eof) then 
  price=rs("price")
  user=rs("username")
  rs("start")="1"
  rs.update
  rs.close
end if
		
	 num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=num
		rs("class")="拍拍区任务"
		rs("content")=session("useridname")&"接手"&user&"的任务"&id&"返回未支付状态"
		rs("price")=0
		rs("jiegou")="返回未支付"
    	rs.update
    	rs.close	
		response.Redirect("remission.asp")
set rs=nothing
conn.close
set conn=nothing
%>