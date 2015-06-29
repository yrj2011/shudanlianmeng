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
'Copyright (C) 2008－2009 捷度传媒信息技术有限公司 版权所有
'*****************************************************************
'------------------------------------------------------------------
id=request.QueryString("id")

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "select * from "&jieducm&"_jieshou where id='"&id&"'  and username='"&session("useridname")&"' and classid='6' and start='1'",conn,1,1
if rs.eof then
    Response.Write("<script>alert('对不起,操作有误无此信息!');window.location.href='remission.asp';</script>")
	response.End()
else
pid=rs("pid")
end if 

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "delete  from "&jieducm&"_jieshou where id='"&id&"' and username='"&session("useridname")&"' and classid='6'",conn,3,3

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "select * from "&jieducm&"_zhongxin where pid='"&pid&"' and classid='6'",conn,3,3
if not (rs.eof) then
  rs("jieshou_num")=rs("jieshou_num")-1
  rs.update
  rs.close
else
    Response.Write("<script>alert('对不起,操作有误无此信息!');window.location.href='remission.asp';</script>")
	response.End()
end if

	 num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=num
		rs("class")="淘宝收藏区"
		rs("content")=session("useridname")&"任务ID:"&pid&"退出任务"
		rs("price")=0
		rs("jiegou")="退出任务"
    	rs.update
    	rs.close	
set rs=nothing
conn.close
set conn=nothing  
response.Redirect("remission.asp")

%>