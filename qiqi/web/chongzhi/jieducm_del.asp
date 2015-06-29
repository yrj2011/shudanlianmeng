<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
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
rs.open "select price,bei from "&jieducm&"_zhongxin where pid='"&id&"' and classid='5' and  username='"&session("useridname")&"' and start='0' ",conn,1,1
if ( rs.eof ) or (rs.bof) then
  Response.Write("<script>alert('任务状态已发生变化!');window.location.href='MyMission.asp';</script>")
  response.end()
else
price=rs("price")
end if

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "delete  from "&jieducm&"_zhongxin where pid='"&id&"'",conn,3,3


 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select cunkuan From "&jieducm&"_register where username='"&session("useridname")&"'" ,Conn,3,3  
		rs("cunkuan")=rs("cunkuan")+price
    	rs.update
    	rs.close	

     num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=num
		rs("class")="充值提现任务区"
		rs("content")=session("useridname")&"取消任务"&id&"成功,返还"&price&"元存款"
		rs("price")="+"&price
		rs("jiegou")="成功"
    	rs.update
    	rs.close	
			call check_jieducm_name(session("useridname"))  
set rs=nothing
conn.close
set conn=nothing
response.Redirect("MyMission.asp")

%>