<!--#include file="../inc/index_conn.asp"-->
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
id=request("id")

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "select * from "&jieducm&"_tixian where username='"&session("useridname")&"' and id="&id&" and start='0'",conn,3,3
if  (rs.eof) then
  Response.Write("<script>alert('本交易已结束!请不要重复操作!平台会记录你的行为!');window.history.back();</script>")
  response.end()
else
  price=rs("ReNum")
  rs("start")=2
  rs.update
  rs.close
end if

 			  	Sql = "select cunkuan from "&jieducm&"_register where username='"&session("useridname")&"'"
				Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,3,3
				if not(rs.eof) then
				  rs("cunkuan")=rs("cunkuan")+price
				  rs.update
				  rs.close
				end if
								
	 num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=num
		rs("class")="撤销提现"
		rs("content")=session("useridname")&"进行撤销提现 "&price&"元"
		rs("price")="+"&price
		rs("jiegou")="撤销成功"
    	rs.update
    	rs.close
				
 Response.Write("<script>alert('您已成功撤销本次提现,\您的金额已返还到你的存款中!');window.location.href='ReMoneyList.asp';</script>")	
set rs=nothing
call CloseConn()
%>