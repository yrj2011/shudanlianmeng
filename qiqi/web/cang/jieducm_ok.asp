<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../taobao/checklogin.asp"-->
<%
'------------------------------------------------------------------
'***************************淘宝信誉互刷系统***********************
'本程序专为淘宝免费互刷打造，是目前市面上功能比较齐全的平台
'程序开发：七  七 传 媒
'官方主页：http://www.778892.com
'技术支持：xsh
'QQ;859680966
'电话：15889679112
'版权：版权属于七七传媒信息技术有限公司，不得拷贝、修改，侵权必究。
'Copyright (C) 2013 七七传媒信息技术有限公司 版权所有
'*****************************************************************
'------------------------------------------------------------------
id=request.QueryString("id")
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "select num from "&jieducm&"_jieshou where pid='"&id&"' and start='1' and classid='6' and username='"&session("useridname")&"'",conn,1,1
if  (rs.eof) or  (rs.bof) then
  Response.Write("<script>alert('对不起,操作有误!此信息状态已发生改变！');window.history.back();</script>")
  response.end()
else
  num=rs("num")
end if

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select * From "&jieducm&"_zhongxin where pid='"&id&"' and classid='6'" ,Conn,1,1  
if not (rs.eof) then
jusername=rs("username")
bei=rs("bei")
Shopkeeper=rs("Shopkeeper")
end if

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select prourl From "&jieducm&"_xinyu where shopname='"&num&"' and class=4" ,Conn,1,1 
if not (rs.eof) then
jieducm_str=rs("prourl")
end if


jieducm_cangid= mid(jieducm_str,54,15)

if bei="宝贝收藏任务" then
jieducm_idnum=1
else
jieducm_idnum=2
end if
jieducm_ProUrl= "http://favorite.taobao.com/collect_list.htm?itemtype="&jieducm_idnum&"&nuid="&jieducm_cangid&"&pagesize=20"

url=jieducm_ProUrl
wstr=getHTTPPage(url) 
Set objRegExp=New RegExp 
objRegExp.IgnoreCase=true
objRegExp.Global=true
objRegExp.Pattern="<span class=""J_WangWang"" data-nick=""(.*?)"">"
set Matches=objRegExp.Execute(wstr)
jieducm_cang=Matches(0).SubMatches(0)
if jieducm_cang<>Shopkeeper then
   Response.Write("<script>alert('系统检测不到您已收藏!');history.back();</script>")
   response.End()
end if
   
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select jifei,fabudian From "&jieducm&"_register where username='"&session("useridname")&"'" ,Conn,3,3 
rs("jifei")=rs("jifei")+ stupfajifei
rs("fabudian")=rs("fabudian")+(stupcang*0.8)
rs.update
rs.close

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "select start from "&jieducm&"_jieshou where pid='"&id&"' and start='1' and classid='6' and username='"&session("useridname")&"'",conn,3,3
if  not rs.eof then
  rs("start")=2
  rs.update
  rs.close
end if

	    num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	    Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=num
		rs("class")="淘宝收藏区"
		rs("content")=session("useridname")&"确认收藏任务ID:"&id&"你得到"&stupcang*0.8&"个发布点,你也得到"&stupfajifei&"个积分"
		rs("price")=0
		rs("jiegou")="确认已收藏"
    	rs.update
    	rs.close
			 
 	    Set rs=server.createobject("ADODB.RECORDSET")
	    rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=jusername
    	rs("num")=num
		rs("class")="淘宝收藏区"
		rs("content")=session("useridname")&"确认收藏了你的任务ID:"&id&""
		rs("price")=0
		rs("jiegou")="确认已收藏"
    	rs.update
    	rs.close
set rs=nothing
conn.close
set conn=nothing
response.Redirect("ReMission.asp")
%>