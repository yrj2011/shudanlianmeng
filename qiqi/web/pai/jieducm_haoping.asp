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
rs.open "select * from "&jieducm&"_jieshou where pid='"&id&"' and start='3' and username='"&session("useridname")&"'",conn,3,3
if (rs.eof) or  (rs.bof) then
    Response.Write("<script>alert('对不起,操作有误!');window.location.href='remission.asp';</script>")
	response.End()
else

nowj=rs("now")
bei=rs("bei")
if bei="马上收货好评"  then
j=0
elseif bei="一天后收货好评" then
j=24
elseif bei="二天后收货好评"  then
j=48
elseif bei="三天后收货好评" then
j=72
end if

sss= DateDiff( "h", nowj, Now())
if cint(j)-cint(sss)>=1 then
  Response.Write("<script>alert('对不起，此任务收货好评时间还未到!');window.location.href='remission.asp';</script>")
  response.End()
end if
  rs("start")="4"
  rs("now")=now
  rs.update
  rs.close
end if

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "select start,username from "&jieducm&"_zhongxin where pid='"&id&"'",conn,3,3
if not (rs.eof) then
  user=rs("username")
  rs("start")="4"
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
		rs("content")=session("useridname")&"接手"&user&"的任务ID:"&id&"确定已好评"
		rs("price")=0
		rs("jiegou")="确定已好评"
    	rs.update
    	rs.close	
		response.Redirect("remission.asp")     
set rs=nothing
conn.close
set conn=nothing
%>