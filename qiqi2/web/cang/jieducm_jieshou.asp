<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../inc/function.asp"-->
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
num=request("xiaohaoname")
ip=Request.ServerVariables("REMOTE_ADDR")
if isjie=1 then
 Response.Write("<script>alert('你已被管理员禁止了接任务功能!');history.back();</script>")
 response.End()
elseif num="" then 
    Response.Write("<script>alert('接手淘宝号不能为空');history.back();</script>")
    response.End()		 
end if		 

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "select Shopkeeper,prourl,username,num,jieshou_num from "&jieducm&"_zhongxin where pid='"&id&"'",conn,1,1
if rs.eof then
    Response.Write("<script>alert('对不起,此任务已不存在!');history.back();</script>")
	response.End()
else
 username2=rs("username")
 prourlj=rs("prourl")
 jieducm_num=rs("num")
 jieshou_num=rs("jieshou_num")
end if

if jieducm_num=jieshou_num then
    Response.Write("<script>alert('对不起,此任务已结束!');history.back();</script>")
	response.End()
elseif username2=session("useridname") then
    Response.Write("<script>alert('不能接自己的任务!');history.back();</script>")
	response.End()
end if 

call jname(username2)

sql="SELECT  count(*) as totalnum FROM  "&jieducm&"_jieshou where username='"&session("useridname")&"' and num='"&num&"' and proUrl='"&ProUrlj&"' and classid='6'"
Set rs = Server.CreateObject("Adodb.RecordSet")
rs.Open Sql,conn,1,1
if not rs.eof then	
jh=rs("totalnum")
end if
if jh>=1 then
 Response.Write("<script>alert('相同商品地址只能接手一次!');history.back();</script>")				 
 response.End()
end if

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "select * from "&jieducm&"_zhongxin where pid='"&id&"' ",conn,3,3
if not (rs.eof) then
  Shopkeeper=rs("Shopkeeper")
  ProUrl=rs("ProUrl")
  bei=rs("bei")
  rs("jieshou_num")=rs("jieshou_num")+1
  rs.update   
  rs.close
end if


     Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_jieshou " ,Conn,3,3  
	    rs.addnew
    	rs("Shopkeeper")=Shopkeeper
		rs("ProUrl")=ProUrl
		rs("bei")=bei		
		rs("username")=session("useridname")
		rs("username2")=username2
		rs("pid")=id
		rs("start")="1"
		rs("num")=num
		rs("classid")=6
		rs("ip")=ip
    	rs.update
    	rs.close

	 num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=num
		rs("class")="淘宝收藏区"
		rs("content")=session("useridname")&"接手"&user&"的任务"&id&"成功"
		rs("price")=0
		rs("jiegou")="等待收藏"
    	rs.update
    	rs.close	

set rs=nothing
conn.close
set conn=nothing
Response.Write("<script>alert('你已领取了这个任务,请收藏商品或店铺地址,\n任务完成后,你将获得发布点!');window.location.href='remission.asp';</script>")
response.End()
%>