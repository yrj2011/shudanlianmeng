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
'Copyright (C) 2008－2011 捷度传媒信息技术有限公司 版权所有
'*****************************************************************
'------------------------------------------------------------------
dim i
i=0
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
rs.open "select prourl,username from "&jieducm&"_zhongxin where pid='"&id&"' and start='0' ",conn,1,1
if rs.eof then
    Response.Write("<script>alert('对不起,此任务已被其它人接手!');history.back();</script>")
	response.End()
else
 username2=rs("username")
 prourlj=rs("prourl")
 Shopkeeper=rs("Shopkeeper")
end if

sql="SELECT  level FROM "&jieducm&"_register where username='"&username2&"' "
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
jieducm_level1=rs("level1")
end if 

if jieducm_level1="0" then
  Response.Write("<script>alert('此会员账号已被锁定!不能再接此会员任务！');history.back();</script>")
  response.End()
end if

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "select * from "&jieducm&"_zhongxin where pid='"&id&"' and start='0'",conn,3,3
if not (rs.eof) then
  classid=rs("classid")
  Price=rs("Price")
  Shopkeeper=rs("Shopkeeper")
  ProUrl=rs("ProUrl")
  bei=rs("bei")
  baohu2=rs("baohu2")
  user=rs("username")
  isprice=rs("isprice")
  rs("start")=1
  rs.update   
  rs.close
else
  Response.Write("<script>alert('对不起,此任务已被其它人接手!');history.back();</script>")				 
  response.End()
end if


     Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_jieshou where pid='"&id&"'" ,Conn,3,3  
	  if rs.eof then 
	    rs.addnew
		rs("Price")=Price
    	rs("Shopkeeper")=Shopkeeper
		rs("ProUrl")=ProUrl
		rs("bei")=bei
		rs("baohu2")=baohu2
		rs("username")=session("useridname")
		rs("username2")=username2
		rs("pid")=id
		rs("start")="1"
		rs("num")=num
		rs("classid")=classid
		rs("ip")=ip
		rs("isprice")=isprice
    	rs.update
    	rs.close
	  else
	    Response.Write("<script>alert('对不起,此任务已被其它人接手!');history.back();</script>")				 
        response.End()
	  end if
	  
	 num=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=num
		rs("class")="充值提现任务区"
		rs("content")=session("useridname")&"接手"&user&"的任务"&id&"成功"
		rs("price")=0
		rs("jiegou")="未支付"
    	rs.update
    	rs.close	
    Response.Write("<script>alert('你已领取了这个任务,请拍下商品,\n联系发布方,按任务价支付,\n请放心支付,发布方发部任务时已扣下了保证金,\n任务完成后,你将获的存款发布点积分!');window.location.href='remission.asp';</script>")
response.End()

set rs=nothing
conn.close
set conn=nothing
%>