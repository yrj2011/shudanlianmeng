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
    Response.Write("<script>alert('接手拍拍号不能为空');history.back();</script>")
    response.End()		 
end if		 

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "select Shopkeeper,prourl,username,Limit from "&jieducm&"_zhongxin where pid='"&id&"' and start='0' ",conn,3,3
if rs.eof then
    Response.Write("<script>alert('对不起,此任务已被其它人接手!');history.back();</script>")
	response.End()
else
 username2=rs("username")
 prourlj=rs("prourl")
 Shopkeeper=Replace(Trim(rs("Shopkeeper")),"'","''")
 Limit=rs("Limit")
end if
if username2=session("useridname") then
Response.Write("<script>alert('不能接手自己的任务!');history.back();</script>")
response.End()
end if
call jname(username2)

if Limit=2 and jifei<100 then
  Response.Write("<script>alert('本任务限制100积分以上的会员才可以接手!');history.back();</script>")
  response.End()
end if
if Limit=3 and jifei<300 then
  Response.Write("<script>alert('本任务限制300积分以上的会员才可以接手!');history.back();</script>")
  response.End()
end if
if Limit=4 and vip="0" then
  Response.Write("<script>alert('本任务限制只能是VIP会员才可以接手!');history.back();</script>")
  response.End()
end if

sql="SELECT  count(*) as totalj FROM "&jieducm&"_jieshou where classid='2' and username='"&session("useridname")&"' and start<>'5'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
totalzx=rs("totalj")
end if 

if int(totalzx)>=int(stupjieweiok) and vip<>"1" then
	Response.Write("<script>alert('你是普通会员只能接手"&stupjieweiok&"个未完成的任务！\n请先完成已接的任务!才能再继续接任务！');history.back();</script>")
    response.End()
elseif int(totalzx)>=int(stupjieweiok+5) and vip="1" then
	Response.Write("<script>alert('VIP会员只能接手"&stupjieweiok+5&"个未完成的任务！\n请先完成已接的任务!才能再继续接任务！');history.back();</script>")
    response.End()
end if 

sql="SELECT  level1,vip FROM "&jieducm&"_register where username='"&username2&"' "
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
vipf=rs("vip")
jieducm_level1=rs("level1")
end if 

if jieducm_level1="0" then
  Response.Write("<script>alert('此会员账号已被锁定!不能再接此会员任务！');history.back();</script>")
  response.End()
end if


sql="SELECT  count(*) as totalnum FROM  "&jieducm&"_jieshou where username='"&session("useridname")&"' and  num='"&num&"' and proUrl='"&ProUrlj&"' and  (datediff(month,[now],getdate())=0) and classid='2'"
Set rs = Server.CreateObject("Adodb.RecordSet")
rs.Open Sql,conn,1,1
if not rs.eof then	
jh=rs("totalnum")
end if
if jh>=1 then
 Response.Write("<script>alert('相同商品地址一个月只能接手一次!');history.back();</script>")				 
 response.End()
end if


sql="SELECT  count(*) as monthnum FROM  "&jieducm&"_jieshou where Shopkeeper='"&Shopkeeper&"' and username='"&session("useridname")&"' and  num='"&num&"' and classid='2' and  (datediff(month,[now],getdate())=0) "
Set rs = Server.CreateObject("Adodb.RecordSet")
rs.Open Sql,conn,1,1
if not rs.eof then	
jieducm_monthnum=rs("monthnum")
end if
if jieducm_monthnum>=stupgeshu then
 Response.Write("<script>alert('你的小号已经和这个大号本月已交易过"&stupgeshu&"次了,\n再交易下去对方会发怒的!\n请您更换小号,或者接别人的任务吧!');history.back();</script>")				 
 response.End()
end if

sql="SELECT  count(*) as monthip FROM  "&jieducm&"_jieshou where Shopkeeper='"&Shopkeeper&"' and username='"&session("useridname")&"' and  ip='"&ip&"' and classid='2' and  (datediff(month,[now],getdate())=0) "
Set rs = Server.CreateObject("Adodb.RecordSet")
rs.Open Sql,conn,1,1
if not rs.eof then	
jieducm_monthip=rs("monthip")
end if

if jieducm_monthip>=3 then
 Response.Write("<script>alert('一个月内每个IP只能接手3次同一个店铺的商品!');history.back();</script>")				 
 response.End()
end if

if vipf="1" then

sql="SELECT  count(*) as datanum FROM  "&jieducm&"_jieshou where shopkeeper='"&Shopkeeper&"' and username='"&session("useridname")&"'  and num='"&num&"' and classid='2' and  ( datediff(day,[now],getdate())=0)"
Set rs = Server.CreateObject("Adodb.RecordSet")
rs.Open Sql,conn,1,1
if not rs.eof then	
jieducm_datamnum=rs("datanum")
end if
if jieducm_datamnum>=5 then
 Response.Write("<script>alert('每日每个淘宝买号最多只能接手同一个店铺5个商品链接!');history.back();</script>")				 
 response.End()
end if


sql="SELECT  count(*) as jieducm_aobaonum FROM  "&jieducm&"_jieshou where  num='"&num&"' and classid='2' and  ( datediff(day,[now],getdate())=0)"
Set rs = Server.CreateObject("Adodb.RecordSet")
rs.Open Sql,conn,1,1
if not rs.eof then	
jieducm_tbdatamnum=rs("jieducm_aobaonum")
end if
if jieducm_tbdatamnum>=20 then
 Response.Write("<script>alert('每日每个淘宝买号最多一天只能接20个商品链接!');history.back();</script>")				 
 response.End()
end if

sql="SELECT  count(*) as dataip FROM  "&jieducm&"_jieshou where Shopkeeper='"&Shopkeeper&"' and username='"&session("useridname")&"' and  ip='"&ip&"' and classid='2' and  (datediff(day,[now],getdate())=0) "
Set rs = Server.CreateObject("Adodb.RecordSet")
rs.Open Sql,conn,1,1
if not rs.eof then	
jieducm_dataip=rs("dataip")
end if

if jieducm_dataip>=1 then
 Response.Write("<script>alert('每日每个IP只能接手一次同一个店铺的商品!');history.back();</script>")				 
 response.End()
end if

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
		rs("classid")=2
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
		rs("class")="拍拍区任务"
		rs("content")=session("useridname")&"接手"&user&"的任务"&id&"成功"
		rs("price")=0
		rs("jiegou")="未支付"
    	rs.update
    	rs.close	

set rs=nothing
conn.close
set conn=nothing
Response.Write("<script>alert('你已领取了这个任务,请拍下商品,\n联系发布方,按任务价支付,\n请放心支付,发布方发布任务时已扣下了保证金,\n任务完成后,你将获得存款发布点积分!');window.location.href='remission.asp';</script>")
response.End()
%>