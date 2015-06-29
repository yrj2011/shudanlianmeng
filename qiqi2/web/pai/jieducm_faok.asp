<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../taobao/checklogin.asp"-->
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
On Error Resume Next 
Server.ScriptTimeOut=9999999 
action=request.form("action")
if action="ok" then
Price=Replace(Trim(Request.Form("Price1")),"'","''")
ProUrl=Replace(Trim(Request.Form("ProUrl1")),"'","''")
Shopkeeper=Replace(Trim(Request.Form("Shopkeeper")),"'","''")
bei=Replace(Trim(Request.Form("bei")),"'","''")
baohu2=request.Form("baohu2")
isprice=request.form("isprice")
depot=request.form("depot")
play=request.form("play")
tips=Replace(Trim(Request.Form("tips")),"'","''")
Limit=request.Form("Limit")
title=Replace(Trim(Request.Form("title")),"'","''")

if cunkuan<1 then
 Response.Write("<script>alert('存款低于1元时不能再发任务！');window.location.href='../user/md5_pay.asp';</script>")
 response.End()
elseif fabudian<1 then
 Response.Write("<script>alert('发布点低于1个时不能再发任务！');window.location.href='../user/md5_pay.asp';</script>")
 response.End() 
elseif price="" then
	Response.Write("<script>alert('商品任务价不能为空!');history.back();</script>")
    response.End()
elseif int(price)<1  or int(price)>1000 then
	Response.Write("<script>alert('目前本平台只支持发布1-1000元的任务 ，请重新输入!');history.back();</script>")
    response.End()
elseif isfa="1" then
	Response.Write("<script>alert('你已被管理员禁止了发任务功能!');history.back();</script>")
	response.End()
elseif Shopkeeper="" then
	Response.Write("<script>alert('你还没有选择掌柜名!');history.back();</script>")
    response.End()	
elseif prourl="" then
	Response.Write("<script>alert('商品地址不能为空!');history.back();</script>")
    response.End()
elseif bei="" then
	Response.Write("<script>alert('你还没有选择备注!');history.back();</script>")
    response.End()
elseif not isNumeric(price) then
          Response.Write("<script>alert('输入有误请重新输入！');history.back();</script>")
          response.End()
end if

sql="SELECT  count(*) as totalz FROM "&jieducm&"_zhongxin where classid='2' and username='"&session("useridname")&"' and start<>'5'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
totalzx=rs("totalz")
end if 

if int(totalzx)>=int(stupfaweiok) and vip<>"1" then
	Response.Write("<script>alert('你是普通会员只能发布"&stupfaweiok&"个未完成的任务！\n请先完成已发的任务!才能再继续发任务！');history.back();</script>")
    response.End()
elseif int(totalzx)>=int(stupfaweiok+5) and vip="1" then
	Response.Write("<script>alert('你VIP会员只能发布"&stupfaweiok+5&"个未完成的任务！\n请先完成已发的任务!才能再继续发任务！');history.back();</script>")
    response.End()
end if 



if price>=1 and price<=40 then
   fabu=1
elseif price>40 and price<=100 then
   fabu=2
elseif price>100 and price<=200 then
   fabu=3
elseif price>200 and price<=500 then
  fabu=4
elseif price>500 and price<=1000 then
  fabu=5
end if 
if bei="一天后收货好评" then
fabu=fabu*2
elseif bei="二天后收货好评"  then
fabu=fabu*2+1
elseif bei="三天后收货好评"  then
fabu=fabu*2+2
end if

price2=price*100
cunkuan2=cunkuan*100
fabudian2=fabudian*100
fabu2=fabu*100

if(int(price2)>int(cunkuan2)) then
    Response.Write("<script>alert('你发布的任务价钱不能大于你的存款!');history.back();</script>")
    response.End()
elseif(int(fabudian2)<int(fabu2)) then
   Response.Write("<script>alert('你的发布点数量不足,请尽快充值或换取发布点!');history.back();</script>")
   response.End()			
end if

sql="SELECT  count(*) as totalz FROM "&jieducm&"_zhongxin where classid='2' and username='"&session("useridname")&"' and start<>'5'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
totalzx=rs("totalz")
end if 

if int(totalzx)>=int(stupfaweiok) and vip<>"1" then
	Response.Write("<script>alert('你是普通会员只能发布"&stupfaweiok&"个未完成的任务！\n请先完成已发的任务!才能再继续发任务！');history.back();</script>")
    response.End()
elseif int(totalzx)>=int(stupfaweiok+5) and vip="1" then
	Response.Write("<script>alert('你VIP会员只能发布"&stupfaweiok+5&"个未完成的任务！\n请先完成已发的任务!才能再继续发任务！');history.back();</script>")
    response.End()
end if 


randomize
if bei="马上收货好评"  then
company=""
numbers=""
else
Cid=int(7*rnd)+1
if cid=1 then
  company="申通E物流"
  if cid>4 then
     top="268"
  else
     top="888"
  end if
  rnum=int(900000000*rnd)+100000000
  numbers=top&rnum
elseif cid=2 then
  company="圆通速递"
  if cid=1 then
    top="12"
  elseif cid=2 then
    top="21"
  elseif cid=3 then
    top="22"
  else
    top="82"
  end if
  rnum=int(90000000*rnd)+10000000
  numbers=top&rnum
elseif cid=3 then
  company="中通快递"
 if cid=1 then
  top="6180"
 elseif cid=2 then
  top="6800"
 elseif cid=3 then
  top="2008"
 elseif cid=4 then
  top="5711"
 else
  top="0100"
 end if
  rnum=int(90000000*rnd)+10000000
  numbers=top&rnum
elseif cid=4 then
  company="宅急送"
  if cid=1 then
   top="1"
  elseif cid=2 then
   top="2"
  elseif cid=3 then
    top="6"
  elseif cid=4 then
    top="7"
  else
    top="8"
  end if
  rnum=int(900000000*rnd)+100000000
  numbers=top&rnum
elseif cid=5 then
  company="联邦快递"
  top="120"
  rnum=int(900000000*rnd)+100000000
  numbers=top&rnum
elseif cid=6 then
  company="德邦物流"
  top="1"
  rnum=int(9000000*rnd)+1000000
  numbers=top&rnum
elseif cid=7 then
  company="邮政平邮"
  if cid>4 then
  top="KA"
  else
  top="PA"
  end if
  rnum=int(90000000000*rnd)+10000000000
  numbers=top&rnum
elseif cid=8 then
  company="EMS快递"
  c=Chr(Int(Rnd * 26 + 65))
  rnum=int(900000000*rnd)+100000000
  numbers="E"&c&rnum&"CN"
end if
end if
		
ranNum=int(90000*rnd)+10000	
pid=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select * From "&jieducm&"_zhongxin  where pid='"&pid&"'" ,Conn,3,3
if  rs.eof then  
rs.addnew
rs("classid")=2
rs("Price")=Price
rs("Shopkeeper")=Shopkeeper
rs("ProUrl")=ProUrl
rs("bei")=bei	
if baohu2="" then
rs("baohu2")=0
else
rs("baohu2")=baohu2
end if
rs("username")=session("useridname")
rs("pid")=pid
rs("now")=now
rs("isprice")=isprice
rs("play")=play
rs("tips")=tips
rs("company")=company
rs("numbers")=numbers
rs("Limit")=Limit
rs("jieducm_fb")=fabu
rs.update
rs.close
else
Response.Write("<script>alert('出错了，请返回重新发布!');history.back();</script>")
response.End()
end if

if depot=1 then
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select * From "&jieducm&"_depot " ,Conn,3,3  
rs.addnew
rs("classid")=2
rs("Price")=Price
rs("Shopkeeper")=Shopkeeper
rs("ProUrl")=ProUrl
rs("bei")=bei
if baohu2="" then
rs("baohu2")=0
else
rs("baohu2")=baohu2
end if
rs("username")=session("useridname")
rs("now")=now
rs("isprice")=isprice
rs("pid")=pid
rs("num")=1
rs("play")=play
rs("tips")=tips
rs("company")=company
rs("numbers")=numbers
rs("Limit")=Limit
rs("title")=title
rs("jieducm_fb")=fabu
rs.update
rs.close
end if
      
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select cunkuan,fabudian From "&jieducm&"_register where username='"&session("useridname")&"'" ,Conn,3,3  
rs("cunkuan")=rs("cunkuan")-price
rs("fabudian")=rs("fabudian")-fabu
rs.update
rs.close
		 
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=pid
		rs("class")="拍拍区任务"
		rs("content")=session("useridname")&"发布任务"&pid&"成功,存款减少"&price&"元,发布点减少了"&fabu&"个"
		rs("price")="-"&price
		rs("jiegou")="成功"
    	rs.update
    	rs.close	
call check_jieducm_name(session("useridname")) 		
set rs=nothing
conn.close
set conn=nothing
Response.Write("<script>alert('发布成功!等待接手人！');window.location.href='MyMission.asp';</script>")	
response.End()
end if
%>