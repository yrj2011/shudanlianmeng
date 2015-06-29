<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../taobao/checklogin.asp"-->
<%
dim From_url,Serv_url
From_url = Cstr(Request.ServerVariables("HTTP_REFERER"))
Serv_url = Cstr(Request.ServerVariables("SERVER_NAME"))
if mid(From_url,8,len(Serv_url)) <> Serv_url or Request.ServerVariables("HTTP_REFERER")="" then
response.write "请求超时,请关闭本页面并重新登陆即可!"
response.end
end If
response.contenttype = "text/html"
response.expires = 0 
response.expiresabsolute = now() - 1 
response.addHeader "pragma","no-cache" 
response.addHeader "cache-control","private" 
response.cachecontrol = "no-cache"
Response.ContentType = "text/html; charset=gb2312"
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
On Error Resume Next 
Server.ScriptTimeOut=9999999

if request.Form("fa")="ok" then				
Dim Price(6)
Dim ProUrl(6)
Price(1)=request.Form("Price1")
ProUrl(1)=Request.Form("ProUrl1")
Price(2)=request.Form("Price2")
ProUrl(2)=Request.Form("ProUrl2")
Price(3)=request.Form("Price3")
ProUrl(3)=Request.Form("ProUrl3")
Price(4)=request.Form("Price4")
ProUrl(4)=Request.Form("ProUrl4")
Price(5)=request.Form("Price5")
ProUrl(5)=Request.Form("ProUrl5")
Shopkeeper=Replace(Trim(request.Form("Shopkeeper")),"'","''")
bei=Request.Form("bei")
baohu2=request.Form("baohu2")
isprice=request.form("isprice")
num=request.Form("num")
depot=request.form("depot")
tips=Replace(Trim(request.Form("tips")),"'","''")
play=request.Form("play")
Limit=request.Form("Limit")
title=Replace(Trim(Request.Form("title")),"'","''")
addfabu=request.Form("addfabu")
ping=Replace(Trim(Request.Form("ping")),"'","''")
pingtxt=Replace(Trim(Request.Form("pingtxt")),"'","''")
outtime=Replace(Trim(Request.Form("Timing")),"'","''")
jieducm_IfComeSet=request.Form("jieducm_IfComeSet")
IfComeSet=request.Form("IfComeSet")
ComeKey=Replace(Trim(Request.Form("ComeKey")),"'","''")
ComeNote=Replace(Trim(Request.Form("ComeNote")),"'","''")
if outtime="" then
outtime=now
end if
fabu2=0
k=0
if num="" then
	Response.Write("<script>alert('请先选择连发几个任务!');history.back();</script>")
    response.End()
elseif  instr (ProUrl(1),"shop") then
 Response.Write("<script>alert('请填写正确宝贝地址！');history.back();</script>")
 response.End()
elseif Price(1)="" then
	Response.Write("<script>alert('任务1商品任务价不能为空!');history.back();</script>")
    response.End()
elseif (Price(1))<0.3  or (Price(0.3))>1000 then
    Response.Write("<script>alert('目前本平台只支持发布0.3-1000元的任务 ，请重新输入!');history.back();</script>")
    response.End()		
elseif Shopkeeper="" then
	Response.Write("<script>alert('你还没有选择掌柜名!');history.back();</script>")
    response.End()	
elseif ProUrl(1)="" then
	Response.Write("<script>alert('任务1商品地址不能为空!');history.back();</script>")
    response.End()
elseif bei="" then
	Response.Write("<script>alert('你还没有选择备注!');history.back();</script>")
    response.End()
elseif not isNumeric(addfabu) then
  Response.Write("<script>alert('输入有误请重新输入！');history.back();</script>")
  response.End()
elseif int(addfabu)<0 then
  Response.Write("<script>alert('输入有误请重新输入！');history.back();</script>")
  response.End()
end if

sql="SELECT  count(*) as totalz FROM "&jieducm&"_zhongxin where classid='1' and username='"&session("useridname")&"' and start<>'5'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
totalzx=rs("totalz")
end if 

if int(totalzx+num)>int(stupfaweiok+5)  then
	Response.Write("<script>alert('你VIP会员只能发布"&stupfaweiok+5&"个未完成的任务！\n请先完成已发的任务!才能继续发任务！');history.back();</script>")
    response.End()
end if 

for i=1 to num
 if Price(i)="" then
	Response.Write("<script>alert('任务"&i&"的商品价不能为空!');history.back();</script>")
    response.End()
 elseif ProUrl(i)="" then
	Response.Write("<script>alert('任务"&i&"的商品地址不能为空!');history.back();</script>")
    response.End()
 elseif (Price(i))<0.3  or (Price(i))>1000 then
    Response.Write("<script>alert('目前本平台只支持发布0.3-1000元的任务 ，请重新输入!');history.back();</script>")
    response.End()
elseif  instr (ProUrl(i),"meal") and addfabu=0 then
 Response.Write("<script>alert('发布套餐任务的宝贝请先增加发布点！');history.back();</script>")
 response.End()
end if
 
url= ProUrl(i)
wstr=getHTTPPage(url) 
Set objRegExp=New RegExp 
objRegExp.IgnoreCase=true
objRegExp.Global=FALSE
objRegExp.Pattern="<a class=""hCard fn"" href=""[^""]+""[^>]*>(.*?)</a>"
set Matches=objRegExp.Execute(wstr)
jieducm_sk=Matches(0).SubMatches(0)
 
Set objRegExp=New RegExp 
objRegExp.IgnoreCase=true
objRegExp.Global=FALSE
objRegExp.Pattern="<a class=""hCard fn"" title=""(.*?)"""
set Matches=objRegExp.Execute(wstr)
jieducm_sk3=Matches(0).SubMatches(0)


if Price(i)>=0.3 and Price(i)<=40 then
   fabu=2
elseif Price(i)>41 and Price(i)<=100 then
   fabu=3
elseif Price(i)>101 and Price(i)<=200 then
   fabu=4
elseif Price(i)>201 and Price(i)<=500 then
  fabu=5
elseif Price(i)>501 and Price(i)<=1000 then
  fabu=6
end if 

if bei="一天后收货好评" then
fabu=fabu*1+1
elseif bei="二天后收货好评"  then
fabu=fabu*1+2
elseif bei="三天后收货好评"  then
fabu=fabu*1+3
elseif bei="四天后收货好评"  then
fabu=fabu*1+4
elseif bei="五天后收货好评"  then
fabu=fabu*1+5
end if
if jieducm_IfComeSet=1 then
fabu=fabu+1
end if
fabu2=int(fabu2)+int(fabu)+int(addfabu)
k=k+Price(i) 
Next 

jieducm_ck=cunkuan*100
jieducm_price=k*100

fabudian2=fabudian*100
fabu3=fabu2*100

if isfa=1 then
    Response.Write("<script>alert('你已被管理员禁止了发任务功能!');history.back();</script>")
    response.End()
elseif int(jieducm_ck)< int(jieducm_price) then
 Response.Write("<script>alert('你发布的任务价钱不能大于你的存款!');history.back();</script>")
 response.End()
elseif(int(fabudian2)<int(fabu3)) then
	Response.Write("<script>alert('你的发布点数量不足,请尽快充值或换取发布点!');history.back();</script>")
	response.End()	
end if

for i=1 to num				 
randomize
if Price(i)>=0.3 and Price(i)<=40 then
   fabu=2
elseif Price(i)>40 and Price(i)<=100 then
   fabu=3
elseif Price(i)>100 and Price(i)<=200 then
   fabu=4
elseif Price(i)>200 and Price(i)<=500 then
  fabu=5
elseif Price(i)>500 and Price(i)<=1000 then
  fabu=6
end if 

if bei="一天后收货好评" then
fabu=fabu*1+1
elseif bei="二天后收货好评"  then
fabu=fabu*1+2
elseif bei="三天后收货好评"  then
fabu=fabu*1+3
elseif bei="四天后收货好评"  then
fabu=fabu*1+4
elseif bei="五天后收货好评"  then
fabu=fabu*1+5
end if

if jieducm_IfComeSet=1 then
fabu=fabu+1
end if

ranNum=int(90000*rnd)+10000	
pid=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select * From "&jieducm&"_zhongxin where pid='"&pid&"'" ,Conn,3,3  
if  rs.eof then
rs.addnew
rs("classid")=1
rs("Price")=Price(i)
rs("Shopkeeper")=Shopkeeper
rs("ProUrl")=ProUrl(i)
rs("bei")=bei	
if baohu2="" then
rs("baohu2")=0
else
rs("baohu2")=baohu2
end if
rs("username")=session("useridname")
rs("pid")=pid
rs("now")=outtime
rs("isprice")=isprice
rs("play")=play
rs("tips")=tips
rs("Limit")=Limit
rs("addfabu")=addfabu
rs("ping")=ping
rs("pingtxt")=pingtxt
rs("jieducm_fb")=fabu
if jieducm_IfComeSet=1 then
rs("IfComeSet")=IfComeSet
rs("ComeKey")=ComeKey
rs("ComeNote")=ComeNote
end if
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
rs("classid")=1
rs("Price")=Price(i)
rs("Shopkeeper")=Shopkeeper
rs("ProUrl")=ProUrl(i)
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
rs("addfabu")=addfabu
rs("ping")=ping
rs("pingtxt")=pingtxt
rs("jieducm_fb")=fabu
if jieducm_IfComeSet=1 then
rs("IfComeSet")=IfComeSet
rs("ComeKey")=ComeKey
rs("ComeNote")=ComeNote
end if
rs.update
rs.close
end if

Next  
 
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select cunkuan,fabudian From "&jieducm&"_register where username='"&session("useridname")&"' " ,Conn,3,3  
rs("fabudian")=rs("fabudian")-fabu2
rs("cunkuan")=rs("cunkuan")-k
rs.update
rs.close
call check_jieducm_name(session("useridname"))
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select * From "&jieducm&"_record" ,Conn,3,3  
rs.addnew
rs("username")=session("useridname")
rs("num")=pid
rs("class")="淘宝VIP连发任务"
rs("content")=session("useridname")&"发布vip连发任务成功,存款减少"&k&"元，发布点减少了"&fabu2&"个"
rs("price")="-"&k
rs("jiegou")="成功"
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing		
Response.Write("<script>alert('发布成功!等待接手人！');window.location.href='MyMission.asp';</script>")	
response.End()
end if
%>