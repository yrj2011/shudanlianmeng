<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="checklogin.asp"-->
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
On Error Resume Next 
Server.ScriptTimeOut=9999999 
if request.Form("fa")="ok" then
Price=Request.Form("Price1")
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
if  instr (ProUrl,"shop") then
 Response.Write("<script>alert('请填写正确宝贝地址！');history.back();</script>")
 response.End()
elseif  instr (ProUrl,"meal") then
 Response.Write("<script>alert('只有VIP会员才可以发布套餐任务的宝贝！');history.back();</script>")
 response.End()
elseif cunkuan<0.1 then
 Response.Write("<script>alert('存款低于0.1元时不能再发任务！');window.location.href='../user/md5_pay.asp';</script>")
 response.End()
elseif fabudian<1 then
 Response.Write("<script>alert('发布点低于1个时不能再发任务！');window.location.href='../user/md5_pay.asp';</script>")
 response.End() 
elseif price="" then
	Response.Write("<script>alert('商品任务价不能为空!');history.back();</script>")
    response.End()
elseif int(price)<0.1  or int(price)>1000 then
	Response.Write("<script>alert('目前本平台只支持发布0.1-1000元的任务 ，请重新输入!');history.back();</script>")
    response.End()
elseif Shopkeeper="" then
	Response.Write("<script>alert('你还没有选择掌柜名!');history.back();</script>")
    response.End()	
elseif prourl="" then
	Response.Write("<script>alert('商品地址不能为空!');history.back();</script>")
    response.End()
elseif ping="" then
	Response.Write("<script>alert('你还没有选择好评方式!');history.back();</script>")
    response.End()
elseif ping="自定义评语"  and pingtxt="" then
	Response.Write("<script>alert('请输入自定义评语!');history.back();</script>")
    response.End()
elseif bei="" then
	Response.Write("<script>alert('你还没有选择备注!');history.back();</script>")
    response.End()
elseif not isNumeric(price) then
  Response.Write("<script>alert('输入有误请重新输入！');history.back();</script>")
  response.End()
end if

sql="SELECT  count(*) as totalz FROM "&jieducm&"_zhongxin where classid='1' and username='"&session("useridname")&"' and start<>'5'"
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

url= ProUrl
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

								
if price>=0.1 and price<=40 then
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
if jieducm_IfComeSet=1 then
fabu=fabu+1
end if
jieducm_ck=cunkuan*100
jieducm_price=price*100
fabudian2=fabudian*100
fabu2=fabu*100
if int(jieducm_ck)< int(jieducm_price) then
	 Response.Write("<script>alert('你发布的任务价钱不能大于你的存款!');history.back();</script>")
	 response.End()
elseif(int(fabudian2)<int(fabu2)) then
	Response.Write("<script>alert('你的发布点数量不足,请尽快充值或换取发布点!');history.back();</script>")
	response.End()  
elseif isfa="1" then
	Response.Write("<script>alert('你已被管理员禁止了发任务功能!');history.back();</script>")
	response.End()
end if
          
randomize	
ranNum=int(90000*rnd)+10000	
pid=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select * From "&jieducm&"_zhongxin  where pid='"&pid&"'" ,Conn,3,3
if  rs.eof then  
rs.addnew
rs("classid")=1
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
rs("now")=outtime
rs("isprice")=isprice
rs("play")=play
rs("tips")=tips
rs("Limit")=Limit
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

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select cunkuan,fabudian From "&jieducm&"_register where username='"&session("useridname")&"'" ,Conn,3,3  
rs("cunkuan")=rs("cunkuan")-price
rs("fabudian")=rs("fabudian")-fabu
rs.update
rs.close
call check_jieducm_name(session("useridname"))	
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
rs.addnew
rs("username")=session("useridname")
rs("num")=pid
rs("class")="淘宝区任务"
rs("content")=session("useridname")&"发布任务"&pid&"成功,存款减少"&price&"元，发布点减少了"&fabu&"个"
rs("price")="-"&price
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