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
On Error Resume Next 
Server.ScriptTimeOut=9999999 
if request.Form("fa")="ok" then
Shopkeeper=Replace(Trim(Request.Form("Shopkeeper")),"'","''")
info=Replace(Trim(Request.Form("info[remarks]")),"'","''")
ProUrl=Replace(Trim(Request.Form("ProUrl")),"'","''")
product=Replace(Trim(Request.Form("product[number]")),"'","''")
depot=request.form("depot")
title=Replace(Trim(Request.Form("title")),"'","''")
tips=Replace(Trim(Request.Form("tips")),"'","''")

if  instr(ProUrl,"shop") and info="宝贝收藏任务" then
 Response.Write("<script>alert('请填写正确宝贝地址！');history.back();</script>")
 response.End()
elseif  instr(ProUrl,"shop")<>8 and info="店铺收藏任务" then
 Response.Write("<script>alert('请填写正确店铺地址！');history.back();</script>")
 response.End()
elseif Shopkeeper="" then
	Response.Write("<script>alert('你还没有选择掌柜名!');history.back();</script>")
    response.End()	
elseif prourl="" then
	Response.Write("<script>alert('请输入店铺或宝贝地址!');history.back();</script>")
    response.End()
elseif product="" then
	Response.Write("<script>alert('收藏数量不能为空!');history.back();</script>")
    response.End()
elseif product<1 then
	Response.Write("<script>alert('收藏数量不能小于1!');history.back();</script>")
    response.End()
elseif not isNumeric(product) then
  Response.Write("<script>alert('输入有误请重新输入！');history.back();</script>")
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


Set objRegExp=New RegExp 
objRegExp.IgnoreCase=true
objRegExp.Global=FALSE
objRegExp.Pattern="<a href="".*?"" class=""hCard fn"">(.*?)</a>"
set Matches=objRegExp.Execute(wstr)
jieducm_sk2=Matches(0).SubMatches(0)

fabu=product*stupcang
fabudian2=fabudian*100
fabu2=fabu*100
if(int(fabudian2)<int(fabu2)) then
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
rs("classid")=6
rs("Shopkeeper")=Shopkeeper
rs("ProUrl")=ProUrl
rs("bei")=info
rs("username")=session("useridname")
rs("pid")=pid
rs("num")=product
rs("now")=now
rs("tips")=tips
rs("jieshou_num")=0
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
rs("classid")=6
rs("Shopkeeper")=Shopkeeper
rs("ProUrl")=ProUrl
rs("username")=session("useridname")
rs("now")=now
rs("bei")=info
rs("pid")=pid
rs("num")=1
rs("num2")=product
rs("title")=title
rs("tips")=tips
rs.update
rs.close
end if

Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select fabudian From "&jieducm&"_register where username='"&session("useridname")&"'" ,Conn,3,3  
rs("fabudian")=rs("fabudian")-fabu
rs.update
rs.close
call check_jieducm_name(session("useridname"))	
Set rs=server.createobject("ADODB.RECORDSET")
rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
rs.addnew
rs("username")=session("useridname")
rs("num")=pid
rs("class")="淘宝收藏区"
rs("content")=session("useridname")&"发布任务"&pid&"成功,发布点减少了"&fabu&"个"
rs("price")=0
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