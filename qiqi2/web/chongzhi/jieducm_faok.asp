<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
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
On Error Resume Next 
Server.ScriptTimeOut=9999999 
if request.Form("fa")="ok" then
Price=Replace(Trim(Request.Form("Price1")),"'","''")
ProUrl=Replace(Trim(Request.Form("ProUrl1")),"'","''")
Shopkeeper=Replace(Trim(Request.Form("Shopkeeper")),"'","''")
bei=Replace(Trim(Request.Form("bei")),"'","''")
baohu2=request.Form("baohu2")
tips=Replace(Trim(Request.Form("tips")),"'","''")

if cunkuan<100 then
 Response.Write("<script>alert('存款低于100元时不能发布提现任务！');window.location.href='../user/md5_pay.asp';</script>")
 response.End()
elseif  instr (ProUrl,"shop") then
 Response.Write("<script>alert('请填写正确宝贝地址！');history.back();</script>")
 response.End()
elseif price="" then
	Response.Write("<script>alert('提现金额不能为空!');history.back();</script>")
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
end if

sql="SELECT  count(*) as totalz FROM "&jieducm&"_zhongxin where classid='5' and username='"&session("useridname")&"' and ( datediff(day,[now],getdate())=0) "
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
totalzx=rs("totalz")
end if 

if int(totalzx)>=int(1) and vip<>"1" then
	Response.Write("<script>alert('普通会员每天只能申请一次提现！');history.back();</script>")
    response.End()
elseif int(totalzx)>=int(3) and vip="1" then
	Response.Write("<script>alert('VIP会员每天只可以申请三次提现！！');history.back();</script>")
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

wstr=getHTTPPage(url) 
Set objRegExp=New RegExp 
objRegExp.IgnoreCase=true
objRegExp.Global=FALSE
objRegExp.Pattern="<a class=""hCard fn"" title=""(.*?)"""
set Matches=objRegExp.Execute(wstr)
jieducm_sk3=Matches(0).SubMatches(0)

if (jieducm_sk=shopkeeper)   or (jieducm_sk3=shopkeeper) then
else
	Response.Write("<script>alert('此掌柜名与填写的商品地址不符!请重新输入!');history.back();</script>")
    response.End()
end if


jieducm_ck=cunkuan*100
jieducm_price=price*100

if int(jieducm_ck)< int(jieducm_price) then
	 Response.Write("<script>alert('你提现金额不能大于你的存款!');history.back();</script>")
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
		rs("classid")=5
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
		rs("tips")=tips
    	rs.update
    	rs.close
		else
		   Response.Write("<script>alert('出错了，请返回重新发布!');history.back();</script>")
	       response.End()
		end if


      Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select cunkuan From "&jieducm&"_register where username='"&session("useridname")&"' " ,Conn,3,3  
    	rs("cunkuan")=rs("cunkuan")-price
    	rs.update
    	rs.close
		
 	 Set rs=server.createobject("ADODB.RECORDSET")
	  rs.open "Select * From "&jieducm&"_record " ,Conn,3,3  
	    rs.addnew
		rs("username")=session("useridname")
    	rs("num")=pid
		rs("class")="充值提现任务区"
		rs("content")=session("useridname")&"发布提现任务"&pid&"成功,存款减少"&price&"元"
		rs("price")="-"&price
		rs("jiegou")="成功"
    	rs.update
    	rs.close
	call check_jieducm_name(session("useridname"))  			
set rs=nothing 
conn.close
set conn=nothing
Response.Write("<script>alert('发布提现任务成功!等待接手人！');window.location.href='MyMission.asp';</script>")	
response.End()
end if
%>