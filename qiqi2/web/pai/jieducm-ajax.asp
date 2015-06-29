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

%>
<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../taobao/checklogin.asp"-->
<!--#INCLUDE FILE="../background.asp"-->
<%open=saferequest("open","get")
if open="indexajex" then%>
<!--#INCLUDE FILE="index1.asp"-->
<%elseif open="myajex" then%>
<!--#INCLUDE FILE="MyMission1.asp"-->
<%elseif open="Reajex" then%>
<!--#INCLUDE FILE="ReMission1.asp"-->
<%end if%>
<%response.charset="gb2312"%>
 <%rs.close
 set rs=nothing
 conn.close
 set conn=nothing%>