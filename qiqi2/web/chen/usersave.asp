<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/function.asp"-->
<!--#INCLUDE FILE="inc/md5.asp"-->

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
Set rs=server.createobject("ADODB.RECORDSET")
if request.Form<>"" then
  if request("password")<>"" then
	 password=md5(Request.Form("password"))
  end if
	 mail=request.Form("email")
     HomePage=Trim(Request.Form("HomePage"))
	 qq=request.Form("qq")
	 face=request.Form("face")
	 sex=request.Form("sex")
	 tjr=Request.Form("tjr")
	 phone=Request.Form("phone")
	 rname=Request.Form("rname")
	 shopnote=Request.Form("shopnote")
	 address=Request.Form("address")
	 czm=Request.Form("czm")
	 weiti=request("weiti")
	 daai=request("daai")
	 dst=request("dst")
	 login_ip=request.Form("login_ip")
	 if homepage="http://" then homepage="" 
	  rs.open "Select * From "&jieducm&"_register where username='"&request("username")&"' " ,Conn,3,3  
	  if request("password")<>"" then
		rs("password1")=password
	  end if
    	rs("mail")=mail
		rs("homepage")=homepage
		rs("qq")=qq
		rs("sex")=sex
		rs("tjr")=tjr
		rs("phone")=phone
		rs("rname")=rname
		rs("shopnote")=shopnote		
		rs("address")=address
		rs("czm")=czm
		rs("weiti")=weiti
		rs("daai")=daai
		rs("login_ip")=login_ip
		rs("dst")=dst
    	rs.update
    	rs.close
		
		
	Set rsr=server.createobject("ADODB.RECORDSET")
	  rsr.open "Select * From "&jieducm&"_recordm " ,Conn,3,3  
	    rsr.addnew
		rsr("username")=session("AdminName")
		rsr("class")="修改用户信息"
		rsr("content")=session("AdminName")&"修改了"&request("username")&"的信息"		
		rsr("jiegou")="修改成功"
    	rsr.update
    	rsr.close
		username=request("username")
     Response.Write("<script>alert('信息修改成功!');window.location.href='usermanage.asp?action=sear&text="&username&"';</script>")

set rs=nothing
conn.close
set conn=nothing
end if
%>