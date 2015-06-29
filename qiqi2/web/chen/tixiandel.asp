<!--#include file="inc/conn.asp"-->

<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/function.asp"-->

<!--#INCLUDE FILE="../background.asp"-->

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
id=request.QueryString("id")
Set rs1=server.createobject("ADODB.RECORDSET")
rs1.open "select * from "&jieducm&"_tixian where id="&id&"",conn,3,3
if not (rs1.eof) then
  rs1("start")="0"  
  rs1.update
  rs1.close
  Response.Write("<script>alert('确定此用户提现金额\n还没有打到该用户支付宝账号上!');window.location.href='admin_bankroll.asp';</script>")
end if
%>