<!--#include file="../inc/index_conn.asp"-->
<!--#INCLUDE FILE="../config.asp"-->
<!--#include file="checklogin.asp"-->
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
Conn.execute("update "&jieducm&"_register set  snow='"&now()&"' where username='"&session("useridname")&"'")
Conn.execute("Delete  from "&jieducm&"_online where username='"&session("useridname")&"'")
session("useridname")=""
session("czm")=""
response.Redirect("../index.asp")
%>
