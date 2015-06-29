<%@language=vbscript codepage=936 %>
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
option explicit
response.buffer=true	
Const PurviewLevel=0
%>
<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/function.asp"-->
<%
dim sql,rs,jieducm_rgtotal,jieducm_viptotal,jieducm_adminbuycar,jieducm_zhtotal,jieducm_jstotal,jieducm_zhtotalok
dim jieducm_fbtotal,jieducm_cartotal,jieducm_chian_total,jieducm_alipay_total
%>
<html>
<head>
<title>后台管理首页</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="Admin_Style.css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center>
<tr align="center">
    <td height=25 colspan=2 class="topbg"><strong>后台管理首页</strong></tr>
<tr>
    <td width="100" class="tdbg" height=23><strong>管理快捷方式：</strong></td>
    <td class="tdbg"><a href="Admin_ArticleAdd1.asp">添加文章 | </a><a href="Admin_ArticleManage.asp">文章管理</a>      </td>
</tr>
</table>
<table cellpadding="2" cellspacing="1" border="0" width="100%" class="border" align=center>
  <tr align="center"> 
    <td height=25 colspan=3 class="topbg"><strong>今日平台详情</strong></td> 
  <tr> 
    <td width="21%"  class="tdbg" height=23>今日注册会员数：<font color="#ff6600">
    <%
sql="SELECT count(*) as jieducm_rgnum FROM "&jieducm&"_register where ( datediff(day,[now],getdate())=0)"
Set rs = Server.CreateObject("Adodb.RecordSet")
rs.Open Sql,conn,1,1
if not rs.eof then			
jieducm_rgtotal=rs("jieducm_rgnum")
end if
rs.close
set rs=nothing
response.Write(jieducm_rgtotal)
	%></font>个</td>
    <td width="29%" class="tdbg">今日发布任务数：<font color="#ff6600">
    <%
sql="SELECT count(*) as jieducm_zhnum FROM "&jieducm&"_zhongxin where ( datediff(day,[now],getdate())=0)"
Set rs = Server.CreateObject("Adodb.RecordSet")
rs.Open Sql,conn,1,1
if not rs.eof then			
jieducm_zhtotal=rs("jieducm_zhnum")
end if
rs.close
set rs=nothing
response.Write(jieducm_zhtotal)
	%></font>个</td>
    <td width="50%" class="tdbg">今日卖出发布点：<font color="#ff6600">
<%
sql="SELECT  sum(fabudian) as adminbuyfb FROM "&jieducm&"_chongjilu where class='购买发布点' and (datediff(day,[now],getdate())=0)"
Set rs = Server.CreateObject("Adodb.RecordSet")
rs.Open Sql,conn,1,1
if not rs.eof then			
jieducm_fbtotal=rs("adminbuyfb")
end if
rs.close
set rs=nothing

if not isnull(jieducm_fbtotal) then
response.Write(jieducm_fbtotal)
else
response.Write("0")
end if
	%></font>个</td>
  </tr>
  <tr> 
    <td width="21%" class="tdbg" height=23>今日加入vip会员数：<font color="#ff6600">
<%
sql="SELECT count(*) as jieducm_vipnum FROM "&jieducm&"_chongjilu where class='购买vip卡' and ( datediff(day,[now],getdate())=0)"
Set rs = Server.CreateObject("Adodb.RecordSet")
rs.Open Sql,conn,1,1
if not rs.eof then			
jieducm_viptotal=rs("jieducm_vipnum")
end if
rs.close
set rs=nothing
response.Write(jieducm_viptotal)
	%></font>个</td>
    <td width="29%" class="tdbg">今日接数任务数：<font color="#ff6600">
    <%
sql="SELECT count(*) as jieducm_jsnum FROM "&jieducm&"_jieshou where ( datediff(day,[now],getdate())=0)"
Set rs = Server.CreateObject("Adodb.RecordSet")
rs.Open Sql,conn,1,1
if not rs.eof then			
jieducm_jstotal=rs("jieducm_jsnum")
end if
rs.close
set rs=nothing
response.Write(jieducm_jstotal)
	%></font>个</td>
    <td width="50%" class="tdbg">今日充值卡充值：<font color="#ff6600">
<%
sql="SELECT  sum(cunkuan) as jieducm_buycar2 FROM "&jieducm&"_chongjilu where class='充值卡充值' and (datediff(day,[now],getdate())=0)"
Set rs = Server.CreateObject("Adodb.RecordSet")
rs.Open Sql,conn,1,1
if not rs.eof then			
jieducm_cartotal=rs("jieducm_buycar2")
end if
rs.close
set rs=nothing

if not isnull(jieducm_cartotal) then
response.Write(jieducm_cartotal)
else
response.Write("0")
end if
	%></font>元</td>
  </tr>
  <tr> 
    <td width="21%" class="tdbg" height=23>今日卖出点卡：<font color="#ff6600"><%
sql="SELECT   sum(cunkuan) as adminbuycar FROM "&jieducm&"_chongjilu where class='购买点卡' and ( datediff(day,[now],getdate())=0)"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not( rs.eof or rs.bof) then			
jieducm_adminbuycar=rs("adminbuycar")
end if
if not isnull(jieducm_adminbuycar) then
response.Write(jieducm_adminbuycar)
else
response.Write("0")
end if
	%></font>元</td>
    <td width="29%" class="tdbg">今日完成任务数：<font color="#ff6600">
    <%
sql="SELECT count(*) as jieducm_zhnumok FROM "&jieducm&"_zhongxin where start='5' and ( datediff(day,[now],getdate())=0)"
Set rs = Server.CreateObject("Adodb.RecordSet")
rs.Open Sql,conn,1,1
if not rs.eof then			
jieducm_zhtotalok=rs("jieducm_zhnumok")
end if
rs.close
set rs=nothing
response.Write(jieducm_zhtotalok)
	%></font>个</td>
    <td width="50%" class="tdbg">今日网银充值：<font color="#ff6600">
<%
sql="SELECT  sum(cunkuan) as jieducm_china FROM "&jieducm&"_chongjilu where class='网银充值' and (datediff(day,[now],getdate())=0)"
Set rs = Server.CreateObject("Adodb.RecordSet")
rs.Open Sql,conn,1,1
if not rs.eof then			
jieducm_chian_total=rs("jieducm_china")
end if
rs.close
set rs=nothing

if not isnull(jieducm_chian_total) then
response.Write(jieducm_chian_total)
else
response.Write("0")
end if
	%>
    </font>元&nbsp;&nbsp;
<%
sql="SELECT  sum(Money) as jieducm_alipay FROM AutoCZ where  Status=1 and (datediff(day,[EndDate],getdate())=0)"
Set rs = Server.CreateObject("Adodb.RecordSet")
rs.Open Sql,conn,1,1
if not rs.eof then			
jieducm_alipay_total=rs("jieducm_alipay")
end if
rs.close
set rs=nothing
%>
	
	 今日支付宝自动充值：<font color="#ff6600"><%if not isnull(jieducm_alipay_total) then
response.Write(jieducm_alipay_total)
else
response.Write("0")
end if%></font>元</td>
  </tr>
  
  <tr>
    <td height=23 colspan="3" class="tdbg">&nbsp;</td>
  </tr>
</table>
<table cellpadding="2" cellspacing="1" border="0" width="100%" class="border" align=center>
  <tr align="center"> 
    <td height=25 colspan=2 class="topbg"><strong>管 理 帮 助</strong></td> 
  <tr> 
    <td width="50%"  class="tdbg" height=23>类型：本程序专为淘宝免费互刷打造，是目前市面上功能比较齐全的平台</td>
    <td width="50%" class="tdbg">产品负责：网站事业部  </td>
  </tr>
  <tr> 
    <td width="50%" class="tdbg" height=23>产品开发：<a href="http://www.778892.com" target="_blank"><a href="mailto:wangxiaoyu-4@163.com">七七</a>网信息技术有限公司</a></td>
    <td width="50%" class="tdbg">电话：</td>
  </tr>
  <tr> 
    <td width="50%" class="tdbg" height=23>技术支持：<a href="http://www.778892.com">七七</a>网</td>
    <td width="50%" class="tdbg">QQ：<a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=859680966&site=qq&menu=yes">859680966</a></td>
  </tr>
  
  <tr>
    <td height=23 colspan="2" class="tdbg">版权：版权属于七七信息技术有限公司，不得拷贝、修改，侵权必究。<a href="Admin_ServerInfo.asp"></a></td>
  </tr>
</table>

<table cellpadding="2" cellspacing="1" border="0" width="100%" class="border" align=center>
  <tr align="center"> 
    <td height=25 class="topbg">&nbsp;</td>
  </tr>
  <TR>
    <TD class="tdbg">    
</table>
<table cellpadding="2" cellspacing="1" border="0" width="100%" class="border" align=center>
  <tr align="center"> 
    <td height=25 class="topbg">&nbsp;</td>
  </tr>
  
</table>
<!--#include file="Admin_fooder.asp"-->
<%
conn.close
set conn=nothing%>
</body>
</html>
