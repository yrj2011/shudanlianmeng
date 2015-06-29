<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/function.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<LINK 
href="../images/mission.css" type=text/css rel=stylesheet>
<LINK 
href="../images/Default.css" 
type=text/css rel=stylesheet>
<LINK href="../images/Bbs.css" type=text/css rel=stylesheet>
<LINK href="../images/style.css" type=text/css rel=stylesheet>
<title>七七传媒</title>
</head>

<body>
<div align="center"><p>
<%	
username=request("username")

sql="SELECT  sum(cunkuan) as admincunkuan FROM "&jieducm&"_chongjilu where class='后台充值' and username='"&username&"'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
admincunkuan=rs("admincunkuan")
end if


sql="SELECT  sum(fabudian) as adminfabudian FROM "&jieducm&"_chongjilu where class='后台充值' and username='"&username&"'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
adminfabudian=rs("adminfabudian")
end if

sql="SELECT  sum(price) as totalprice FROM "&jieducm&"_jieshou where username='"&username&"' and classid='1' and start='5'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
totalprice =rs("totalprice")
end if	

sql="SELECT  sum(price) as jieducm_totalprice FROM "&jieducm&"_jieshou where username='"&username&"' and classid='5' and start='5'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
jieducm_totalprice =rs("jieducm_totalprice")
end if	

sql="SELECT  sum(price) as jieducm_fanum FROM "&jieducm&"_zhongxin where username='"&username&"' and classid='1'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
jieducm_fanum =rs("jieducm_fanum")
end if

sql="SELECT  sum(price) as jieducm_remoney FROM "&jieducm&"_zhongxin where username='"&username&"' and classid='5'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
jieducm_remoney =rs("jieducm_remoney")
end if	

if not isnull(jieducm_remoney) then
jieducm_remoney=jieducm_remoney
else
jieducm_remoney=0
end if
	

sql="SELECT  sum(ReNum) as jieducm_reok FROM "&jieducm&"_tixian where username='"&username&"' and start<>'2'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
jieducm_reok =rs("jieducm_reok")
end if
if not isnull(jieducm_reok) then
jieducm_reok=jieducm_reok
else
jieducm_reok=0
end if

sql="SELECT  sum(fabudian) as jieducm_buyfb FROM "&jieducm&"_chongjilu where class='购买发布点' and username='"&username&"'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
jieducm_buyfb=rs("jieducm_buyfb")
end if


sql="SELECT  sum(cunkuan) as jieducm_buycunkuan FROM "&jieducm&"_chongjilu where class='购买点卡' and username='"&username&"'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
jieducm_buycunkuan=rs("jieducm_buycunkuan")
end if


sql="SELECT  sum(cunkuan) as jieducm_buyvip FROM "&jieducm&"_chongjilu where class='购买vip卡' and username='"&username&"'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
jieducm_buyvip=rs("jieducm_buyvip")
end if
%>              
 </p>
  <p>&nbsp;</p>
  <TABLE cellSpacing=1 cellPadding=1 width="451" align=center bgColor=#c2ddef 
border=0>
  <TBODY>
  
  <TR style="mso-yfti-irow: 7">
    <TD 
      height=30 colspan="2" align=middle vAlign=center bgColor=#ffffff class=t1><div align="center">会员：<%=username%>的资金数据</div></TD>
    </TR>
  <TR style="mso-yfti-irow: 7">
    <TD width="156" 
      height=30 align=middle vAlign=center bgColor=#ffffff class=t1><div align="right">后台充值总额:</div></TD>
    <TD width="288" height=30 align=left vAlign=center bgColor=#ffffff>&nbsp;<%=admincunkuan%>元  发布点：<%=adminfabudian%>个</TD>
  </TR>
  <TR style="mso-yfti-irow: 7">
    <TD class=t1 vAlign=center align=middle bgColor=#ffffff 
      height=30><div align="right">已接淘宝任务总额:</div></TD>
    <TD vAlign=center align=left bgColor=#ffffff height=30>&nbsp;<%=totalprice%>元</TD>
  </TR>
  <TR style="mso-yfti-irow: 7">
    <TD class=t1 vAlign=center align=middle bgColor=#ffffff 
      height=30><div align="right">已接充值提现任务总额:</div></TD>
    <TD vAlign=center align=left bgColor=#ffffff height=30>&nbsp;<%=jieducm_totalprice%>元</TD>
  </TR>
  
  <TR style="mso-yfti-irow: 7">
    <TD class=t1 vAlign=center align=middle bgColor=#ffffff 
      height=30><div align="right">已发任务总额：</div></TD>
    <TD vAlign=center align=left bgColor=#ffffff height=30>&nbsp;<%=jieducm_fanum%>元</TD>
  </TR>
  <TR style="mso-yfti-irow: 7">
    <TD class=t1 vAlign=center align=middle bgColor=#ffffff 
      height=30><div align="right">已提现任务总额:</div></TD>
    <TD vAlign=center align=left bgColor=#ffffff height=30>&nbsp;<%=jieducm_remoney+jieducm_reok%>元</TD>
  </TR>
  <TR style="mso-yfti-irow: 7">
    <TD class=t1 vAlign=center align=middle bgColor=#ffffff 
      height=30><div align="right">购买发布点总额:</div></TD>
    <TD vAlign=center align=left bgColor=#ffffff height=30>&nbsp;<%=jieducm_buyfb%>个</TD>
  </TR>
  <TR style="mso-yfti-irow: 7">
    <TD class=t1 vAlign=center align=middle bgColor=#ffffff 
      height=30><div align="right">购买点卡总额:</div></TD>
    <TD vAlign=center align=left bgColor=#ffffff height=30>&nbsp;<%=jieducm_buycunkuan%>元</TD>
  </TR>
  <TR style="mso-yfti-irow: 7">
    <TD class=t1 vAlign=center align=middle bgColor=#ffffff 
      height=30><div align="right">购买vip卡总额:</div></TD>
    <TD vAlign=center align=left bgColor=#ffffff height=30><div align="left">&nbsp;<%=jieducm_buyvip%>元  
	</div></TD>
  </TR></TBODY></TABLE>
</div>
</body>
</html>
<%
conn.close
set conn=nothing
%>