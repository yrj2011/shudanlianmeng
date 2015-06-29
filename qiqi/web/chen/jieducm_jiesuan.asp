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
admincunkuan=0
adminfabudain=0
adminjifei=0
adminchina=0
adminalipay1=0
adminbuyfb=0
adminbuycar=0
sql="SELECT  sum(cunkuan) as admincunkuan FROM "&jieducm&"_chongjilu where class='后台充值' and (datediff(day,[now],getdate())=0)"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
admincunkuan=rs("admincunkuan")
end if 

sql="SELECT  sum(fabudian) as adminfabudian FROM "&jieducm&"_chongjilu where class='后台充值' and (datediff(day,[now],getdate())=0)"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
adminfabudian=rs("adminfabudian")
end if 

sql="SELECT  sum(jifei) as adminjifei FROM "&jieducm&"_chongjilu where class='后台充值' and (datediff(day,[now],getdate())=0)"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
adminjifei=rs("adminjifei")
end if 

sql="SELECT  sum(cunkuan) as adminchina FROM "&jieducm&"_chongjilu where class='网银充值' and (datediff(day,[now],getdate())=0)"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
adminchina=rs("adminchina")
end if 

sql="SELECT  sum(cunkuan) as adminebao FROM "&jieducm&"_chongjilu where class='易宝充值' and (datediff(day,[now],getdate())=0)"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
adminebao=rs("adminebao")
end if 

sql="SELECT  sum(cunkuan) as adminalipay FROM "&jieducm&"_chongjilu where class='支付宝充值' and (datediff(day,[now],getdate())=0)"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
adminalipay1=rs("adminalipay")
end if 

sql="SELECT  sum(fabudian) as adminbuyfb FROM "&jieducm&"_chongjilu where class='购买发布点' and (datediff(day,[now],getdate())=0)"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
adminbuyfb=rs("adminbuyfb")
end if
 
sql="SELECT  sum(cunkuan) as adminbuycar FROM "&jieducm&"_chongjilu where class='购买点卡' and (datediff(day,[now],getdate())=0)"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
jieducm_adminbuycar=rs("adminbuycar")
end if 

sql="SELECT  sum(cunkuan) as adminbuycar2 FROM "&jieducm&"_chongjilu where class='充值卡充值' and (datediff(day,[now],getdate())=0)"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
jieducm_adminbuycar2=rs("adminbuycar2")
end if 


sql="SELECT  count(*) as dateren FROM "&jieducm&"_zhongxin where (datediff(day,[now],getdate())=0) and start='5'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
daterenxin=rs("dateren")
end if 


sql="SELECT  sum(cunkuan) as totalcunkuan FROM "&jieducm&"_chongjilu where class='后台充值' "
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
totalcunkuan=rs("totalcunkuan")
end if 

sql="SELECT  sum(fabudian) as totalfabudian FROM "&jieducm&"_chongjilu where class='后台充值' "
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
totalfabudian=rs("totalfabudian")
end if 

sql="SELECT  sum(jifei) as totaljifei FROM "&jieducm&"_chongjilu where class='后台充值' "
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
totaljifei=rs("totaljifei")
end if 

sql="SELECT  sum(cunkuan) as totalchong FROM "&jieducm&"_chongjilu where class='充值卡充值' "
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
totalchongzhi=rs("totalchong")
end if 


sql="SELECT  sum(cunkuan) as totalchina FROM "&jieducm&"_chongjilu where class='网银充值' "
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
totalchina=rs("totalchina")
end if 

sql="SELECT  sum(cunkuan) as totalebao FROM "&jieducm&"_chongjilu where class='易宝充值' "
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
totalebao=rs("totalebao")
end if 


sql="SELECT  sum(cunkuan) as totalalipay FROM "&jieducm&"_chongjilu where class='支付宝充值' "
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
totalalipay=rs("totalalipay")
end if 

sql="SELECT  sum(fabudian) as totalbuyfb FROM "&jieducm&"_chongjilu where class='购买发布点'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
totalbuyfb=rs("totalbuyfb")
end if

sql="SELECT  sum(cunkuan) as totalbuycar FROM "&jieducm&"_chongjilu where class='购买点卡' "
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1

if not rs.eof then			
jieducm_totalbuycar=rs("totalbuycar")
end if

sql="SELECT  sum(cunkuan) as totalusermoney FROM "&jieducm&"_register "
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
totalusermoney=rs("totalusermoney")
end if

sql="SELECT  sum(fabudian) as totaluserfabudian FROM "&jieducm&"_register "
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
totaluserfabudian=rs("totaluserfabudian")
end if


sql="SELECT  sum(price) as totalprice FROM "&jieducm&"_zhongxin where start<>'5' "
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
totalprice =rs("totalprice")
end if

sql="SELECT sum(ReNum) as totaltixian FROM "&jieducm&"_tixian  where start='1'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then			
totaltixian =rs("totaltixian")
end if

sql="SELECT  sum(Money) as jieducm_alipay FROM AutoCZ where  Status=1 and (datediff(day,[EndDate],getdate())=0)"
Set rs = Server.CreateObject("Adodb.RecordSet")
rs.Open Sql,conn,1,1
if not rs.eof then			
jieducm_alipay_total=rs("jieducm_alipay")
end if

sql="SELECT  sum(Money) as jieducm_alipay_total FROM AutoCZ where   Status=1 "
Set rs = Server.CreateObject("Adodb.RecordSet")
rs.Open Sql,conn,1,1
if not rs.eof then			
jieducm_alipay_totalok=rs("jieducm_alipay_total")
end if

rs.close
set rs=nothing%>              
 </p>
  <p>&nbsp;</p>
  <TABLE cellSpacing=1 cellPadding=1 width="451" align=center bgColor=#c2ddef 
border=0>
  <TBODY>
  <TR>
    <TD class=t1 vAlign=center align=middle width=130 bgColor=#ffffff 
      height=30><div align="right">今日后台充值:</div></TD>
    <TD width="314" height=30 align=left vAlign=center bgColor=#ffffff>
      <div align="left">存款：<%=admincunkuan%>，发布点：<%=adminfabudian%>积分：<%=adminjifei%></div></TD></TR>
  <TR>
    <TD class=t1 vAlign=center align=middle width=130 bgColor=#f6fcff 
      height=30><div align="right">今日网银充值:</div></TD>
    <TD vAlign=center align=left bgColor=#f6fcff height=30><%=adminchina%>元</TD></TR>
  <TR>
    <TD class=t1 vAlign=center align=middle width=130 bgColor=#ffffff 
      height=30><div align="right">今日支付宝充值:</div></TD>
    <TD vAlign=center align=left bgColor=#ffffff height=30><%=adminalipay1%>元</TD></TR>
  <TR>
    <TD class=t1 vAlign=center align=middle bgColor=#f6fcff 
      height=30><div align="right">今日支付宝自动充值:</div></TD>
    <TD vAlign=center align=left bgColor=#f6fcff height=30><%=jieducm_alipay_total%>元</TD>
  </TR>
  <TR>
    <TD class=t1 vAlign=center align=middle bgColor=#f6fcff 
      height=30><div align="right">今日易宝充值:</div></TD>
    <TD vAlign=center align=left bgColor=#f6fcff height=30><div align="left"><%=adminebao%>元</div></TD>
  </TR>
  <TR>
    <TD class=t1 vAlign=center align=middle width=130 bgColor=#f6fcff 
      height=30><div align="right">今日卖出发布点:</div></TD>
    <TD vAlign=center align=left bgColor=#f6fcff height=30><%=adminbuyfb%>个</TD></TR>
  <TR>
    <TD class=t1 vAlign=center align=middle bgColor=#ffffff 
      height=30><div align="right">今日卖出点卡:</div></TD>
    <TD vAlign=center align=left bgColor=#ffffff height=30><%=jieducm_adminbuycar%>元</TD>
  </TR>
  <TR>
    <TD class=t1 vAlign=center align=middle width=130 bgColor=#ffffff 
      height=30><div align="right">今日充值卡充值:</div></TD>
    <TD vAlign=center align=left bgColor=#ffffff height=30><%=jieducm_adminbuycar2%>元</TD></TR>
  <TR style="mso-yfti-irow: 6">
    <TD class=t1 vAlign=center align=middle width=130 bgColor=#f6fcff 
      height=30><div align="right">今日已完成好评:</div></TD>
    <TD vAlign=center align=left bgColor=#f6fcff height=30><%=daterenxin%>个</TD></TR>
  <TR style="mso-yfti-irow: 7">
    <TD class=t1 vAlign=center align=middle bgColor=#ffffff 
      height=30><div align="right"></div></TD>
    <TD vAlign=center align=left bgColor=#ffffff height=30>&nbsp;</TD>
  </TR>
  <TR style="mso-yfti-irow: 7">
    <TD class=t1 vAlign=center align=middle bgColor=#ffffff 
      height=30><div align="right">后台充值总额:</div></TD>
    <TD vAlign=center align=left bgColor=#ffffff height=30>存款:<%=totalcunkuan%> 发布点：<%=totalfabudian%> 积分：<%=totaljifei%></TD>
  </TR>
  <TR style="mso-yfti-irow: 7">
    <TD class=t1 vAlign=center align=middle bgColor=#ffffff 
      height=30><div align="right">网银充值总额:</div></TD>
    <TD vAlign=center align=left bgColor=#ffffff height=30><%=totalchina%>元</TD>
  </TR>
  <TR style="mso-yfti-irow: 7">
    <TD class=t1 vAlign=center align=middle bgColor=#ffffff 
      height=30><div align="right">支付宝充值总额:</div></TD>
    <TD vAlign=center align=left bgColor=#ffffff height=30><%=totalalipay%>元</TD>
  </TR>
  <TR style="mso-yfti-irow: 7">
    <TD class=t1 vAlign=center align=middle bgColor=#ffffff 
      height=30><div align="right">支付宝自动充值总额:</div></TD>
    <TD vAlign=center align=left bgColor=#ffffff height=30><%=jieducm_alipay_totalok%>元</TD>
  </TR>
  <TR style="mso-yfti-irow: 7">
    <TD class=t1 vAlign=center align=middle bgColor=#ffffff 
      height=30><div align="right">易宝充值总额:</div></TD>
    <TD vAlign=center align=left bgColor=#ffffff height=30><%=totalebao%>元</TD>
  </TR>
  <TR style="mso-yfti-irow: 7">
    <TD class=t1 vAlign=center align=middle bgColor=#ffffff 
      height=30><div align="right">充值卡充值总额：</div></TD>
    <TD vAlign=center align=left bgColor=#ffffff height=30><%=totalchongzhi%>元</TD>
  </TR>
  <TR style="mso-yfti-irow: 7">
    <TD class=t1 vAlign=center align=middle bgColor=#ffffff 
      height=30><div align="right">卖出发布点总额:</div></TD>
    <TD vAlign=center align=left bgColor=#ffffff height=30><%=totalbuyfb%>个</TD>
  </TR>
  <TR style="mso-yfti-irow: 7">
    <TD class=t1 vAlign=center align=middle bgColor=#ffffff 
      height=30><div align="right">卖出点卡总额:</div></TD>
    <TD vAlign=center align=left bgColor=#ffffff height=30>
	<%if not isnull(jieducm_totalbuycar) then
	 response.Write(formatnumber((jieducm_totalbuycar),2))
	 else
	 response.Write("0")
	end if%>
	元</TD>
  </TR>
  <TR style="mso-yfti-irow: 7">
    <TD class=t1 vAlign=center align=middle bgColor=#ffffff 
      height=30><div align="right">平台会员账户总额:</div></TD>
    <TD vAlign=center align=left bgColor=#ffffff height=30><div align="left">
	<%if not isnull(totalusermoney) then
	 response.Write(formatnumber((totalusermoney),2))
	else
	response.Write("0")
	end if%>
	元  
	发布点：
	<%if not isnull(totaluserfabudian) then
	 response.Write(formatnumber((totaluserfabudian),2))
	else
	response.Write("0")
	end if%>
</div></TD>
  </TR>
  <TR style="mso-yfti-irow: 7">
    <TD class=t1 vAlign=center align=middle width=130 bgColor=#ffffff 
      height=30><div align="right">进行中的任务总额:</div></TD>
    <TD vAlign=center align=left bgColor=#ffffff height=30>
	<%if not isnull(totalprice) then
	 response.Write(formatnumber((totalprice),2))
	else
	response.Write("0")
	end if%>
	元</TD></TR>
  <TR style="mso-yfti-irow: 8">
    <TD class=t1 vAlign=center align=middle width=130 bgColor=#f6fcff 
      height=30><div align="right">提现完成总额:</div></TD>
    <TD vAlign=center align=left bgColor=#f6fcff height=30>
	<%if not isnull(totaltixian) then
	 response.Write(formatnumber((totaltixian),2))
	else
	response.Write("0")
	end if%>
	元</TD></TR></TBODY></TABLE>
</div>
</body>
</html>
<%
conn.close
set conn=nothing
%>