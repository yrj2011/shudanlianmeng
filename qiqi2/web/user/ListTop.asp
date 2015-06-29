<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#INCLUDE FILE="../config.asp"-->
<!--#include file="checklogin.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD>
<title>会员中心-刷客排行榜-<%=stupname%> -<%=stuptitle%></title>
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<LINK href="../css/top.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2900.5659" name=GENERATOR></HEAD>
<BODY>
<!--#include file=../inc/jieducm_top.asp-->
<DIV id=box>
  <DIV class=clear></DIV>
<DIV id=contenter>
<DIV class=block>
<DIV class=block_top>任务发布排行榜</DIV>
<DIV class=block_contenter>
  <DIV class=grid_box>
<UL id=ax1>
<%  
Sql = "select top 10  username,count(*) as jieducm_num  from  "&jieducm&"_zhongxin   group by username order by jieducm_num desc"
Set rs = Server.CreateObject("Adodb.RecordSet")
rs.Open Sql,conn,1,1
IF not rs.Eof  Then
Do While Not rs.Eof
%>
<LI><SPAN><%=left(rs("username"),1)%>***<%=right(rs("username"),1)%></SPAN>次数：<%=rs("jieducm_num")%> </LI>
<%  
rs.MoveNext
Loop
End IF
 %>
  
 </UL>
</DIV>
<DIV class=block_bottom></DIV></DIV></DIV>
<DIV class=block>
<DIV class=block_top>任务接手排行榜</DIV>
<DIV class=block_contenter>
  <DIV class=grid_box>
<UL id=bx1>
<%  
Sql = "select top 10  username,count(*) as jieducm_num  from  "&jieducm&"_jieshou   group by username order by jieducm_num desc"
Set rs = Server.CreateObject("Adodb.RecordSet")
rs.Open Sql,conn,1,1
IF not rs.Eof  Then
Do While Not rs.Eof
%>
<LI><SPAN><%=left(rs("username"),1)%>***<%=right(rs("username"),1)%></SPAN>次数：<%=rs("jieducm_num")%> </LI>
<%  
rs.MoveNext
Loop
End IF
 %>
 </UL>
<UL id=bx2 style="DISPLAY: none"></UL></DIV>
<DIV class=block_bottom></DIV></DIV></DIV>
<DIV class=block>
<DIV class=block_top>个人消费排行榜</DIV>
<DIV class=block_contenter>
  <DIV class=grid_box>
<UL id=cx1>
<%  
Sql = "select top 10  username,sum(fabudian) as jieducm_num  from  "&jieducm&"_chongjilu   group by username order by jieducm_num desc"
Set rs = Server.CreateObject("Adodb.RecordSet")
rs.Open Sql,conn,1,1
IF not rs.Eof  Then
Do While Not rs.Eof
%>
<LI><SPAN><%=left(rs("username"),1)%>***<%=right(rs("username"),1)%></SPAN>发布点：<%=rs("jieducm_num")%> </LI>
<%  
rs.MoveNext
Loop
End IF
 %>
 </UL>
<UL id=cx2 style="DISPLAY: none"></UL></DIV>
<DIV class=block_bottom></DIV></DIV></DIV>
<DIV class=block>
<DIV class=block_top>网站积分排行榜</DIV>
<DIV class=block_contenter>
  <DIV class=grid_box>
<UL id=dx1>
<%  
Sql = "select top 10  username,jifei  from  "&jieducm&"_register  order by jifei desc"
Set rs = Server.CreateObject("Adodb.RecordSet")
rs.Open Sql,conn,1,1
IF not rs.Eof  Then
Do While Not rs.Eof
%>
<LI><SPAN><%=left(rs("username"),1)%>***<%=right(rs("username"),1)%></SPAN>积分：<%=rs("jifei")%> </LI>
<%  
rs.MoveNext
Loop
End IF
 %>
  </UL>
</DIV>
<DIV class=block_bottom></DIV></DIV></DIV>
<DIV class=block>
<DIV class=block_top>网站推广排行榜</DIV>
<DIV class=block_contenter>
  <DIV class=grid_box>
<UL id=ex1>
<%  
Sql = "select top 10  tjr,count(*) as jieducm_num  from  "&jieducm&"_register where tjr<>''  group by tjr order by jieducm_num desc"
Set rs = Server.CreateObject("Adodb.RecordSet")
rs.Open Sql,conn,1,1
IF not rs.Eof  Then
Do While Not rs.Eof
%>
<LI><SPAN><%=left(rs("tjr"),1)%>***<%=right(rs("tjr"),1)%></SPAN>推广人数：<%=rs("jieducm_num")%> </LI>
<%  
rs.MoveNext
Loop
End IF
 %>
  </UL>
</DIV>
<DIV class=block_bottom></DIV></DIV></DIV>
<DIV class=block>
<DIV class=block_top>充值排行榜</DIV>
<DIV class=block_contenter>
  <DIV class=grid_box>
<UL id=fx1>

<%  
Sql = "select top 10  username,sum(cunkuan) as jieducm_num  from  "&jieducm&"_chongjilu where class='充值卡充值' or class='网银充值' or class='支付宝充值'  or class='易宝充值' group by username order by jieducm_num desc"
Set rs = Server.CreateObject("Adodb.RecordSet")
rs.Open Sql,conn,1,1
IF not rs.Eof  Then
Do While Not rs.Eof
%>
<LI><SPAN><%=left(rs("username"),1)%>***<%=right(rs("username"),1)%></SPAN>金额：￥<%=rs("jieducm_num")%> </LI>
<%  
rs.MoveNext
Loop
End IF
 %>
 </UL>
</DIV>
<DIV class=block_bottom></DIV></DIV></DIV>
</DIV><!-- 页脚 -->
</DIV>
 <!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
  <%rs.close
  set rs=nothing
  call closeconn()%>
</BODY></HTML>
