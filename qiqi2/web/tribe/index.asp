<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#INCLUDE FILE="../config.asp"-->
<!--#include file="../user/checklogin.asp"-->
<!--#INCLUDE FILE="../background.asp"-->

<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD>
<title><%=stupname%> -<%=stuptitle%></title>
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>

<META content="MSHTML 6.00.2900.2180" name=GENERATOR></HEAD>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!--#include file=../inc/jieducm_top.asp-->

<DIV 
style="CLEAR: both; OVERFLOW: hidden; WIDTH: 906px; BACKGROUND-COLOR: #f3f8fe">

  <DIV style="CLEAR: both; MARGIN: 20px; LINE-HEIGHT: 160%">&nbsp;<SPAN 
style="FONT-WEIGHT: bolder; FONT-SIZE: 18px; COLOR: red">刷客部落</SPAN><BR>
  <A href="#" ><SPAN style="FONT-WEIGHT: bolder; FONT-SIZE: 14px; COLOR: red"></SPAN></A><BR>
<form id="form1" name="form1" method="post" action="tribeok.asp">
<input name="action" type="hidden" value="del">
<DIV id="ctl00_AllBaseContentPlaceHolder_hidden3">  
<%
sql="select * from "&jieducm&"_tribe where id="&tribeid&""
set rs=Server.CreateObject("Adodb.RecordSet")
rs.open sql,conn,1,1
if not rs.eof then
tribename2=rs("tribename")
 %>
 <input name="tribeidid" type="hidden" value="<%=rs("id")%>">
<SPAN style="FONT-WEIGHT: bolder; FONT-SIZE: 14px; COLOR: red">* 你已加入的部落是   <a href="TribeInfo.asp?id=<%=tribeid%>">【<%=tribename2%>】</a></SPAN>
 <INPUT id="Button2"  type="submit" value="我要退出此部落" name="Button2" onClick="return confirm('确定退出此部落吗？');">
<% end if%>
</DIV></form>

<% action=request("action")
if action="ok" then
tid=request("tid")

sql = "select * from "&jieducm&"_tribe where username='"&session("useridname")&"'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then
	Response.Write("<script>alert('你已经是酋长了!不能再加入部落！');window.location.href='index.asp';</script>")
    response.End()			
end if

if tribeid<>0 then
	Response.Write("<script>alert('你已经加入部落了!不能再加入其它部落！');window.location.href='index.asp';</script>")
    response.End()	
end if

sql = "select * from "&jieducm&"_tribe where id="&tid&""
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then
	tribename=rs("tribename")	
	bulouid=rs("id")	
end if

%>
<form id="form1" name="form1" method="post" action="tribeok.asp">
<input name="action" type="hidden" value="add">
<div id="ctl00_AllBaseContentPlaceHolder_hidden2">
<input name="bulouid" type="hidden" value="<%=bulouid%>">
<span style=" font-size:18px; color:Red; font-weight:bolder;">
* 你已申请加入的部落是：<%=tribename%><br />
请输入激活码：<input name="tribecode" type="text" id="tribecode" /> 
<input  type="submit" name="Button1" value="确认"  id="Button1" /> 
该部落的激活码是:778892 <a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=<%=qq%>><%=rs3("tribename")%>&site=qq&menu=yes"></A></span></div>
 </form>
 <% end if%>
  </DIV>


</DIV>

<DIV style="CLEAR: both; BORDER-RIGHT: #abbec8 1px solid; BORDER-TOP: #abbec8 1px solid; MARGIN-TOP: 5px; BORDER-LEFT: #abbec8 1px solid; WIDTH: 906px;height:906px; BORDER-BOTTOM: #abbec8 1px solid; BACKGROUND-COLOR: #f3f8fe; TEXT-ALIGN: center">
<DIV style="WIDTH: 100%; PADDING-TOP: 20px">
   <%
	Sql = "select * from "&jieducm&"_tribe order by id asc"		  
	Set Rs = Server.CreateObject("Adodb.RecordSet")
	Rs.Open Sql,conn,1,1
	if not rs.eof then
	dim maxperpage,url,PageNo
	 url="index.asp"
	 rs.pagesize=20
	 PageNo=REQUEST("PageNo")
	 if PageNo="" or PageNo=0 then PageNo=1
	 RS.AbsolutePage=PageNo
	 TSum=rs.pagecount
	 maxperpage=rs.pagesize
	 RowCount=rs.PageSize
	   PageNo=PageNo+1
	   PageNo=PageNo-1
	 if CINT(PageNo)>1 then
	    if CINT(PageNo)>CINT(TSum) then
		  response.Write("对不起没有您想要的页数")
          Response.End
	    end if
	 end if
		end if    
     if PageNo<0 then
	    response.Write("没有这一页!")
		Response.End
	 End if
	j=1
	if PageNo=1 then i=rs.RECORDCOUNT else i=rs.RECORDCOUNT-(30*(PageNo-1))
	 DO WHILE NOT rs.EOF AND RowCount>0%>
	 <form id="form<%=j%>" name="form<%=j%>" method="post" action="index.asp?action=ok">
	 <input name="tid" type="hidden" value="<%=rs("id")%>">
  <DIV style="BORDER-RIGHT: #abbec8 1px solid; BORDER-TOP: #abbec8 1px solid; FLOAT: left; MARGIN: 6px 6px 10px; BORDER-LEFT: #abbec8 1px solid; WIDTH: 22%; BORDER-BOTTOM: #abbec8 1px solid; HEIGHT: 340px">
<DIV class=PicTitle style="MARGIN-TOP: 20px">
<A href="TribeInfo.asp?id=<%=rs("id")%>"><IMG height=120 src="../<%=rs("pic")%>" width=160></A>
</DIV>
<DIV style="MARGIN-TOP: 5px; HEIGHT: 30px; TEXT-ALIGN: left">
<DIV style="FLOAT: left; WIDTH: 32%; TEXT-ALIGN: right">部落：</DIV>
<DIV style="FLOAT: left; WIDTH: 68%"><%=rs("tribename")%></DIV></DIV>
<DIV style="HEIGHT: 30px; TEXT-ALIGN: left">
<DIV style="FLOAT: left; WIDTH: 32%; TEXT-ALIGN: right">酋长：</DIV>
<DIV style="FLOAT: left; WIDTH: 68%"><%=rs("username")%>
<%
Sql = "select qq,jifei from "&jieducm&"_register where username='"&rs("username")&"'"
Set Rs1 = Server.CreateObject("Adodb.RecordSet")
Rs1.Open Sql,conn,1,1
if not rs1.eof then
 sjifei=rs1("jifei")
 qq=rs1("qq")
end if
call xinyu(sjifei)
%>

</DIV>
</DIV>
<DIV style="HEIGHT: 30px; TEXT-ALIGN: left">
<input type="hidden" name="quiid" id="quiid" value="<%=rs("id")%>">
<DIV style="FLOAT: left; WIDTH: 32%; TEXT-ALIGN: right">繁华度：</DIV>
<DIV style="FLOAT: left; WIDTH: 68%"><%call tribecount(rs("id"))%> <A href="TribeInfo.asp?id=<%=rs("id")%>">查看部落-&gt;&gt;</A></DIV></DIV>
<DIV style="HEIGHT: 50px; TEXT-ALIGN: left">

<DIV style="FLOAT: center; WIDTH: 100%; TEXT-ALIGN: center">
联系酋长<a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=<%=qq%>&site=qq&menu=yes"><img src="http://wpa.qq.com/pa?p=1:<%=qq%>:41"  border="0"  /> </a>
</DIV></br >
<DIV style="FLOAT: center;color:red;WIDTH: 80%">部落激活码：778892</DIV></DIV>
</br >
<DIV style="HEIGHT: 30px; TEXT-ALIGN: center">
<label>
<input type="submit" name="Submit" value="加入此部落" onClick="return confirm('确定要加入此部落吗？');">
</label>
</DIV>
  </DIV>  </FORM>
 <%
	RowCount = RowCount - 1
	  i=i-1
	  j=j+1
      rs.MoveNext 
      Loop
 %>
</DIV>
</DIV>
<DIV 
style="WIDTH: 98%; LINE-HEIGHT: 40px; PADDING-TOP: 10px; HEIGHT: 40px; TEXT-ALIGN: center">
 <% call showpage(url,rs.RECORDCOUNT,maxperpage,false,true,"条主题") %></DIV>
<!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
<%call closeconn()%>
</BODY></HTML>
