<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#INCLUDE FILE="../inc/function.asp"-->
<!--#INCLUDE FILE="../config.asp"-->
<!--#include file="../user/checklogin.asp"-->
<!--#INCLUDE FILE="../background.asp"-->

<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD>
<title><%=stupname%> -<%=stuptitle%></title>
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<META content="MSHTML 6.00.2900.2180" name=GENERATOR>
<style type="text/css">
<!--
.STYLE1 {color: #FF0033}
.STYLE2 {color: #FF0000}
.STYLE3 {
	color: #FF3300;
	font-size: 14px;
}
.STYLE7 {
	color: #3366FF;
	font-size: 13px;
}
.STYLE11 {
	font-size: 18px;
	font-weight: bold;
}
.STYLE12 {color: #FF0066}
.STYLE13 {color: #FF3300}
.STYLE14 {color: #FF3333}
.STYLE15 {color: #000000}
.STYLE16 {color: #FF6633}
-->
</style>
</HEAD>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!--#include file=../inc/jieducm_top.asp-->
<%
id=request("id")
sql="select * from "&jieducm&"_tribe where id="&id&""
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then
username=rs("username")
tribename2=rs("tribename")
pic=rs("pic")
tribeinfo=rs("tribeinfo")
end if
%>

<DIV style="CLEAR: both; OVERFLOW: hidden; WIDTH: 950px; BACKGROUND-COLOR: #f3f8fe">
<DIV style="CLEAR: both; MARGIN: 20px; LINE-HEIGHT: 160%">&nbsp;<SPAN style="FONT-WEIGHT: bolder; FONT-SIZE: 18px; COLOR: red"><a href="../tribe"><span class="STYLE12">刷</span><span class="STYLE1">客</span><span class="STYLE13">部</span><span class="STYLE14">落</span></a><img src="../images/jiaodian_biao.gif" width="21" height="17"></SPAN>
<SPAN style="FONT-WEIGHT: bolder; FONT-SIZE: 18px; COLOR: red">【<%=tribename2%>】</SPAN></DIV>
</DIV>
<DIV style="CLEAR: both; BORDER-RIGHT: #abbec8 1px solid; BORDER-TOP: #abbec8 1px solid; MARGIN-TOP: 5px; BORDER-LEFT: #abbec8 1px solid; WIDTH: 948px; BORDER-BOTTOM: #abbec8 1px solid; BACKGROUND-COLOR: #f3f8fe; TEXT-ALIGN: center">
<DIV style="MARGIN: 20px; WIDTH: 900px">
<DIV style="FONT-WEIGHT: bolder; FONT-SIZE: 16px; FLOAT: left; WIDTH: 200px; LINE-HEIGHT: 150%; TEXT-ALIGN: center">
  <p><IMG height=120 src="../<%=pic%>" width=160><BR>
    <span class="STYLE2"><img src="../admin/images/Soft_ontop.gif" width="9" height="15"><span class="STYLE11">部落徽章</span></span><img src="../images/Soft_ontop.gif" width="9" height="15"><br>
   
  </p>
  </DIV>
<DIV 
style="FONT-WEIGHT: bolder; FONT-SIZE: 16px; FLOAT: left; WIDTH: 240px; LINE-HEIGHT: 150%; TEXT-ALIGN: left">部落：【<%=tribename2%>】<BR>酋长：<%=username%>&nbsp;
<%
Sql = "select qq,jifei from "&jieducm&"_register where username='"&username&"'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not rs.eof then
 qq=rs("qq")
 jifei=rs("jifei")		
end if
call xinyu(jifei)
%>
<BR>繁华度：<%call tribecount(id)%><BR>
<a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=<%=qq%>&site=qq&menu=yes"><img src="http://wpa.qq.com/pa?p=1:<%=qq%>:4"  border="0"  /> 联系酋长获取部落激活码</A><BR>
<a href="tribeinfo.asp?yu=ok&qid=<%=id%>"></a> <BR>
</DIV>
<DIV 
style="FLOAT: left; WIDTH: 450px; LINE-HEIGHT: 150%; TEXT-ALIGN: left"><FONT 
style="FONT-WEIGHT: bolder; FONT-SIZE: 16px">部落介绍：</FONT><BR><%=tribeinfo%> </DIV></DIV></DIV>

<DIV style="CLEAR: both; BORDER-RIGHT: #abbec8 1px solid; BORDER-TOP: #abbec8 1px solid; MARGIN-TOP: 5px; BORDER-LEFT: #abbec8 1px solid; WIDTH: 948px; BORDER-BOTTOM: #abbec8 1px solid; BACKGROUND-COLOR: #f3f8fe; TEXT-ALIGN: center">
<DIV 
style="CLEAR: both; FONT-WEIGHT: bolder; FONT-SIZE: 16px; WIDTH: 948px; LINE-HEIGHT: 150%; BORDER-BOTTOM: #abbec8 1px solid">【<%=tribename2%>】 
部落交流 </DIV>
  <%
	Sql = "select * from "&jieducm&"_tribebook where tribeid="&id&" order by id desc"			  
	Set Rs = Server.CreateObject("Adodb.RecordSet")
	Rs.Open Sql,conn,1,1
	if not rs.eof then					
	dim maxperpage,url,PageNo
	 url="TribeInfo.asp?id="&id&""
	 rs.pagesize=15
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
	
	if PageNo=1 then i=rs.RECORDCOUNT else i=rs.RECORDCOUNT-(30*(PageNo-1))
	 DO WHILE NOT rs.EOF AND RowCount>0
	 
	 %>
<DIV style="MARGIN-TOP: 5px; WIDTH: 98%">
  <DIV style="CLEAR: both; BORDER-RIGHT: #cccccc 1px dashed; PADDING-RIGHT: 15px; BORDER-TOP: #cccccc 1px dashed; PADDING-LEFT: 15px; MARGIN-BOTTOM: 10px; PADDING-BOTTOM: 15px; BORDER-LEFT: #cccccc 1px dashed; WIDTH: 90%; PADDING-TOP: 15px; BORDER-BOTTOM: #cccccc 1px dashed">
<DIV style="CLEAR: both">
<DIV 
style="FLOAT: left; LINE-HEIGHT: 150%">&nbsp;
<%
Sql1 = "select * from "&jieducm&"_register where username='"&rs("username")&"'"
Set Rs1 = Server.CreateObject("Adodb.RecordSet")
Rs1.Open Sql1,conn,1,1
if not rs1.eof then
jifei=rs1("jifei") 
qq=rs1("qq")
vip=rs1("vip")

end if
%>
<a href="../user/send_message.asp?sendname=<%=rs("username")%>" title="发送站内信息" target="_blank"><%=rs("username")%></a>
<%if vip="1" then %><img src="../images/jieducm_vip.gif"  alt="尊贵VIP,信誉值：<%call vipxinyu(rs("username"))%>"  border="0"/><% end if%>
积分：<%
response.Write(jifei)
response.Write("分")
call xinyu(jifei)%>
<a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=<%=qq%>&site=qq&menu=yes"><img src="http://wpa.qq.com/pa?p=1:<%=qq%>:4"  border="0" alt="<%=qq%>" /> </a>
&nbsp;&nbsp;(部落：【<%=tribename2%>】)</DIV>
<DIV style="CLEAR: right; FLOAT: right">  留言发表于:<%=rs("now")%></DIV></DIV>
<DIV 
style="CLEAR: both; BORDER-TOP: #cccccc 1px dashed; FONT-SIZE: 15px; COLOR: #666666; LINE-HEIGHT: 200%; PADDING-TOP: 5px; TEXT-ALIGN: left">
<P><%=rs("content")%></P></DIV>
<DIV 
style="CLEAR: both; BORDER-TOP: #cccccc 1px dashed; FONT-SIZE: 12px; COLOR: #666666; LINE-HEIGHT: 200%; PADDING-TOP: 5px; TEXT-ALIGN: right"><% if username=session("useridname") then%><a href="bookadd.asp?id=<%=rs("id")%>&quid=<%=id%>&action=del">删除</a> <% end if%></DIV></DIV>
</DIV>
   <%
RowCount = RowCount - 1
	  i=i-1
      rs.MoveNext 
      Loop  %>

<DIV 
style="CLEAR: both; MARGIN-TOP: 5px; WIDTH: 948px">&nbsp;&nbsp;
<% call showpage(url,rs.RECORDCOUNT,maxperpage,false,true,"条主题") %></DIV></DIV>

<DIV 
style="CLEAR: both; BORDER-RIGHT: #abbec8 1px solid; BORDER-TOP: #abbec8 1px solid; MARGIN-TOP: 5px; BORDER-LEFT: #abbec8 1px solid; WIDTH: 948px; BORDER-BOTTOM: #abbec8 1px solid; BACKGROUND-COLOR: #f3f8fe; TEXT-ALIGN: center">
<DIV 
style="CLEAR: both; FONT-WEIGHT: bolder; FONT-SIZE: 16px; WIDTH: 948px; LINE-HEIGHT: 150%; BORDER-BOTTOM: #abbec8 1px solid"><STRONG>签 
写 留 言 ：</STRONG></DIV>
  
<DIV style="MARGIN-TOP: 5px; WIDTH: 98%">
  <DIV style="CLEAR: both; BORDER-RIGHT: #cccccc 1px dashed; PADDING-RIGHT: 15px; BORDER-TOP: #cccccc 1px dashed; PADDING-LEFT: 15px; MARGIN-BOTTOM: 10px; PADDING-BOTTOM: 15px; BORDER-LEFT: #cccccc 1px dashed; WIDTH: 90%; PADDING-TOP: 15px; BORDER-BOTTOM: #cccccc 1px dashed">
<DIV style="CLEAR: both">
<DIV 
style="FLOAT: left; LINE-HEIGHT: 150%">&nbsp;留言内容：</DIV>
<DIV style="CLEAR: right; FLOAT: right"></DIV>
</DIV>
<FORM id=aspnetForm name=aspnetForm action="bookadd.asp" method=post >
  <input type="hidden" name="id" id="id" value="<%=id%>">
  <input name="action" type="hidden" value="add">
<DIV style="CLEAR: both; BORDER-TOP: #cccccc 1px dashed; FONT-SIZE: 15px; COLOR: #666666; LINE-HEIGHT: 200%; PADDING-TOP: 5px; TEXT-ALIGN: left">
<P><textarea name="content" style="display:none"></textarea> 

<iframe id="Editor" name="Editor" src="../HtmlEditor/index.html?ID=content" frameborder="0" marginheight="0"     marginwidth="0" scrolling="No" style="height:320px;width:100%"></iframe> </P></DIV>

<DIV 
style="CLEAR: both; BORDER-TOP: #cccccc 1px dashed; FONT-SIZE: 12px; COLOR: #666666; LINE-HEIGHT: 200%; PADDING-TOP: 5px; TEXT-ALIGN: center"><INPUT id=Button1 type=submit value=" 签 写 留 言 " name=Button1></DIV></FORM></DIV>
</DIV>


<DIV style="CLEAR: both; MARGIN-TOP: 5px; WIDTH: 948px">&nbsp;&nbsp;
</DIV></DIV>
<!--#INCLUDE FILE=../inc/jieducm_bottom.asp-->
<%
rs.close()
set rs=nothing
call closeconn()%>
</BODY></HTML>
