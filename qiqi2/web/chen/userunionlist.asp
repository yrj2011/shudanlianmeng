<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/function.asp"-->
<!--#INCLUDE FILE="../background.asp"-->
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
sub fareu(userna)
Sql1 = "select  count(*) as renwu  from "&jieducm&"_zhongxin where username='"&userna&"' and start='5'"
Set Rs1 = Server.CreateObject("Adodb.RecordSet")
Rs1.Open Sql1,conn,1,1
if NOT rs1.EOF  then
	renwu=rs1("renwu")
end if
response.Write(renwu)
rs1.close	
end sub	

sub jiereu(username)
Sql2 = "select  count(*) as jiewu  from "&jieducm&"_jieshou where username='"&username&"' and start='5' "
Set Rs1 = Server.CreateObject("Adodb.RecordSet")
Rs1.Open Sql2,conn,1,1
if NOT rs1.EOF  then
jiewu=rs1("jiewu")
end if
response.Write(jiewu)
rs1.close
end sub

sub vipcar(username)
Sql2 = "select  count(*) as vipcarnum  from "&jieducm&"_chongjilu where class='购买vip卡' and username='"&username&"' "
Set Rs1 = Server.CreateObject("Adodb.RecordSet")
Rs1.Open Sql2,conn,1,1
if NOT rs1.EOF  then
vipcarn=rs1("vipcarnum")
end if
response.Write(vipcarn)
rs1.close
end sub

sub diancar(username)
Sql2 = "select  count(*) as diancarnum  from "&jieducm&"_chongjilu where class='购买点卡' and username='"&username&"' "
Set Rs1 = Server.CreateObject("Adodb.RecordSet")
Rs1.Open Sql2,conn,1,1
if NOT rs1.EOF  then
diancarn=rs1("diancarnum")
end if
response.Write(diancarn)
rs1.close
end sub

sub diancarprice(username)
Sql2 = "select  sum(cunkuan) as totalcunkuan  from "&jieducm&"_chongjilu where class='购买点卡' and username='"&username&"' "
Set Rs1 = Server.CreateObject("Adodb.RecordSet")
Rs1.Open Sql2,conn,1,1
if NOT rs1.EOF  then
totalck=rs1("totalcunkuan")
end if
totalck=totalck*0.1
if isnull(totalck) then
response.Write("0")
else
response.Write(totalck)
end if
rs1.close
end sub

sub vipcarfb(username)
Sql2 = "select  count(*) as vipnumfb  from "&jieducm&"_chongjilu where class='购买vip卡' and username='"&username&"' "
Set Rs1 = Server.CreateObject("Adodb.RecordSet")
Rs1.Open Sql2,conn,1,1
if NOT rs1.EOF  then
vipcarfbd=rs1("vipnumfb")
end if
vipcarfbd=vipcarfbd*20
response.Write(vipcarfbd)
rs1.close
end sub

sub jifeitjr(username)
Sql2 = "select  count(*) as totalzx  from "&jieducm&"_zhongxin where start='5' and username='"&username&"' "
Set Rs1 = Server.CreateObject("Adodb.RecordSet")
Rs1.Open Sql2,conn,1,1
if NOT rs1.EOF  then
totalzxnum=rs1("totalzx")
end if
totalzxnum=totalzxnum*stupfajifei
rs1.close

Sql2 = "select  count(*) as totaljs  from "&jieducm&"_jieshou where start='5' and username='"&username&"' "
Set Rs1 = Server.CreateObject("Adodb.RecordSet")
Rs1.Open Sql2,conn,1,1
if NOT rs1.EOF  then
totaljsnum=rs1("totaljs")
end if
totaljsnum=totaljsnum*stupfajifei
rs1.close

jifeitotal=totaljsnum+totalzxnum+stuptjrzhu
response.Write(jifeitotal)
end sub

username=request("username")

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>发言排名</title>
<link rel="stylesheet" type="text/css" href="Admin_Style.css">
</head>

<body>
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center>
  <tr>
	<td colspan="2" align="center" class="title"><strong>惩罚管理</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td width="70" height="30" ><strong>管理导航：</strong></td>
    <td>&nbsp;</td>
  </tr>
</table>
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center><tr></tr>
</table>
<table width='100%' border="0" cellpadding="0" cellspacing="0" ><tr>
    <form name="myform" method="Post" action="chengfadel.asp" onSubmit="return ConfirmDel();">
      <table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center>
        <tr>
          <td colspan="2" align="left" class="title"></td>
        </tr>
        <tr valign="middle" class="tdbg">
          <td height="22"></td>
          <td width="200" height="22" align="right"></td>
        </tr>
      </table>
     <td><table class="border" border="0" cellspacing="1" width="100%" cellpadding="0">
          <tr class="title" height="22"> 
            <td width="104" height="22" align="center" ><strong>用户名</strong></td>
            <td width="83" align="center" ><strong>积分</strong></td>
            <td width="83" align="center" ><strong>发布点</strong></td>
            <td width="138" align="center" ><strong>存款</strong></td>
            <td width="232" align="center" ><strong>购买点卡</strong></td>
            <td width="178" align="center" ><strong>购买VIP卡</strong></td>
            <td width="178" align="center" ><strong>接手任务</strong></td>
            <td width="178" align="center" ><strong>发布任务</strong></td>
            <td width="94" align="center" ><strong>注册时间</strong></td>
            <td width="94" align="center" ><strong>QQ</strong></td>
            <!--<td width="40" align="center" ><strong>已审核</strong></td>-->
          </tr>
       
		   	<%	
	sql="SELECT * FROM "&jieducm&"_register where tjr='"&username&"' order by id desc"		
	Set Rs = Server.CreateObject("Adodb.RecordSet")
	Rs.Open Sql,conn,1,1
	if rs.eof then
	response.Write("<tr class=tdbg ><td  colspan=9 align=center>暂无记录</td></tr>")
	end if
	if not rs.eof then					
	dim maxperpage,url,PageNo
	 url="userunionlist.asp?username="&username&""
	 rs.pagesize=40
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
 DO WHILE NOT rs.EOF AND RowCount>0%>
		  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#cccccc'" style="padding: 0px 2px;">
            <td width="104" align="center"><%=rs("username")%></td>
            <td width="83" align="center"><%call jifeitjr(rs("username"))%></td>
            <td width="83" align="center"><%call vipcarfb(rs("username"))%></td>        
            <td width="138" align="center"><%call diancarprice(rs("username"))%> </td>
            <td width="232" align="center"><%call diancar(rs("username"))%> </td>          
            <td width="178" align="center"><%call vipcar(rs("username"))%></td>
            <td width="178" align="center"><%call jiereu(rs("username"))%></td>
            <td width="178" align="center"><%call fareu(rs("username"))%></td>
       
            <td width="94" align="center"><%=rs("now")%></td>
            <td width="94" align="center"><a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=<%=rs("qq")%>&site=qq&menu=yes"><img src="http://wpa.qq.com/pa?p=1:<%=rs("qq")%>:4"  border="0" /> </a></td>   
          </tr>
		  
            <%
      rs.MoveNext
	  RowCount = RowCount - 1 
      Loop%>
          </table>
<table width="100%" border="0" cellpadding="0" cellspacing="1" class="border">
  <tr class="tdbg">
    <td height="30" colspan="2">&nbsp;&nbsp;&nbsp;&nbsp;</td>
    </tr>
  <tr bgcolor="#FFFFFF"><td  ><div align="center"><br>
    <% call showpage(url,rs.RECORDCOUNT,maxperpage,false,true,"条主题") %>
  </div></td></tr>
</table>
</td>
</form></tr>

</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="56" align="center" background=images/admin_bg_bottom.gif><font color=#cccccc>Copyright &copy; 2013</font> <a href='http://www.778892.com' target="_blank"><font color=#ff6600><strong>七七传媒-七七互刷平台</strong></font></a> <font color=#cccccc>. All Rights Reserved.</font></td>
  </tr>
</table>

<%rs.close
set rs=nothing
conn.close
set conn=nothing%>
</body>
</html>
