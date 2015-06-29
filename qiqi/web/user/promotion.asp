<!--#include file="../inc/index_conn.asp"-->
<!--#INCLUDE FILE="../inc/function.asp"-->
<!--#INCLUDE FILE="../config.asp"-->
<!--#include file="checklogin.asp"-->
<!--#include file="../background.asp"-->
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
sub jieducm_fanum(username_fa)
jieducm_oknum=0
sql="select count(*) as faok from "&jieducm&"_zhongxin where  username='"&username_fa&"'"
Set Rsk = Server.CreateObject("Adodb.RecordSet")
Rsk.Open Sql,conn,1,1
if not rsk.eof then
jieducm_oknum=rsk("faok")
end if
rsk.close
response.Write(jieducm_oknum)
end sub

sub jieducm_jienumk(username_jie)
jieducm_jienum=0
sql="select count(*) as jieducm_shehe from "&jieducm&"_jieshou where  username='"&rs("username")&"'  "
Set Rsj = Server.CreateObject("Adodb.RecordSet")
Rsj.Open Sql,conn,1,1
if not Rsj.eof then
jieducm_jienum=Rsj("jieducm_shehe")
end if
rsj.close
response.Write(jieducm_jienum)
end sub
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">				
<head>
<title>会员中心-<%=stupname%> -<%=stuptitle%></title>
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<LINK href="../images/Default.css" type=text/css rel=stylesheet>

</head>
<body>
<!--#include file=../inc/jieducm_top.asp-->
	
 	
<div id="Container">
<div > 
<table align="center">
<tr><td>      		
<table id="menuToolBar" width="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
    <td valign="top" bgcolor="FFFFFF">
    <!--#include file="userleft.asp"--></TD>
</TR>
</table>
</td><td valign="top">
<DIV style="CLEAR: both; WIDTH: 730px; BACKGROUND-COLOR: #f3f8fe">
<DIV style="CLEAR: both; MARGIN-TOP: 20px; LINE-HEIGHT: 150%">
<TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
  <TBODY>
  <TR>
    <TD>
      <TABLE cellSpacing=0 cellPadding=0 align=left border=0>
        <TBODY>
        <TR>
          <TD width=1><IMG height=25 
            src="../images/mytaobaobj1_3.gif" 
            width=3></TD>
          <TD align=middle width=120 
          background="../images/mytaobaobj1_4.gif"><FONT 
            color=#ffffff><SPAN>我推荐的用户</SPAN></FONT></TD>
          <TD width=1><IMG height=25 
            src="../images/mytaobaobj1_6.gif" 
            width=3></TD></TR></TBODY></TABLE>
      <TABLE cellSpacing=0 cellPadding=0 align=left border=0>
        <TBODY>
        <TR>
          <TD><SPAN style="PADDING-LEFT: 20px">从<%=nowe%>后推荐新人来<%=stupname%>的</SPAN></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
  <TBODY>
  <TR>
    <TD bgColor=#1e88c1 height=4></TD></TR></TBODY></TABLE></DIV></DIV>
<DIV 
style="CLEAR: both; BORDER-RIGHT: #abbec8 1px solid; BORDER-TOP: #abbec8 1px solid; MARGIN-TOP: 10px; BORDER-LEFT: #abbec8 1px solid; WIDTH: 730px; BORDER-BOTTOM: #abbec8 1px solid; BACKGROUND-COLOR: #f3f8fe">
<TABLE 
style="BORDER-TOP-WIDTH: 0px; BORDER-LEFT-WIDTH: 0px; BORDER-BOTTOM-WIDTH: 0px; MARGIN-LEFT: 20px; WIDTH: 90%; BORDER-RIGHT-WIDTH: 0px" 
align=left>
  <THEAD style="BORDER-BOTTOM: #cccccc 1px solid; HEIGHT: 40px">
  <TR>
    <TD width="12%" style="FONT-WEIGHT: bolder; ">用户名 </TD>
    <TD width="12%" style="FONT-WEIGHT: bolder; ">联系方式 </TD>
    <TD width="16%" style="FONT-WEIGHT: bolder; ">发任务数</TD>
    <TD width="16%" style="FONT-WEIGHT: bolder;">接任务数</TD>
    <TD width="16%" style="FONT-WEIGHT: bolder; ">注册时间 </TD>
    <TD width="9%" style="FONT-WEIGHT: bolder; ">现在积分 </TD>
    <TD width="19%" style="FONT-WEIGHT: bolder; ">现在发布点 </TD>
  </TR></THEAD>
  <TBODY>
  <%
 	          Sql = "select * from "&jieducm&"_register  where tjr='"&session("useridname")&"'"
				Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				if  not rs.eof then
	dim maxperpage,url,PageNo
	 url="promotion.asp"
	 rs.pagesize=20
	 PageNo=REQUEST("PageNo")
	 if PageNo="" or PageNo=0 then PageNo=1
	 RS.AbsolutePage=PageNo
	 TSum=rs.pagecount
	 maxperpage=rs.pagesize
	 tCount=rs.RECORDCOUNT
	 RowCount=rs.PageSize
	   PageNo=PageNo+1
	   PageNo=PageNo-1
	 if CINT(PageNo)>1 then
	    if CINT(PageNo)>CINT(TSum) then
		  response.Write("对不起没有您想要的页数")
          Response.End
	    end if
	 end if
		   
     if PageNo<0 then
	    response.Write("没有这一页!")
		Response.End
	 End if
	    
	if PageNo=1 then i=rs.RECORDCOUNT else i=rs.RECORDCOUNT-(30*(PageNo-1))
	   DO WHILE NOT rs.EOF AND RowCount>0%>
  <TR>
    <TD class=centers><%=rs("username")%> </TD>
    <TD class=centers>
	<a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=<%=rs("qq")%>&site=qq&menu=yes"><img src="http://wpa.qq.com/pa?p=1:<%=rs("qq")%>:4"  border="0"/> </a> </TD>
    <TD class=centers><%call jieducm_fanum(rs("username"))%>个</TD>
    <TD class=centers><%call jieducm_jienumk(rs("username"))%>个</TD>
    <TD class=centers><%=rs("now")%> </TD>
    <TD class=centers><%=rs("jifei")%></TD>
    <TD class=centers><%=rs("fabudian")%></TD>
  </TR>
    
                <%
			  	RowCount = RowCount - 1
	  i=i-1
      rs.MoveNext 
      Loop
				
			  %>
  <TR>
    <TD align=middle colSpan=6>  <% call showpage(url,rs.RECORDCOUNT,maxperpage,false,true,"条主题") %></TD></TR>
    
    <%else%>
    暂无推广
    <%end if%>
    </TBODY></TABLE>
</DIV>
<DIV style="CLEAR: both; WIDTH: 730px; BACKGROUND-COLOR: #f3f8fe">
<DIV style="CLEAR: both; MARGIN-TOP: 20px; LINE-HEIGHT: 150%">
<TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
  <TBODY>
  <TR>
    <TD>
      <TABLE cellSpacing=0 cellPadding=0 align=left border=0>
        <TBODY>
        <TR>
          <TD width=1><IMG height=25 
            src="../images/mytaobaobj1_3.gif" 
            width=3></TD>
          <TD align=middle width=200 
          background="../images/mytaobaobj1_4.gif"><FONT 
            color=#ffffff><SPAN>推广排行榜（推广积分排列）</SPAN></FONT></TD>
          <TD width=1><IMG height=25 
            src="../images/mytaobaobj1_6.gif" 
            width=3></TD></TR></TBODY></TABLE>
      <TABLE cellSpacing=0 cellPadding=0 align=left border=0>
        <TBODY>
        <TR>
          <TD><SPAN 
            style="PADDING-LEFT: 20px"></SPAN></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
  <TBODY>
  <TR>
    <TD bgColor=#1e88c1 height=4></TD></TR></TBODY></TABLE></DIV></DIV>
<DIV 
style="CLEAR: both; BORDER-RIGHT: #abbec8 1px solid; BORDER-TOP: #abbec8 1px solid; MARGIN-TOP: 10px; BORDER-LEFT: #abbec8 1px solid; WIDTH: 730px; BORDER-BOTTOM: #abbec8 1px solid; BACKGROUND-COLOR: #f3f8fe">
<TABLE 
style="BORDER-TOP-WIDTH: 0px; BORDER-LEFT-WIDTH: 0px; BORDER-BOTTOM-WIDTH: 0px; MARGIN-LEFT: 20px; WIDTH: 90%; BORDER-RIGHT-WIDTH: 0px" 
align=left>
  <THEAD style="BORDER-BOTTOM: #cccccc 1px solid; HEIGHT: 40px">
  <TR>
    <TD style="FONT-WEIGHT: bolder; WIDTH: 20%">用户名 </TD>
    <TD style="FONT-WEIGHT: bolder; WIDTH: 20%">推荐人数 </TD>
    <TD style="FONT-WEIGHT: bolder; WIDTH: 20%">当前推广积分 </TD>
    </TR></THEAD>
  <TBODY>
  <%  
 		       Sql = "select top 15  tjr,count(*) as yu  from  "&jieducm&"_register  where tjr<>''  group by tjr order by yu desc"
				Set Rst = Server.CreateObject("Adodb.RecordSet")
				Rst.Open Sql,conn,1,1
				IF Rst.Eof or rst.bof Then
					
				Else
				i=1
				Do While Not Rst.Eof
			  %>
  <TR>
    <TD class=centers><%=left(rst("tjr"),3)%>***</TD>
    <TD class=centers><%=rst("yu")%>人 </TD>
    <TD class=centers><%=rst("yu")*10%> </TD>
    </TR>
   <%            i=i+1
			  	Rst.MoveNext
				Loop
				End IF
			  %>
    </TBODY></TABLE>
</DIV>
  </ul>  
  </td></tr></table>         	
  </div>
</div>

<!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->

</body>
</html>


