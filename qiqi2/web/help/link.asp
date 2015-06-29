<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="linkfunction.asp"-->
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
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML><HEAD>
<title><%=stupname%> -<%=stuptitle%></title>
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="images/link.css" type=text/css rel=stylesheet>
<link href="images/css.css" rel="stylesheet" type="text/css">
<LINK href="help.css" type=text/css rel=stylesheet>

<STYLE type=text/css>.style1 {
	FONT-WEIGHT: bold; FONT-SIZE: 14px
}
</STYLE>

<META content="MSHTML 6.00.2900.3492" name=GENERATOR></HEAD>
<BODY>
<!--#include file="top.asp"-->

<TABLE cellSpacing=0 cellPadding=0 width=840 align=center border=0>
  <TBODY>
  <TR>
    <TD height=10></TD></TR>
  <TR>
    <TD style="BORDER-BOTTOM: #eee 1px solid; BACKGROUND-REPEAT: no-repeat" 
    background=images/sub_header.gif height=23>
      <DIV style="PADDING-RIGHT: 20px; TEXT-ALIGN: right"><A class=blue 
      href="index.asp">帮助首页</A> &gt; 友情链接</DIV></TD></TR>
  <TR>
    <TD vAlign=top>
      <DIV 
      style="PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px">
      <TABLE cellSpacing=0 cellPadding=0 width="100%" bgColor=#f3f3f3 
        border=0><TBODY>
       
        <TR>
          <TD width="19%">
            <DIV 
            style="PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px"><IMG 
            src="images/sub_header2.gif"></DIV></TD>
          <TD width="81%">
            <DIV 
            style="PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 10px; LINE-HEIGHT: 18px; PADDING-TOP: 10px">&nbsp;&nbsp;为了共同的发展，本站诚招各类友情链接。要求：有一定的访问量。与本站链接的网站必须无不良信息，无违反国家法律的信息。<BR>
            &nbsp;<BR>本站名称：<SPAN class=orange14b><%=stupname%>  
        </SPAN>本站网址：<SPAN class=orange14b><%=stupurl%></SPAN><BR>
            </DIV></TD></TR> 
           
            <TR>
          <TD><span class="topdh8"><STRONG>本站LOGO：</STRONG></span></TD>
          <TD><IMG src="../<%=stuplogo%>" ></TD>
        </TR> 
            <TR>
              <TD><span class="topdh8"><STRONG>本站文字链接代码：<BR>
                    <FONT 
            color=#ff0000>演示：</FONT></STRONG><BR>
                <A 
            href="../index.asp" target=_blank><%=stupname%></A></span></TD>
              <TD><TEXTAREA name=textarea rows=5 cols=60><a href="<%=stupurl%>" target="_blank" title="免费推广店铺、免费淘宝刷钻"><%=stupname%></a></TEXTAREA></TD>
            </TR>
            <TR>
              <TD><span class="topdh8"><STRONG>本站LOGO链接代码：<BR>
                    <FONT 
            color=#ff0000>演示：</FONT></STRONG>&nbsp;&nbsp;<br>
              </span></TD>
              <TD><TEXTAREA name=textarea2 rows=5 cols=60><a href="<%=stupurl%>" target="_blank" title="免费推广店铺、免费淘宝刷钻"><IMG src="http://www.renrenlai.com/uploadpic/jieducm_2014322942835858.jpg"  border="0"></a>
</TEXTAREA></TD>
            </TR>
            <TR>
              <TD><STRONG><FONT color=#ff6600>广 告 服 务</FONT></STRONG></TD>
              <TD>可在本站投放广告，具体事宜请联系客服。<a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=<%=stupqq%>&site=qq&menu=yes"><IMG height=16 alt=联系充值客服 
src="images/qq1_online.gif" width=23 border=0> 技术咨询</A></TD>
            </TR></TBODY></TABLE>
      </DIV></TD></TR></TBODY></TABLE>
<TABLE style="FONT-SIZE: 12px" cellSpacing=0 cellPadding=0 width=840 
align=center border=0>
  <TBODY>
  <TR>
    <TD><IMG height=30 src="images/link_header01.gif" width=105 
      useMap=#Map border=0></TD></TR>
  <TR>
    <TD bgColor=#6699cc height=3></TD></TR>
  <TR>
    <TD class=text>
      <DIV 
      style="PADDING-RIGHT: 5px; PADDING-LEFT: 5px; PADDING-BOTTOM: 5px; PADDING-TOP: 5px">
        <div align="left">
          <%
			  	Sql = "select * from "&jieducm&"_FriendLinks where  IsOK=1 and LinkType=2 order by id desc"
				Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				IF Rs.Eof Then
					Response.Write("暂无信息")
				Else
					Do While Not Rs.Eof
			  %>
          [<a href="<%=rs("siteurl")%>" target="_blank" title="<%=rs("sitename")%>"><%=rs("SiteName")%></a>]
          
          <%
			  	Rs.MoveNext
				Loop
				End IF
			  %>
        </div>
      </DIV></TD></TR>
  <TR>
    <TD height=10></TD></TR></TBODY></TABLE>
<TABLE style="FONT-SIZE: 12px" cellSpacing=0 cellPadding=0 width=840 
align=center border=0>
  <TBODY>
  <TR>
    <TD><IMG height=30 src="images/link_header03.gif" width=105 
      useMap=#Map border=0></TD></TR>
  <TR>
    <TD bgColor=#ff921c height=3></TD></TR>
  <TR>
    <TD class=text>
      <DIV 
      style="PADDING-RIGHT: 5px; PADDING-LEFT: 5px; PADDING-BOTTOM: 5px; PADDING-TOP: 5px">
        <div align="left">
          <%
			  	Sql = "select * from "&jieducm&"_FriendLinks where  IsOK=1 and LinkType=1 order by id desc"
				Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				IF Rs.Eof Then
					Response.Write("暂无信息")
				Else
					Do While Not Rs.Eof
			  %>
          <a href="<%=rs("siteurl")%>" target="_blank" title="<%=rs("sitename")%>"><img src=<%=rs("LogoUrl")%> width=88 height=31 border=0></a>
          
          <%
			  	Rs.MoveNext
				Loop
				End IF
			  %>
        </div>
      </DIV></TD></TR></TBODY></TABLE>
        <!--#include file="bottom.asp"-->
<%rs.close
set rs=nothing
conn.close
set conn=nothing%>
</BODY></HTML>
