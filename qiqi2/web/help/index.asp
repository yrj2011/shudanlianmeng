<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../config.asp"-->
<%
'------------------------------------------------------------------
'***************************淘宝信誉互刷系统***********************
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
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML><HEAD>
<title><%=stupname%> -<%=stuptitle%></title>
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="help.css" type=text/css rel=stylesheet>
<link href="images/css.css" rel="stylesheet" type="text/css">


<STYLE type=text/css>.style1 {
	FONT-WEIGHT: bold; FONT-SIZE: 14px
}
</STYLE>


<META content="MSHTML 6.00.2900.3492" name=GENERATOR>
</HEAD>
<BODY>
<!--#include file="top.asp"-->
<TABLE cellSpacing=0 cellPadding=0 width=840 align=center border=0>
  <TBODY>
  <TR>
    <TD class=newm_01 height=8></TD></TR>
  <TR>
    <TD class=newm_bg height=8>
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
        <TBODY>
        <TR>
          <TD>
            <DIV style="PADDING-LEFT: 70px"><IMG 
            src="images/new_01.gif"></DIV></TD>
          <TD><A 
           
            href="../register.asp" target=_blank><IMG 
            id=Image1 src="images/newbut_01.gif" border=0 
            name=Image1></A>&nbsp;<A 
           
            href="#" target=_blank><IMG 
            id=Image2 src="images/newbut_02.gif" border=0 
            name=Image2></A>&nbsp;&nbsp;<A 
           
            href="help8.asp" target=_blank><IMG 
            id=Image4 height=50 src="images/newbut_04.gif" width=144 
            border=0 name=Image4></A></TD>
        </TR></TBODY></TABLE></TD></TR>
  <TR>
    <TD class=newm_02 height=8></TD></TR>
  <TR>
    <TD height=15></TD></TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width=840 align=center border=0>
  <TBODY>
  <TR>
    <TD vAlign=top width=600>
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
        <TBODY>
        <TR>
          <TD class=content_01 height=7></TD></TR>
        <TR>
          <TD class=content_bg vAlign=top>
            <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
              <TBODY>
              <TR>
                <TD>
                  <DIV style="PADDING-LEFT: 25px; PADDING-TOP: 10px"><IMG 
                  src="images/content_03.gif">
                    
                    </DIV></TD></TR>
              <TR>
                <TD class=content>
                  <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
                    <TBODY>
                    <TR>
                      <TD>
                        <DIV 
                        style="PADDING-LEFT: 40px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px"><IMG 
                        src="images/arrow.gif">&nbsp;<SPAN 
                        class=content_header><A class=blue 
                        href="help10.asp?action=82">初识刷客</A></SPAN> 
                        <BR>
                        <%
			  	Sql = "select Top 7 NewsPath,Title,ArticleID,ClassID,UpdateTime from "&jieducm&"_Article where classid in (82) order by ArticleID desc"
				Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				IF Rs.Eof Then
					Response.Write("暂无信息")
				Else
					Do While Not Rs.Eof
			  %>

                        <SPAN class=blue><a href="new.asp?id=<%=rs("ArticleID")%>" target="_blank" class="blue"><%=left(Replace(Rs("Title"),"&nbsp;"," "),8)%></a></SPAN>
                        <%
			  	Rs.MoveNext
				Loop
				End IF
			  %>
                        <BR>
                        <IMG src="images/arrow.gif">&nbsp;<SPAN 
                        class=content_header><A class=blue 
                        href="help10.asp?action=84">刷客类型</A></SPAN> 
                        <BR>
                          <%
			  	Sql = "select Top 7 NewsPath,Title,ArticleID,ClassID,UpdateTime from "&jieducm&"_Article where classid in (84) order by ArticleID desc"
				Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				IF Rs.Eof Then
					Response.Write("暂无信息")
				Else
					Do While Not Rs.Eof
			  %>

                        <SPAN class=blue><a href="new.asp?id=<%=rs("ArticleID")%>" target="_blank" class="blue"><%=left(Replace(Rs("Title"),"&nbsp;"," "),8)%></a></SPAN>
                        <%
			  	Rs.MoveNext
				Loop
				End IF
			  %> 
                        <br>
                        <span style="PADDING-BOTTOM: 10px; PADDING-TOP: 10px"><span 
                        class=content_header><IMG src="images/arrow.gif">&nbsp;<A class=blue 
                        href="help10.asp?action=86">刷客相关</A></span><br>
                          <%
			  	Sql = "select Top 7 NewsPath,Title,ArticleID,ClassID,UpdateTime from "&jieducm&"_Article where classid in (86) order by ArticleID desc"
				Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				IF Rs.Eof Then
					Response.Write("暂无信息")
				Else
					Do While Not Rs.Eof
			  %>

                        <SPAN class=blue><a href="new.asp?id=<%=rs("ArticleID")%>" target="_blank" class="blue"><%=left(Replace(Rs("Title"),"&nbsp;"," "),8)%></a></SPAN>
                        <%
			  	Rs.MoveNext
				Loop
				End IF
			  %>
                        
                        <BR>
                        </DIV></TD>
                      <TD width="50%">
                        <DIV 
                        style="PADDING-BOTTOM: 10px; PADDING-TOP: 10px"><IMG 
                        src="images/arrow.gif">&nbsp;<SPAN 
                        class=content_header><A class=blue 
                        href="help10.asp?action=83">互刷帮助</A></SPAN> 
                        <BR>
                        
                          <%
			  	Sql = "select Top 7 NewsPath,Title,ArticleID,ClassID,UpdateTime from "&jieducm&"_Article where classid in (83) order by ArticleID desc"
				Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				IF Rs.Eof Then
					Response.Write("暂无信息")
				Else
					Do While Not Rs.Eof
			  %>

                        <SPAN class=blue><a href="new.asp?id=<%=rs("ArticleID")%>" target="_blank" class="blue"><%=left(Replace(Rs("Title"),"&nbsp;"," "),8)%></a></SPAN>
                        <%
			  	Rs.MoveNext
				Loop
				End IF
			  %>
                        <BR><IMG src="images/arrow.gif">&nbsp;<SPAN 
                        class=content_header><A class=blue 
                        href="help10.asp?action=85">刷客宝典</A></SPAN> 
                        <BR>
                          <%
			  	Sql = "select Top 7 NewsPath,Title,ArticleID,ClassID,UpdateTime from "&jieducm&"_Article where classid in (85) order by ArticleID desc"
				Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				IF Rs.Eof Then
					Response.Write("暂无信息")
				Else
					Do While Not Rs.Eof
			  %>

                        <SPAN class=blue><a href="new.asp?id=<%=rs("ArticleID")%>" target="_blank" class="blue"><%=left(Replace(Rs("Title"),"&nbsp;"," "),8)%></a></SPAN>
                        <%
			  	Rs.MoveNext
				Loop
				End IF
			  %>
                        <BR><IMG src="images/arrow.gif">&nbsp;<SPAN 
                        class="blue style1"><A class=blue 
                        href="help10.asp?action=87">高手必读</A></SPAN> 
                        <BR>  <%
			  	Sql = "select Top 7 NewsPath,Title,ArticleID,ClassID,UpdateTime from "&jieducm&"_Article where classid in (87) order by ArticleID desc"
				Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				IF Rs.Eof Then
					Response.Write("暂无信息")
				Else
					Do While Not Rs.Eof
			  %>

                        <SPAN class=blue><a href="new.asp?id=<%=rs("ArticleID")%>" target="_blank" class="blue"><%=left(Replace(Rs("Title"),"&nbsp;"," "),8)%></a></SPAN>
                        <%
			  	Rs.MoveNext
				Loop
				End IF
			  %>
                      <BR>
                      </DIV></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR>
        <TR>
          <TD class=content_02 height=7></TD></TR></TBODY></TABLE></TD>
    <TD vAlign=top width=260>
      <TABLE class=qa_bg cellSpacing=0 cellPadding=0 width="100%" border=0>
        <TBODY>
        <TR>
          <TD>
            <DIV 
            style="PADDING-RIGHT: 30px; PADDING-LEFT: 40px; PADDING-TOP: 100px">
            <TABLE class=blue cellSpacing=1 cellPadding=1 width="100%" 
              border=0><TBODY>
                <%
			  	Sql = "select Top 7 NewsPath,Title,ArticleID,ClassID,UpdateTime from "&jieducm&"_Article where classid in (88) order by ArticleID desc"
				Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				IF Rs.Eof Then
					Response.Write("暂无信息")
				Else
					Do While Not Rs.Eof
			  %>

                       
                       
              <TR>
                <TD class=dash>·<a href="new.asp?id=<%=rs("ArticleID")%>" target="_blank" class="blue"><%=left(Replace(Rs("Title"),"&nbsp;"," "),8)%></a></TD></TR> <%
			  	Rs.MoveNext
				Loop
				End IF
			  %>
              <TR>
                <TD class=dash>&nbsp;</TD>
              </TR>
              <TR>
                <TD class=blue align=right>
                  <DIV style="PADDING-RIGHT: 10px"><a href="help10.asp?action=88">更多</a></DIV></TD></TR></TBODY></TABLE>
            </DIV></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
<!--#include file="bottom.asp"-->
<%rs.close
set rs=nothing
conn.close
set conn=nothing%>
</BODY></HTML>
