<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../config.asp"-->
<%
'------------------------------------------------------------------
'***************************�Ա�������ˢϵͳ***********************
'������רΪ�Ա���ѻ�ˢ���죬��Ŀǰ�����Ϲ��ܱȽ���ȫ��ƽ̨
'���򿪷����ݶȴ�ý
'�ٷ���ҳ��http://www.jieducm.cn
'����֧�֣�robin 
'QQ;616591415
'�绰��13598006137
'��Ȩ����Ȩ���ڽݶȴ�ý��Ϣ�������޹�˾�����ÿ������޸ģ���Ȩ�ؾ���
'Copyright (C) 2008��2011 �ݶȴ�ý��Ϣ�������޹�˾ ��Ȩ����
'*****************************************************************
'------------------------------------------------------------------
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML><HEAD><title><%=stupname%> -<%=stuptitle%></title>
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<META http-equiv=Content-Type content="text/html; charset=gb2312"><LINK 
href="content.css" type=text/css rel=stylesheet><LINK 
href="help.css" type=text/css rel=stylesheet><LINK 
href="text.css" type=text/css rel=stylesheet>
<link href="images/css.css" rel="stylesheet" type="text/css">
<STYLE type=text/css>.style1 {
	COLOR: #ff6600
}
</STYLE>

<META content="MSHTML 6.00.2900.3492" name=GENERATOR></HEAD>
<BODY>
<!--#include file="top.asp"-->
<TABLE class=14hao cellSpacing=0 cellPadding=0 width=840 align=center 
  border=0><TBODY>
  <TR>
    <TD vAlign=top width=230>
      <DIV id=left>
      <DIV id=content_nav></DIV>
      <!--#include file="left.asp"-->
      </DIV></TD>
                    <%
  id=request("id")
Set rsnews=Server.CreateObject("ADODB.RecordSet") 
sql="update "&jieducm&"_Article set hits=hits+1 where ArticleID="&request("id")
conn.execute sql

	   set rs=server.createobject("adodb.recordset")
		rs.open "select  * from "&jieducm&"_Article where ArticleID="&id&" ",conn,1,1
		if rs.eof then
		  Response.Write("<script>alert('�������ô���!');window.location.href='index.asp';</script>")
		else
		   
                Sqlc = "select * from "&jieducm&"_articleclass where ClassID="&rs("ClassID")&" "
				Set Rsc = Server.CreateObject("Adodb.RecordSet")
				Rsc.Open Sqlc,conn,1,1
				IF  Rsc.Eof Then
					Response.Write("������Ϣ")
				Else
				  classid=rsc("classid")
				  classname=rsc("classname")
                end if
		%> 
    <TD vAlign=top width=610>
      <DIV id=right>
      <DIV id=right_top><IMG src="images/arrow3.gif">&nbsp;&nbsp; <A 
      class=blue href="index.asp">������ҳ</A>��<A 
            href="help10.asp?action=<%=classid%>" class="blue"><%=classname%></A> </DIV>
      <DIV>
      <TABLE cellSpacing=0 cellPadding=0 width=610 border=0>
        <TBODY>
        <TR>
          <TD height=5></TD></TR>
        <TR>
          <TD class=header bgColor=#f0f0f0><IMG 
            src="images/arrow2.gif"> <SPAN 
          class=blue12b><%=rs("title")%></SPAN></TD></TR>
        <TR>
          <TD class=header>
            <TABLE class=14hao cellSpacing=0 cellPadding=0 width="100%" 
border=0>
              <TBODY>
              
              <TR>
                <TD 
                  class=gs><p><DIV class=article id=contentInfo 
style="DISPLAY: block; VERTICAL-ALIGN: baseline"><%=rs("content")%></DIV></p></TD>
              </TR></TBODY></TABLE></TD></TR></TBODY></TABLE></DIV></DIV></TD></TR></TBODY></TABLE>
<!--#include file="bottom.asp"--></BODY></HTML>
<% end if%>