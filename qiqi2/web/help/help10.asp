<!--#include file="../inc/index_conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../background.asp"-->
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
action=request("action") 
Sqlc = "select * from "&jieducm&"_articleclass where ClassID="&action&" "
				Set Rsc = Server.CreateObject("Adodb.RecordSet")
				Rsc.Open Sqlc,conn,1,1
				IF  Rsc.Eof Then
					Response.Write("������Ϣ")
					response.End()
				Else
				  classid=rsc("classid")
				  classname=rsc("classname")
                end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">

<HTML><HEAD><title><%=stupname%> -<%=stuptitle%></title>
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="content.css" type=text/css rel=stylesheet>
<LINK href="help.css" type=text/css rel=stylesheet>
<LINK href="text.css" type=text/css rel=stylesheet>
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
     
      <!--#include file="left.asp"-->
     </TD>
    <TD vAlign=top width=610>
      <DIV id=right>
      <DIV id=right_top><IMG src="images/arrow3.gif">&nbsp;&nbsp; <A  class=blue href="index.asp">������ҳ</A>��<A href="news.asp?action=<%=classid%>" class="blue"><%=classname%></A></DIV>
      <DIV>
      <TABLE cellSpacing=0 cellPadding=0 width=610 border=0>
        <TBODY>
        <TR>
          <TD height=5></TD></TR>
  <%
	     sql="select * from "&jieducm&"_article where classid in ("&action&") order by Articleid desc"
			     Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				if not rs.eof then					
	dim maxperpage,url,PageNo
	 url="help10.asp"
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
		  response.Write("�Բ���û������Ҫ��ҳ��")
          Response.End
	    end if
	 end if
		end if    
     if PageNo<0 then
	    response.Write("û����һҳ!")
		Response.End
	 End if
	
	if PageNo=1 then i=rs.RECORDCOUNT else i=rs.RECORDCOUNT-(30*(PageNo-1))
	  DO WHILE NOT rs.EOF AND RowCount>0%>
              <TR>
          <TD class=header bgColor=#f0f0f0><IMG src="images/arrow2.gif"> <SPAN class=blue12b>
		  <A href="new.asp?id=<%=rs("ArticleID")%>" title="<%=rs("title")%>" target="_blank" class="blue12b"><%=left(rs("title"),20)%></A>
		  </SPAN></TD></TR>
        <TR>
  <%
      rs.MoveNext
	  RowCount = RowCount - 1 
      Loop%>
          
          <TD class=header><% call showpage(url,rs.RECORDCOUNT,maxperpage,false,true,"������") %>
                      </TD>
        </TR></TBODY></TABLE></DIV></DIV></TD></TR></TBODY></TABLE>
        <br>
<!--#include file="bottom.asp"-->

<%rs.close
set rs=nothing
conn.close
set conn=nothing%>
</BODY></HTML>
