<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#INCLUDE FILE="../config.asp"-->
<!--#INCLUDE FILE="../inc/function.asp"-->
<!--#include file="checklogin.asp"-->
<!--#include file="../background.asp"-->
<%
'------------------------------------------------------------------
'***************************�Ա�������ˢϵͳ**************************************
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
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD>
<title><%=stupname%> -<%=stuptitle%></title>
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<link href="../css/style.css" rel="stylesheet" type="text/css" />
<LINK href="../css/center.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2900.2180" name=GENERATOR>
</HEAD>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">   
<!--#include file=../inc/jieducm_top.asp-->
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
<tr>
	    <td>
        <table width="910"  border="0" cellspacing="0" cellpadding="0">
	          <tr>
	          
	            <td width="140" valign="top">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td height="5"> </td>
    </tr>
</table>



<table id="menuToolBar" width="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
    <td valign="top" bgcolor="FFFFFF">
    <!--#include file="userleft.asp"--></TD>
</TR>
</table></td>
	            <td width="15">&nbsp;</td>
	          
	            <td valign="top">
                <table width="100%" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td height="5"></td>
    </tr>
</table>
 <DIV class=pp9>
      <DIV style="PADDING-BOTTOM: 15px; WIDTH: 97%">
      <DIV class=pp7>
        <div align="left">�����ڵ�λ���ǣ�<%=stupname%> &gt;&gt; ��Ա���� &gt;&gt;���ܵ����� &gt;&gt; </div>
      </DIV>
      <DIV class=pp8>
        <div align="left"><STRONG>���ܵ�����</STRONG></div>
      </DIV>
      <table width="100%" border="0" cellpadding="0" cellspacing="0"  class="margin6">
          <tr>
            <td width="16%" height="20" align="center">&nbsp;<span class="redcolor">&nbsp;������</span></td>
            <td width="13%">&nbsp;</td>
            <td width="12%">&nbsp;</td>
            <td width="15%">&nbsp;</td>
   
      <td width="16%">&nbsp;</td>
            <td width="14%">&nbsp;</td>
            <td width="14%">&nbsp;</td>
          </tr>
          <tr>
            <td height="26" align="center" class="font12h">����</td>
            <td align="center" class="font12h">������</td>
            <td align="center" class="font12h">����ID </td>
            <td align="center" class="font12h">����ʱ��</td>
            <td align="center" class="font12h">�Է��Ƿ��ѱ��</td>
            <td align="center" class="font12h">�ٷ��Ƿ������</td>
            <td align="center" class="font12h">����</td>
          </tr>
        </table>
		
		<table width="100%" border="0" cellspacing="0" cellpadding="0"  id="con_three_1">
 <%	 

	     sql="select  * from "&jieducm&"_toushu where name='"&session("useridname")&"' order by id desc"
			     Set Rs = Server.CreateObject("Adodb.RecordSet")
				Rs.Open Sql,conn,1,1
				if not rs.eof then
					
	dim maxperpage,url,PageNo	
	 url="complaintre.asp"
	 rs.pagesize=9
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
	  %>
       <% DO WHILE NOT rs.EOF AND RowCount>0%>  
            <tr>
              <td width="16%" height="55" align="center" class="borderc"><%=left(rs("title"),10)%></td>
              <td width="13%" align="center" class="borderc"><%=rs("username")%></td>
              <td width="12%" align="center" class="borderc"><%=rs("pid")%></td>
              <td width="16%" align="center" class="borderc"><%=year(rs("now"))%>-<%=month(rs("now"))%>-<%=day(rs("now"))%>&nbsp;<%=hour(rs("now"))%>:<%=minute(rs("now"))%> </td>
              <td width="15%" align="center" class="borderc"><%if rs("bianjie")="��" then%>��<%else%>��<%end if%></td>
              <td width="14%" align="center" class="borderc"><%if rs("guang")="��" then%>��<%else%>��<%end if%></td>
              <td width="14%" align="center" class="borderc"><span class="font12h">
              <%if rs("bianjie2")="��" then%><a href="complainbian.asp?id=<%=rs("id")%>">���</a><%else%>�ѱ��<%end if%>
              &nbsp; <a href="complaintloogk.asp?id=<%=rs("id")%>&amp;my=2">�鿴</a></span></td>
            </tr>
             <%
			  	RowCount = RowCount - 1
	  i=i-1
      rs.MoveNext 
      Loop
				
			  %>
          </table>
      </DIV>
 </DIV></td>
	          </tr>
	        </table>	    </td>
	  </tr>
</table>

<!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
<%call closeconn()%>
</BODY></HTML>
