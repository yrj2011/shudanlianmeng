<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#INCLUDE FILE="../config.asp"-->
<!--#INCLUDE FILE="../inc/function.asp"-->
<!--#include file="checklogin.asp"-->

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
<style type="text/css">
.set {
	clear:both;
}
.set td {
	border-bottom:1px solid #d5d5d5;
	line-height:16px;
	vertical-align:middle;
	padding-left:30px;
}
.set thead td {
	margin-bottom:4px;
	color:#666;
	background:#F8F8F8;
	font-weight:bold;
}
.set tbody tr:hover td {
	background-color:#F8F8F8;
}
.set_bar td {
	background:#F6FBFF url(../images/ico_1.gif) 15px center no-repeat;
	font-weight:bold;
}

.STYLE18 {
	color: #3333FF;
	font-weight: bold;
}
.STYLE23 {color: #3300FF}
.STYLE24 {color: #3333FF}
.STYLE25 {color: #3366CC}
.STYLE26 {color: #3366CC; font-weight: bold; }
.STYLE31 {
	color: #000000;
	font-weight: bold;
}
</style>
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
    <td valign="top" bgcolor="FFFFFF"><!--#include file="userleft.asp"--></TD>
</TR>
</table></td>
	            <td width="15">&nbsp;</td>
	          
	            <td valign="top">
                <table width="100%" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td height="5"></td>
    </tr>
</table>
                <div class=pp9>
                  <div style="PADDING-BOTTOM: 15px; WIDTH: 97%">
                    <div class=pp7>�����ڵ�λ���ǣ�<%=stupname%> &gt;&gt; ��Ա���� &gt;&gt;�ٷ��ͷ����� &gt;&gt; </div>
                    <div class=pp8>
                      <table width="100%" border="0">
                        <tr>
                          <td width="78%"><strong>�ٷ�������Ա����</strong></td>
                          <td width="22%"><div align="center"><a href="complist.asp?action=1" class="STYLE11">�鿴ȫ��</a></div></td>
                        </tr>
                      </table>
                    </div>
                    <table  border="0" cellspacing="1" width="100%" cellpadding="0" class="set">
                      <tr class="title" height="22">
                        <td width="156" height="22" align="center" ><span class="STYLE2"><strong>�û���</strong></span></td>
                        <td width="123" align="center" ><span class="STYLE1"><strong>��ֵ��Ŀ</strong></span></td>
                        <td width="123" align="center" ><span class="STYLE1"><strong>��������</strong></span></td>
                        <td width="164" align="center" ><span class="STYLE2"><strong>����Ա</strong></span></td>
                        <td width="164" align="center" ><span class="STYLE1"><strong>���ʱ��</strong></span></td>
                        <td width="323" align="center" ><span class="STYLE1"><strong>����<span class="STYLE5">/</span>��ֵ����</strong></span></td>
                      </tr>
                      <%
sql="SELECT top 10 * FROM "&jieducm&"_chengfa where leibie=1 order by id desc"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
IF Rs.Eof Then
	Response.Write("������Ϣ")
Else
Do While Not Rs.Eof
%>
                      <tr class="tdbg" >
                        <td width="156" align="center"><%=rs("username")%></td>
                        <td width="123" align="center"><%if rs("class")=1 then
			 response.Write("���") 
			elseif rs("class")=2 then
			 response.Write("������")
			 elseif rs("class")=3 then
			 response.Write("����")
		   end if%></td>
                        <td width="123" align="center"><%=rs("num")%></td>
                        <td width="164" align="center"><%=rs("manage")%></td>
                        <td width="164" align="center"><%=rs("now")%> </td>
                        <td width="323" align="center">&nbsp;<%=rs("content")%></td>
                      </tr>
                      <%
Rs.MoveNext
Loop
End IF
%>
                    </table>
                    <div class=pp8>
                      <table width="100%" border="0">
                        <tr>
                          <td width="78%"><strong><span class="STYLE5">�ٷ�����</span><span class="STYLE5">��Ա����</span></strong></td>
                          <td width="22%"><div align="center" class="STYLE5"><a href="complist.asp?action=2"><strong>�鿴ȫ��</strong></a></div></td>
                        </tr>
                      </table>
                    </div>
                    <table  border="0" cellspacing="1" width="100%" cellpadding="0" class="set">
                      <tr class="title" height="22">
                        <td width="156" height="22" align="center" ><span class="STYLE5"><strong>�û���</strong></span></td>
                        <td width="123" align="center" ><span class="STYLE5"><strong>������Ŀ</strong></span></td>
                        <td width="123" align="center" ><span class="STYLE5"><strong>�������� </strong></span></td>
                        <td width="164" align="center" ><span class="STYLE5"><strong>����Ա</strong></span></td>
                        <td width="164" align="center" ><span class="STYLE5"><strong>����ʱ��</strong></span></td>
                        <td width="323" align="center" ><span class="STYLE5"><strong>����/����ԭ��</strong></span></td>
                      </tr>
                      <%
sql="SELECT top 10 * FROM "&jieducm&"_chengfa where leibie=2 order by id desc"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
IF Rs.Eof Then
	'Response.Write("������Ϣ")
Else
Do While Not Rs.Eof
%>
                      <tr class="tdbg" >
                        <td width="156" align="center"><%=rs("username")%></td>
                        <td width="123" align="center"><%if rs("class")=1 then
			 response.Write("���") 
			elseif rs("class")=2 then
			 response.Write("������")
			 elseif rs("class")=3 then
			 response.Write("����")
		   end if%></td>
                        <td width="123" align="center"><%=rs("num")%></td>
                        <td width="164" align="center"><%=rs("manage")%></td>
                        <td width="164" align="center"><%=rs("now")%> </td>
                        <td width="323" align="center">&nbsp;<%=rs("content")%></td>
                      </tr>
                      <%
Rs.MoveNext
Loop
End IF
%>
                    </table>
                  </div>
                </div></td>
	          </tr>
	        </table>	    </td>
	  </tr>
</table>

<!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
<%call closeconn()%>
</BODY></HTML>
