<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!--#include file="../inc/index_conn.asp"-->
<!--#INCLUDE FILE="../inc/function.asp"-->
<!--#INCLUDE FILE="../config.asp"-->
<!--#include file="../user/checklogin.asp"-->
<!--#INCLUDE FILE="../background.asp"-->
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
sql = "select * from "&jieducm&"_tribe where username='"&session("useridname")&"'"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if  rs.eof then
	Response.Write("<script>alert('�㲻��������!���ܹ����䣡');window.location.href='index.asp';</script>")
    response.End()
else
   tribeid=rs("id")			
end if
 %>
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD>
<title><%=stupname%> -<%=stuptitle%></title>
<meta content="<%=stupdesc%>" name="description"/>
<meta content="<%=stupkey%>" name="keywords"/>
<link href="../css/style.css" rel="stylesheet" type="text/css" />
<LINK href="../css/center.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2900.2180" name=GENERATOR>
<SCRIPT src="../js/Dialog.js" type=text/javascript></SCRIPT>
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
    <!--#include file="../user/userleft.asp"--></TD>
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
                    <div class=pp7>�����ڵ�λ���ǣ�<%=stupname%> &gt;&gt; ��Ա���� &gt;&gt;������ &gt;&gt; </div>
                    <div class=pp8><strong>������</strong></div>
                    <br>
                    <br>
					  <table width="95%" border="0" align="center" cellpadding="3" cellspacing="0" bgcolor="#DEEEFA">
                <tr>
                  <td><STRONG>&nbsp;������</STRONG></td>
                </tr>
              </table>
			  <table width="100%" border="0" cellpadding="0" cellspacing="0"  class="margin6">
          <tr>
            <td width="27%" align="center"><span class="font12h">���</span></td>
            <td width="21%" height="20" align="center"><span class="font12h">�û���</span>&nbsp;<span class="redcolor">&nbsp;</span></td>
            <td width="29%"><div align="center"><span class="font12h"><strong>������</strong></span></div></td>
            <td width="23%"><div align="center"><span class="font12h">����</span></div></td>
            </tr>
	<%	
   	sql="SELECT * FROM  "&jieducm&"_register  where tribeid in ("&tribeid&") order by id desc"
	Set Rs = Server.CreateObject("Adodb.RecordSet")
	Rs.Open Sql,conn,1,1
	if not rs.eof then
	dim maxperpage,url,PageNo
	 url="manage.asp"
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
	 j=1
	if PageNo=1 then i=rs.RECORDCOUNT else i=rs.RECORDCOUNT-(30*(PageNo-1))
	DO WHILE NOT rs.EOF AND RowCount>0%>
          <tr>
            <td align="center" ><%=j%></td>
            <td height="26" align="center" ><%=rs("username")%></td>
            <td align="center" >
			<%
sqls="SELECT count(*) as tribetotal FROM "&jieducm&"_tribebook where username='"&rs("username")&"'"
Set Rss = Server.CreateObject("Adodb.RecordSet")
Rss.Open Sqls,conn,1,1
if not rss.eof then			
tribetotal=rss("tribetotal")
end if
rss.close
set rss=nothing
response.Write(tribetotal)
			%>
			</td>
            <td align="center" ><a href="#1" onClick="javascript:showDialog('��ȷ��Ҫ�߳��˻�Ա��',1,7,'managedel.asp?action=del&id=<%=rs("id")%>')" title="�߳��˻�Ա��">�߳�</a></td>
            </tr>
			  <%
			  j=j+1
      rs.MoveNext
	  RowCount = RowCount - 1 
      Loop%>
	   <tr>
            <td height="26" colspan="4" align="center" class="font12h"> <% call showpage(url,rs.RECORDCOUNT,maxperpage,false,true,"������") %></td>
            </tr>
        </table>
                  </div>
                </div></td>
	          </tr>
	        </table>	   
			 </td>
	  </tr>
</table>

<!--#INCLUDE FILE="../inc/jieducm_bottom.asp"-->
<%
rs.close()
set rs=nothing
call closeconn()%>
</BODY></HTML>
