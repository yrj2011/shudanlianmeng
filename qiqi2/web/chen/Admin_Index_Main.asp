<%@language=vbscript codepage=936 %>
<%
'------------------------------------------------------------------
'***************************�Ա�������ˢϵͳ***********************
'������רΪ�Ա���ѻ�ˢ���죬��Ŀǰ�����Ϲ��ܱȽ���ȫ��ƽ̨
'���򿪷�����  �� �� ý
'�ٷ���ҳ��http://www.778892.com
'����֧�֣�xsh
'QQ;859680966
'�绰��15889679112
'��Ȩ����Ȩ�������ߴ�ý��Ϣ�������޹�˾�����ÿ������޸ģ���Ȩ�ؾ���
'Copyright (C) 2013 ���ߴ�ý��Ϣ�������޹�˾ ��Ȩ����
'*****************************************************************
'------------------------------------------------------------------
option explicit
response.buffer=true	
Const PurviewLevel=0
%>
<!--#include file="inc/conn.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<!--#include file="inc/function.asp"-->
<%
dim sql,rs,jieducm_rgtotal,jieducm_viptotal,jieducm_adminbuycar,jieducm_zhtotal,jieducm_jstotal,jieducm_zhtotalok
dim jieducm_fbtotal,jieducm_cartotal,jieducm_chian_total,jieducm_alipay_total
%>
<html>
<head>
<title>��̨������ҳ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="Admin_Style.css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="border" align=center>
<tr align="center">
    <td height=25 colspan=2 class="topbg"><strong>��̨������ҳ</strong></tr>
<tr>
    <td width="100" class="tdbg" height=23><strong>�����ݷ�ʽ��</strong></td>
    <td class="tdbg"><a href="Admin_ArticleAdd1.asp">������� | </a><a href="Admin_ArticleManage.asp">���¹���</a>      </td>
</tr>
</table>
<table cellpadding="2" cellspacing="1" border="0" width="100%" class="border" align=center>
  <tr align="center"> 
    <td height=25 colspan=3 class="topbg"><strong>����ƽ̨����</strong></td> 
  <tr> 
    <td width="21%"  class="tdbg" height=23>����ע���Ա����<font color="#ff6600">
    <%
sql="SELECT count(*) as jieducm_rgnum FROM "&jieducm&"_register where ( datediff(day,[now],getdate())=0)"
Set rs = Server.CreateObject("Adodb.RecordSet")
rs.Open Sql,conn,1,1
if not rs.eof then			
jieducm_rgtotal=rs("jieducm_rgnum")
end if
rs.close
set rs=nothing
response.Write(jieducm_rgtotal)
	%></font>��</td>
    <td width="29%" class="tdbg">���շ�����������<font color="#ff6600">
    <%
sql="SELECT count(*) as jieducm_zhnum FROM "&jieducm&"_zhongxin where ( datediff(day,[now],getdate())=0)"
Set rs = Server.CreateObject("Adodb.RecordSet")
rs.Open Sql,conn,1,1
if not rs.eof then			
jieducm_zhtotal=rs("jieducm_zhnum")
end if
rs.close
set rs=nothing
response.Write(jieducm_zhtotal)
	%></font>��</td>
    <td width="50%" class="tdbg">�������������㣺<font color="#ff6600">
<%
sql="SELECT  sum(fabudian) as adminbuyfb FROM "&jieducm&"_chongjilu where class='���򷢲���' and (datediff(day,[now],getdate())=0)"
Set rs = Server.CreateObject("Adodb.RecordSet")
rs.Open Sql,conn,1,1
if not rs.eof then			
jieducm_fbtotal=rs("adminbuyfb")
end if
rs.close
set rs=nothing

if not isnull(jieducm_fbtotal) then
response.Write(jieducm_fbtotal)
else
response.Write("0")
end if
	%></font>��</td>
  </tr>
  <tr> 
    <td width="21%" class="tdbg" height=23>���ռ���vip��Ա����<font color="#ff6600">
<%
sql="SELECT count(*) as jieducm_vipnum FROM "&jieducm&"_chongjilu where class='����vip��' and ( datediff(day,[now],getdate())=0)"
Set rs = Server.CreateObject("Adodb.RecordSet")
rs.Open Sql,conn,1,1
if not rs.eof then			
jieducm_viptotal=rs("jieducm_vipnum")
end if
rs.close
set rs=nothing
response.Write(jieducm_viptotal)
	%></font>��</td>
    <td width="29%" class="tdbg">���ս�����������<font color="#ff6600">
    <%
sql="SELECT count(*) as jieducm_jsnum FROM "&jieducm&"_jieshou where ( datediff(day,[now],getdate())=0)"
Set rs = Server.CreateObject("Adodb.RecordSet")
rs.Open Sql,conn,1,1
if not rs.eof then			
jieducm_jstotal=rs("jieducm_jsnum")
end if
rs.close
set rs=nothing
response.Write(jieducm_jstotal)
	%></font>��</td>
    <td width="50%" class="tdbg">���ճ�ֵ����ֵ��<font color="#ff6600">
<%
sql="SELECT  sum(cunkuan) as jieducm_buycar2 FROM "&jieducm&"_chongjilu where class='��ֵ����ֵ' and (datediff(day,[now],getdate())=0)"
Set rs = Server.CreateObject("Adodb.RecordSet")
rs.Open Sql,conn,1,1
if not rs.eof then			
jieducm_cartotal=rs("jieducm_buycar2")
end if
rs.close
set rs=nothing

if not isnull(jieducm_cartotal) then
response.Write(jieducm_cartotal)
else
response.Write("0")
end if
	%></font>Ԫ</td>
  </tr>
  <tr> 
    <td width="21%" class="tdbg" height=23>���������㿨��<font color="#ff6600"><%
sql="SELECT   sum(cunkuan) as adminbuycar FROM "&jieducm&"_chongjilu where class='����㿨' and ( datediff(day,[now],getdate())=0)"
Set Rs = Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,conn,1,1
if not( rs.eof or rs.bof) then			
jieducm_adminbuycar=rs("adminbuycar")
end if
if not isnull(jieducm_adminbuycar) then
response.Write(jieducm_adminbuycar)
else
response.Write("0")
end if
	%></font>Ԫ</td>
    <td width="29%" class="tdbg">���������������<font color="#ff6600">
    <%
sql="SELECT count(*) as jieducm_zhnumok FROM "&jieducm&"_zhongxin where start='5' and ( datediff(day,[now],getdate())=0)"
Set rs = Server.CreateObject("Adodb.RecordSet")
rs.Open Sql,conn,1,1
if not rs.eof then			
jieducm_zhtotalok=rs("jieducm_zhnumok")
end if
rs.close
set rs=nothing
response.Write(jieducm_zhtotalok)
	%></font>��</td>
    <td width="50%" class="tdbg">����������ֵ��<font color="#ff6600">
<%
sql="SELECT  sum(cunkuan) as jieducm_china FROM "&jieducm&"_chongjilu where class='������ֵ' and (datediff(day,[now],getdate())=0)"
Set rs = Server.CreateObject("Adodb.RecordSet")
rs.Open Sql,conn,1,1
if not rs.eof then			
jieducm_chian_total=rs("jieducm_china")
end if
rs.close
set rs=nothing

if not isnull(jieducm_chian_total) then
response.Write(jieducm_chian_total)
else
response.Write("0")
end if
	%>
    </font>Ԫ&nbsp;&nbsp;
<%
sql="SELECT  sum(Money) as jieducm_alipay FROM AutoCZ where  Status=1 and (datediff(day,[EndDate],getdate())=0)"
Set rs = Server.CreateObject("Adodb.RecordSet")
rs.Open Sql,conn,1,1
if not rs.eof then			
jieducm_alipay_total=rs("jieducm_alipay")
end if
rs.close
set rs=nothing
%>
	
	 ����֧�����Զ���ֵ��<font color="#ff6600"><%if not isnull(jieducm_alipay_total) then
response.Write(jieducm_alipay_total)
else
response.Write("0")
end if%></font>Ԫ</td>
  </tr>
  
  <tr>
    <td height=23 colspan="3" class="tdbg">&nbsp;</td>
  </tr>
</table>
<table cellpadding="2" cellspacing="1" border="0" width="100%" class="border" align=center>
  <tr align="center"> 
    <td height=25 colspan=2 class="topbg"><strong>�� �� �� ��</strong></td> 
  <tr> 
    <td width="50%"  class="tdbg" height=23>���ͣ�������רΪ�Ա���ѻ�ˢ���죬��Ŀǰ�����Ϲ��ܱȽ���ȫ��ƽ̨</td>
    <td width="50%" class="tdbg">��Ʒ������վ��ҵ��  </td>
  </tr>
  <tr> 
    <td width="50%" class="tdbg" height=23>��Ʒ������<a href="http://www.778892.com" target="_blank"><a href="mailto:wangxiaoyu-4@163.com">����</a>����Ϣ�������޹�˾</a></td>
    <td width="50%" class="tdbg">�绰��</td>
  </tr>
  <tr> 
    <td width="50%" class="tdbg" height=23>����֧�֣�<a href="http://www.778892.com">����</a>��</td>
    <td width="50%" class="tdbg">QQ��<a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=859680966&site=qq&menu=yes">859680966</a></td>
  </tr>
  
  <tr>
    <td height=23 colspan="2" class="tdbg">��Ȩ����Ȩ����������Ϣ�������޹�˾�����ÿ������޸ģ���Ȩ�ؾ���<a href="Admin_ServerInfo.asp"></a></td>
  </tr>
</table>

<table cellpadding="2" cellspacing="1" border="0" width="100%" class="border" align=center>
  <tr align="center"> 
    <td height=25 class="topbg">&nbsp;</td>
  </tr>
  <TR>
    <TD class="tdbg">    
</table>
<table cellpadding="2" cellspacing="1" border="0" width="100%" class="border" align=center>
  <tr align="center"> 
    <td height=25 class="topbg">&nbsp;</td>
  </tr>
  
</table>
<!--#include file="Admin_fooder.asp"-->
<%
conn.close
set conn=nothing%>
</body>
</html>
