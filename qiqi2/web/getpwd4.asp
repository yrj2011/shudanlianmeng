<!--#include file="inc/index_conn.asp"-->
<!--#INCLUDE FILE="inc/function.asp"-->
<!--#include file=md5.asp -->
<%
From_url = Cstr(Request.ServerVariables("HTTP_REFERER"))
Serv_url = Cstr(Request.ServerVariables("SERVER_NAME"))
if mid(From_url,8,len(Serv_url)) <> Serv_url or Request.ServerVariables("HTTP_REFERER")="" then
response.write "����ʱ!"
response.end
end If
username=Replace(Trim(Request.form("username")),"'","''")
passwd=md5(trim(request.form("passwd")))

if username="" then
 Response.Write("<script>alert('�û�������Ϊ�գ�');history.back();</script>")
 response.End()
elseif passwd="" then
 Response.Write("<script>alert('���벻��Ϊ�գ�');history.back();</script>")
 response.End()
end if

if  checkenter(username) then
else
	response.write "<script language=javascript>alert('�벻Ҫ����Ƿ��ַ���');history.back();</script>"
	response.end
end if
if  checkenter(passwd) then
else
	response.write "<script language=javascript>alert('�벻Ҫ����Ƿ��ַ���');history.back();</script>"
	response.end
end if

set rs=Server.CreateObject("Adodb.Recordset")
sql="select daai from "&jieducm&"_register where username='"&username&"'"
rs.open sql,conn,1,1
if not rs.eof then
daai=rs("daai")
end if

If session("jieducm_daai")<>daai Then
 Response.Write("<script>alert('�벻Ҫ����Ƿ��ַ���');window.close();</script>")
 response.End()
end if

set rs=Server.CreateObject("Adodb.Recordset")
sql="select password1 from "&jieducm&"_register where username='"&username&"'"
rs.open sql,conn,1,3
If rs.eof Then
response.write "<script language=javascript>alert('����û���û��ע���أ��뵽��ҳע��ɣ�');history.back();</script>"
response.end
else
rs("password1")=passwd
rs.update
end if
rs.close
set rs=nothing
conn.close   
set conn=nothing
%>
<html><head><title>ȡ������</title>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/css.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.STYLE1 {color: #FFFF00}
-->
</style>
</head>
<body topmargin="20" leftmargin="0" oncontextmenu="self.event.returnValue=false">
<span class="STYLE1"></span>
<table width="320" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#6FABCB" style="border-collapse: collapse">
<tr>
<td height="100" bgcolor="F8FAFB"> 
<p align="center">���������óɹ����뷵����ҳ*<a href="index.asp">���µ�¼</a>*��</p>
</td>
</tr>
<tr>
<td bgcolor="F8FAFB">
<p align="center"><a href="javascript:window.close()" class="1">�رմ���</a></p></td>
</tr>
</table>
</body>
</html>
<%rs.close
set rs=nothing
 call CloseConn()
%>