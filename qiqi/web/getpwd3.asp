<!--#include file="inc/index_conn.asp"-->
<!--#INCLUDE FILE="inc/function.asp"-->
<!--#include file="config.asp"-->
<%
From_url = Cstr(Request.ServerVariables("HTTP_REFERER"))
Serv_url = Cstr(Request.ServerVariables("SERVER_NAME"))
if mid(From_url,8,len(Serv_url)) <> Serv_url or Request.ServerVariables("HTTP_REFERER")="" then
response.write "����ʱ!"
response.end
end If

username=Replace(Trim(Request.form("username")),"'","''")
daai=Replace(Trim(Request.form("daai")),"'","''")
if username="" then
 Response.Write("<script>alert('�û�������Ϊ�գ�');history.back();</script>")
 response.End()
end if
if  checkenter(username) then
else
	response.write "<script language=javascript>alert('�벻Ҫ����Ƿ��ַ���');history.back();</script>"
	response.end
end if

if daai="" then
 Response.Write("<script>alert('�𰸲���Ϊ�գ�');history.back();</script>")
 response.End()
end if
set rs=Server.CreateObject("Adodb.Recordset")
sql="select * from "&jieducm&"_register where username='"&username&"'"
rs.open sql,conn,1,1
If rs.eof Then
 Response.Write("<script>alert('����û���û��ע���أ��뵽��ҳע��ɣ�');window.close();</script>")
 response.End()
else
jieducm_daai=Replace(Trim(rs("daai")),"'","''")
End If

If jieducm_daai<>daai Then
 Response.Write("<script>alert('����д�Ĵ𰸲���ȷ��');window.close();</script>")
 response.End()
else
session("jieducm_daai")=jieducm_daai
End If%>
<html><head><title>ȡ������</title>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/css.css" rel="stylesheet" type="text/css">
<script LANGUAGE="javascript">
<!--
function FrmAddLink_onsubmit() {
 if (document.FrmAddLink.passwd.value=="")
	{
	  alert("�Բ��������������룡")
	  document.FrmAddLink.passwd.focus()
	  return false
	 }
else if (document.FrmAddLink.passwd.value.length < 6)
	{
	  alert("Ϊ�˰�ȫ����������Ӧ�ó�һ�㣡")
	  document.FrmAddLink.passwd.focus()
	  return false
	 }
else if (document.FrmAddLink.passwd.value.length > 16)
	{
	  alert("��������̫���˰ɣ�")
	  document.FrmAddLink.passwd.focus()
	  return false
	 }
else if (document.FrmAddLink.passwd2.value=="")
	{
	  alert("�Բ�������������֤���룡")
	  document.FrmAddLink.passwd2.focus()
	  return false
	 }
else if (document.FrmAddLink.passwd2.value !== document.FrmAddLink.passwd.value)
	{
	  alert("�Բ�����������������벻һ�£�")
	  document.FrmAddLink.passwd2.focus()
	  return false
	 }
}
//-->
</script>
</head>
<body topmargin="20" leftmargin="0" oncontextmenu="self.event.returnValue=false">
<form method="POST" name="FrmAddLink" language="javascript" onSubmit="return FrmAddLink_onsubmit()" action="getpwd4.asp">      
<table border="0" cellpadding="2" cellspacing="1" width="400" bgcolor="#6FABCB" style="border-collapse: collapse" align="center">
<TR><TD bgColor=#ffffff height=1></TD></TR>
<TR><TD align=center>�������������������롣</TD>
</TR>
<tr> 
<td bgcolor="F8FAFB">
<p align="center">�� �� �룺

<input class=wenbenkuang type="password" name="passwd" size="15" maxlenght=18><br>
��֤���룺
<input class=wenbenkuang type="password" name="passwd2" size="15" maxlenght=18><br><input name="username" type="hidden" value="<%=username%>">
<input class=go-wenbenkuang type="submit" value=" �������� " name="submit"></p>
<p align="center"><a href="javascript:window.close()" class="1">�رմ���</a></p>
</td>
</tr>
</table>
</form>
</body>
</html>
<%   
rs.close
set rs=nothing
conn.close   
set conn=nothing 
%>
