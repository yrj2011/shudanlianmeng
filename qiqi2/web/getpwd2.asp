<!--#include file="inc/index_conn.asp"-->
<!--#include file="config.asp" -->
<%
From_url = Cstr(Request.ServerVariables("HTTP_REFERER"))
Serv_url = Cstr(Request.ServerVariables("SERVER_NAME"))
if mid(From_url,8,len(Serv_url)) <> Serv_url or Request.ServerVariables("HTTP_REFERER")="" then
response.write "����ʱ!"
response.end
end If
username=Replace(Trim(Request.Form("username")),"'","''")
if username="" then
 Response.Write("<script>alert('�û�������Ϊ�գ�');history.back();</script>")
 response.End()
end if
set rs=Server.CreateObject("Adodb.Recordset")
sql="select * from "&jieducm&"_register where username='"&username&"' "
rs.open sql,conn,1,1
If rs.eof Then
%>
<script language="javascript">
alert("����û�û��ע�ᣬ��ע�ᣡ");
window.close();
</script>
<%
response.End()
End If
%>
<html><head><title>ȡ������ | �ش�����</title>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/css.css" rel="stylesheet" type="text/css">
<script language="javascript">
<!--
function form1_onsubmit() {
if (document.form1.daai.value=="")
	{
	alert("��������������𰸡�")
	document.form1.daai.focus()
	return false
	}
}
// --></script>
</head>
<body topmargin="20" leftmargin="0" oncontextmenu="self.event.returnValue=false">
<form method="POST" name="form1" language="javascript" onSubmit="return form1_onsubmit()" action="getpwd3.asp">
<table border="0" cellpadding="2" cellspacing="1" width="400" bgcolor="#6FABCB" style="border-collapse: collapse" align="center">
<TR><TD bgColor=#ffffff height=1></TD></TR>
<TR><TD align=center height=30>�ڶ�������ش��������⡣</TD></TR>
<tr bgcolor="#ffffff"> 
<td height="100" bgcolor="F8FAFB">  
<table border="0" cellpadding="0" width="99%" cellspacing="0" height="1" style="border-collapse: collapse" bordercolor="#111111">
<%if rs("weiti")<>"" then%>
<tr> 
<td width="100%" height="22">
�� �⣺
<font color="red"><%=rs("weiti")%></font> 

</td>
</tr>
<tr> 
<td width="100%" height="1" align="center">�� ����
<input name="username" type="hidden" value="<%=username%>">
<input name="daai" type="text" class=wenbenkuang id="daai" size="18">
<input class=go-wenbenkuang name="imageField" type="submit" value=" ��һ�� " onFocus="this.blur()"><%end if%>
<br>
���������ʾ������Ϊ�գ�����*<a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=<%=stupqq%>&site=qq&menu=yes">����Ա</a>*��ϵȡ����������!</td>
</tr>

</table>
<p align="center"><a href="javascript:window.close()" class="1">�رմ���</a></p>
</td>
</tr>
</table>
</form>
</div>
</body>
</html>
<%   
rs.close
set rs=nothing
conn.close   
set conn=nothing 
%>