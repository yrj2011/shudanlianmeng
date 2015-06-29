<!--#include file="inc/index_conn.asp"-->
<!--#include file="config.asp" -->
<%
From_url = Cstr(Request.ServerVariables("HTTP_REFERER"))
Serv_url = Cstr(Request.ServerVariables("SERVER_NAME"))
if mid(From_url,8,len(Serv_url)) <> Serv_url or Request.ServerVariables("HTTP_REFERER")="" then
response.write "请求超时!"
response.end
end If
username=Replace(Trim(Request.Form("username")),"'","''")
if username="" then
 Response.Write("<script>alert('用户名不能为空！');history.back();</script>")
 response.End()
end if
set rs=Server.CreateObject("Adodb.Recordset")
sql="select * from "&jieducm&"_register where username='"&username&"' "
rs.open sql,conn,1,1
If rs.eof Then
%>
<script language="javascript">
alert("这个用户没有注册，请注册！");
window.close();
</script>
<%
response.End()
End If
%>
<html><head><title>取回密码 | 回答问题</title>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/css.css" rel="stylesheet" type="text/css">
<script language="javascript">
<!--
function form1_onsubmit() {
if (document.form1.daai.value=="")
	{
	alert("请输入您的问题答案。")
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
<TR><TD align=center height=30>第二步：请回答下列问题。</TD></TR>
<tr bgcolor="#ffffff"> 
<td height="100" bgcolor="F8FAFB">  
<table border="0" cellpadding="0" width="99%" cellspacing="0" height="1" style="border-collapse: collapse" bordercolor="#111111">
<%if rs("weiti")<>"" then%>
<tr> 
<td width="100%" height="22">
问 题：
<font color="red"><%=rs("weiti")%></font> 

</td>
</tr>
<tr> 
<td width="100%" height="1" align="center">答 案：
<input name="username" type="hidden" value="<%=username%>">
<input name="daai" type="text" class=wenbenkuang id="daai" size="18">
<input class=go-wenbenkuang name="imageField" type="submit" value=" 下一步 " onFocus="this.blur()"><%end if%>
<br>
如果您的提示问题或答案为空，请与*<a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=<%=stupqq%>&site=qq&menu=yes">管理员</a>*联系取回您的密码!</td>
</tr>

</table>
<p align="center"><a href="javascript:window.close()" class="1">关闭窗口</a></p>
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